"""
Pipeline de atualização automática de Preço Unitário (PU) em site institucional.

Executa em 3 etapas:
  1. Dispara uma macro VBA (via Excel COM) que recalcula os históricos de PU
     de cada ativo a partir das calculadoras internas.
  2. Faz upload dos arquivos de histórico atualizados no CMS (via Playwright),
     substituindo o anexo de cada post correspondente.
  3. Notifica um canal de chat (Microsoft Teams / Power Automate) com o
     resultado da execução.

Todas as credenciais e caminhos são lidos de variáveis de ambiente — veja
`.env.example` na raiz do projeto para a lista completa.
"""

import os
import glob
import sys
import time
import json
import threading
import urllib.request
import tkinter as tk
from tkinter import simpledialog, messagebox
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright

__version__ = "1.4.0"
# Changelog:
# 1.4.0 — corrigidos os seletores do repeater de anexos (causa raiz real do
#         falso-positivo): os ícones de remover/editar só ficam visíveis via
#         CSS :hover, então o Playwright headless nunca satisfazia
#         wait_for(visible). Agora a verificação usa o valor do input oculto
#         que guarda o ID do anexo (o dado que de fato é salvo) e a
#         navegação até a linha aninhada real do repeater. run() deixou de
#         notificar o chat sozinho: agora retorna a lista de falhas, e a
#         tela de progresso pergunta ao usuário antes de notificar quando
#         há falha — só notifica automaticamente no sucesso total.
# 1.3.0 — fix de falso-positivo de sucesso no upload: o script relatava
#         "concluído" mesmo quando o anexo não era realmente trocado. Agora
#         cada ativo é verificado de verdade (recarrega o post e compara a
#         linha do repeater antes/depois); falhas reais são coletadas e
#         propagadas para a UI e para a notificação. Também trata o painel
#         de confirmação do editor de blocos e salva screenshot de
#         diagnóstico em caso de falha.
# 1.2.0 — fix de erro de cálculo (#VALOR) na macro VBA: passou a abrir a
#         planilha de dados de mercado por conta própria, com recálculo
#         forçado e aguardado antes de ler valores.
# 1.1.0 — planilha de dados de mercado permanece aberta durante toda a
#         Etapa 1 (evita reabrir a cada ativo).
# 1.0.0 — versão inicial: macro VBA + upload no CMS + notificação de chat.

# --------------------------------------------------------------------------
# Configuração — tudo via variável de ambiente (ver .env.example)
# --------------------------------------------------------------------------
PASTA_CALCULADORAS = os.environ.get("PASTA_CALCULADORAS", r"C:\caminho\para\Calculadoras")
MACRO_WORKBOOK = os.environ.get("MACRO_WORKBOOK", r"C:\caminho\para\Macro.xlsm")
MACRO_NAME = "Planilha1.Atualizar_Historicos_PU_Auto"
# Planilha de referência (ex.: taxas de mercado) que as calculadoras consomem
# via link externo. Precisa estar aberta na mesma instância do Excel antes
# da macro rodar.
BASE_DADOS_MERCADO = os.environ.get("BASE_DADOS_MERCADO", r"C:\caminho\para\BaseDadosMercado.xlsm")

URL_LOGIN = os.environ.get("CMS_URL_LOGIN", "https://seusite.exemplo.com/wp-login.php")
CMS_USER = os.environ.get("CMS_USER", "")
CMS_PASS = os.environ.get("CMS_PASS", "")

# Webhook de um Workflow (Power Automate) que posta em um canal de chat.
CHAT_WEBHOOK_URL = os.environ.get("CHAT_WEBHOOK_URL", "")

# Data-base passada para a macro VBA (preenchida em runtime pela tela de progresso)
DATA_BASE = None

# Lista de ativos monitorados: nome de exibição, padrão de arquivo a
# procurar na pasta local, e ID do post no CMS que recebe o anexo.
# Os identificadores reais dos ativos e os IDs de post foram substituídos
# por valores de exemplo — configure via ATIVOS_CONFIG (JSON) ou edite aqui.
CALCULADORAS = json.loads(os.environ.get("ATIVOS_CONFIG", json.dumps([
    {
        "nome": "Debênture A",
        "padrao_arquivo": "Historico_PU_Debenture_A*.xls*",
        "post_id": 1001,
    },
    {
        "nome": "Debênture B",
        "padrao_arquivo": "Historico_PU_Debenture_B*.xls*",
        "post_id": 1002,
    },
    {
        "nome": "Nota Comercial C",
        "padrao_arquivo": "Historico_PU_NotaComercial_C*.xls*",
        "post_id": 1003,
    },
    {
        "nome": "CRI D",
        "padrao_arquivo": "Historico_PU_CRI_D*.xls*",
        "post_id": 1004,
    },
])))

def encontrar_arquivo_mais_recente(padrao):
    caminho_busca = os.path.join(PASTA_CALCULADORAS, padrao)
    arquivos = glob.glob(caminho_busca)
    if not arquivos:
        return None
    # Pega o arquivo modificado mais recentemente
    return max(arquivos, key=os.path.getmtime)

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846

def _com_retry(func, *args, tentativas=30, espera=2.0, **kwargs):
    """Chama func(*args, **kwargs) com retry em RPC_E_CALL_REJECTED.
    O Excel rejeita chamadas COM quando está ocupado (abrindo arquivos,
    mostrando diálogo de Protected View etc.)."""
    import pywintypes
    ultimo_erro = None
    for i in range(tentativas):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as e:
            ultimo_erro = e
            hr = e.args[0] if e.args else None
            if hr in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER):
                print(f"  -> Excel ocupado (tentativa {i + 1}/{tentativas}); aguardando {espera:.0f}s...")
                time.sleep(espera)
                continue
            raise
    raise RuntimeError(f"Excel não respondeu após {tentativas} tentativas: {ultimo_erro}")

def abrir_excel_com_base_dados():
    """Inicia instância COM do Excel e abre a planilha de dados de mercado.
    Ela PRECISA permanecer aberta durante TODA a Etapa 1 — as calculadoras
    leem taxas dela via link externo, e fechá-la antes da macro terminar de
    gravar gera erro de cálculo (#VALOR) no histórico."""
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError:
        raise RuntimeError("pywin32 não está instalado. Rode: pip install pywin32")

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1  # msoAutomationSecurityLow: habilita macros sem prompt
    except Exception:
        pass
    try:
        excel.AskToUpdateLinks = False
    except Exception:
        pass
    # Garante recálculo automático — se herdar cálculo manual de alguma
    # sessão anterior, as fórmulas linkadas não recalculam ao abrir as
    # calculadoras, e o histórico recebe erro de cálculo.
    try:
        excel.Calculation = -4105  # xlCalculationAutomatic
    except Exception:
        pass

    print("  -> Abrindo planilha de dados de mercado...")
    try:
        wb_dados = _com_retry(excel.Workbooks.Open, BASE_DADOS_MERCADO, 0, False)
    except Exception as e:
        print(f"  -> [AVISO] Não consegui abrir a planilha de dados de mercado: {e}")
        print("            Prosseguindo, mas alguns valores podem ficar incorretos.")
        wb_dados = None

    return excel, wb_dados


def executar_macro_historico(excel, data_str):
    """Etapa 1: dispara a macro VBA. Recebe a instância do Excel já com a
    planilha de dados de mercado aberta. NÃO fecha o Excel — quem chamou
    deve chamar fechar_excel() somente quando estiver pronto para iniciar a
    Etapa 2."""
    print(f"Etapa 1/3: rodando macro VBA para a data {data_str}...")

    wb = None
    try:
        print("  -> Abrindo pasta de trabalho da macro...")
        wb = _com_retry(excel.Workbooks.Open, MACRO_WORKBOOK, 0, False)  # UpdateLinks=0, ReadOnly=False

        # Apaga log anterior pra distinguir execução nova de leftover
        log_path = os.path.join(os.environ.get("TEMP", "."), "pu_macro_log.txt")
        try:
            if os.path.exists(log_path):
                os.remove(log_path)
        except Exception:
            pass

        macro_alvo = f"'{os.path.basename(MACRO_WORKBOOK)}'!{MACRO_NAME}"
        print("  -> Disparando macro (pode levar minutos)...")
        resultado = _com_retry(excel.Run, macro_alvo, data_str)
        print(f"  -> Retorno COM: {resultado!r}")

        # Application.Run de Function em módulo de planilha pode não propagar o retorno
        # pelo COM. A macro escreve o log em %TEMP%\pu_macro_log.txt — fonte autoritativa.
        if not isinstance(resultado, str) or len(resultado) == 0:
            if os.path.exists(log_path):
                try:
                    with open(log_path, "r", encoding="cp1252") as f:
                        resultado = f.read().strip()
                    print(f"  -> Log lido de {log_path}:\n{resultado}")
                except Exception as e:
                    print(f"  -> [AVISO] Não consegui ler {log_path}: {e}")
            else:
                print(f"  -> [AVISO] Macro não gerou log em {log_path}.")

        if isinstance(resultado, str) and resultado.startswith("ERRO"):
            raise RuntimeError(resultado)
        if not isinstance(resultado, str) or len(resultado) == 0:
            print("  -> [AVISO] Sem retorno e sem log. Assumindo execução OK; "
                  "confira manualmente se os históricos foram atualizados.")
            resultado = "OK (sem log)"

        return resultado
    finally:
        if wb is not None:
            try:
                _com_retry(wb.Close, False, tentativas=5, espera=1.0)
            except Exception:
                pass


def fechar_excel(excel, wb_dados):
    """Encerra a instância COM do Excel. Deve ser chamada IMEDIATAMENTE
    antes da Etapa 2 (Playwright) — só nesse momento é seguro fechar a
    planilha de dados de mercado, com a Etapa 1 totalmente concluída."""
    import pythoncom

    print("  -> Fechando planilha de dados de mercado e encerrando Excel...")
    if wb_dados is not None:
        try:
            _com_retry(wb_dados.Close, False, tentativas=5, espera=1.0)
        except Exception:
            pass
    if excel is not None:
        try:
            _com_retry(excel.Quit, tentativas=5, espera=1.0)
        except Exception:
            pass
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

def notificar_chat(titulo, data_str, quantidade=4, falhas=None):
    """Envia mensagem a um canal de chat via Workflow (Power Automate).
    O Workflow espera o payload no formato Adaptive Card."""
    if not CHAT_WEBHOOK_URL:
        print("  -> [AVISO] CHAT_WEBHOOK_URL não configurada; pulando notificação.")
        return

    falhas = falhas or []
    sucesso_total = len(falhas) == 0
    emoji = "✅" if sucesso_total else "⚠️"
    cor = "Good" if sucesso_total else "Warning"
    texto_corpo = f"Atualização concluída em {data_str}.\n\n- **Ativos atualizados:** {quantidade} arquivos importados no site"
    if falhas:
        lista_falhas = "\n".join(f"  - {nome}: {motivo}" for nome, motivo in falhas)
        texto_corpo += f"\n\n**⚠️ Falharam ({len(falhas)}):**\n{lista_falhas}"

    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "type": "AdaptiveCard",
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.4",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": f"{emoji} {titulo}",
                            "size": "Medium",
                            "weight": "Bolder",
                            "color": cor
                        },
                        {
                            "type": "TextBlock",
                            "text": texto_corpo,
                            "wrap": True
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Abrir site",
                            "url": URL_LOGIN.replace("/wp-login.php", "")
                        }
                    ]
                },
            }
        ],
    }

    try:
        req = urllib.request.Request(
            CHAT_WEBHOOK_URL,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            print(f"  -> Chat notificado (HTTP {resp.status}).")
    except Exception as e:
        print(f"  -> [AVISO] Falha ao notificar o chat: {e}")

def run():
    """Executa as Etapas 1 e 2 (macro VBA + upload no CMS) e retorna a
    lista de falhas por ativo (vazia = todos com sucesso).
    Não notifica o chat — quem chama decide isso (a Etapa 3, de
    notificação, é responsabilidade do chamador, para permitir perguntar
    ao usuário antes de enviar diagnóstico em caso de falha)."""
    # Etapa 1: macro VBA atualiza os históricos.
    # A planilha de dados de mercado precisa permanecer aberta durante toda
    # a execução da macro — só fechamos imediatamente antes da Etapa 2.
    if DATA_BASE:
        excel = None
        wb_dados = None
        try:
            excel, wb_dados = abrir_excel_com_base_dados()
            executar_macro_historico(excel, DATA_BASE)
        except Exception as e:
            print(f"  -> [FALHA] Macro VBA não concluiu: {e}")
            print("Abortando antes do upload.")
            raise
        finally:
            # Mesmo em caso de erro, encerramos o Excel para não deixar
            # processo órfão segurando arquivos.
            fechar_excel(excel, wb_dados)
    else:
        print("Data-base não informada. Pulando etapa da macro.")

    print("\nEtapa 2/3: iniciando upload no site...")

    # Inicializa o Playwright
    with sync_playwright() as p:
        # headless=True: navegador roda em segundo plano (sem janela)
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()

        print("Acessando página de login...")
        page.goto(URL_LOGIN)

        # Fazendo login
        page.fill("#user_login", CMS_USER)
        page.fill("#user_pass", CMS_PASS)
        page.click("#wp-submit")
        page.wait_for_load_state("networkidle")

        if "wp-login" in page.url:
            print("Erro no login. Verifique as credenciais.")
            browser.close()
            raise RuntimeError("Erro no login do CMS. Verifique as credenciais.")

        print("Login realizado com sucesso!\n")

        pasta_screenshots = os.path.join(os.environ.get("TEMP", "."), "pu_falhas_screenshots")
        falhas = []  # lista de (nome, motivo)

        for calc in CALCULADORAS:
            print(f"Processando: {calc['nome']}")
            arquivo_local = encontrar_arquivo_mais_recente(calc['padrao_arquivo'])

            if not arquivo_local:
                print(f"  -> [ERRO] Nenhum arquivo encontrado na pasta para o padrão '{calc['padrao_arquivo']}'")
                continue

            print(f"  -> Encontrou arquivo: {os.path.basename(arquivo_local)}")

            # Navega direto para a tela de edição do post
            base_url = URL_LOGIN.replace("/wp-login.php", "")
            url_post = f"{base_url}/wp-admin/post.php?post={calc['post_id']}&action=edit&lang=pt-br"
            page.goto(url_post)
            page.wait_for_load_state("networkidle")

            try:
                # Repeater de campos customizados (ACF): precisamos encontrar
                # a linha (row) que tem um input com o valor "PU"
                print("  -> Buscando a linha que contém 'PU' no Título Anexo...")

                # Seletor para a linha externa do repeater ("Tipo de Anexos" = "PU")
                row = page.locator('tr.acf-row:has(input[value="PU"]), .acf-row:has(input[value="PU"])').first

                # Aguarda a linha existir
                row.wait_for(state="attached", timeout=10000)
                row.scroll_into_view_if_needed()

                # Assinatura da linha ANTES da troca, usada depois para confirmar
                # que o anexo realmente mudou (não basta a UI "parecer" ter salvo).
                assinatura_antes = row.inner_text()

                # A linha "PU" contém um repeater ANINHADO com a linha real do
                # campo de arquivo. Os ícones de remover/editar desse campo só
                # ficam visíveis via CSS :hover — nunca satisfazem um
                # wait_for(visible) em modo headless. Por isso verificamos o
                # valor do input OCULTO que guarda o ID do anexo (o dado que
                # de fato é salvo), em vez de esperar visibilidade.
                inner_row = row.locator('tr.acf-row:not(.acf-clone)').first
                inner_row.wait_for(state="attached", timeout=10000)
                id_input = inner_row.locator('input[data-name="id"]').first
                remove_btn = inner_row.locator('a[data-name="remove"]').first
                add_btn = inner_row.locator('a[data-name="add"]').first

                valor_antigo = id_input.get_attribute("value") or ""

                # Passo 1: remover o arquivo existente, se houver.
                # O clique é feito via JS (evaluate) porque o link só é
                # visualmente alcançável no hover; o clique funciona mesmo com o
                # elemento CSS-oculto, pois apenas dispara o evento que o campo
                # escuta via delegação.
                print("  -> Verificando se existe arquivo anterior para ser substituído...")
                if valor_antigo:
                    remove_btn.evaluate("node => node.click()")
                    print("  -> Arquivo anterior removido.")
                    for _ in range(20):
                        if not (id_input.get_attribute("value") or ""):
                            break
                        page.wait_for_timeout(250)
                    else:
                        raise RuntimeError("clique em remover não limpou o input oculto do anexo")
                else:
                    print("  -> Campo já estava vazio.")

                # Passo 2: Clicar no botão "Adicionar Arquivo"
                print("  -> Clicando em Adicionar Arquivo...")
                add_btn.evaluate("node => node.click()")

                # Passo 3: Fazer upload no modal de mídia
                print("  -> Fazendo upload do novo arquivo...")
                page.wait_for_selector(".media-modal-content", state="visible")

                # Clica na aba "Enviar arquivos" (Upload)
                aba_enviar = page.locator('button.media-menu-item:has-text("Enviar arquivos")')
                aba_enviar.click()
                page.wait_for_timeout(1000)

                # Injeta o arquivo direto no input file invisível
                print("  -> Injetando o arquivo no sistema...")
                page.locator('.media-modal input[type="file"]').first.set_input_files(arquivo_local)

                # Esperar a barra de progresso terminar e o botão "Selecionar" ficar habilitado
                print("  -> Aguardando o site processar o arquivo (upload)...")
                page.wait_for_selector('.media-button-select:not([disabled])', state="visible", timeout=60000)

                page.click('.media-button-select')

                print("  -> Aguardando o campo vincular o novo arquivo...")
                # A prova real de vínculo é o input oculto de ID mudar de valor —
                # não depende de nenhum elemento ficar visualmente visível.
                for _ in range(60):
                    valor_novo = id_input.get_attribute("value") or ""
                    if valor_novo and valor_novo != valor_antigo:
                        break
                    page.wait_for_timeout(500)
                else:
                    raise RuntimeError(
                        "input oculto do anexo não mudou de valor após selecionar o arquivo "
                        f"(continua '{valor_antigo}')"
                    )

                # Passo 4: Salvar o post
                print("  -> Salvando post...")
                if page.locator("#publish").is_visible():
                    page.click("#publish")
                    print("  -> Aguardando notificação de salvamento (editor clássico)...")
                    page.wait_for_selector('#message.updated', state="visible", timeout=45000)
                else:
                    page.click('.editor-post-publish-button')
                    # No editor de blocos, se o post ainda não estiver publicado, o
                    # primeiro clique só abre o painel "Pronto para publicar?" — é
                    # preciso um segundo clique para confirmar de fato. Se esse
                    # painel aparecer, confirmamos; se o post já estava publicado,
                    # o clique único já salva e o painel nunca aparece.
                    painel_confirmar = page.locator('.editor-post-publish-panel__header-publish-button .editor-post-publish-button')
                    try:
                        painel_confirmar.wait_for(state="visible", timeout=3000)
                        painel_confirmar.click()
                        print("  -> Painel de confirmação de publicação detectado e confirmado.")
                    except Exception:
                        pass  # post já estava publicado; um clique foi suficiente
                    print("  -> Aguardando notificação de salvamento (editor de blocos)...")
                    page.wait_for_selector('.components-snackbar', state="visible", timeout=45000)

                # Esperamos só por garantia que as requisições pendentes cessem
                page.wait_for_load_state("networkidle")

                # Verificação real: recarrega a página e confere que a linha do
                # repeater mudou de fato. Sem isso, a UI pode "parecer" ter
                # salvo (botões, snackbars) sem o anexo ter sido realmente trocado.
                print("  -> Verificando se o anexo realmente mudou...")
                page.goto(url_post)
                page.wait_for_load_state("networkidle")
                row_verif = page.locator('tr.acf-row:has(input[value="PU"]), .acf-row:has(input[value="PU"])').first
                row_verif.wait_for(state="attached", timeout=10000)
                assinatura_depois = row_verif.inner_text()

                if assinatura_depois == assinatura_antes:
                    raise RuntimeError(
                        "post foi salvo mas o anexo da linha 'PU' não mudou "
                        "(verificação pós-publicação falhou)"
                    )

                print(f"  -> [SUCESSO] Ativo atualizado e verificado!\n")

            except Exception as e:
                motivo = str(e)
                print(f"  -> [FALHA] Ocorreu um erro ao atualizar: {motivo}\n")
                falhas.append((calc['nome'], motivo))
                try:
                    os.makedirs(pasta_screenshots, exist_ok=True)
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    caminho_print = os.path.join(pasta_screenshots, f"{calc['post_id']}_{ts}.png")
                    page.screenshot(path=caminho_print, full_page=True)
                    print(f"  -> Screenshot de diagnóstico salvo em {caminho_print}")
                except Exception as e_print:
                    print(f"  -> [AVISO] Não consegui salvar screenshot: {e_print}")

        browser.close()

    print("=== PROCESSO FINALIZADO ===")
    return falhas

class TelaCarregamento(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Atualização de Preços")
        self.geometry("450x250")
        self.configure(bg="#2E3440")
        self.eval('tk::PlaceWindow . center')

        # Frames para a animação do personagem (usando emojis)
        self.frames_personagem = ["⏳ 🚶", "⏳ 🏃", "⏳ 🏃💨", "⏳ 🚶"]
        self.frame_atual = 0
        self.executando = True

        self.lbl_titulo = tk.Label(self, text="Importando Ativos...", font=("Segoe UI", 14, "bold"), bg="#2E3440", fg="#ECEFF4")
        self.lbl_titulo.pack(pady=20)

        self.lbl_personagem = tk.Label(self, text=self.frames_personagem[0], font=("Segoe UI", 40), bg="#2E3440", fg="#88C0D0")
        self.lbl_personagem.pack(pady=10)

        self.lbl_status = tk.Label(self, text="Iniciando...", font=("Segoe UI", 10), bg="#2E3440", fg="#D8DEE9", wraplength=400)
        self.lbl_status.pack(pady=10)

        self.lbl_tempo = tk.Label(self, text="Tempo: 00:00", font=("Segoe UI", 9), bg="#2E3440", fg="#EBCB8B")
        self.lbl_tempo.pack(pady=5)

        self.inicio_tempo = time.time()
        self.atualizar_tempo()

        # Redireciona o print() para atualizar o status na interface
        self.stdout_antigo = sys.stdout
        sys.stdout = self

        self.animar()

        # Inicia o processo em uma thread separada para não travar a interface gráfica
        threading.Thread(target=self.rodar_processo, daemon=True).start()

    def write(self, texto):
        if texto.strip():
            self.after(0, self.atualizar_status, texto.strip())
            self.stdout_antigo.write(texto)

    def flush(self):
        pass

    def atualizar_status(self, texto):
        if "=== PROCESSO FINALIZADO ===" not in texto:
            self.lbl_status.config(text=texto)

    def atualizar_tempo(self):
        if self.executando:
            decorrido = int(time.time() - self.inicio_tempo)
            minutos, segundos = divmod(decorrido, 60)
            self.lbl_tempo.config(text=f"Tempo: {minutos:02d}:{segundos:02d}")
            self.after(1000, self.atualizar_tempo)

    def animar(self):
        if self.executando:
            self.frame_atual = (self.frame_atual + 1) % len(self.frames_personagem)
            self.lbl_personagem.config(text=self.frames_personagem[self.frame_atual])
            self.after(300, self.animar)  # Atualiza a cada 300ms

    def rodar_processo(self):
        try:
            falhas = run()
            self.after(0, self.finalizar, falhas)
        except Exception as e:
            self.after(0, self.finalizar_erro, str(e))

    def finalizar(self, falhas):
        """Chamado na thread principal ao fim de run(). Se todos os ativos
        tiverem sido atualizados, notifica o chat normalmente. Se algum
        falhou, pergunta ao usuário antes de enviar diagnóstico — para não
        notificar sucesso quando na verdade houve falha."""
        total = len(CALCULADORAS)
        hoje_str = datetime.now().strftime("%Y-%m-%d")

        if not falhas:
            notificar_chat("Preços atualizados", hoje_str, total, falhas=[])
            self.finalizar_sucesso()
            return

        self.executando = False
        qtd_falha = len(falhas)
        self.lbl_personagem.config(text="⚠️", fg="#EBCB8B")
        self.lbl_titulo.config(text="Falha parcial na Importação")
        self.lbl_status.config(text=f"Falha no upload de {qtd_falha}/{total} ativos.")
        sys.stdout = self.stdout_antigo

        enviar = messagebox.askyesno(
            "Falha no Upload",
            f"Falha no Upload de {qtd_falha}/{total} ativos.\nDeseja enviar diagnóstico no chat?",
            parent=self,
        )
        if enviar:
            qtd_sucesso = total - qtd_falha
            notificar_chat("Preços — falha no upload", hoje_str, qtd_sucesso, falhas=falhas)
            print("  -> Diagnóstico enviado ao chat.")
        else:
            print("  -> Usuário optou por não enviar diagnóstico.")

    def finalizar_sucesso(self):
        self.executando = False
        self.lbl_personagem.config(text="🎉 ✅")
        self.lbl_personagem.config(fg="#A3BE8C")
        self.lbl_titulo.config(text="Importação Concluída!")
        self.lbl_status.config(text="Todos os ativos foram atualizados no site.")
        sys.stdout = self.stdout_antigo

    def finalizar_erro(self, erro):
        self.executando = False
        self.lbl_personagem.config(text="❌")
        self.lbl_personagem.config(fg="#BF616A")
        self.lbl_titulo.config(text="Erro na Importação")
        self.lbl_status.config(text=f"Falha: {erro}")
        sys.stdout = self.stdout_antigo

def pedir_data_base():
    """Mini-janela que pergunta a data dd/mm/aaaa antes da execução."""
    root = tk.Tk()
    root.withdraw()
    hoje = datetime.now().strftime("%d/%m/%Y")
    while True:
        valor = simpledialog.askstring(
            "Data-base",
            "Informe a DATA (dd/mm/aaaa) para a atualização dos históricos:",
            initialvalue=hoje,
            parent=root,
        )
        if valor is None:
            root.destroy()
            return None
        valor = valor.strip()
        try:
            datetime.strptime(valor, "%d/%m/%Y")
            root.destroy()
            return valor
        except ValueError:
            messagebox.showerror("Data inválida", "Use o formato dd/mm/aaaa.", parent=root)

if __name__ == "__main__":
    DATA_BASE = pedir_data_base()
    if not DATA_BASE:
        print("Operação cancelada pelo usuário.")
        sys.exit(0)
    app = TelaCarregamento()
    app.mainloop()
