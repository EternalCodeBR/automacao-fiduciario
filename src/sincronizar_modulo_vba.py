# -*- coding: utf-8 -*-
"""
Substitui o codigo do modulo de planilha dentro da pasta de trabalho da
macro pelo conteudo atual de 'vba/Modulo_Planilha.cls' (CP1252).

Motivacao: colar codigo VBA diretamente do editor do VBA para o chat (ou
vice-versa) corrompe acentuacao, porque o VBA usa CP1252 e a maioria dos
editores modernos usa UTF-8. Este script le o .cls sempre como CP1252 e
injeta o codigo via COM, eliminando o passo manual de copy/paste.

Requisitos:
- Excel deve estar com "Confiar no acesso ao modelo de objeto do projeto do VBA" habilitado
  (Central de Confiabilidade -> Configuracoes de Macro).
- Feche todas as instancias do Excel antes de rodar.

Variaveis de ambiente:
- CLS_FILE:   caminho do .cls fonte (default: vba/Modulo_Planilha.cls, relativo ao repo)
- XLSM_FILE:  caminho da pasta de trabalho .xlsm que contem o modulo
- MODULE_NAME: nome do modulo VBA a substituir (default: Planilha1)
"""

import os
import sys
import time
import win32com.client
import pywintypes

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLS_FILE = os.environ.get("CLS_FILE", os.path.join(REPO_ROOT, "vba", "Modulo_Planilha.cls"))
XLSM_FILE = os.environ.get("XLSM_FILE", r"C:\caminho\para\Macro.xlsm")
MODULE_NAME = os.environ.get("MODULE_NAME", "Planilha1")


def carregar_corpo_cls(path):
    """Le o .cls em CP1252 e devolve apenas o corpo (sem o cabecalho de classe)."""
    with open(path, "rb") as f:
        raw = f.read()
    texto = raw.decode("cp1252")
    linhas = texto.splitlines()
    # Pular cabecalho ate (e inclusive) a ultima linha 'Attribute VB_...'
    inicio = 0
    for i, l in enumerate(linhas):
        if l.strip().startswith("Attribute VB_"):
            inicio = i + 1
    # Pular linhas em branco logo apos os Attributes
    while inicio < len(linhas) and linhas[inicio].strip() == "":
        inicio += 1
    corpo = "\r\n".join(linhas[inicio:])
    return corpo


def main():
    if not os.path.exists(CLS_FILE):
        print(f"[ERRO] Nao achei: {CLS_FILE}")
        sys.exit(1)
    if not os.path.exists(XLSM_FILE):
        print(f"[ERRO] Nao achei: {XLSM_FILE}")
        sys.exit(1)

    corpo = carregar_corpo_cls(CLS_FILE)
    print(f"[INFO] Corpo do modulo: {len(corpo)} caracteres, "
          f"{corpo.count(chr(10))+1} linhas")

    print("[INFO] Iniciando Excel...")
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1  # msoAutomationSecurityLow

    try:
        print("[INFO] Abrindo workbook...")
        wb = excel.Workbooks.Open(XLSM_FILE, UpdateLinks=0)
        time.sleep(1.0)

        try:
            vbproj = wb.VBProject
        except pywintypes.com_error as e:
            print("[ERRO] Nao consegui acessar VBProject.")
            print("       Habilite 'Confiar no acesso ao modelo de objeto do "
                  "projeto do VBA' na Central de Confiabilidade.")
            print(f"       Detalhe: {e}")
            wb.Close(SaveChanges=False)
            excel.Quit()
            sys.exit(2)

        # Localizar o modulo alvo
        comp = None
        for c in vbproj.VBComponents:
            if c.Name == MODULE_NAME:
                comp = c
                break
        if comp is None:
            print(f"[ERRO] Modulo '{MODULE_NAME}' nao encontrado no VBProject.")
            wb.Close(SaveChanges=False)
            excel.Quit()
            sys.exit(3)

        cm = comp.CodeModule
        n = cm.CountOfLines
        print(f"[INFO] Modulo '{MODULE_NAME}' tem {n} linhas. Limpando...")
        if n > 0:
            cm.DeleteLines(1, n)

        print("[INFO] Injetando codigo novo...")
        cm.AddFromString(corpo)
        time.sleep(0.5)

        print("[INFO] Salvando workbook (mantendo .xlsm)...")
        wb.Save()
        wb.Close(SaveChanges=False)
        print("[OK] Modulo sincronizado com sucesso.")

    finally:
        try:
            excel.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
