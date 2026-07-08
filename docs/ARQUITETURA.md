# Arquitetura — detalhes técnicos

## Visão geral

| Campo | Valor |
|---|---|
| Linguagem | Python 3.8+ |
| Bibliotecas principais | Playwright, pywin32 (COM), tkinter |
| Macro VBA | módulo de planilha (`vba/Modulo_Planilha.cls`) |
| Ambiente de destino | CMS WordPress (genérico neste repositório) |
| Notificação | webhook de chat (Adaptive Card via Power Automate) |

## Por que essa ordem de etapas?

| Etapa | Depende de | Motivo |
|---|---|---|
| 1 — VBA | Excel COM + planilha de dados de mercado aberta | Os históricos precisam dos PUs do dia gerados pelas calculadoras |
| 2 — Upload | Histórico atualizado pela Etapa 1 + credenciais do CMS | É o histórico atualizado que vai para o site |
| 3 — Notificação | Sucesso (ou falha) da Etapa 2 | Só faz sentido notificar depois que se sabe o resultado real do upload |

## Etapa 1 — macro VBA

Para cada ativo configurado em `GetPairs()`:

1. Garante que a planilha de dados de mercado está aberta (reaproveita se
   o Python já abriu; senão abre e recalcula).
2. Abre a calculadora do ativo, habilitando eventos temporariamente para
   que o auto-refresh da planilha dispare.
3. Abre o histórico correspondente e lê a última data registrada.
4. Localiza essa data na calculadora para identificar o bloco de linhas
   novo a copiar.
5. **Bloco de recálculo garantido** antes de ler qualquer valor:
   atualização de links externos → `RefreshAll` → recálculo automático →
   `CalculateFullRebuild` → espera ativa até o motor de cálculo do Excel
   sinalizar que terminou.
6. Copia os valores (não fórmulas) da calculadora para o histórico,
   preservando a formatação da última linha existente.
7. Salva e fecha os arquivos.

### Causa raiz do bug de erro de cálculo (#VALOR)

Quando a macro abria a calculadora via COM/headless:

- Os links externos para a planilha de dados de mercado não eram
  atualizados na abertura;
- Eventos desabilitados impediam o auto-refresh da calculadora de disparar;
- O bloco de valores era lido antes do recálculo terminar.

**Fix:** a macro passou a abrir a planilha de dados de mercado por conta
própria, habilita eventos temporariamente durante a abertura das
calculadoras, e força + aguarda o recálculo completo antes de ler valores.

## Etapa 2 — upload no CMS

Cada post tem um campo de repetição (ACF Repeater) com uma linha cujo
"Título Anexo" é `"PU"`. O script:

1. Localiza essa linha via seletor CSS;
2. Remove o arquivo anterior (clique disparado via JavaScript, já que o
   botão só é visualmente alcançável em hover);
3. Abre o modal de mídia e injeta o novo arquivo direto no input de
   arquivo;
4. Aguarda o processamento e confirma a seleção;
5. Salva o post (compatível com editor clássico e editor de blocos).

### Causa raiz do falso-positivo de sucesso

A verificação original dependia de sinais visuais (botão reaparecer,
snackbar de confirmação). Só que os ícones de remover/editar do campo de
arquivo ficam **visíveis apenas via `:hover`** — uma condição que o
Playwright headless nunca satisfaz. Isso fazia `wait_for(state="visible")`
sempre expirar silenciosamente (engolido por um `try/except` amplo
demais), e o script seguia adiante achando que tinha dado certo mesmo
quando o clique de remoção nunca havia sido processado visualmente pela
página.

**Fix:** parar de depender de visibilidade CSS e verificar o **valor do
input oculto que guarda o ID do anexo** (o dado que de fato é persistido)
antes e depois da troca — só considera sucesso se o valor mudou. Depois de
salvar, o script ainda recarrega a página e compara a linha do repeater
antes/depois como segunda camada de verificação.

A causa raiz foi confirmada com um script Playwright **somente leitura**
(login + navegação + dump do HTML da página, sem clicar em nada que
altere estado) — mais rápido e confiável do que adivinhar seletores.

## Etapa 3 — notificação

Envia um Adaptive Card via POST para um webhook de Workflow. Se a
execução teve falhas parciais, a lista de ativos que falharam (e o motivo
de cada falha) é incluída no corpo da mensagem. A notificação é
responsabilidade de quem chama `run()`, não da função em si — isso permite
que a camada de UI pergunte ao usuário antes de notificar quando há falha
parcial, evitando reportar "sucesso" para o time quando na verdade algo
deu errado.

## Riscos e segurança

- Credenciais (usuário/senha do CMS, webhook de notificação) são lidas de
  variáveis de ambiente — nunca hardcoded no código-fonte.
- O webhook de notificação funciona como uma credencial (permite postar no
  canal); deve ser tratado como segredo e rotacionado se houver suspeita
  de vazamento.
- O código-fonte VBA deve ser editado sempre no arquivo `.cls` (nunca via
  copy/paste do editor VBA para um chat ou vice-versa) para evitar
  corrupção de acentuação por incompatibilidade de encoding (CP1252 vs.
  UTF-8).
