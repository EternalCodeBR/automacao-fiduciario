# Automação da atualização diária de preços

## O problema

Todo dia útil, alguém da área de Riscos precisava atualizar manualmente o
preço de referência (PU) de 4 ativos de renda fixa e publicá-lo no site da
empresa. Eram **6 etapas manuais** — recalcular cada calculadora, copiar os
dados para os históricos, acessar o painel do site, substituir o arquivo de
cada ativo, salvar cada publicação e avisar a equipe — consumindo **de 30 a
45 minutos por dia**.

- Um erro de cópia ou de arquivo encadearia em um preço errado publicado.

## O que eu fiz

Construí um pipeline que faz essa atualização sozinho, do início ao fim, em
**uma única execução** dividida em 3 etapas verificáveis:

1. **Recálculo**: atualiza as calculadoras e os históricos de PU a partir
   das fontes internas, aguardando a confirmação de que todos os valores
   foram recalculados antes de prosseguir (proteção contra dado inconsistente).
2. **Publicação verificada**: substitui o arquivo de cada ativo no site,
   confirmando cada upload antes de salvar.
3. **Notificação**: avisa a equipe automaticamente ao final — com o motivo
   exato quando algo falha.

## Resultado

- **6 etapas manuais foram eliminadas**: uma rotina de 30–45 minutos por
  dia, com vários pontos de falha, passou a rodar sozinha, do recálculo até
  a divulgação do PU no site.
- A equipe passou a ser avisada automaticamente, com o motivo exato
  quando algo falha, em vez de descobrir o problema horas depois pelo gestor do fundo.

---
