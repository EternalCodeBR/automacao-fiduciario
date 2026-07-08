# Automação da atualização diária de preços (renda fixa)

## O problema

Todo dia útil, alguém da área de Riscos precisava atualizar manualmente o
preço de referência (PU) de vários ativos de renda fixa e publicá-lo no
site da empresa. Um processo de várias etapas, repetido diariamente.

- Um erro de cópia ou de arquivo encadearia em um preço errado publicado.

## O que eu fiz

Construí um sistema que faz essa atualização sozinho, do início ao fim:
recalcula os preços a partir das fontes internas e publica no site, sem
intervenção manual.

## Resultado

- Uma rotina manual diária, com vários pontos de falha, passou a rodar
  sozinha, do recálculo até a divulgação do PU.
- A equipe passou a ser avisada automaticamente, com o motivo exato
  quando algo falha, em vez de descobrir o problema horas depois pelo gestor do fundo. 
  
---

Esse é um entre outros projetos de automação que já construí para áreas
de Riscos (precificação de swaps, marcação a mercado, coleta de dados de
mercado). Fico à disposição para conversar sobre qualquer um deles em
mais detalhe.

