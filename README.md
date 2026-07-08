# Automação da atualização diária de preços (renda fixa)

*Projeto real, hoje em produção em uma distribuidora de valores mobiliários
(DTVM). Nomes de ativos, sistemas internos e credenciais foram removidos
deste repositório — aqui fica a história do problema e da solução, não o
código.*

## O problema

Todo dia útil, alguém da área de Riscos precisava atualizar manualmente o
preço de referência (PU) de vários ativos de renda fixa e publicá-lo no
site da empresa. Um processo de várias etapas, repetido diariamente, com
duas consequências que ninguém queria assumir:

- Um erro de cópia ou de arquivo virava um preço errado publicado
  publicamente.
- Pior: às vezes o sistema *parecia* ter salvo a atualização, mas não
  tinha. Isso só era percebido horas depois — quando alguém de fora
  notava o preço desatualizado no site.

## O que eu fiz

Construí um sistema que faz essa atualização sozinho, do início ao fim:
recalcula os preços a partir das fontes internas e publica no site, sem
intervenção manual.

Mas o ponto que mais importa não é ter automatizado — é o que descobri
*ao* automatizar. O processo antigo confiava em sinais da tela (um botão
que reaparecia, uma mensagem de "salvo com sucesso") para decidir se
tinha dado certo. Só que esses sinais podiam aparecer mesmo quando o
dado real não tinha sido trocado.

Rastreei a causa raiz e reconstruí a verificação para checar o que
realmente importa — não a aparência da tela, mas o valor de fato
publicado. Hoje o sistema só declara sucesso quando o preço publicado
efetivamente mudou, e avisa a equipe, com o motivo exato, sempre que algo
não confere.

## Resultado

- Uma rotina manual diária, com vários pontos de falha, passou a rodar
  sozinha, do recálculo até a publicação.
- O erro de falso-positivo — o mais arriscado dos dois, porque era
  silencioso — foi eliminado e validado em uso real.
- A equipe passou a ser avisada automaticamente, com o motivo exato
  quando algo falha, em vez de descobrir o problema horas depois por
  fora.

## Por que conto essa história

Não é sobre a ferramenta. É sobre o hábito: quando alguém me diz "está
funcionando", eu confiro o dado, não a tela. Automação sem essa checagem
só troca um erro manual por um erro automático — e mais difícil de
perceber, porque parece que deu certo.

---

Esse é um entre outros projetos de automação que já construí para áreas
de Riscos (precificação de swaps, marcação a mercado, coleta de dados de
mercado). Fico à disposição para conversar sobre qualquer um deles em
mais detalhe.

**Márcio Oliveira** — [github.com/EternalCodeBR](https://github.com/EternalCodeBR)
