# Processo de construção da página

Registo de como a página foi construída, para efeitos do Trabalho de Projeto
e para quem a venha a manter.

## 1. Levantamento dos dados

- Ponto de partida: o levantamento da informação da acreditação dos ciclos de
  estudos feito pelo NACC (ficheiro
  `dados/fonte/Levantamento da informação da acreditação dos CE_MP.xlsx`).
- Cada ciclo foi completado e conferido com a consulta pública da A3ES
  (processo, data e sentido da decisão, anos de acreditação, fim da validade)
  e com o registo da DGES (n.º e data de registo, ligação à ficha do curso).
- Os campos `origem` e `validado` de cada registo indicam de onde veio a
  informação e se já foi conferida.
- O registo institucional (acreditação institucional da UMinho pela A3ES)
  está na folha **Institucional**.

## 2. Desenho da página

- Protótipo do módulo de Acreditação proposto em
  `docs/historico/Guia_Modulo_Acreditacao_UMinho.docx`; a primeira versão
  da página está em `docs/historico/modulo_acreditacao_UMinho.html`.
- Identidade visual próxima da dos portais da UMinho (bordeaux, tipografia
  sem serifa), para que o protótipo se leia como uma página institucional
  plausível, sempre com a indicação de que não é oficial.
- Conteúdo:
  - indicadores de topo (ciclos no registo, acreditados, acreditados com
    condições, acreditações a expirar em até 18 meses);
  - cartão da acreditação institucional;
  - **Lista**: tabela com pesquisa, filtros (unidade orgânica, grau, estado)
    e ordenação por coluna, com barra do tempo de acreditação que resta;
  - **Panorama**: gráficos (oferta por grau e escola, situação da
    acreditação, calendário de fim das acreditações, tipos de processo,
    completude das fichas, etc.), que respondem aos filtros.
- A página é um único ficheiro HTML, sem dependências externas: funciona
  offline e pode ser alojada em qualquer servidor estático.

## 3. Alterações de outubro de 2026

- Cabeçalho com o logótipo da Universidade do Minho e a identificação da
  EEG, mais a faixa «Protótipo de investigação … Não é uma página oficial da
  Universidade do Minho».
- Nova opção **«A expirar em até 18 meses»** no filtro de estado, com o
  mesmo critério do indicador de topo (fim da acreditação entre hoje e daqui
  a 18 meses).

## 4. Separação entre dados e página (este repositório)

- Os dados, que estavam escritos dentro do HTML, passaram para
  `dados/acreditacao_uminho.xlsx`.
- `scripts/gerar_pagina.py` lê o Excel, valida cabeçalhos e datas, e injeta
  os dados e o logótipo no modelo `modelo/pagina.html`, produzindo
  `index.html`.
- Verificação: a página gerada a partir do Excel reproduz a versão publicada
  (mesmos 260 ciclos, mesmos indicadores, mesmos 43 ciclos no filtro dos 18
  meses).
- No GitHub, a ação `.github/workflows/gerar-pagina.yml` repete a geração
  sempre que o Excel muda.

## Notas metodológicas

- «A expirar em até 18 meses» conta a partir da data em que a página é
  aberta (não da data de atualização), em meses, sem arredondamento.
- Acreditações já expiradas não entram nesse grupo; aparecem na tabela com
  a indicação «expirou».
- A acreditação plena corresponde a 6 anos; períodos mais curtos aparecem
  destacados no Panorama.
