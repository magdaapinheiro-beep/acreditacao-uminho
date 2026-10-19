# Acreditação da Oferta Educativa — Universidade do Minho (protótipo)

Página de consulta do estado de avaliação e acreditação institucional e dos
ciclos de estudos da Universidade do Minho junto da A3ES e da DGES.

> **Protótipo de investigação.** Trabalho de Projeto do Mestrado em Direção
> Intermédia na Administração Pública (EEG, UMinho). Não é uma página oficial
> da Universidade do Minho. A faixa que o indica no topo da página deve ser
> mantida em qualquer publicação.

## Estrutura

```
acreditacao-uminho/
├── index.html                    ← página final (gerada; é o que o GitHub Pages publica)
├── dados/
│   ├── acreditacao_uminho.xlsx   ← FONTE DE DADOS: é aqui que se edita
│   └── fonte/                    ← levantamento original do NACC (referência)
├── modelo/pagina.html            ← modelo da página (aspeto, gráficos, filtros)
├── assets/logo-uminho.png        ← logótipo, embutido na página ao gerar
├── scripts/
│   ├── gerar_pagina.py           ← Excel + modelo → index.html
│   └── colunas.py                ← correspondência colunas do Excel ↔ campos
├── docs/
│   ├── PROCESSO.md               ← como a página foi construída
│   └── historico/                ← versões e documentos anteriores
└── .github/workflows/gerar-pagina.yml  ← regenera a página no GitHub
```

## Atualizar os dados

1. Abrir `dados/acreditacao_uminho.xlsx` e editar as folhas **Ciclos** e/ou
   **Institucional** (uma linha por registo; não alterar os cabeçalhos).
   A folha **Leia-me** tem o dicionário de dados.
2. Na folha **Configuração**, pôr a data de atualização (célula amarela).
3. Gerar a página:

   ```bash
   pip install -r requirements.txt     # só da primeira vez
   python3 scripts/gerar_pagina.py
   ```

4. Abrir `index.html` no browser para confirmar e fazer commit/push.

Se o repositório estiver no GitHub, o passo 3 é opcional: basta carregar o
Excel novo (pelo site do GitHub, em *Add file → Upload files*, ou por push) e
a ação `Gerar página` regenera o `index.html` sozinha.

## Publicar no GitHub Pages

1. Criar um repositório no GitHub (ex.: `acreditacao-uminho`) e enviar esta pasta:

   ```bash
   cd acreditacao-uminho
   git init -b main
   git add .
   git commit -m "Versão inicial"
   git remote add origin https://github.com/<utilizador>/acreditacao-uminho.git
   git push -u origin main
   ```

   (Em alternativa, usar o GitHub Desktop: *Add existing repository* → esta pasta → *Publish*.)
2. No repositório: **Settings → Pages → Build and deployment → Source: Deploy
   from a branch → Branch: `main` / `(root)`** → *Save*.
3. Em **Settings → Actions → General → Workflow permissions**, escolher
   *Read and write permissions* (para a ação poder gravar o `index.html`).
4. A página fica em `https://<utilizador>.github.io/acreditacao-uminho/`.

## Mudar o aspeto da página

O aspeto, os textos fixos, os gráficos e os filtros estão em
`modelo/pagina.html`. Os dados entram no marcador `/*__DADOS__*/` e o
logótipo no marcador `__LOGO__` — não os apagar. Depois de editar o modelo,
correr de novo `scripts/gerar_pagina.py`.

## Fontes

- A3ES — consulta pública de processos: <https://si.a3es.pt>
- DGES — registo de ciclos de estudos: <https://www.dges.gov.pt>
- Levantamento do NACC (UMinho), em `dados/fonte/`.
