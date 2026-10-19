# Acreditação da Oferta Educativa — Universidade do Minho (protótipo)

Página de consulta do estado de avaliação e acreditação institucional e dos
ciclos de estudos da Universidade do Minho junto da A3ES e da DGES.

**Página publicada:** <https://magdaapinheiro-beep.github.io/acreditacao-uminho/>

> **Protótipo de investigação.** Trabalho de Projeto do Mestrado em Direção
> Intermédia na Administração Pública (UMinho). Não é uma página oficial da
> Universidade do Minho. A faixa que o indica no topo da página deve ser
> mantida em qualquer publicação.

Só são publicados os ciclos das unidades orgânicas que autorizaram a
divulgação, isto é, os que constam do levantamento em vigor na pasta `fontes/`.

## Ficheiros

| Item | Para que serve |
|------|----------------|
| `fontes/` | **O levantamento do NACC em vigor** (Excel). É a fonte oficial: o que não consta daqui não é publicado. Deve estar aqui só um levantamento. |
| `deliberacoes/` | PDF das deliberações da A3ES (opcional), cada um com o nome do n.º de processo (ex.: `ACEF_2425_1000796.pdf`). |
| `correcoes.csv` | Erros conhecidos do levantamento ainda não corrigidos na origem: `proc;campo;valor;motivo`. Aplicados a cada atualização. |
| `dados.csv`, `institucional.csv` | Tabela consolidada (gerada a partir de `fontes/`; não é preciso editar). |
| `atualizar.py` | Lê o levantamento e os PDF, verifica os dados e consolida-os. |
| `importar_tabela.py`, `importar_pdf.py` | Importação do Excel/CSV e dos PDF (chamados pelo `atualizar.py`). |
| `publicar_web.py` | Produz o `index.html` publicado (cabeçalho com o logótipo, indicação de protótipo, botão «Copiar tabela»). |
| `template.html`, `logo_uminho.png` | Modelo da página e logótipo. |
| `index.html` | **A página publicada** (gerada; não editar à mão). |
| `.github/workflows/gerar-pagina.yml` | Refaz a página no GitHub sempre que os dados mudam. |

## Atualizar a página (sem instalar nada)

1. No GitHub, abrir a pasta `fontes/`, apagar o levantamento anterior e
   carregar o novo (*Add file → Upload files*).
2. A ação **Gerar página** corre sozinha: importa o levantamento, verifica-o,
   refaz o `index.html` e grava-o. Em cerca de um minuto a página fica
   atualizada, com a data do dia.
3. Ver o **relatório de verificação** no separador *Actions* → última execução:
   ciclos sem ligação à A3ES/DGES, ligações trocadas (corrigidas sozinhas),
   correções registadas aplicadas e ciclos retirados. O que for assinalado
   deve ser corrigido no levantamento de origem.

## Atualizar no computador (alternativa)

```bash
pip install -r requirements.txt        # só da primeira vez
python3 atualizar.py && python3 publicar_web.py
```

Abrir `index.html` no navegador para confirmar e depois fazer commit/push.
Antes de editar localmente, fazer `git pull` (a ação grava commits próprios).

## Validação humana

O automatismo acelera a publicação mas não dispensa a validação: ciclos
importados de PDF ficam com `validado=nao` em `dados.csv` até serem
conferidos por uma pessoa. O registo da DGES não consta das deliberações da
A3ES e tem de vir do levantamento.

## Fontes

- A3ES — consulta pública de processos: <https://si.a3es.pt>
- DGES — registo de ciclos de estudos: <https://www.dges.gov.pt>
- Levantamento do NACC (UMinho), em `fontes/`.
