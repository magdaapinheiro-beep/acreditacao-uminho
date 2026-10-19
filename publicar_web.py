# -*- coding: utf-8 -*-
"""
publicar_web.py -- Gera a pagina publicada no GitHub Pages (index.html)
======================================================================

Parte da pagina ja gerada pelo atualizar.py (publicar/modulo_acreditacao_UMinho.html)
e prepara-a para publicacao publica como prototipo:

  - troca a barra de navegacao e o trilho que imitam o portal da UMinho
    (menus) por um cabecalho proprio com o logotipo (logo_uminho.png),
    sempre marcado como «Prototipo» e com a nota de que nao e uma pagina
    oficial da Universidade do Minho;
  - troca o «Exportar CSV» por «Copiar tabela», que copia a tabela filtrada
    para colar diretamente no Excel;
  - acerta as margens laterais para telemovel e fixa o tema claro.

Uso:
    python3 atualizar.py && python3 publicar_web.py

Resultado: index.html na raiz do repositorio -- e o ficheiro que o GitHub
Pages publica. No GitHub, a acao «Gerar página» corre estes dois passos
sozinha sempre que os dados mudam.
"""

import io, os, re, sys, base64

AQUI    = os.path.dirname(os.path.abspath(__file__))
ORIGEM  = os.path.join(AQUI, "publicar", "modulo_acreditacao_UMinho.html")
DESTINO = os.path.join(AQUI, "index.html")
LOGO    = os.path.join(AQUI, "logo_uminho.png")

CABECA = """<!doctype html>
<html lang="pt">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Acreditação da Oferta Educativa — Universidade do Minho (protótipo)</title>
<meta name="robots" content="noindex">
"""

CSS_EXTRA = """
  /* ---------- ajustes para a versao publicada ---------- */
  :root{color-scheme:light}
  html,body{background:var(--paper)}
  .wrap{padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
  .proto-bar{background:var(--navbar);color:#fff;font-size:13px}
  .proto-bar .wrap{padding-block:9px;display:flex;flex-wrap:wrap;gap:4px 14px;align-items:baseline}
  .proto-bar b{font-weight:700;letter-spacing:.02em}
  .proto-bar span{opacity:.88}
  button:focus-visible,a:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid var(--brand);outline-offset:2px}
  /* ---------- cabeçalho com o logótipo da UMinho ---------- */
  .um-head{background:#8A2022;color:#fff}
  .um-in{display:flex;align-items:center;justify-content:space-between;gap:16px;padding-block:14px}
  .um-id{display:flex;align-items:center;gap:14px;min-width:0}
  .um-logo{width:52px;height:52px;display:block;flex:none;border-radius:3px}
  .um-name{display:flex;flex-direction:column;line-height:1.25;min-width:0;padding-left:14px;border-left:1px solid rgba(255,255,255,.35)}
  .um-name b{font-size:17px;font-weight:700;letter-spacing:.01em}
  .um-name span{font-size:12.5px;opacity:.88}
  .um-tag{flex:none;font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;border:1px solid rgba(255,255,255,.6);border-radius:99px;padding:4px 11px}
  .um-note{background:#6E1A20;font-size:12.5px;color:rgba(255,255,255,.9)}
  .um-note .wrap{padding-block:7px}
  @media (max-width:560px){ .um-logo{width:42px;height:42px} .um-name b{font-size:15px} .um-name span{font-size:11.5px} .um-tag{display:none} }
</style>"""

BARRA = """<header class="um-head">
  <div class="wrap um-in">
    <div class="um-id">
      <img class="um-logo" src="__LOGO__" alt="Universidade do Minho" width="52" height="52">
      <div class="um-name">
        <b>Universidade do Minho</b>
      </div>
    </div>
    <span class="um-tag">Protótipo</span>
  </div>
  <div class="um-note">
    <div class="wrap">Protótipo de investigação · Trabalho de Projeto do Mestrado em Direção Intermédia na Administração Pública (UMinho). Não é uma página oficial da Universidade do Minho.</div>
  </div>
</header>

<div class="mastwrap">"""


def logo_data_uri():
    """Logotipo embutido na pagina (a pagina funciona como ficheiro unico)."""
    if not os.path.exists(LOGO):
        sys.exit("Falta o logotipo em %s. Nada foi escrito." % LOGO)
    with open(LOGO, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode("ascii")


COPIAR_JS = """document.getElementById("btnCsv").addEventListener("click",ev=>{
  const btn=ev.currentTarget;
  const cel=v=>String(v??"").replace(/[\\t\\r\\n]+/g," ");
  const head=["Ciclo de estudos","Grau","Unidade Orgânica","Estado A3ES","N.º processo","Data decisão","Prazo (anos)","Fim acreditação","Registo DGES","Data registo","Link A3ES","Link DGES"];
  const lines=filtrados().map(c=>[c.nome,c.grau,c.uo,c.estado,c.proc,c.dataDecisao,c.anos||"",c.fim,c.registo,c.dataRegisto,c.a3es,c.dges].map(cel).join("\\t"));
  const texto=head.join("\\t")+"\\n"+lines.join("\\n");
  const aviso=t=>{btn.textContent=t;setTimeout(()=>{btn.textContent="Copiar tabela";},2200);};
  const recurso=()=>{
    const ta=document.createElement("textarea");ta.value=texto;ta.setAttribute("readonly","");
    ta.style.position="fixed";ta.style.opacity="0";document.body.appendChild(ta);ta.select();
    let ok=false;try{ok=document.execCommand("copy");}catch(e){}
    ta.remove();aviso(ok?`Copiadas ${lines.length} linhas`:"Não foi possível copiar");
  };
  try{
    navigator.clipboard.writeText(texto).then(()=>aviso(`Copiadas ${lines.length} linhas`),recurso);
  }catch(e){recurso();}
});
render();
</script>"""


def substituir(html, padrao, novo, nome, flags=re.S):
    novo_html, n = re.subn(padrao, lambda m: novo, html, count=1, flags=flags)
    if n != 1:
        sys.exit("publicar_web.py: nao encontrei «%s» no HTML gerado. "
                 "O template mudou? Nada foi escrito." % nome)
    return novo_html


def main():
    if not os.path.exists(ORIGEM):
        sys.exit("Falta %s -- correr primeiro: python3 atualizar.py" % ORIGEM)
    with io.open(ORIGEM, "r", encoding="utf-8") as f:
        html = f.read()

    html = substituir(html, r'^\s*<meta charset[^>]*>\s*<meta name="viewport"[^>]*>\s*<title>.*?</title>\s*',
                      CABECA, "cabeca")
    html = substituir(html, r"\s*</style>", CSS_EXTRA, "fim do CSS")
    # barra do portal + trilho -> cabecalho com logotipo, marcado como prototipo
    html = substituir(html, r'<div class="pt-nav">.*?<div class="mastwrap">',
                      BARRA.replace("__LOGO__", logo_data_uri()), "barra do portal")
    html = substituir(html, r'<button class="export" id="btnCsv" type="button">Exportar CSV</button>',
                      '<button class="export" id="btnCsv" type="button" '
                      'title="Copia a tabela filtrada para colar no Excel">Copiar tabela</button>',
                      "botao CSV")
    html = substituir(html, r'document\.getElementById\("btnCsv"\)\.addEventListener.*?</script>\s*$',
                      COPIAR_JS, "exportacao CSV")
    if "uminho.pt\"" in html or "pt-nav" in html.split("</style>")[-1]:
        sys.exit("publicar_web.py: ainda ha elementos do portal no HTML. Nada foi escrito.")

    with io.open(DESTINO, "w", encoding="utf-8") as f:
        f.write(html)
    print("Pagina publicada gerada (%d KB):\n  %s" % (len(html.encode("utf-8")) // 1024, DESTINO))


if __name__ == "__main__":
    main()
