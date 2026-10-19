# -*- coding: utf-8 -*-
"""
atualizar.py — Gerador do módulo de Acreditação da Universidade do Minho
========================================================================

O que faz
---------
Lê a tabela de dados (dados.csv), a acreditação institucional (institucional.csv)
e a pasta de PDFs (deliberacoes/) e (re)constrói a página web do módulo em
publicar/modulo_acreditacao_UMinho.html, com a data de "última atualização"
acertada automaticamente. Cada ciclo de estudos cuja deliberação em PDF tenha
sido depositada na pasta deliberacoes/ passa a ligar diretamente para esse PDF.

Como usar (não requer instalação de nada além do Python 3)
----------------------------------------------------------
1. Manter o dados.csv atualizado (uma linha por ciclo de estudos).
2. Depositar os PDF das deliberações na pasta deliberacoes/, nomeando cada
   ficheiro com o n.º de processo (ex.: ACEF_2425_1000796.pdf). Pontos, barras
   ou espaços no nome são ignorados na correspondência.
3. Executar:
       python3 atualizar.py
   ou, para atualizar sozinho sempre que se mexa nos ficheiros:
       python3 atualizar.py --watch

O resultado fica em publicar/ (página + cópia dos PDF), pronto a abrir num
navegador ou a publicar.
"""

import io, os, re, csv, json, sys, time, shutil, datetime

AQUI   = os.path.dirname(os.path.abspath(__file__))
DADOS  = os.path.join(AQUI, "dados.csv")
INST   = os.path.join(AQUI, "institucional.csv")
DELIB  = os.path.join(AQUI, "deliberacoes")
FONTES = os.path.join(AQUI, "fontes")
TPL    = os.path.join(AQUI, "template.html")
SAIDA  = os.path.join(AQUI, "publicar")

CAMPOS_NUM = {"anos"}


def ler_csv(caminho):
    if not os.path.exists(caminho):
        return []
    with io.open(caminho, "r", encoding="utf-8-sig", newline="") as f:
        return [dict(r) for r in csv.DictReader(f, delimiter=";")]


def normalizar(s):
    """Remove tudo o que não seja letra/algarismo, para comparar processos."""
    return re.sub(r"[^A-Za-z0-9]", "", s or "").upper()


def indexar_pdfs():
    """Devolve {proc_normalizado: nome_do_ficheiro} para os PDF em deliberacoes/."""
    idx = {}
    if not os.path.isdir(DELIB):
        return idx
    for nome in sorted(os.listdir(DELIB)):
        if nome.lower().endswith(".pdf"):
            idx[normalizar(os.path.splitext(nome)[0])] = nome
    return idx


def calcular_fim(data_decisao, anos):
    """Se a data de fim não vier preenchida, estima 31 de julho de (ano+anos)."""
    try:
        ano = int(str(data_decisao)[:4])
        return "%d-07-31" % (ano + int(anos))
    except Exception:
        return ""


def construir(ciclos, inst, pdfs):
    usados = set()
    saida_ciclos = []
    for c in ciclos:
        r = {k: (v or "").strip() for k, v in c.items()}
        # tipos
        try:
            r["anos"] = int(r.get("anos") or 0) or ""
        except ValueError:
            r["anos"] = ""
        if not r.get("fim") and r.get("dataDecisao") and r.get("anos"):
            r["fim"] = calcular_fim(r["dataDecisao"], r["anos"])
        # ligar ao PDF depositado, se existir
        chave = normalizar(r.get("proc"))
        pdf = pdfs.get(chave)
        if not pdf:  # tolera nomes que apenas contenham o processo
            for k, v in pdfs.items():
                if chave and chave in k:
                    pdf = v; break
        if pdf:
            r["a3es"] = "deliberacoes/" + pdf
            r["temPdf"] = True
            usados.add(pdf)
        saida_ciclos.append(r)
    orfaos = [p for p in pdfs.values() if p not in usados]
    return saida_ciclos, orfaos


def gerar(importar_pdfs=True):
    if importar_pdfs:
        # 1) tabelas (Excel/CSV): a tabela é a fonte oficial e corre primeiro
        try:
            import importar_tabela
            importar_tabela.importar(verbose=True)
        except Exception as e:
            print("  ! importacao de tabela nao executada:", e)
        # 2) PDF da A3ES: completa campos ainda vazios e liga cada ciclo ao seu PDF
        try:
            import importar_pdf
            importar_pdf.importar(verbose=True)
        except Exception as e:
            print("  ! importacao de PDF nao executada:", e)
    ciclos = ler_csv(DADOS)
    inst   = ler_csv(INST)
    pdfs   = indexar_pdfs()
    if not os.path.exists(TPL):
        raise SystemExit("Falta o template.html em %s" % AQUI)

    ciclos, orfaos = construir(ciclos, inst, pdfs)
    for i in inst:
        for k in list(i):
            i[k] = (i[k] or "").strip()

    hoje = datetime.date.today().isoformat()
    agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    with io.open(TPL, "r", encoding="utf-8") as f:
        html = f.read()
    html = html.replace("__ATUALIZADO__", hoje)
    html = html.replace("__GERADO__", agora)
    html = html.replace("__INSTITUCIONAL__", json.dumps(inst, ensure_ascii=False, indent=1))
    html = html.replace("__CICLOS__", json.dumps(ciclos, ensure_ascii=False, indent=1))

    os.makedirs(SAIDA, exist_ok=True)
    destino = os.path.join(SAIDA, "modulo_acreditacao_UMinho.html")
    with io.open(destino, "w", encoding="utf-8") as f:
        f.write(html)
    # copia com o nome "index.html" para alojamento web (GitHub Pages, Netlify, ...)
    with io.open(os.path.join(SAIDA, "index.html"), "w", encoding="utf-8") as f:
        f.write(html)

    # copiar os PDF para junto da página publicada (para as ligações funcionarem)
    destino_pdfs = os.path.join(SAIDA, "deliberacoes")
    if os.path.isdir(DELIB):
        shutil.rmtree(destino_pdfs, ignore_errors=True)
        shutil.copytree(DELIB, destino_pdfs)

    # relatório
    com_pdf = sum(1 for c in ciclos if c.get("temPdf"))
    print("[%s] pagina atualizada." % agora)
    print("  - ciclos de estudos ....... %d" % len(ciclos))
    print("  - com PDF ligado .......... %d" % com_pdf)
    print("  - sem PDF (so link A3ES) .. %d" % (len(ciclos) - com_pdf))
    if orfaos:
        print("  ! PDF sem linha correspondente no dados.csv (verificar o n.o de processo):")
        for p in orfaos:
            print("      -", p)
    print("  -> ", destino)
    return destino


def assinatura():
    """Impressão digital do estado dos ficheiros, para o modo --watch."""
    marcas = []
    for caminho in (DADOS, INST):
        try: marcas.append((caminho, os.path.getmtime(caminho)))
        except OSError: pass
    for pasta in (DELIB, FONTES):
        if os.path.isdir(pasta):
            for nome in os.listdir(pasta):
                p = os.path.join(pasta, nome)
                try: marcas.append((p, os.path.getmtime(p)))
                except OSError: pass
    return tuple(sorted(marcas))


def main():
    imp = "--sem-importar" not in sys.argv  # por defeito, extrai dos PDF
    if "--watch" in sys.argv:
        print("A vigiar dados.csv, institucional.csv e a pasta deliberacoes/ ...")
        print("(Ctrl+C para parar.)")
        gerar(importar_pdfs=imp)
        ultima = assinatura()
        try:
            while True:
                time.sleep(2)
                atual = assinatura()
                if atual != ultima:
                    ultima = atual
                    try:
                        gerar(importar_pdfs=imp)
                    except Exception as e:
                        print("  ! erro:", e)
        except KeyboardInterrupt:
            print("\nTerminado.")
    else:
        gerar(importar_pdfs=imp)


if __name__ == "__main__":
    main()
