# -*- coding: utf-8 -*-
"""
importar_pdf.py — Extração automática de dados a partir dos PDF da A3ES
=======================================================================

Lê os PDF depositados em deliberacoes/ e extrai, de cada um, a informação do
ciclo de estudos, atualizando o dados.csv automaticamente. É chamado pelo
atualizar.py, mas também pode ser executado à parte:

    python3 importar_pdf.py

O n.º de processo é obtido do NOME do ficheiro (fiável): nomeie cada PDF com o
processo, p.ex. ACEF_2425_1000796.pdf  ->  ACEF/2425/1000796. Os restantes
campos são lidos do TEXTO do PDF (melhor esforço).

IMPORTANTE
----------
- Do PDF da A3ES NÃO consta o registo da DGES (n.º e data): esses campos ficam
  em branco e devem ser completados no dados.csv.
- A extração de texto exige `pypdf` (pip install pypdf) ou, em alternativa, o
  Ghostscript (gs). PDF digitalizados (imagem) precisam de OCR e não são lidos.
- Cada linha importada fica marcada com origem=pdf e validado=nao: a informação
  DEVE ser conferida por uma pessoa antes de publicar (página de acesso público).
"""

import io, os, re, csv, sys, subprocess, datetime

AQUI  = os.path.dirname(os.path.abspath(__file__))
DELIB = os.path.join(AQUI, "deliberacoes")
DADOS = os.path.join(AQUI, "dados.csv")

COLUNAS = ["nome","grau","uo","estado","proc","dataDecisao","anos","fim",
           "registo","dataRegisto","a3es","dges","origem","validado"]

MESES = {"janeiro":1,"fevereiro":2,"marco":3,"março":3,"abril":4,"maio":5,
         "junho":6,"julho":7,"agosto":8,"setembro":9,"outubro":10,
         "novembro":11,"dezembro":12}

# Unidades orgânicas da UMinho (para normalizar a coluna uo)
UOS = [
    ("EAAD","Escola de Arquitetura, Arte e Design"),
    ("EC","Escola de Ciências"),
    ("ED","Escola de Direito"),
    ("EE","Escola de Engenharia"),
    ("EEG","Escola de Economia e Gestão"),
    ("ELACH","Escola de Letras, Artes e Ciências Humanas"),
    ("EM","Escola de Medicina"),
    ("EP","Escola de Psicologia"),
    ("ESE","Escola Superior de Enfermagem"),
    ("ICS","Instituto de Ciências Sociais"),
    ("IE","Instituto de Educação"),
    ("I3Bs","Instituto de Investigação em Biomateriais, Biodegradáveis e Biomiméticos"),
]

GRAUS = ["Mestrado Integrado","Doutoramento","Licenciatura","Mestrado"]


def extrair_texto(caminho):
    """Extrai texto do PDF via pypdf; se falhar, tenta o Ghostscript."""
    try:
        from pypdf import PdfReader
        r = PdfReader(caminho)
        t = "\n".join((p.extract_text() or "") for p in r.pages)
        if t.strip():
            return t
    except Exception:
        pass
    try:
        out = subprocess.run(
            ["gs","-q","-dNOPAUSE","-dBATCH","-sDEVICE=txtwrite","-sOutputFile=-", caminho],
            capture_output=True, text=True, timeout=90)
        if out.returncode == 0:
            return out.stdout
    except Exception:
        pass
    return ""


def proc_do_nome(nome_ficheiro):
    """ACEF_2425_1000796.pdf -> ACEF/2425/1000796"""
    base = os.path.splitext(nome_ficheiro)[0]
    partes = [p for p in re.split(r"[^A-Za-z0-9]+", base) if p]
    if len(partes) >= 2 and re.match(r"^[A-Za-z]{2,6}$", partes[0]):
        return partes[0].upper() + "/" + "/".join(partes[1:])
    return ""


def limpar(s):
    return re.sub(r"\s+", " ", (s or "")).strip()


def _data_pt(txt):
    """Converte '18 de Março de 2011' ou '31/07/2019' em AAAA-MM-DD."""
    m = re.search(r"(\d{1,2})\s+de\s+([A-Za-zçãéíóúâ]+)\s+de\s+(\d{4})", txt, re.I)
    if m:
        mes = MESES.get(m.group(2).lower().replace("ç", "c"))
        if mes:
            return "%s-%02d-%02d" % (m.group(3), mes, int(m.group(1)))
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", txt)
    if m:
        return "%s-%s-%s" % (m.group(3), m.group(2), m.group(1))
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", txt)
    return m.group(0) if m else ""


def parse_estado(texto):
    """Lê a caixa assinalada [X] (usa a última = recomendação final)."""
    achados = re.findall(
        r"\[\s*X\s*\]\s*A\s+(n[ãa]o\s+acredita[çc][ãa]o|acredita[çc][ãa]o\s+condicional|acredita[çc][ãa]o)\s+(?:condicional\s+)?do\s+ciclo",
        texto, re.I)
    if achados:
        ult = achados[-1].lower()
        if "não" in ult or "nao" in ult:
            return "Não acreditado"
        if "condicional" in ult:
            return "Acreditado com condições"
        return "Acreditado"
    # Apresentação do pedido: "Decisão do Conselho de Administração\n Acreditar"
    m = re.search(r"Decis[ãa]o do Conselho de Administra[çc][ãa]o\.?\s*\n\s*([^\n|]+)", texto, re.I)
    if m:
        d = m.group(1).lower()
        if "não acredit" in d or "nao acredit" in d:
            return "Não acreditado"
        if "condic" in d:
            return "Acreditado com condições"
        if "acredit" in d:
            return "Acreditado"
    return ""


def parse_anos(texto):
    achados = re.findall(r"Per[ií]odo de acredita[çc][ãa]o\.?\s*\n?\s*(\d{1,2})", texto, re.I)
    if achados:
        return achados[-1]
    m = re.search(r"(\d{1,2})\s+anos", texto, re.I)
    return m.group(1) if m else ""


def parse_grau(texto):
    """Lê o valor junto ao rótulo 'Grau (PT)'."""
    m = re.search(r"\bGrau\b\.?\s*\(PT\)\s*:?\s*\n?\s*([^\n]+)", texto, re.I)
    janela = (m.group(1) if m else texto).lower()
    if re.search(r"doutor|3[.ºo]?\s*ciclo|phd", janela): return "Doutoramento"
    if "mestrado integrado" in janela: return "Mestrado Integrado"
    if "mestrado" in janela or "2[.ºo]? ciclo" in janela: return "Mestrado"
    if "licenciatura" in janela or "1[.ºo]? ciclo" in janela: return "Licenciatura"
    # recurso: procurar no texto todo
    for g in GRAUS:
        if g.lower() in texto.lower():
            return g
    return ""


def parse_uo(texto):
    m = re.search(r"Unidade org[âa]nica[^\n:]*:\s*\n?\s*([^\n]+)", texto, re.I)
    alvo = m.group(1) if (m and "sem resposta" not in m.group(1).lower()) else ""
    procura = (alvo + " " + texto).lower()
    for cod, nome in UOS:
        if nome.lower() in procura:
            return "%s — %s" % (cod, nome)
    return ""


def parse_nome(texto):
    """Designação do ciclo de estudos (PT). O valor pode vir antes ou depois do rótulo."""
    lab = r"Designa[çc][ãa]o do ciclo de estudos\.?\s*\(PT\)\s*:?"
    # (a) valor na mesma linha ou na linha seguinte
    m = re.search(lab + r"\s*\n?\s*([^\n]+)", texto, re.I)
    if m:
        v = limpar(m.group(1))
        if v and not re.match(r"(?i)designa|1\.3|\(EN\)|\[sem", v):
            return re.sub(r"\s*\([^)]*\)\s*$", "", v)
    # (b) valor na linha ANTERIOR ao rótulo (PDF multi-coluna)
    linhas = texto.split("\n")
    for i, l in enumerate(linhas):
        if re.search(lab, l, re.I) and i > 0:
            v = limpar(linhas[i - 1])
            if v and not re.match(r"(?i)designa|1\.[23]|\(EN\)|\[sem|unidade", v):
                return re.sub(r"\s*\([^)]*\)\s*$", "", v)
    return ""


def parse_area(texto):
    m = re.search(r"Área cient[íi]fica predominante do ciclo de estudos\.?\s*\(PT\)\s*:?\s*\n?\s*([^\n]+)", texto, re.I)
    if m:
        v = limpar(m.group(1))
        if v and "predominante" not in v.lower():
            return v
    return ""


def parse_registo(texto):
    """O n.º/data de registo da DGES surge, por vezes, em texto livre."""
    m = re.search(r"registad[oa]\s+(?:na|pela)\s+DGES\s+com\s+o\s+n[úu]mero\s+([RN]/A[-\w./ ]+?)\s+em\s+([\d/]+ de [A-Za-zçãéíóú]+ de \d{4}|\d{2}/\d{2}/\d{4}|\d{1,2} de [A-Za-zçãéíóú]+ de \d{4})", texto, re.I)
    if m:
        return limpar(m.group(1)), _data_pt(m.group(2))
    m = re.search(r"\b(R/A-\w+\s*\d+/\d{4}(?:/AL\d+)?)\b", texto)
    return (limpar(m.group(1)) if m else ""), ""


def parse_data(texto):
    """Data da decisão (melhor esforço). 'A partir de' dá o início da validade."""
    m = re.search(r"Data da decis[ãa]o\.?\s*\n?\s*([\d/]{8,10}|\d{1,2} de [A-Za-zçãéíóú]+ de \d{4})", texto, re.I)
    if m:
        return _data_pt(m.group(1))
    return ""


def parse_inicio(texto):
    m = re.search(r"A partir de:?\s*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.I)
    return _data_pt(m.group(1)) if m else ""


def calcular_fim(data_decisao, anos):
    try:
        return "%d-07-31" % (int(str(data_decisao)[:4]) + int(anos))
    except Exception:
        return ""


def ler_csv():
    linhas, campos = [], list(COLUNAS)
    if os.path.exists(DADOS):
        with io.open(DADOS, "r", encoding="utf-8-sig", newline="") as f:
            r = csv.DictReader(f, delimiter=";")
            campos = list(dict.fromkeys(list(r.fieldnames or []) + COLUNAS))
            linhas = [dict(x) for x in r]
    return linhas, campos


def gravar_csv(linhas, campos):
    with io.open(DADOS, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=campos, delimiter=";", extrasaction="ignore")
        w.writeheader()
        for l in linhas:
            w.writerow({k: l.get(k, "") for k in campos})


def importar(verbose=True):
    if not os.path.isdir(DELIB):
        return 0
    linhas, campos = ler_csv()
    for c in COLUNAS:
        if c not in campos:
            campos.append(c)
    por_proc = {re.sub(r"[^A-Za-z0-9]", "", (l.get("proc") or "")).upper(): l for l in linhas}

    novos, atualizados, ignorados = 0, 0, []
    for ficheiro in sorted(os.listdir(DELIB)):
        if not ficheiro.lower().endswith(".pdf"):
            continue
        proc = proc_do_nome(ficheiro)
        if not proc:
            ignorados.append((ficheiro, "nome sem n.º de processo"))
            continue
        texto = extrair_texto(os.path.join(DELIB, ficheiro))
        if not texto.strip():
            ignorados.append((ficheiro, "sem texto (PDF digitalizado? falta pypdf/gs)"))
            # ainda assim regista a ligação ao PDF, com proc do nome
            texto = ""
        reg, dreg = parse_registo(texto)
        extraido = {
            "proc": proc,
            "nome": parse_nome(texto),
            "grau": parse_grau(texto),
            "uo": parse_uo(texto),
            "estado": parse_estado(texto),
            "dataDecisao": parse_data(texto),
            "anos": parse_anos(texto),
            "registo": reg,
            "dataRegisto": dreg,
        }
        base = extraido["dataDecisao"] or parse_inicio(texto)
        extraido["fim"] = calcular_fim(base, extraido["anos"])
        extraido["a3es"] = "deliberacoes/" + ficheiro

        chave = re.sub(r"[^A-Za-z0-9]", "", proc).upper()
        if chave in por_proc:
            linha = por_proc[chave]
            # atualiza só campos vazios (preserva o que já foi validado à mão)
            mudou = False
            for k, v in extraido.items():
                if v and not (linha.get(k) or "").strip():
                    linha[k] = v; mudou = True
            # a ligação ao PDF é sempre atualizada
            if linha.get("a3es") != extraido["a3es"]:
                linha["a3es"] = extraido["a3es"]; mudou = True
            if mudou:
                atualizados += 1
        else:
            nova = {k: "" for k in campos}
            nova.update(extraido)
            nova["origem"] = "pdf"
            nova["validado"] = "nao"
            linhas.append(nova)
            por_proc[chave] = nova
            novos += 1

    gravar_csv(linhas, campos)
    if verbose:
        print("Importação de PDF:")
        print("  - novos ciclos ......... %d" % novos)
        print("  - ciclos atualizados ... %d" % atualizados)
        if ignorados:
            print("  ! PDF ignorados:")
            for f, motivo in ignorados:
                print("      - %s (%s)" % (f, motivo))
    return novos + atualizados


if __name__ == "__main__":
    importar()
