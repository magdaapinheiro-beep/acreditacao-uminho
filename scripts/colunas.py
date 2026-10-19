"""Correspondência entre as colunas do Excel e os campos usados pela página."""

# (cabeçalho no Excel, campo na página, tipo, descrição)
CICLOS = [
    ("Nome do ciclo de estudos", "nome", "texto", "Designação oficial do curso."),
    ("Grau", "grau", "texto", "Licenciatura, Mestrado Integrado, Mestrado ou Doutoramento."),
    ("Unidade orgânica", "uo", "texto", "Sigla e nome da escola/instituto, no formato «EEG — Escola de Economia, Gestão e Ciência Política»."),
    ("Estado da acreditação", "estado", "texto", "Ex.: «Acreditado», «Acreditado com condições». Alimenta o filtro de estado."),
    ("N.º do processo A3ES", "proc", "texto", "Ex.: ACEF/2526/1200076. O prefixo (ACEF, NCE, PERA…) define o tipo de processo."),
    ("Data da decisão A3ES", "dataDecisao", "data", "Data da deliberação do Conselho de Administração da A3ES."),
    ("Anos de acreditação", "anos", "inteiro", "Duração da acreditação concedida (6 = acreditação plena)."),
    ("Fim da acreditação", "fim", "data", "Último dia de validade. Usado para o indicador e o filtro «a expirar em até 18 meses»."),
    ("N.º de registo DGES", "registo", "texto", "Ex.: R/A-Cr 146/2013/AL02."),
    ("Data do registo DGES", "dataRegisto", "data", "Data do registo do ciclo de estudos na DGES."),
    ("Ligação A3ES", "a3es", "ligacao", "URL da página do processo na consulta pública da A3ES."),
    ("Ligação DGES", "dges", "ligacao", "URL da ficha do curso na DGES."),
    ("Origem do dado", "origem", "texto", "De onde veio o registo (ex.: «excel» = levantamento do NACC)."),
    ("Validado", "validado", "texto", "«sim» quando o registo foi conferido com a A3ES/DGES."),
]

INSTITUCIONAL = [
    ("Entidade", "entidade", "texto", "Instituição avaliada."),
    ("Tipo de avaliação", "tipo", "texto", "Ex.: Acreditação institucional (A3ES)."),
    ("Estado", "estado", "texto", "Resultado da avaliação."),
    ("Nível (cor)", "nivel", "texto", "ok = verde, warn = âmbar, bad = vermelho, neutral = cinzento."),
    ("Validade", "validade", "texto", "Período de validade, ex.: 2023-2029."),
    ("Referência do processo", "ref", "texto", "N.º do processo na A3ES."),
    ("Ligação A3ES", "linkA3es", "ligacao", "URL do processo na A3ES."),
    ("Ligação DGES", "linkDges", "ligacao", "URL da ficha da instituição na DGES."),
]

# campos que a página espera sempre presentes (mesmo vazios)
SEMPRE = {"a3es", "dges"}
