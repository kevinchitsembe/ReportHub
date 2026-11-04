"""
diagnosticar_template.py
-------------------------
Corre isto no PC/servidor onde estão os templates (ambiente Windows),
apontando para o template ORIGINAL (não o ficheiro gerado).

Não altera nada — só lista e mostra o conteúdo das partes do .xlsx que
podem estar a causar o aviso "We found a problem with some content"
(ligações de dados externas e desenhos/formas).

Uso:
    python diagnosticar_template.py "C:\\caminho\\para\\o\\template.xlsx"
"""

import sys
import zipfile

if len(sys.argv) < 2:
    print("❌ Uso: python diagnosticar_template.py \"CAMINHO_DO_TEMPLATE.xlsx\"")
    sys.exit(1)

path = sys.argv[1]

with zipfile.ZipFile(path) as z:
    nomes = z.namelist()

    print("=" * 60)
    print("Partes relevantes encontradas no ficheiro:")
    print("=" * 60)
    relevantes = [
        n for n in nomes
        if "connections" in n or "drawing" in n or "externalLink" in n or "queryTable" in n
    ]
    if not relevantes:
        print("(nenhuma encontrada)")
    for n in relevantes:
        print(" -", n)

    if "xl/connections.xml" in nomes:
        print("\n" + "=" * 60)
        print("Conteúdo de xl/connections.xml (ligação de dados externa)")
        print("=" * 60)
        print(z.read("xl/connections.xml").decode("utf-8", errors="replace"))
    else:
        print("\n(não existe xl/connections.xml — sem ligação de dados externa registada)")

    for n in nomes:
        if n.startswith("xl/drawings/drawing") and n.endswith(".xml"):
            print("\n" + "=" * 60)
            print(f"Conteúdo de {n}")
            print("=" * 60)
            print(z.read(n).decode("utf-8", errors="replace"))

    # Também lista quaisquer queryTables (outra forma comum de ligação externa)
    query_tables = [n for n in nomes if "queryTable" in n]
    for n in query_tables:
        print("\n" + "=" * 60)
        print(f"Conteúdo de {n}")
        print("=" * 60)
        print(z.read(n).decode("utf-8", errors="replace"))

print("\n" + "=" * 60)
print("Fim do diagnóstico. Cola o output completo de volta na conversa.")
print("=" * 60)