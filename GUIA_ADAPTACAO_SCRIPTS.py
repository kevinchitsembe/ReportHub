"""
GUIA DE ADAPTAÇÃO DOS SCRIPTS EXISTENTES
─────────────────────────────────────────
Cada script .py precisa de uma pequena alteração:
substituir o input() manual por sys.argv.

ANTES (como estava):
    data_ref = input("📅 Introduza a data de referência (formato: YYYYMMDD): ")

DEPOIS (para scripts com data única):
    import sys
    data_ref = sys.argv[1]          # ex: "20250430"

DEPOIS (para scripts com período):
    import sys
    data_inicio = sys.argv[1]       # ex: "20250401"
    data_fim    = sys.argv[2]       # ex: "20250430"

────────────────────────────────────────────────────────────────
EXEMPLO COMPLETO: cartoes_bloqueados.py adaptado
────────────────────────────────────────────────────────────────
"""

import sys
import xlwings as xw
import pandas as pd
import oracledb
from datetime import datetime
import os

# ── Receber data via argumento (substituí o input()) ──────────
if len(sys.argv) < 2:
    print("❌ Uso: python cartoes_bloqueados.py YYYYMMDD")
    sys.exit(1)

data_ref = sys.argv[1]

# ── Validação básica do formato ───────────────────────────────
try:
    datetime.strptime(data_ref, "%Y%m%d")
except ValueError:
    print(f"❌ Data inválida: '{data_ref}'. Use o formato YYYYMMDD.")
    sys.exit(1)

print(f"📅 Data de referência: {data_ref}")

# ── Resto do script mantém-se igual ──────────────────────────
def conectar_oracle():
    try:
        dsn = oracledb.makedsn("X", 1521, service_name="X")
        conn = oracledb.connect(user="X", password="X", dsn=dsn)
        print("✅ Conexão com Oracle estabelecida com sucesso!")
        return conn
    except oracledb.Error as e:
        print(f"❌ Erro ao conectar ao Oracle: {e}")
        return None

# ... (restante do script igual ao original)
