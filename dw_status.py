"""
dw_status.py
------------
Duas verificações relacionadas com a frescura dos dados do DW:

1. verificar_status_dw() — estado geral (disponível/atrasado), com base
   nos parâmetros DTLOAD (BANKA) e DTLOAD_BANKA (MIS) em PARAMMB.PARAMETERS.
   Mostra um badge verde/vermelho no cabeçalho da app.

2. obter_data_liquidacao_disponivel() — só para um relatório específico,
   mostra a data de liquidação mais recente disponível (sem noção de
   "disponível/não disponível", só a data em si).

Ambas cacheadas com TTL curto (5 min) — as queries são leves, mas o
Streamlit volta a correr o script a cada interacção do utilizador, por
isso o cache evita bater na Oracle a cada clique. O cache é partilhado
por todos os utilizadores no mesmo processo, por isso o impacto na app
é insignificante.
"""

import streamlit as st
from datetime import datetime, timedelta
from db import connection

TTL_SEGUNDOS = 300  # 5 minutos

# ─────────────────────────────────────────────
# Estado geral do DW (badge disponível/atrasado)
# ─────────────────────────────────────────────

# (PARAMETER, PARAMETER_TYPE) que têm de estar todos em D0 (hoje)
# Confirmado directamente na tabela: DTLOAD é do tipo BANKA,
# DTLOAD_BANKA é do tipo MIS (troca de nomes invertida da intuição).
PARAMETROS_ESPERADOS = {("DTLOAD", "BANKA"), ("DTLOAD_BANKA", "MIS")}

QUERY_STATUS_DW = """
SELECT PARAMETER, PARAMETER_TYPE, PARAMETER_VALUE, PARAMETER_DESCRIPTION
FROM PARAMMB.PARAMETERS
WHERE TRIM(UPPER(PARAMETER)) IN ('DTLOAD_BANKA','DTLOAD')
  AND TRIM(UPPER(PARAMETER_TYPE)) IN ('BANKA','MIS')
"""


def _parse_data_parametro(valor):
    """Tenta interpretar o valor do parâmetro como data (YYYYMMDD ou
    YYYY-MM-DD, com ou sem hora a seguir). Devolve um date, ou None se
    não conseguir interpretar."""
    if valor is None:
        return None
    valor = str(valor).strip()

    for tamanho, fmt in ((8, "%Y%m%d"), (10, "%Y-%m-%d")):
        pedaco = valor[:tamanho]
        try:
            return datetime.strptime(pedaco, fmt).date()
        except ValueError:
            continue

    return None


@st.cache_data(ttl=TTL_SEGUNDOS)
def verificar_status_dw() -> dict:
    """
    Devolve:
        {"disponivel": True/False/None, "data_ref": "DD/MM/AAAA"}

    disponivel=True  → DTLOAD (BANKA) e DTLOAD_BANKA (MIS) estão ambos
                        na data de hoje
    disponivel=False → pelo menos um dos dois está atrasado
    disponivel=None  → a verificação em si falhou (erro Oracle, parâmetro
                        em falta, ou valor ilegível) — tratado à parte de
                        False, para não afirmar "atrasado" quando na
                        verdade não conseguimos verificar

    data_ref é SEMPRE a data do parâmetro menos 1 dia (o parâmetro marca
    D0 de carregamento, mas os dados em si são de D-1) — isto é só uma
    questão de rótulo visual, a lógica de disponível/atrasado acima não muda.
    """
    hoje = datetime.now().date()
    hoje_str = hoje.strftime("%d/%m/%Y")

    try:
        with connection() as conn:
            cursor = conn.cursor()
            cursor.execute(QUERY_STATUS_DW)
            linhas = cursor.fetchall()
    except Exception as e:
        print(f"⚠️ [dw_status] Erro ao verificar disponibilidade do DW: {e}")
        return {"disponivel": None, "data_ref": hoje_str}

    encontrados = {
        (linha[0].strip().upper(), linha[1].strip().upper()): linha[2]
        for linha in linhas
    }

    em_falta = PARAMETROS_ESPERADOS - set(encontrados.keys())
    if em_falta:
        print(f"⚠️ [dw_status] Parâmetro(s) não encontrado(s) em PARAMMB.PARAMETERS: {em_falta}")
        return {"disponivel": None, "data_ref": hoje_str}

    datas = {}
    for chave, valor_bruto in encontrados.items():
        data_parametro = _parse_data_parametro(valor_bruto)
        if data_parametro is None:
            print(f"⚠️ [dw_status] Valor de {chave} não reconhecido como data: {valor_bruto!r}")
            return {"disponivel": None, "data_ref": hoje_str}
        datas[chave] = data_parametro

    if all(d == hoje for d in datas.values()):
        data_exibida = hoje - timedelta(days=1)
        return {"disponivel": True, "data_ref": data_exibida.strftime("%d/%m/%Y")}
    else:
        data_mais_antiga = min(datas.values())
        data_exibida = data_mais_antiga - timedelta(days=1)
        return {"disponivel": False, "data_ref": data_exibida.strftime("%d/%m/%Y")}


# ─────────────────────────────────────────────
# Data de liquidação disponível para um relatório específico
# (não é um estado disponível/indisponível — só mostra a data)
# ─────────────────────────────────────────────

QUERY_DATA_LIQUIDACAO = """
SELECT /*+ PARALLEL */ MAX(DATA_LIQUIDACAO)
FROM HUB.MOVS_EDST1
"""


@st.cache_data(ttl=TTL_SEGUNDOS)
def obter_data_liquidacao_disponivel() -> dict:
    """
    Devolve a data de liquidação mais recente disponível em HUB.MOVS_EDST1.

    {"data_liquidacao": "DD/MM/AAAA"} em caso de sucesso,
    {"data_liquidacao": None} se a tabela estiver vazia ou a verificação falhar.
    """
    try:
        with connection() as conn:
            cursor = conn.cursor()
            cursor.execute(QUERY_DATA_LIQUIDACAO)
            linha = cursor.fetchone()
    except Exception as e:
        print(f"⚠️ [dw_status] Erro ao obter data de liquidação: {e}")
        return {"data_liquidacao": None}

    if not linha or linha[0] is None:
        return {"data_liquidacao": None}

    valor = linha[0]
    try:
        data_formatada = valor.strftime("%d/%m/%Y")
    except AttributeError:
        # fallback, caso venha como string em vez de datetime
        data_formatada = str(valor)

    return {"data_liquidacao": data_formatada}