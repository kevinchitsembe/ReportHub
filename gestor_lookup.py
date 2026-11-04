"""
gestor_lookup.py
----------------
Consulta a tabela DWMB.lkp_gestor para validar códigos de gestor e UN
ao criar/editar utilizadores na app.

Tabela: DWMB.lkp_gestor
Colunas usadas: COD_GESTOR, COD_UNID_NEG
"""

from db import connection


def validar_gestor(cod_gestor: str) -> dict | None:
    """
    Confirma que o código de gestor existe na tabela e devolve a sua UN.

    Devolve:
        {"cod_gestor": "123", "cod_unid_neg": "33", "nome": "..."}  ou None
    """
    cod_gestor = str(cod_gestor).strip()

    query = """
        SELECT DISTINCT COD_GESTOR, COD_UNID_NEG
        FROM DWMB.lkp_gestor
        WHERE TO_CHAR(COD_GESTOR) = :cod_gestor
    """

    try:
        with connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, {"cod_gestor": cod_gestor})
            linha = cursor.fetchone()

            if not linha:
                return None

            return {
                "cod_gestor":   str(linha[0]),
                "cod_unid_neg": str(linha[1]),
            }
    except Exception as e:
        print(f"❌ Erro ao validar gestor: {e}")
        return None


def validar_un(cod_unid_neg: str) -> bool:
    """
    Confirma que existe pelo menos um gestor associado a esta UN
    (ou seja, que a UN é válida no DW).
    """
    cod_unid_neg = str(cod_unid_neg).strip()

    query = """
        SELECT COUNT(*)
        FROM DWMB.lkp_gestor
        WHERE TO_CHAR(COD_UNID_NEG) = :cod_un
    """

    try:
        with connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, {"cod_un": cod_unid_neg})
            (total,) = cursor.fetchone()
            return total > 0
    except Exception as e:
        print(f"❌ Erro ao validar UN: {e}")
        return False


def listar_gestores_da_un(cod_unid_neg: str) -> list:
    """Devolve todos os códigos de gestor pertencentes a uma UN."""
    cod_unid_neg = str(cod_unid_neg).strip()

    query = """
        SELECT DISTINCT COD_GESTOR
        FROM DWMB.lkp_gestor
        WHERE TO_CHAR(COD_UNID_NEG) = :cod_un
    """

    try:
        with connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, {"cod_un": cod_unid_neg})
            return [str(r[0]) for r in cursor.fetchall()]
    except Exception as e:
        print(f"❌ Erro ao listar gestores da UN: {e}")
        return []