"""
db.py
-----
Connection pool Oracle partilhado por todos os scripts.
Importar em cada script:

    from db import get_connection, release_connection

Ou usar como context manager:

    with get_connection() as conn:
        cursor = conn.cursor()
        ...
    # a conexão é automaticamente devolvida ao pool
"""

import oracledb
from contextlib import contextmanager

# ── Configuração ──────────────────────────────────────────────
# Alterar estas credenciais para as do ambiente
_DB_HOST    = "SVDCPBDP13"
_DB_PORT    = 1521
_DB_SERVICE = "ORAMIS"
_DB_USER    = "M001781"
_DB_PASS    = "6AIhPOS4AInA"

# ── Limites do pool ───────────────────────────────────────────
# min  : conexões abertas ao arrancar (0 = lazy, só abre quando precisar)
# max  : máximo de conexões simultâneas ao Oracle
# increment: quantas abre de cada vez quando precisa de mais
_POOL_MIN       = 2
_POOL_MAX       = 15
_POOL_INCREMENT = 2
_POOL_TIMEOUT   = 45   # segundos à espera de uma conexão livre antes de dar erro

_pool = None


def _criar_pool() -> oracledb.ConnectionPool:
    dsn = oracledb.makedsn(_DB_HOST, _DB_PORT, service_name=_DB_SERVICE)
    return oracledb.create_pool(
        user=_DB_USER,
        password=_DB_PASS,
        dsn=dsn,
        min=_POOL_MIN,
        max=_POOL_MAX,
        increment=_POOL_INCREMENT,
        getmode=oracledb.POOL_GETMODE_WAIT,
        wait_timeout=_POOL_TIMEOUT * 1000,  # milissegundos
    )


def get_pool() -> oracledb.ConnectionPool:
    """Devolve o pool global, criando-o na primeira chamada."""
    global _pool
    if _pool is None:
        _pool = _criar_pool()
    return _pool


def get_connection() -> oracledb.Connection:
    """
    Adquire uma conexão do pool.
    Lembrar de chamar release_connection() quando terminar,
    ou usar o context manager abaixo.
    """
    return get_pool().acquire()


def release_connection(conn: oracledb.Connection):
    """Devolve a conexão ao pool (não a fecha)."""
    try:
        get_pool().release(conn)
    except Exception:
        pass


@contextmanager
def connection():
    """
    Context manager — garante que a conexão é sempre devolvida ao pool,
    mesmo que ocorra um erro.

    Uso:
        with connection() as conn:
            df = pd.read_sql_query(query, conn)
    """
    conn = None
    try:
        conn = get_connection()
        yield conn
    finally:
        if conn is not None:
            release_connection(conn)


def pool_status() -> dict:
    """Devolve estatísticas do pool (útil para debug)."""
    p = get_pool()
    return {
        "abertos":     p.opened,
        "ocupados":    p.busy,
        "disponiveis": p.opened - p.busy,
        "maximo":      p.max,
    }