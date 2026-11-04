"""
auth.py
-------
Lógica de autenticação e gestão de utilizadores.
Passwords guardadas com hash bcrypt.
"""

import json
import bcrypt
from pathlib import Path

USERS_FILE = Path(__file__).parent / "users.json"

DEPARTAMENTOS = [
    "DIG",
    "CC",
    "DOP",
    "DMCC",
    "DRC",
    "Cont",
    "DQOM"
]
# ─────────────────────────────────────────────
# Leitura / escrita do ficheiro
# ─────────────────────────────────────────────
def carregar_users() -> dict:
    if not USERS_FILE.exists():
        return {}
    with open(USERS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar_users(users: dict):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2, ensure_ascii=False)


# ─────────────────────────────────────────────
# Autenticação
# ─────────────────────────────────────────────
def verificar_login(username: str, password: str) -> dict | None:
    """
    Verifica credenciais. Devolve o objecto do utilizador ou None.
    """
    users = carregar_users()
    user  = users.get(username.strip().lower())
    if not user:
        return None
    if not user.get("activo", True):
        return None
    if bcrypt.checkpw(password.encode("utf-8"), user["password_hash"].encode("utf-8")):
        return user
    return None


# ─────────────────────────────────────────────
# Gestão de utilizadores (só superuser)
# ─────────────────────────────────────────────
def criar_utilizador(username: str, password: str, nome: str,
                     relatorios: list, superuser: bool = False,
                     tipo_acesso: str = "normal", cod_gestor: str = None,
                     cod_un: str = None, departamento: str = None) -> tuple[bool, str]:
    """
    tipo_acesso:
        "normal"  → sem filtro de dados (vê tudo nos relatórios atribuídos)
        "gestor"  → filtra pelo seu COD_GESTOR (só os seus clientes)
        "gerente" → filtra pela UN inteira (todos os clientes da agência)
    """
    users = carregar_users()
    key   = username.strip().lower()

    if key in users:
        return False, "Utilizador já existe."
    if len(password) < 6:
        return False, "A password deve ter pelo menos 6 caracteres."
    if tipo_acesso == "gestor" and not cod_gestor:
        return False, "Indique o código de gestor."
    if tipo_acesso == "gerente" and not cod_un:
        return False, "Indique o código da UN."

    password_hash = bcrypt.hashpw(
        password.encode("utf-8"), bcrypt.gensalt()
    ).decode("utf-8")

    users[key] = {
        "username":      key,
        "nome":          nome.strip(),
        "password_hash": password_hash,
        "superuser":     superuser,
        "relatorios":    relatorios,
        "activo":        True,
        "tipo_acesso":   tipo_acesso,          # "normal" | "gestor" | "gerente"
        "cod_gestor":    str(cod_gestor) if cod_gestor else None,
        "cod_un":        str(cod_un) if cod_un else None,
        "departamento":  departamento,
    }
    guardar_users(users)
    return True, "Utilizador criado com sucesso."


def editar_utilizador(username: str, nome: str = None,
                      relatorios: list = None, activo: bool = None,
                      tipo_acesso: str = None, cod_gestor: str = None,
                      cod_un: str = None, departamento: str = None) -> tuple[bool, str]:
    users = carregar_users()
    key   = username.strip().lower()

    if key not in users:
        return False, "Utilizador não encontrado."
    if users[key].get("superuser") and activo is False:
        return False, "Não é possível desactivar o superuser."

    if nome        is not None: users[key]["nome"]        = nome.strip()
    if relatorios  is not None: users[key]["relatorios"]  = relatorios
    if activo      is not None: users[key]["activo"]      = activo
    if tipo_acesso is not None: users[key]["tipo_acesso"]  = tipo_acesso
    if cod_gestor  is not None: users[key]["cod_gestor"]   = str(cod_gestor) if cod_gestor else None
    if cod_un      is not None: users[key]["cod_un"]       = str(cod_un) if cod_un else None
    if departamento is not None: users[key]["departamento"] = departamento

    guardar_users(users)
    return True, "Utilizador actualizado."


def alterar_password(username: str, nova_password: str) -> tuple[bool, str]:
    users = carregar_users()
    key   = username.strip().lower()

    if key not in users:
        return False, "Utilizador não encontrado."
    if len(nova_password) < 6:
        return False, "A password deve ter pelo menos 6 caracteres."

    users[key]["password_hash"] = bcrypt.hashpw(
        nova_password.encode("utf-8"), bcrypt.gensalt()
    ).decode("utf-8")

    guardar_users(users)
    return True, "Password alterada com sucesso."


def apagar_utilizador(username: str) -> tuple[bool, str]:
    users = carregar_users()
    key   = username.strip().lower()

    if key not in users:
        return False, "Utilizador não encontrado."
    if users[key].get("superuser"):
        return False, "Não é possível apagar o superuser."

    del users[key]
    guardar_users(users)
    return True, "Utilizador apagado."


# ─────────────────────────────────────────────
# Permissões
# ─────────────────────────────────────────────
def relatorios_permitidos(user: dict, todos_relatorios: list) -> list:
    """
    Devolve a lista de relatórios que o utilizador pode ver.
    Superuser vê todos. Os restantes vêem apenas os ids em 'relatorios'.
    """
    if user.get("superuser"):
        return todos_relatorios
    ids_permitidos = set(user.get("relatorios", []))
    return [r for r in todos_relatorios if r["id"] in ids_permitidos]


def filtro_dados_user(user: dict) -> dict:
    """
    Devolve os parâmetros de filtragem de dados a passar aos scripts.

    Devolve:
        {"tipo": "normal",  "valor": None}              → sem filtro (vê tudo)
        {"tipo": "gestor",  "valor": "12345"}            → filtra por COD_GESTOR
        {"tipo": "gerente", "valor": "33"}                → filtra por COD_UNID_NEG (UN inteira)

    Superuser nunca é filtrado, mesmo que tenha tipo_acesso definido.
    """
    if user.get("superuser"):
        return {"tipo": "normal", "valor": None}

    tipo = user.get("tipo_acesso", "normal")

    if tipo == "gestor":
        return {"tipo": "gestor", "valor": user.get("cod_gestor")}
    if tipo == "gerente":
        return {"tipo": "gerente", "valor": user.get("cod_un")}

    return {"tipo": "normal", "valor": None}


# ─────────────────────────────────────────────
# Inicialização — cria superuser se não existir
# ─────────────────────────────────────────────
def inicializar_superuser():
    """
    Cria o superuser padrão se o users.json estiver vazio.
    Credenciais iniciais: admin / admin123  (alterar após primeiro login)
    """
    users = carregar_users()
    if not any(u.get("superuser") for u in users.values()):
        criar_utilizador(
            username="admin",
            password="admin123",
            nome="Administrador",
            relatorios=[],
            superuser=True
        )
        print("⚠️  Superuser criado com credenciais padrão: admin / admin123")