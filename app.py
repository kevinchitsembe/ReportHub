"""
app.py
------
Aplicação Streamlit — Mapas Banco de Moçambique
"""

import streamlit as st
import json
import os
import time
import glob
from pathlib import Path
from datetime import date, timedelta
from tasks import executar_relatorio_task
from job_store import guardar_job, limpar_job, estado_job
from auth import (
    verificar_login, inicializar_superuser,
    carregar_users, criar_utilizador, editar_utilizador,
    alterar_password, apagar_utilizador, relatorios_permitidos,
    filtro_dados_user, DEPARTAMENTOS
)
from gestor_lookup import validar_gestor, validar_un
from dw_status import verificar_status_dw, obter_data_liquidacao_disponivel
DATA_MINIMA = date(2008, 1, 1)  # ajusta se precisares de recuar mais
# ─────────────────────────────────────────────
# Inicialização
# ─────────────────────────────────────────────
USERS_FILE = Path(__file__).parent / "users.json"

if not USERS_FILE.exists():
    import bcrypt as _bcrypt
    _hash = _bcrypt.hashpw(b"admin123", _bcrypt.gensalt()).decode("utf-8")
    with open(USERS_FILE, "w", encoding="utf-8") as _f:
        json.dump({
            "admin": {
                "username":      "admin",
                "nome":          "Administrador",
                "password_hash": _hash,
                "superuser":     True,
                "relatorios":    [],
                "activo":        True,
            }
        }, _f, indent=2, ensure_ascii=False)
else:
    inicializar_superuser()

# ─────────────────────────────────────────────
# Configuração da página
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Moza | ReportHub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;600;700&family=Source+Code+Pro:wght@400;500&display=swap');
    html, body, [class*="css"] { font-family: 'Open Sans', sans-serif; }

    .stApp { background-color: #FFFFFF; color: #1A1A1A; }
    .block-container { padding-top: 1.5rem !important; padding-bottom: 2rem; max-width: 100% !important; }

    /* ── Forçar sidebar sempre visível, sem botão de colapso ── */
    [data-testid="stSidebarCollapseButton"],
    [data-testid="collapsedControl"] {
        display: none !important;
        visibility: hidden !important;
        pointer-events: none !important;
    }
    section[data-testid="stSidebar"] {
        transform: none !important;
        visibility: visible !important;
        display: block !important;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background-color: #1A1A1A;
        min-width: 270px !important;
        max-width: 270px !important;
    }
    [data-testid="stSidebar"] * { color: #FFFFFF; }

    .sidebar-logo {
        text-align: center;
        padding: 1.2rem 0 0.8rem 0;
        border-bottom: 1px solid #333;
        margin-bottom: 1rem;
    }

    /* Botão refresh do logo — invisível, sobrepõe o logo */
    [data-testid="stSidebar"] [data-testid="stButton-btn_refresh"] > button {
        background: transparent !important;
        border: none !important;
        color: transparent !important;
        font-size: 0.01rem !important;
        height: 4px !important;
        padding: 0 !important;
        margin: 0 !important;
        min-height: 0 !important;
        cursor: pointer !important;
    }
    [data-testid="stSidebar"] [data-testid="stButton-btn_refresh"] > button:hover {
        background: transparent !important;
        border: none !important;
    }

    .sidebar-user {
        font-size: 0.75rem; color: #AAAAAA;
        padding: 0 0 0.8rem 0;
        border-bottom: 1px solid #333;
        margin-bottom: 1rem;
    }
    .sidebar-user strong {
        color: #FFFFFF; font-size: 0.85rem;
        display: block; margin-bottom: 0.2rem;
    }
    .sidebar-super {
        font-size: 0.58rem; background: #E30613; color: white;
        padding: 0.1rem 0.35rem; border-radius: 3px; font-weight: 700;
        letter-spacing: 0.04em; margin-left: 0.3rem;
    }

    .sidebar-section {
        font-size: 0.62rem; font-weight: 700; color: #666;
        text-transform: uppercase; letter-spacing: 0.08em;
        margin: 0.8rem 0 0.4rem 0;
        white-space: nowrap;
    }

    /* ── Search box ── */
    [data-testid="stSidebar"] .stTextInput input {
        background: #2a2a2a !important;
        border: 1px solid #444 !important;
        border-radius: 4px !important;
        color: #FFFFFF !important;
        font-size: 0.8rem !important;
        padding: 0.35rem 0.7rem !important;
    }
    [data-testid="stSidebar"] .stTextInput input::placeholder { color: #555 !important; }
    [data-testid="stSidebar"] .stTextInput input:focus { border-color: #E30613 !important; outline: none !important; }
    [data-testid="stSidebar"] .stTextInput label { display: none !important; }

    /* ── Item relatório na sidebar ── */
    .rel-item-wrap { margin-bottom: 0.4rem; }
    .rel-item-header {
        padding: 0.5rem 0.7rem 0.2rem 0.7rem;
        font-size: 0.8rem; font-weight: 600; color: #DDDDDD;
        line-height: 1.3;
    }
    .rel-item-header.active { color: #FFFFFF; }
    .rel-badge {
        display: inline-block;
        font-size: 0.58rem; font-family: 'Source Code Pro', monospace;
        padding: 0.08rem 0.35rem; border-radius: 2px;
        margin-top: 0.2rem;
        background: #333; color: #999;
    }
    .rel-badge-p { background: #2a1515; color: #E30613; }

    /* Botão seleccionar na sidebar */
    [data-testid="stSidebar"] .stButton > button {
        background: transparent !important;
        border: 1px solid #333 !important;
        color: #AAAAAA !important;
        font-size: 0.75rem !important;
        font-weight: 400 !important;
        padding: 0.25rem 0.7rem !important;
        width: 100% !important;
        text-align: left !important;
        margin: 0.15rem 0 0.4rem 0 !important;
        border-radius: 3px !important;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background: #E30613 !important;
        border-color: #E30613 !important;
        color: #FFFFFF !important;
    }

    /* ── Cabeçalho principal ── */
    .header-wrapper {
        border-bottom: 3px solid #E30613;
        padding: 0 0 0.8rem 0; margin-bottom: 1.5rem;
    }
    .header-wrapper h1 { font-size: 1.3rem; font-weight: 700; color: #1A1A1A; margin: 0; }
    .header-wrapper p  { font-size: 0.78rem; color: #6B6B6B; margin: 0.15rem 0 0 0; }

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"] { gap: 0.5rem; border-bottom: 2px solid #E0E0E0; }
    .stTabs [data-baseweb="tab"] {
        font-size: 0.82rem; font-weight: 600; color: #6B6B6B;
        padding: 0.45rem 0.9rem; border-radius: 4px 4px 0 0;
    }
    .stTabs [aria-selected="true"] { color: #E30613 !important; border-bottom: 2px solid #E30613 !important; }

    /* ── Painel de execução (centralizado e menor) ── */
    .exec-wrap {
        max-width: 580px;
        margin: 0 auto;
        padding: 0 1rem;
    }
    .exec-wrap .stDateInput,
    .exec-wrap .stDateInput > div,
    .exec-wrap .stDateInput > div > div {
        max-width: 100% !important;
    }
    .exec-wrap .stButton > button,
    .exec-wrap .stDownloadButton > button {
        max-width: 100% !important;
    }
    .exec-panel {
        background: #F8F8F8; border: 1px solid #E0E0E0;
        border-top: 3px solid #E30613; border-radius: 4px;
        padding: 1rem 1.2rem 0.8rem 1.2rem; margin-bottom: 1rem;
    }
    .exec-title    { font-size: 0.92rem; font-weight: 700; color: #1A1A1A; margin: 0; }
    .exec-subtitle { font-size: 0.75rem; color: #6B6B6B; margin: 0.15rem 0 0 0; }

    /* ── Log ── */
    .log-box {
        background: #1A1A1A; border: 1px solid #333; border-radius: 4px;
        padding: 0.8rem 1rem; font-family: 'Source Code Pro', monospace;
        font-size: 0.72rem; color: #CCCCCC; max-height: 260px;
        overflow-y: auto; white-space: pre-wrap; word-break: break-all; line-height: 1.7;
    }
    .log-line-ok   { color: #5CB85C; }
    .log-line-err  { color: #E30613; }
    .log-line-warn { color: #F0AD4E; }
    .log-line-info { color: #CCCCCC; }

    /* ── Empty state ── */
    .empty-state {
        max-width: 600px; margin: 3rem auto;
        height: 220px; display: flex; flex-direction: column;
        align-items: center; justify-content: center;
        color: #BBBBBB; font-size: 0.83rem;
        border: 1px dashed #DDDDDD; border-radius: 6px; background: #FAFAFA;
        gap: 0.5rem;
    }
    .empty-icon { font-size: 1.8rem; }

    /* ── Tabela utilizadores ── */
    .user-table { width: 100%; border-collapse: collapse; font-size: 0.82rem; margin-bottom: 1rem; }
    .user-table th {
        background: #1A1A1A; color: #FFF; font-weight: 600;
        padding: 0.55rem 0.8rem; text-align: left;
        font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.05em;
    }
    .user-table td { padding: 0.55rem 0.8rem; border-bottom: 1px solid #E0E0E0; color: #1A1A1A; }
    .user-table tr:hover td { background: #F9F9F9; }
    .tag-activo   { background: #d4edda; color: #155724; padding: 0.1rem 0.4rem; border-radius: 3px; font-size: 0.67rem; font-weight: 600; }
    .tag-inactivo { background: #f8d7da; color: #721c24; padding: 0.1rem 0.4rem; border-radius: 3px; font-size: 0.67rem; font-weight: 600; }
    .tag-super    { background: #E30613; color: #fff; padding: 0.1rem 0.4rem; border-radius: 3px; font-size: 0.67rem; font-weight: 700; }

    /* ── Checkboxes ── */
    .stCheckbox label p { color: #1A1A1A !important; font-size: 0.83rem !important; }

    /* ── Botões principais (login + conteúdo) ── */
    .stButton > button {
        background: #E30613 !important; color: #FFFFFF !important;
        border: none !important; border-radius: 4px;
        font-family: 'Open Sans', sans-serif; font-size: 0.83rem; font-weight: 700;
        transition: none !important;
    }
    .stButton > button:hover    { background: #E30613 !important; color: #FFFFFF !important; }
    .stButton > button:active   { background: #E30613 !important; color: #FFFFFF !important; }
    .stButton > button:disabled { background: #CCCCCC !important; color: #888 !important; }

    /* ── Botões sidebar (Seleccionar + Logout) — sobrepõe o selector acima ── */
    [data-testid="stSidebar"] .stButton > button {
        background: transparent !important;
        border: 1px solid #444 !important;
        color: #CCCCCC !important;
        font-size: 0.75rem !important;
        font-weight: 600 !important;
        padding: 0.28rem 0.7rem !important;
        width: 100% !important;
        border-radius: 3px !important;
        margin: 0.1rem 0 0.5rem 0 !important;
        transition: none !important;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background: #E30613 !important;
        border-color: #E30613 !important;
        color: #FFFFFF !important;
    }

    .stDownloadButton > button {
        background: #4A4A4A !important; color: #FFF !important;
        border: none !important; border-radius: 4px;
        font-family: 'Open Sans', sans-serif; font-size: 0.83rem; font-weight: 700;
        transition: background 0.15s !important;
    }
    .stDownloadButton > button:hover { background: #333 !important; }

    .stDateInput label { font-size: 0.78rem; font-weight: 600; color: #1A1A1A; }

    /* ── Logo clicável ── */
    .logo-link { cursor: pointer; display: inline-block; }
    .logo-link:hover { opacity: 0.85; }

    hr { border-color: #E0E0E0; }
    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# JS — forçar sidebar sempre aberta e esconder botão de colapso
st.markdown("""
<script>
    const fixSidebar = () => {
        const doc = window.parent.document;
        // Esconder botão de colapso
        doc.querySelectorAll('[data-testid="stSidebarCollapseButton"], [data-testid="collapsedControl"]')
           .forEach(el => el.style.display = 'none');
        // Garantir sidebar visível
        const sb = doc.querySelector('section[data-testid="stSidebar"]');
        if (sb) { sb.style.transform = 'none'; sb.style.visibility = 'visible'; }
    };
    fixSidebar();
    setTimeout(fixSidebar, 300);
    setTimeout(fixSidebar, 800);
    setTimeout(fixSidebar, 1500);
</script>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Utilitários
# ─────────────────────────────────────────────
@st.cache_data
def carregar_config():
    config_path = Path(__file__).parent / "config.json"
    if not config_path.exists():
        return []
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def formatar_linha_log(linha: str) -> str:
    if linha.startswith("✅"):   return f'<span class="log-line-ok">{linha}</span>'
    elif linha.startswith("❌"): return f'<span class="log-line-err">{linha}</span>'
    elif linha.startswith("⚠️"): return f'<span class="log-line-warn">{linha}</span>'
    else:                        return f'<span class="log-line-info">{linha}</span>'


def limpar_temp_antigos(max_minutos: int = 30):
    pasta = Path(__file__).parent / "temp"
    if not pasta.exists(): return
    agora = time.time()

    # Relatórios gerados, dentro de temp/<username>/
    for subpasta in pasta.iterdir():
        if subpasta.is_dir():
            for f in subpasta.glob("*.xlsx"):
                try:
                    if (agora - os.path.getmtime(f)) > max_minutos * 60:
                        f.unlink()
                except Exception:
                    pass

    # Ficheiros de upload (entidades, saldos médios, etc.), gravados
    # directamente na raiz de temp/ — não dentro de nenhuma subpasta
    for padrao in ("*.xlsx", "*.xls", "*.csv"):
        for f in pasta.glob(padrao):  # não recursivo — só o nível de topo
            try:
                if (agora - os.path.getmtime(f)) > max_minutos * 60:
                    f.unlink()
            except Exception:
                pass

# ─────────────────────────────────────────────
# Estado da sessão
# ─────────────────────────────────────────────
for key, default in [
    ("logged_in",     False),
    ("user",          None),
    ("relatorio_sel", None),
    ("task_id",       None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ═════════════════════════════════════════════
# ECRÃ DE LOGIN
# ═════════════════════════════════════════════
if not st.session_state.logged_in:
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        st.markdown("""
        <div style="text-align:center; margin:3rem 0 1.5rem 0;">
            <img src="https://www.mozabanco.co.mz/Content/Images/logo_moza_red.png"
                style="height:48px;" onerror="this.style.display='none'">
        </div>
        <div style="background:#F8F8F8; border-top:3px solid #E30613; border-radius:6px;
                    padding:1.8rem 2rem 1.5rem 2rem; border:1px solid #E0E0E0;">
            <p style="font-size:0.95rem; font-weight:700; color:#1A1A1A; margin:0 0 1.2rem 0; text-align:center;">
                Login
            </p>
        </div>
        """, unsafe_allow_html=True)
        username = st.text_input("Utilizador", placeholder="username")
        password = st.text_input("Password",   placeholder="••••••••", type="password")
        if st.button("Entrar", use_container_width=True):
            if not username or not password:
                st.error("Preencha o utilizador e a password.")
            else:
                user = verificar_login(username, password)
                if user:
                    st.session_state.logged_in = True
                    st.session_state.user      = user
                    st.rerun()
                else:
                    st.error("Credenciais inválidas ou utilizador inactivo.")
    st.stop()


# ═════════════════════════════════════════════
# PLATAFORMA (após login)
# ═════════════════════════════════════════════
user         = st.session_state.user
is_superuser = user.get("superuser", False)
todos_relats = carregar_config()
meus_relats  = relatorios_permitidos(user, todos_relats)


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    # Logo + nome centrados e clicáveis para refresh
    st.markdown("""
    <div class="sidebar-logo">
        <div style="display:flex; flex-direction:column; align-items:center; gap:0.5rem;">
            <img src="https://www.mozabanco.co.mz/Content/Images/logo_moza_red.png"
                 alt="Moza" onerror="this.style.display='none'"
                 style="height:34px; filter:brightness(0) invert(1);">
            <span style="color:#FFFFFF; font-size:0.88rem; font-weight:700; letter-spacing:0.03em;">Moza Banco</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("↺", key="btn_refresh", help="Recarregar", use_container_width=True):
        st.rerun()

    # Utilizador
    super_tag = '<span class="sidebar-super">SUPER</span>' if is_superuser else ""
    st.markdown(f"""
    <div class="sidebar-user">
        <strong>👤 {user.get('nome', user.get('username', ''))}</strong>{super_tag}
        Sessão iniciada
    </div>
    """, unsafe_allow_html=True)

    # Pesquisa
    st.markdown('<p class="sidebar-section">Pesquisar relatório</p>', unsafe_allow_html=True)
    pesquisa = st.text_input("pesquisa", placeholder="Escreva para filtrar...",
                              label_visibility="collapsed", key="pesquisa_input")

    # Filtrar
    termo = pesquisa.strip().lower()
    relats_filtrados = [
        r for r in meus_relats
        if termo in r["titulo"].lower() or termo in r.get("descricao", "").lower()
    ] if termo else meus_relats

    # Sugestões em tempo real (só quando está a pesquisar)
    if termo and relats_filtrados:
        st.markdown(f'<p style="font-size:0.65rem; color:#666; margin:0.2rem 0 0.4rem 0;">'
                    f'{len(relats_filtrados)} resultado(s)</p>', unsafe_allow_html=True)

    # Lista de relatórios
    st.markdown('<p class="sidebar-section">Relatórios disponíveis</p>', unsafe_allow_html=True)

    if not relats_filtrados:
        st.markdown('<p style="font-size:0.75rem; color:#555; padding:0.3rem 0.5rem;">'
                    'Nenhum resultado encontrado.</p>', unsafe_allow_html=True)
    else:
        for rel in relats_filtrados:
            badge_txt = "Data única" if rel["modo_data"] == "simples" else "Período"
            badge_cls = "rel-badge" if rel["modo_data"] == "simples" else "rel-badge rel-badge-p"
            is_active = st.session_state.relatorio_sel == rel["id"]
            act_cls   = "active" if is_active else ""

            # Nome + badge
            st.markdown(f"""
            <div class="rel-item-wrap">
                <div class="rel-item-header {act_cls}">
                    {rel['titulo']}<br>
                    <span class="{badge_cls}">{badge_txt}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Botão seleccionar
            if st.button("Seleccionar", key=f"sb_{rel['id']}", use_container_width=True):
                st.session_state.relatorio_sel = rel["id"]
                st.session_state.task_id       = None
                st.rerun()

    # Logout
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("---")
    if st.button("🚪 Terminar sessão", use_container_width=True):
        limpar_job(user.get("username", ""))
        for k in ["logged_in", "user", "relatorio_sel", "task_id"]:
            st.session_state[k] = False if k == "logged_in" else None
        st.rerun()


# ─────────────────────────────────────────────
# CABEÇALHO
# ─────────────────────────────────────────────
status_dw = verificar_status_dw()

if status_dw["disponivel"] is True:
    badge_dw = (f'<span class="tag-activo" style="font-size:0.78rem; padding:0.3rem 0.7rem;">'
                f'✅ Dados Disponíveis ({status_dw["data_ref"]})</span>')
elif status_dw["disponivel"] is False:
    badge_dw = (f'<span class="tag-inactivo" style="font-size:0.78rem; padding:0.3rem 0.7rem;">'
                f'❌ Dados Não Disponíveis ({status_dw["data_ref"]})</span>')
else:
    badge_dw = ""  # erro na verificação — não afirma nada, evita alarme falso

st.markdown(f"""
<div class="header-wrapper" style="display:flex; justify-content:space-between; align-items:flex-end;">
    <div>
        <h1>ReportHub</h1>
        <p>Selecciona o relatório na barra lateral, define o período e gera o ficheiro Excel</p>
    </div>
    <div>{badge_dw}</div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tabs_labels = ["📊 Relatórios", "⚙️ Gestão de Utilizadores"] if is_superuser else ["📊 Relatórios"]
tab_objects  = st.tabs(tabs_labels)
tab_rel      = tab_objects[0]
tab_gestao   = tab_objects[1] if is_superuser else None


# ═════════════════════════════════════════════
# TAB — RELATÓRIOS
# ═════════════════════════════════════════════
with tab_rel:
    if not meus_relats:
        st.info("Não tens relatórios atribuídos. Contacta o administrador.")

    elif st.session_state.relatorio_sel is None:
        st.markdown("""
        <div class="empty-state">
            <div class="empty-icon">📂</div>
            <div>Seleccione um relatório na barra lateral para começar</div>
        </div>
        """, unsafe_allow_html=True)

    else:
        ids_permitidos = [r["id"] for r in meus_relats]
        if st.session_state.relatorio_sel not in ids_permitidos:
            st.session_state.relatorio_sel = None
            st.rerun()

        rel = next((r for r in todos_relats if r["id"] == st.session_state.relatorio_sel), None)
        if rel is None:
            st.session_state.relatorio_sel = None
            st.rerun()

        # Centrar com colunas nativas do Streamlit
        _, col_mid, _ = st.columns([1, 3, 1])
        with col_mid:
            modo = rel["modo_data"]
            badge_map = {"simples": "Data única", "periodo": "Período", "periodo_upload": "Período + Upload", "simples_upload": "Simples + Upload",
                         "upload_ficheiro": "Upload", "periodo_numero": "filtro_like", "char_sem_data":"filtro_char_s", "sem_filtros": "relatórios sem filtros"}
            badge_txt = badge_map.get(modo, "Período")

            st.markdown(f"""
            <div class="exec-panel">
                <p class="exec-title">{rel['titulo']}</p>
                <p class="exec-subtitle">{rel.get('descricao', '')} &nbsp;·&nbsp; {badge_txt}</p>
            </div>
            """, unsafe_allow_html=True)

            if rel["id"] == "pagamentos_cartao":
                info_liquidacao = obter_data_liquidacao_disponivel()
                if info_liquidacao["data_liquidacao"]:
                    st.markdown(
                        f'<p style="font-size:0.78rem; color:#6B6B6B; margin:0 0 0.8rem 0;">'
                        f'📅 Data de Liquidação: <strong>{info_liquidacao["data_liquidacao"]}</strong></p>',
                        unsafe_allow_html=True
                    )

            hoje = date.today()
            ficheiro_upload_path = None
            username = user.get("username", "")

            # ── Inputs de data ──
            if modo in ("simples", "simples_upload"):
                data_sel = st.date_input("Data de referência",
                                         value=hoje - timedelta(days=1),
                                         min_value=DATA_MINIMA,
                                         format="DD/MM/YYYY", key="data_simples")
                data_inicio_str = data_sel.strftime("%Y%m%d")
                data_fim_str    = None
            
            elif modo in ("upload_ficheiro","char_sem_data","sem_filtros"):
                data_inicio_str = None
                data_fim_str    = None

            else:
                c1, c2 = st.columns(2)
                with c1:
                    data_ini = st.date_input("Data de início",
                                             value=hoje - timedelta(days=30),
                                             min_value=DATA_MINIMA,
                                             format="DD/MM/YYYY", key="data_ini")
                with c2:
                    data_fim = st.date_input("Data de fim",
                                             value=hoje - timedelta(days=1),
                                             min_value=DATA_MINIMA,
                                             format="DD/MM/YYYY", key="data_fim")
                if data_fim < data_ini:
                    st.warning("⚠️ A data de fim não pode ser anterior à data de início.")
                    st.stop()
                data_inicio_str = data_ini.strftime("%Y%m%d")
                data_fim_str    = data_fim.strftime("%Y%m%d")

            valor_filtro_like = None
            valor_filtro_gar = None
            # ── Upload ──
            if modo in ("periodo_upload", "upload_ficheiro", "simples_upload"):
                st.markdown("---")
                st.markdown('<p style="font-size:0.8rem; font-weight:600; color:#1A1A1A; margin-bottom:0.3rem;">Carregar Ficheiro</p>', unsafe_allow_html=True)
                st.markdown('<p style="font-size:0.73rem; color:#6B6B6B; margin-bottom:0.5rem;">Carrega o ficheiro .xlsx ou .csv com os códigos</p>', unsafe_allow_html=True)
                uploaded = st.file_uploader("entidades", type=["xlsx", "xls", "csv"],
                                            label_visibility="collapsed", key="upload_entidades")
                if uploaded:
                    pasta_temp = Path(__file__).parent / "temp"
                    pasta_temp.mkdir(exist_ok=True)
                    caminho_upload = pasta_temp / f"entidades_upload_{int(time.time())}{Path(uploaded.name).suffix}"
                    with open(caminho_upload, "wb") as f:
                        f.write(uploaded.getbuffer())
                    ficheiro_upload_path = str(caminho_upload)
                    st.success(f"✅ {uploaded.name} carregado.")

            elif modo == "periodo_numero":
                            st.markdown("---")
                            st.markdown('<p style="font-size:0.8rem; font-weight:600; color:#1A1A1A; margin-bottom:0.3rem;">Filtro</p>', unsafe_allow_html=True)
                            st.markdown('<p style="font-size:0.73rem; color:#6B6B6B; margin-bottom:0.5rem;">Introduz o número (prefixo) a filtrar</p>', unsafe_allow_html=True)
            
                            valor_filtro_like = st.text_input(
                                "filtro_like", label_visibility="collapsed", key="filtro_like_input"
                            ).strip()
            
                            if valor_filtro_like and not valor_filtro_like.isdigit():
                                st.error("❌ O filtro deve conter apenas dígitos.")

            elif modo == "char_sem_data":
                            st.markdown("---")
                            st.markdown('<p style="font-size:0.8rem; font-weight:600; color:#1A1A1A; margin-bottom:0.3rem;">Tipos de Garantia</p>', unsafe_allow_html=True)
                            st.markdown('<p style="font-size:0.73rem; color:#6B6B6B; margin-bottom:0.5rem;">Introduz os códigos dos tipos de garantia separados por vírgula (ex: 3, 8, 5)</p>', unsafe_allow_html=True)
            
                            valor_filtro_gar = st.text_input(
                                "filtro_gar", 
                                label_visibility="collapsed", 
                                key="filtro_gar_input",
                                placeholder="Exemplo: 3, 8, 5"
                            ).strip()
            
                            if valor_filtro_gar:
                                # Remover espaços e separar por vírgula
                                tipos_gar = [t.strip() for t in valor_filtro_gar.split(',') if t.strip()]
                                
                                # Verificar se todos são dígitos
                                tipos_invalidos = [t for t in tipos_gar if not t.isdigit()]
                                if tipos_invalidos:
                                    st.error(f"❌ Tipos inválidos: {', '.join(tipos_invalidos)}. Devem conter apenas dígitos.")
                                elif len(tipos_gar) == 0:
                                    st.warning("⚠️ Por favor, insira pelo menos um tipo de garantia.")
                                else:
                                    st.success(f"✅ Tipos selecionados: {', '.join(tipos_gar)}")

            st.markdown("---")

            # ── Estado do job actual ──
            job = estado_job(username)
            job_activo = job["existe"] and job["estado"] in ("PENDING", "STARTED", "PROGRESS")
            upload_em_falta = (modo in ("periodo_upload", "upload_ficheiro", "simples_upload") and not ficheiro_upload_path)

            if upload_em_falta:
                st.info("⬆️ Carregue o ficheiro para continuar.")

            # ── Botão Gerar ──
            col_btn, col_cancelar = st.columns([3, 1])
            with col_btn:
                gerar = st.button("▶  Gerar relatório", use_container_width=True,
                                  disabled=job_activo or upload_em_falta)
            with col_cancelar:
                if job_activo:
                    if st.button("✕ Cancelar", use_container_width=True):
                        limpar_job(username)
                        st.session_state.task_id = None
                        st.rerun()

            # ── Submeter novo job ──
            if gerar:
                limpar_temp_antigos(max_minutos=30)
                filtro = filtro_dados_user(user)
                task = executar_relatorio_task.delay(
                    script_path=rel["script"],
                    modo_data=modo,
                    data_inicio=data_inicio_str,
                    data_fim=data_fim_str,
                    ficheiro_upload=ficheiro_upload_path,
                    filtro_tipo=filtro["tipo"],
                    filtro_valor=filtro["valor"],
                    username=username,
                    valor_filtro=valor_filtro_like,
                    valor_filtro_sdata=valor_filtro_gar
                )
                guardar_job(username, task.id, rel["id"])
                st.session_state.task_id = task.id
                st.rerun()

            # ── Mostrar estado do job ──
            if job["existe"]:
                estado_actual = job["estado"]

                # Ícone e mensagem de estado
                if estado_actual in ("PENDING", "STARTED", "PROGRESS"):
                    icone, cor = "⏳", "#F0AD4E"
                    msg = "Na fila..." if estado_actual == "PENDING" else job["status_msg"]
                elif estado_actual == "SUCCESS":
                    icone, cor = "✅", "#5CB85C"
                    msg = "Concluído"
                elif estado_actual == "FAILURE":
                    icone, cor = "❌", "#E30613"
                    msg = job["status_msg"]
                else:
                    icone, cor = "⚠️", "#F0AD4E"
                    msg = job["status_msg"]

                st.markdown(f"""
                <div style="display:flex; align-items:center; gap:0.5rem;
                            margin-bottom:0.8rem; font-size:0.82rem;">
                    <span style="font-size:1rem;">{icone}</span>
                    <span style="color:{cor}; font-weight:600;">{msg}</span>
                </div>
                """, unsafe_allow_html=True)

                # Log
                if job["log"]:
                    html_log = "\n".join(formatar_linha_log(l) for l in job["log"])
                    st.markdown(f'<div class="log-box">{html_log}</div>',
                                unsafe_allow_html=True)

                # Auto-refresh enquanto a executar
                if estado_actual in ("PENDING", "STARTED", "PROGRESS"):
                    time.sleep(2)
                    st.rerun()

                # Botão download quando concluído
                if estado_actual == "SUCCESS" and job["excel_path"]:
                    excel_path = job["excel_path"]
                    if Path(excel_path).exists():
                        st.markdown("---")
                        with open(excel_path, "rb") as f:
                            nome_ficheiro = Path(excel_path).name
                            st.download_button(
                                label=f"⬇️  Descarregar  {nome_ficheiro}",
                                data=f,
                                file_name=nome_ficheiro,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        # Botão para iniciar novo relatório
                        if st.button("↺  Novo relatório", use_container_width=True):
                            limpar_job(username)
                            st.session_state.task_id = None
                            st.rerun()
                    else:
                        st.warning("⚠️ O ficheiro expirou. Gere o relatório novamente.")
                        limpar_job(username)


# ═════════════════════════════════════════════
# TAB — GESTÃO DE UTILIZADORES
# ═════════════════════════════════════════════
if is_superuser and tab_gestao is not None:
    with tab_gestao:
        st.markdown("#### Utilizadores registados")

        users = carregar_users()

        def _linha_html(uname, u):
            estado = '<span class="tag-activo">Activo</span>'   if u.get("activo", True) \
                else '<span class="tag-inactivo">Inactivo</span>'
            tipo   = '<span class="tag-super">SUPER</span>'     if u.get("superuser") else "Normal"
            n_rels = "Todos" if u.get("superuser") else str(len(u.get("relatorios", [])))

            if u.get("superuser"):
                filtro_txt = "—"
            else:
                tipo_acesso = u.get("tipo_acesso", "normal")
                if tipo_acesso == "gestor":
                    filtro_txt = f"Gestor {u.get('cod_gestor', '—')}"
                elif tipo_acesso == "gerente":
                    filtro_txt = f"UN {u.get('cod_un', '—')} (completa)"
                else:
                    filtro_txt = "Sem filtro"

            return f"""
            <tr>
                <td><strong>{u.get('nome','')}</strong></td>
                <td>{uname}</td><td>{tipo}</td>
                <td>{n_rels} relatório(s)</td>
                <td>{filtro_txt}</td>
                <td>{estado}</td>
            </tr>"""

        def _tabela_html(users_dict):
            linhas = "".join(_linha_html(uname, u) for uname, u in users_dict.items())
            return f"""
            <table class="user-table">
                <thead><tr>
                    <th>Nome</th><th>Username</th><th>Tipo</th><th>Acessos</th><th>Filtro de dados</th><th>Estado</th>
                </tr></thead>
                <tbody>{linhas}</tbody>
            </table>
            """

        # Administradores ficam sempre visíveis, sem agrupamento por departamento
        admins = {u: d for u, d in users.items() if d.get("superuser")}
        if admins:
            st.markdown(_tabela_html(admins), unsafe_allow_html=True)

        # Restantes utilizadores, agrupados por departamento
        restantes = {u: d for u, d in users.items() if not d.get("superuser")}
        por_departamento = {}
        for uname, u in restantes.items():
            dep = u.get("departamento") or "Sem departamento"
            por_departamento.setdefault(dep, {})[uname] = u

        for dep in sorted(por_departamento):
            users_dep = por_departamento[dep]
            with st.expander(f"🏢 {dep}  ·  {len(users_dep)} utilizador(es)"):
                st.markdown(_tabela_html(users_dep), unsafe_allow_html=True)

        st.markdown("---")

        accao = st.selectbox("Acção", [
            "➕ Criar utilizador",
            "✏️  Editar utilizador",
            "🔑 Alterar password",
            "🗑️  Apagar utilizador",
        ], label_visibility="collapsed")

        nomes_users_normais = [u for u in users if not users[u].get("superuser")]

        # ── Criar ──
        if accao == "➕ Criar utilizador":
            st.markdown("**Novo utilizador**")

            # Fora do form — precisa de re-renderizar ao mudar a opção
            tipo_acesso_sel = st.selectbox(
                "Tipo de acesso",
                ["normal", "gestor", "gerente"],
                format_func=lambda t: {
                    "normal":  "Normal — vê todos os dados dos relatórios atribuídos",
                    "gestor":  "Gestor — vê apenas os clientes do seu código de gestor",
                    "gerente": "Gerente — vê todos os clientes da sua UN",
                }[t],
                key="novo_tipo_acesso"
            )

            cod_gestor_input = None
            cod_un_input     = None

            if tipo_acesso_sel == "gestor":
                cod_gestor_input = st.text_input("Código de gestor (COD_GESTOR)", key="novo_cod_gestor")
                if cod_gestor_input:
                    info = validar_gestor(cod_gestor_input)
                    if info:
                        st.success(f"✅ Gestor válido — UN {info['cod_unid_neg']}")
                    else:
                        st.error("❌ Código de gestor não encontrado na tabela DWMB.lkp_gestor.")

            elif tipo_acesso_sel == "gerente":
                cod_un_input = st.text_input("Código da UN (COD_UNID_NEG)", key="novo_cod_un")
                if cod_un_input:
                    if validar_un(cod_un_input):
                        st.success(f"✅ UN {cod_un_input} válida.")
                    else:
                        st.error("❌ Código de UN não encontrado na tabela DWMB.lkp_gestor.")

            departamento_sel = st.selectbox("Departamento", DEPARTAMENTOS, key="novo_departamento")

            with st.form("form_criar", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown('<p style="font-size:0.8rem; font-weight:600; color:#1A1A1A; margin-bottom:0.3rem;">Nome completo</p>', unsafe_allow_html=True)
                    novo_nome = st.text_input("Nome completo", label_visibility="collapsed", key="novo_nome")
                    st.markdown('<p style="font-size:0.8rem; font-weight:600; color:#1A1A1A; margin-bottom:0.3rem;">Username</p>', unsafe_allow_html=True)
                    novo_user = st.text_input("Username", label_visibility="collapsed", key="novo_user")
                with c2:
                    nova_pass  = st.text_input("Password",          type="password")
                    nova_pass2 = st.text_input("Confirmar password", type="password")
                st.markdown("**Relatórios permitidos**")
                relats_sel = []
                cols = st.columns(2)
                for i, r in enumerate(todos_relats):
                    with cols[i % 2]:
                        if st.checkbox(r["titulo"], key=f"new_rel_{r['id']}"):
                            relats_sel.append(r["id"])

                if st.form_submit_button("Criar utilizador", use_container_width=True):
                    if nova_pass != nova_pass2:
                        st.error("As passwords não coincidem.")
                    elif tipo_acesso_sel == "gestor" and not cod_gestor_input:
                        st.error("Indique o código de gestor.")
                    elif tipo_acesso_sel == "gestor" and not validar_gestor(cod_gestor_input):
                        st.error("Código de gestor inválido. Verifique antes de criar o utilizador.")
                    elif tipo_acesso_sel == "gerente" and not cod_un_input:
                        st.error("Indique o código da UN.")
                    elif tipo_acesso_sel == "gerente" and not validar_un(cod_un_input):
                        st.error("Código de UN inválido. Verifique antes de criar o utilizador.")
                    else:
                        ok, msg = criar_utilizador(
                            novo_user, nova_pass, novo_nome, relats_sel,
                            tipo_acesso=tipo_acesso_sel,
                            cod_gestor=cod_gestor_input,
                            cod_un=cod_un_input,
                            departamento=departamento_sel
                        )
                        if ok:
                            st.toast(f"✅ {msg}")
                            st.rerun()
                        else:
                            st.error(msg)

        # ── Editar ──
        elif accao == "✏️  Editar utilizador":
            if not nomes_users_normais:
                st.info("Não há utilizadores normais para editar.")
            else:
                user_sel = st.selectbox("Utilizador", nomes_users_normais, key="edit_user_sel")
                u_data   = users[user_sel]

                # Fora do form — precisa de re-renderizar ao mudar a opção
                tipo_actual = u_data.get("tipo_acesso", "normal")
                tipo_acesso_sel = st.selectbox(
                    "Tipo de acesso",
                    ["normal", "gestor", "gerente"],
                    index=["normal", "gestor", "gerente"].index(tipo_actual),
                    format_func=lambda t: {
                        "normal":  "Normal — vê todos os dados dos relatórios atribuídos",
                        "gestor":  "Gestor — vê apenas os clientes do seu código de gestor",
                        "gerente": "Gerente — vê todos os clientes da sua UN",
                    }[t],
                    key=f"edit_tipo_acesso_{user_sel}"
                )

                cod_gestor_input = None
                cod_un_input     = None

                if tipo_acesso_sel == "gestor":
                    cod_gestor_input = st.text_input(
                        "Código de gestor (COD_GESTOR)",
                        value=u_data.get("cod_gestor") or "",
                        key=f"edit_cod_gestor_{user_sel}"
                    )
                    if cod_gestor_input:
                        info = validar_gestor(cod_gestor_input)
                        if info:
                            st.success(f"✅ Gestor válido — UN {info['cod_unid_neg']}")
                        else:
                            st.error("❌ Código de gestor não encontrado na tabela DWMB.lkp_gestor.")

                elif tipo_acesso_sel == "gerente":
                    cod_un_input = st.text_input(
                        "Código da UN (COD_UNID_NEG)",
                        value=u_data.get("cod_un") or "",
                        key=f"edit_cod_un_{user_sel}"
                    )
                    if cod_un_input:
                        if validar_un(cod_un_input):
                            st.success(f"✅ UN {cod_un_input} válida.")
                        else:
                            st.error("❌ Código de UN não encontrado na tabela DWMB.lkp_gestor.")

                with st.form(f"form_editar_{user_sel}"):
                    novo_nome = st.text_input(
                        "Nome completo", value=u_data.get("nome", ""),
                        key=f"edit_nome_{user_sel}"
                    )
                    activo = st.toggle(
                        "Conta activa", value=u_data.get("activo", True),
                        key=f"edit_activo_{user_sel}"
                    )
                    st.markdown("**Relatórios permitidos**")
                    relats_actuais = set(u_data.get("relatorios", []))
                    relats_sel = []
                    cols = st.columns(2)
                    for i, r in enumerate(todos_relats):
                        with cols[i % 2]:
                            if st.checkbox(r["titulo"],
                                           value=r["id"] in relats_actuais,
                                           key=f"edit_rel_{user_sel}_{r['id']}"):
                                relats_sel.append(r["id"])

                    if st.form_submit_button("Guardar alterações", use_container_width=True):
                        if tipo_acesso_sel == "gestor" and not cod_gestor_input:
                            st.error("Indique o código de gestor.")
                        elif tipo_acesso_sel == "gestor" and not validar_gestor(cod_gestor_input):
                            st.error("Código de gestor inválido. Verifique antes de guardar.")
                        elif tipo_acesso_sel == "gerente" and not cod_un_input:
                            st.error("Indique o código da UN.")
                        elif tipo_acesso_sel == "gerente" and not validar_un(cod_un_input):
                            st.error("Código de UN inválido. Verifique antes de guardar.")
                        else:
                            ok, msg = editar_utilizador(
                                user_sel, nome=novo_nome,
                                relatorios=relats_sel, activo=activo,
                                tipo_acesso=tipo_acesso_sel,
                                cod_gestor=cod_gestor_input,
                                cod_un=cod_un_input
                            )
                            if ok:
                                st.toast(f"✅ {msg}")
                                st.rerun()
                            else:
                                st.error(msg)

        # ── Password ──
        elif accao == "🔑 Alterar password":
            user_sel = st.selectbox("Utilizador", list(users.keys()), key="pwd_user_sel")
            with st.form("form_password", clear_on_submit=True):
                nova_pass  = st.text_input("Nova password",      type="password")
                nova_pass2 = st.text_input("Confirmar password", type="password")
                if st.form_submit_button("Alterar password", use_container_width=True):
                    if nova_pass != nova_pass2:
                        st.error("As passwords não coincidem.")
                    else:
                        ok, msg = alterar_password(user_sel, nova_pass)
                        if ok: st.success(msg)
                        else:  st.error(msg)

        # ── Apagar ──
        elif accao == "🗑️  Apagar utilizador":
            if not nomes_users_normais:
                st.info("Não há utilizadores normais para apagar.")
            else:
                user_sel = st.selectbox("Utilizador a apagar", nomes_users_normais)
                st.warning(f"⚠️ Tens a certeza que queres apagar **{user_sel}**? Esta acção é irreversível.")
                if st.button("🗑️ Confirmar eliminação", use_container_width=True):
                    ok, msg = apagar_utilizador(user_sel)
                    if ok:
                        st.toast(f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(msg)