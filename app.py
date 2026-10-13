import base64
import html
import io
import re
from pathlib import Path

import requests
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt

BASE_DIR = Path(__file__).resolve().parent
LOGO_CANDIDATOS = ["logo.png", "logo.jpg", "logo.jpeg", "logo.webp"]
_LOGO = next((BASE_DIR / n for n in LOGO_CANDIDATOS if (BASE_DIR / n).exists()), None)

st.set_page_config(
    page_title="PLANI PRESENTACION LEVEL",
    page_icon=str(_LOGO) if _LOGO else "🎓",
    layout="centered",
)

HEADERS = {
    "User-Agent": "PresentadorEducativo/1.0 (app educativa; contacto: tu-correo@ejemplo.com)"
}
TIMEOUT = 8

# =====================================================================
#  USUARIOS, LOGIN Y PANEL DE ADMINISTRADOR
# =====================================================================
import datetime as dt
import hashlib
import hmac
import json
import secrets as pysecrets

TITULO_APP = "PLANI PRESENTACION LEVEL"
USUARIOS_PATH = BASE_DIR / "usuarios.json"

def _hash(password, salt=None):
    salt = salt or pysecrets.token_hex(16)
    h = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt), 200_000
    )
    return salt, h.hex()


def _nuevo_registro(nombre, password, rol="usuario", vence=None, clave_inicial=False):
    salt, h = _hash(password)
    return {
        "nombre": nombre, "salt": salt, "hash": h, "rol": rol,
        "activo": True, "vence": vence,
        "creado": dt.date.today().isoformat(),
        "clave_inicial": clave_inicial,
    }


def _clave_correcta(password, reg):
    _, h = _hash(password, reg["salt"])
    return hmac.compare_digest(h, reg["hash"])


def guardar_usuarios(usuarios):
    tmp = USUARIOS_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(usuarios, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(USUARIOS_PATH)


def cargar_usuarios():
    if USUARIOS_PATH.exists():
        try:
            return json.loads(USUARIOS_PATH.read_text(encoding="utf-8"))
        except Exception:
            st.error("No se pudo leer usuarios.json. Revisa el archivo antes de continuar.")
            st.stop()
    # Primera ejecución: crear el administrador inicial
    try:
        admin_user = str(st.secrets["admin"]["usuario"]).strip().lower()
        admin_pass = str(st.secrets["admin"]["clave"])
        inicial = False
    except Exception:
        admin_user, admin_pass, inicial = "admin", "admin2026", True
    usuarios = {
        admin_user: _nuevo_registro("Administrador", admin_pass, "admin", None, inicial)
    }
    guardar_usuarios(usuarios)
    return usuarios


def autenticar(usuario, password):
    reg = cargar_usuarios().get(usuario.strip().lower())
    if not reg or not _clave_correcta(password, reg):
        return None, "Usuario o contraseña incorrectos."
    if not reg.get("activo", True):
        return None, "Tu cuenta está desactivada. Contacta al administrador."
    vence = reg.get("vence")
    if vence and dt.date.fromisoformat(vence) < dt.date.today():
        return None, "Tu acceso venció. Contacta al administrador para renovarlo."
    return reg, None


ESTILO_BASE = """
<style>
[data-testid="stHeader"] { background: transparent; }
[data-testid="stForm"] {
    background: #FFFFFF; border: 1px solid #D9E4F5; border-radius: 18px;
    padding: 1.6rem 1.8rem 1.4rem; box-shadow: 0 8px 24px rgba(10,42,94,.10);
}
.stButton > button, .stDownloadButton > button,
[data-testid="stFormSubmitButton"] > button {
    background: linear-gradient(90deg, #0A2A5E, #1A5BB5); color: #FFFFFF; border: none;
    border-radius: 10px; font-weight: 700; padding: .55rem 1.3rem;
    box-shadow: 0 4px 12px rgba(15,63,140,.30);
}
.stButton > button p, .stDownloadButton > button p,
[data-testid="stFormSubmitButton"] > button p { color: #FFFFFF; }
.stButton > button:hover, .stDownloadButton > button:hover,
[data-testid="stFormSubmitButton"] > button:hover { filter: brightness(1.18); color: #FFFFFF; }
.stButton > button:disabled { opacity: .45; }
.plani-logo-ph {
    width: 110px; height: 110px; margin: 0 auto; border-radius: 50%;
    background: #0F3F8C; color: #D4AF37; font-weight: 800; font-size: 1.4rem;
    display: flex; align-items: center; justify-content: center;
}
</style>
"""

ESTILO_LOGIN = """
<style>
.stApp, [data-testid="stAppViewContainer"] {
    background:
        radial-gradient(circle at 12% 8%, rgba(26,91,181,.75) 0%, rgba(26,91,181,0) 42%),
        radial-gradient(circle at 90% 95%, rgba(212,175,55,.18) 0%, rgba(212,175,55,0) 35%),
        linear-gradient(135deg, #071F47 0%, #0A2A5E 45%, #0F3F8C 100%);
}
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] { display: none; }
.block-container { max-width: 520px !important; padding-top: 3rem; }
.plani-title {
    color: #FFFFFF; text-align: center; font-weight: 800; letter-spacing: .08em;
    font-size: clamp(1.3rem, 5vw, 1.8rem); margin: .5rem 0 .2rem;
}
.plani-title::after {
    content: ""; display: block; width: 72px; height: 3px; background: #D4AF37;
    margin: .7rem auto 0; border-radius: 2px;
}
.plani-tag { color: #BFD3F2; text-align: center; margin: .6rem 0 1.5rem; font-size: .95rem; }
.plani-badge {
    background: linear-gradient(135deg, #071F47, #0F3F8C); border-radius: 16px;
    padding: 18px; text-align: center; margin-bottom: 14px;
    border-bottom: 3px solid #D4AF37;
}
.plani-badge img { height: 130px; width: auto; }
.plani-sub { text-align: center; color: #0A2A5E; font-weight: 600; margin: .2rem 0 1rem; }
[data-testid="stForm"] { border: none; box-shadow: 0 24px 60px rgba(0,0,0,.40); }
[data-testid="stForm"] label p { color: #0A2A5E; font-weight: 600; }
</style>
"""

ESTILO_APP = """
<style>
.stApp { background: #EEF3FA; }
.block-container { max-width: 1100px !important; padding-top: 1.5rem; }
h1, h2, h3, h4 { color: #0A2A5E; }

/* Barra lateral */
[data-testid="stSidebar"] { background: linear-gradient(180deg, #071F47 0%, #0F3F8C 100%); }
[data-testid="stSidebar"] * { color: #FFFFFF; }
[data-testid="stSidebar"] input { color: #0A2A5E !important; -webkit-text-fill-color: #0A2A5E; }
[data-testid="stSidebar"] [data-testid="stAlert"] * { color: #0A2A5E !important; }
[data-testid="stSidebar"] [data-testid="stForm"] {
    background: transparent; border: none; box-shadow: none; padding: 0;
}
[data-testid="stSidebar"] .stButton > button,
[data-testid="stSidebar"] [data-testid="stFormSubmitButton"] > button {
    background: rgba(255,255,255,.12); border: 1px solid rgba(255,255,255,.55);
    box-shadow: none; width: 100%;
}
[data-testid="stSidebar"] [data-testid="stExpander"] {
    border: 1px solid rgba(255,255,255,.25); border-radius: 12px;
}
.plani-user {
    display: flex; align-items: center; gap: 12px; padding: 12px 14px; margin-bottom: 12px;
    background: rgba(255,255,255,.10); border-radius: 14px; border-left: 4px solid #D4AF37;
}
.plani-avatar {
    width: 42px; height: 42px; border-radius: 50%; background: #D4AF37; color: #0A2A5E !important;
    font-weight: 800; font-size: 1.2rem; display: flex; align-items: center; justify-content: center;
    flex: 0 0 auto;
}
.plani-user b { display: block; line-height: 1.2; }
.plani-user span { font-size: .8rem; color: #BFD3F2; }

/* Banner superior */
.plani-banner {
    display: flex; align-items: center; gap: 22px; margin-bottom: 1.1rem;
    padding: 14px 26px; border-radius: 20px; border-bottom: 4px solid #D4AF37;
    background: linear-gradient(120deg, #071F47 0%, #0A2A5E 40%, #1A5BB5 100%);
    box-shadow: 0 10px 28px rgba(10,42,94,.28);
}
.plani-banner img { height: 78px; width: auto; flex: 0 0 auto; }
.plani-banner .t {
    color: #FFFFFF; font-weight: 800; letter-spacing: .07em; line-height: 1.2;
    font-size: clamp(1.1rem, 3.2vw, 1.7rem);
}
.plani-banner .s { color: #BFD3F2; margin-top: .25rem; font-size: .92rem; }

/* Pestañas */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px; background: #FFFFFF; padding: 6px; border-radius: 14px;
    box-shadow: 0 2px 10px rgba(10,42,94,.08);
}
.stTabs [data-baseweb="tab"] { border-radius: 10px; padding: 8px 18px; height: auto; }
.stTabs [data-baseweb="tab"] p { color: #0A2A5E; font-weight: 600; }
.stTabs [aria-selected="true"] { background: #0F3F8C; }
.stTabs [aria-selected="true"] p { color: #FFFFFF; }
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] { display: none; }

/* Tarjetas de estadísticas */
.plani-stats {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
    gap: 14px; margin: .6rem 0 1.2rem;
}
.plani-stat {
    background: #FFFFFF; border-radius: 16px; padding: 16px 18px;
    border-left: 5px solid #0F3F8C; box-shadow: 0 4px 14px rgba(10,42,94,.09);
}
.plani-stat b { display: block; font-size: 2rem; line-height: 1.1; color: #0A2A5E; }
.plani-stat span { color: #5A6F94; font-size: .78rem; text-transform: uppercase; letter-spacing: .07em; }
.plani-stat.ok { border-left-color: #1E9E6A; }
.plani-stat.warn { border-left-color: #E0A100; }
.plani-stat.bad { border-left-color: #D64545; }

/* Encabezado del generador */
.plani-hero {
    background: #FFFFFF; border-radius: 16px; padding: 1rem 1.4rem; margin-bottom: 1rem;
    border-left: 6px solid #D4AF37; box-shadow: 0 4px 14px rgba(10,42,94,.09);
}
.plani-hero b { color: #0A2A5E; font-size: 1.25rem; }
.plani-hero p { margin: .3rem 0 0; color: #4A5F85; }
</style>
"""


def _logo_b64():
    for nombre in LOGO_CANDIDATOS:
        ruta = BASE_DIR / nombre
        if ruta.exists():
            try:
                from PIL import Image
                im = Image.open(ruta).convert("RGBA")
                im.thumbnail((400, 400))
                buf = io.BytesIO()
                im.save(buf, "PNG")
                return base64.b64encode(buf.getvalue()).decode()
            except Exception:
                return None
    return None


_logo_b64 = st.cache_resource(show_spinner=False)(_logo_b64)


def _logo_html():
    b64 = _logo_b64()
    if b64:
        return f'<img src="data:image/png;base64,{b64}" alt="PLANI">'
    return '<div class="plani-logo-ph">PLANI</div>'


def pantalla_login():
    st.markdown(ESTILO_BASE + ESTILO_LOGIN, unsafe_allow_html=True)
    st.markdown(f'<div class="plani-title">{TITULO_APP}</div>', unsafe_allow_html=True)
    st.markdown('<div class="plani-tag">Generador de presentaciones educativas</div>',
                unsafe_allow_html=True)
    with st.form("login_form"):
        st.markdown(
            f'<div class="plani-badge">{_logo_html()}</div>'
            '<div class="plani-sub">Inicia sesión para continuar</div>',
            unsafe_allow_html=True,
        )
        user = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Ingresar", use_container_width=True)
    if entrar:
        reg, error = autenticar(user, password)
        if error:
            st.error(error)
        else:
            st.session_state.update(
                autenticado=True,
                usuario=user.strip().lower(),
                rol=reg["rol"],
                nombre=reg.get("nombre") or user.strip(),
            )
            st.rerun()


def banner():
    st.markdown(
        f'<div class="plani-banner">{_logo_html()}<div>'
        f'<div class="t">{TITULO_APP}</div>'
        '<div class="s">Generador de presentaciones educativas con información real</div>'
        '</div></div>',
        unsafe_allow_html=True,
    )


def sidebar_usuario():
    nombre = str(st.session_state.get("nombre") or st.session_state.usuario)
    rol = "Administrador" if st.session_state.get("rol") == "admin" else "Usuario"
    st.sidebar.markdown(
        f'<div class="plani-user"><div class="plani-avatar">{html.escape(nombre[:1].upper())}</div>'
        f'<div><b>{html.escape(nombre)}</b><span>{rol}</span></div></div>',
        unsafe_allow_html=True,
    )


def estado_usuario(reg):
    if not reg.get("activo", True):
        return "Desactivado"
    v = reg.get("vence")
    if v:
        dias = (dt.date.fromisoformat(v) - dt.date.today()).days
        if dias < 0:
            return "Vencido"
        if dias <= 7:
            return "Por vencer"
    return "Activo"


def _es_ultimo_admin(usuarios, usuario):
    admins = [u for u, r in usuarios.items() if r["rol"] == "admin" and r.get("activo", True)]
    return usuarios[usuario]["rol"] == "admin" and admins == [usuario]


def cambiar_mi_clave():
    with st.sidebar.expander("🔑 Cambiar mi contraseña"):
        with st.form("form_mi_clave", clear_on_submit=True):
            actual = st.text_input("Contraseña actual", type="password")
            nueva = st.text_input("Nueva contraseña (mín. 6)", type="password")
            ok = st.form_submit_button("Actualizar")
        if ok:
            usuarios = cargar_usuarios()
            yo = usuarios.get(st.session_state.usuario)
            if not yo or not _clave_correcta(actual, yo):
                st.error("La contraseña actual no es correcta.")
            elif len(nueva) < 6:
                st.error("La nueva contraseña debe tener al menos 6 caracteres.")
            else:
                yo["salt"], yo["hash"] = _hash(nueva)
                yo["clave_inicial"] = False
                guardar_usuarios(usuarios)
                st.success("Contraseña actualizada.")


def panel_admin():
    usuarios = cargar_usuarios()
    yo = st.session_state.usuario
    if "flash" in st.session_state:
        st.success(st.session_state.pop("flash"))
    if usuarios.get(yo, {}).get("clave_inicial"):
        st.warning(
            "Estás usando la contraseña inicial del administrador. "
            "Cámbiala en el menú lateral (🔑 Cambiar mi contraseña)."
        )

    # ---- Resumen ----
    estados = [estado_usuario(r) for r in usuarios.values()]
    st.markdown(
        '<div class="plani-stats">'
        f'<div class="plani-stat"><b>{len(estados)}</b><span>Usuarios totales</span></div>'
        f'<div class="plani-stat ok"><b>{estados.count("Activo") + estados.count("Por vencer")}</b><span>Con acceso</span></div>'
        f'<div class="plani-stat warn"><b>{estados.count("Por vencer")}</b><span>Vencen en 7 días</span></div>'
        f'<div class="plani-stat bad"><b>{estados.count("Vencido") + estados.count("Desactivado")}</b><span>Vencidos o desactivados</span></div>'
        '</div>',
        unsafe_allow_html=True,
    )

    # ---- Crear usuario ----
    st.subheader("➕ Crear usuario")
    with st.form("crear_usuario", clear_on_submit=True):
        c1, c2 = st.columns(2)
        nombre = c1.text_input("Nombre completo")
        nuevo = c2.text_input("Usuario (letras, números, . _ -)")
        c3, c4 = st.columns(2)
        clave = c3.text_input("Contraseña (mín. 6 caracteres)", type="password")
        rol = c4.selectbox("Rol", ["usuario", "admin"])
        sin_venc = st.checkbox("Sin fecha de vencimiento")
        vence = st.date_input(
            "Acceso válido hasta", value=dt.date.today() + dt.timedelta(days=30)
        )
        crear = st.form_submit_button("Crear usuario")
    if crear:
        nuevo = nuevo.strip().lower()
        if not re.fullmatch(r"[a-z0-9_.-]{3,30}", nuevo):
            st.error("Usuario inválido: usa de 3 a 30 letras minúsculas, números, punto, guion o guion bajo.")
        elif nuevo in usuarios:
            st.error("Ese usuario ya existe.")
        elif len(clave) < 6:
            st.error("La contraseña debe tener al menos 6 caracteres.")
        else:
            usuarios[nuevo] = _nuevo_registro(
                nombre.strip() or nuevo, clave, rol,
                None if sin_venc else vence.isoformat(),
            )
            guardar_usuarios(usuarios)
            st.session_state.flash = f"Usuario «{nuevo}» creado."
            st.rerun()

    # ---- Lista ----
    st.subheader("👥 Usuarios")
    st.dataframe(
        [
            {
                "Usuario": u,
                "Nombre": r.get("nombre", ""),
                "Rol": r["rol"],
                "Estado": estado_usuario(r),
                "Vence": r.get("vence") or "Sin vencimiento",
                "Creado": r.get("creado", ""),
            }
            for u, r in usuarios.items()
        ],
        use_container_width=True, hide_index=True,
    )

    # ---- Gestionar ----
    st.subheader("⚙️ Gestionar usuario")
    sel = st.selectbox("Selecciona un usuario", list(usuarios))
    reg = usuarios[sel]
    a, b = st.columns(2)

    with a:
        with st.form(f"form_estado_{sel}"):
            activo = st.checkbox("Cuenta activa", value=reg.get("activo", True))
            sin = st.checkbox("Sin vencimiento", value=not reg.get("vence"))
            fecha = st.date_input(
                "Válido hasta",
                value=dt.date.fromisoformat(reg["vence"]) if reg.get("vence")
                else dt.date.today() + dt.timedelta(days=30),
            )
            guardar = st.form_submit_button("Guardar cambios")
        if guardar:
            if not activo and sel == yo:
                st.error("No puedes desactivar tu propia cuenta.")
            elif not activo and _es_ultimo_admin(usuarios, sel):
                st.error("Debe quedar al menos un administrador activo.")
            else:
                reg["activo"] = activo
                reg["vence"] = None if sin else fecha.isoformat()
                guardar_usuarios(usuarios)
                st.session_state.flash = f"Cambios guardados para «{sel}»."
                st.rerun()

    with b:
        with st.form(f"form_clave_{sel}", clear_on_submit=True):
            nueva = st.text_input("Nueva contraseña (mín. 6)", type="password")
            restablecer = st.form_submit_button("Restablecer contraseña")
        if restablecer:
            if len(nueva) < 6:
                st.error("La contraseña debe tener al menos 6 caracteres.")
            else:
                reg["salt"], reg["hash"] = _hash(nueva)
                reg["clave_inicial"] = False
                guardar_usuarios(usuarios)
                st.session_state.flash = f"Contraseña de «{sel}» restablecida."
                st.rerun()

    confirmar = st.checkbox(f"Confirmo que quiero eliminar a «{sel}»", key=f"conf_{sel}")
    if st.button("🗑️ Eliminar usuario", disabled=not confirmar):
        if sel == yo:
            st.error("No puedes eliminar tu propia cuenta.")
        elif _es_ultimo_admin(usuarios, sel):
            st.error("Debe quedar al menos un administrador activo.")
        else:
            del usuarios[sel]
            guardar_usuarios(usuarios)
            st.session_state.flash = f"Usuario «{sel}» eliminado."
            st.rerun()


# =====================================================================
#  BÚSQUEDA DE INFORMACIÓN REAL
# =====================================================================
SECCIONES_IGNORADAS = {
    "referencias", "véase también", "enlaces externos", "notas",
    "bibliografía", "further reading", "references", "see also",
    "external links", "notes", "bibliography", "citas", "galería",
}


def _get(url, **params):
    r = requests.get(url, params=params, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


@st.cache_data(ttl=3600, show_spinner=False)
def wiki_buscar(consulta, lang, limite=5):
    try:
        data = _get(
            f"https://{lang}.wikipedia.org/w/api.php",
            action="query", list="search", srsearch=consulta,
            srlimit=limite, format="json",
        )
        return [x["title"] for x in data["query"]["search"]]
    except Exception:
        return []


@st.cache_data(ttl=3600, show_spinner=False)
def wiki_pagina(titulo, lang):
    """Devuelve texto completo con secciones, url e imagen principal."""
    try:
        data = _get(
            f"https://{lang}.wikipedia.org/w/api.php",
            action="query", prop="extracts|pageimages|info",
            titles=titulo, explaintext=1, exsectionformat="wiki",
            piprop="thumbnail", pithumbsize=900, inprop="url",
            redirects=1, format="json",
        )
        page = next(iter(data["query"]["pages"].values()))
        if "extract" not in page:
            return None
        return {
            "titulo": page["title"],
            "texto": page["extract"],
            "url": page.get("fullurl", ""),
            "imagen": page.get("thumbnail", {}).get("source"),
            "lang": lang,
        }
    except Exception:
        return None


@st.cache_data(ttl=3600, show_spinner=False)
def commons_imagenes(consulta, limite=6):
    """Imágenes libres de Wikimedia Commons."""
    try:
        data = _get(
            "https://commons.wikimedia.org/w/api.php",
            action="query", generator="search", gsrnamespace=6,
            gsrsearch=consulta, gsrlimit=limite, prop="imageinfo",
            iiprop="url|mime|extmetadata", iiurlwidth=900, format="json",
        )
        out = []
        for p in data.get("query", {}).get("pages", {}).values():
            info = (p.get("imageinfo") or [{}])[0]
            if info.get("mime") not in ("image/jpeg", "image/png"):
                continue
            url = info.get("thumburl") or info.get("url")
            if not url:
                continue
            autor = info.get("extmetadata", {}).get("Artist", {}).get("value", "")
            autor = re.sub(r"<[^>]+>", "", autor).strip()[:60]
            out.append({
                "url": url,
                "pagina": info.get("descriptionurl", ""),
                "autor": autor or "Wikimedia Commons",
            })
        return out
    except Exception:
        return []


@st.cache_data(ttl=3600, show_spinner=False)
def openalex_articulos(consulta, limite=3):
    try:
        data = _get(
            "https://api.openalex.org/works",
            search=consulta, per_page=limite,
            select="title,publication_year,doi,primary_location",
        )
        return [
            {"nombre": f"{w['title']} ({w.get('publication_year', 's.f.')})",
             "url": w.get("doi") or ""}
            for w in data.get("results", []) if w.get("title")
        ]
    except Exception:
        return []


@st.cache_data(ttl=3600, show_spinner=False)
def openlibrary_libros(consulta, limite=3):
    try:
        data = _get(
            "https://openlibrary.org/search.json",
            q=consulta, limit=limite, fields="title,author_name,first_publish_year,key",
        )
        out = []
        for d in data.get("docs", []):
            autor = ", ".join(d.get("author_name", [])[:2]) or "Autor desconocido"
            out.append({
                "nombre": f"{d['title']} — {autor} ({d.get('first_publish_year', 's.f.')})",
                "url": f"https://openlibrary.org{d['key']}",
            })
        return out
    except Exception:
        return []


def parsear_secciones(texto):
    secciones, titulo, buf = [], "Introducción", []
    for linea in texto.splitlines():
        linea = linea.strip()
        m = re.match(r"^(={2,})\s*(.+?)\s*\1$", linea)
        if m:
            if len(m.group(1)) == 2:
                if buf:
                    secciones.append((titulo, " ".join(buf)))
                titulo, buf = m.group(2), []
            continue
        if linea:
            buf.append(linea)
    if buf:
        secciones.append((titulo, " ".join(buf)))
    return [(t, x) for t, x in secciones if t.lower() not in SECCIONES_IGNORADAS]


def puntos_clave(texto, maximo=4):
    texto = re.sub(r"\[\d+\]|\[[a-z]\]", "", texto)
    frases = re.split(r"(?<=[.!?])\s+(?=[A-ZÁÉÍÓÚÑ¿¡0-9])", texto)
    buenas = [f.strip() for f in frases if 40 <= len(f.strip()) <= 260]
    return buenas[:maximo]


def recopilar(tema, n_slides):
    """Reúne contenido real de varias fuentes. Devuelve (slides, fuentes)."""
    fuentes, slides, usados = [], [], set()

    pagina = None
    for lang in ("es", "en"):
        for titulo in wiki_buscar(tema, lang):
            pagina = wiki_pagina(titulo, lang)
            if pagina and len(pagina["texto"]) > 400:
                break
            pagina = None
        if pagina:
            break
    if not pagina:
        return [], [], None

    fuentes.append({
        "nombre": f"Wikipedia ({pagina['lang']}): {pagina['titulo']}",
        "url": pagina["url"],
    })

    # Diapositivas desde las secciones reales del artículo
    for titulo, texto in parsear_secciones(pagina["texto"]):
        pts = puntos_clave(texto)
        if pts:
            slides.append({"titulo": titulo, "puntos": pts, "imagen": None})
        if len(slides) >= n_slides:
            break

    # Si faltan diapositivas, completar con artículos relacionados
    if len(slides) < n_slides:
        for lang in ("es", "en"):
            for titulo in wiki_buscar(tema, lang, limite=8):
                if len(slides) >= n_slides:
                    break
                if titulo == pagina["titulo"] or titulo in usados:
                    continue
                usados.add(titulo)
                rel = wiki_pagina(titulo, lang)
                if not rel:
                    continue
                pts = puntos_clave(rel["texto"].split("\n==")[0], 3)
                if pts:
                    slides.append({"titulo": rel["titulo"], "puntos": pts,
                                   "imagen": rel["imagen"] and {
                                       "url": rel["imagen"], "pagina": rel["url"],
                                       "autor": "Wikipedia"}})
                    fuentes.append({"nombre": f"Wikipedia ({lang}): {rel['titulo']}",
                                    "url": rel["url"]})

    # Imágenes: portada de Wikipedia + búsquedas en Commons por diapositiva
    vistas = set()
    if pagina["imagen"]:
        vistas.add(pagina["imagen"])
    portada = (
        {"url": pagina["imagen"], "pagina": pagina["url"], "autor": "Wikipedia"}
        if pagina["imagen"] else None
    )
    pool = [i for i in commons_imagenes(tema, 8)]
    for s in slides:
        if s["imagen"]:
            continue
        candidatas = commons_imagenes(f"{tema} {s['titulo']}", 4) + pool
        for c in candidatas:
            if c["url"] not in vistas:
                s["imagen"] = c
                vistas.add(c["url"])
                if c["pagina"]:
                    fuentes.append({"nombre": f"Imagen: {c['autor']} (Wikimedia Commons)",
                                    "url": c["pagina"]})
                break
    if portada and slides and not slides[0]["imagen"]:
        slides[0]["imagen"] = portada

    # Fuentes adicionales: artículos académicos y libros
    fuentes += openalex_articulos(tema)
    fuentes += openlibrary_libros(tema)
    return slides, fuentes, portada


# =====================================================================
#  CONSTRUCCIÓN DEL PPTX
# =====================================================================
AZUL = RGBColor(0x1F, 0x3A, 0x5F)
GRIS = RGBColor(0x33, 0x33, 0x33)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)


def descargar_imagen(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code == 200 and len(r.content) > 2000:
            return io.BytesIO(r.content)
    except Exception:
        pass
    return None


def barra_titulo(prs, slide, texto):
    from pptx.enum.shapes import MSO_SHAPE
    barra = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.2))
    barra.fill.solid()
    barra.fill.fore_color.rgb = AZUL
    barra.line.fill.background()
    tf = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(12.3), Inches(0.9)).text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = texto
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = BLANCO


def poner_imagen(slide, img, x, y, max_w, max_h):
    from PIL import Image
    flujo = descargar_imagen(img["url"])
    if not flujo:
        return False
    try:
        w, h = Image.open(flujo).size
        flujo.seek(0)
        escala = min(max_w / w, max_h / h)
        ancho, alto = int(w * escala), int(h * escala)
        slide.shapes.add_picture(flujo, x + (max_w - ancho) // 2, y, width=ancho, height=alto)
        cap = slide.shapes.add_textbox(x, y + alto + Inches(0.05), max_w, Inches(0.4)).text_frame
        cap.word_wrap = True
        cap.paragraphs[0].text = f"Imagen: {img['autor']}"
        cap.paragraphs[0].font.size = Pt(10)
        cap.paragraphs[0].font.color.rgb = RGBColor(0x77, 0x77, 0x77)
        return True
    except Exception:
        return False


def construir_pptx(tema, slides, fuentes, portada, incluir_yt):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    vacio = prs.slide_layouts[6]

    # Portada
    s = prs.slides.add_slide(vacio)
    fondo = s.background.fill
    fondo.solid()
    fondo.fore_color.rgb = AZUL
    tf = s.shapes.add_textbox(Inches(0.7), Inches(2.3), Inches(7), Inches(2.5)).text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = tema
    tf.paragraphs[0].font.size = Pt(44)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = BLANCO
    if portada:
        poner_imagen(s, portada, Inches(8), Inches(1.5), Inches(4.8), Inches(4.5))

    # Contenido
    for d in slides:
        s = prs.slides.add_slide(vacio)
        barra_titulo(prs, s, d["titulo"])
        con_img = bool(d["imagen"])
        ancho_txt = Inches(7.4) if con_img else Inches(12.3)
        tf = s.shapes.add_textbox(Inches(0.5), Inches(1.5), ancho_txt, Inches(5.5)).text_frame
        tf.word_wrap = True
        for i, punto in enumerate(d["puntos"]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + punto
            p.font.size = Pt(18)
            p.font.color.rgb = GRIS
            p.space_after = Pt(12)
        if con_img:
            if not poner_imagen(s, d["imagen"], Inches(8.2), Inches(1.6), Inches(4.7), Inches(4.8)):
                pass

    # Fuentes
    s = prs.slides.add_slide(vacio)
    barra_titulo(prs, s, "Fuentes y referencias")
    tf = s.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.3), Inches(5.6)).text_frame
    tf.word_wrap = True
    vistos, i = set(), 0
    for f in fuentes:
        clave = (f["nombre"], f["url"])
        if clave in vistos or i >= 12:
            continue
        vistos.add(clave)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"{f['nombre']}" + (f" — {f['url']}" if f["url"] else "")
        p.font.size = Pt(12)
        p.font.color.rgb = GRIS
        i += 1

    if incluir_yt:
        s = prs.slides.add_slide(vacio)
        barra_titulo(prs, s, "Videos recomendados en YouTube")
        q = tema.replace(" ", "+")
        tf = s.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.3), Inches(4)).text_frame
        tf.word_wrap = True
        tf.paragraphs[0].text = f"• Explicación del tema: https://www.youtube.com/results?search_query={q}"
        p = tf.add_paragraph()
        p.text = f"• Clases y tutoriales: https://www.youtube.com/results?search_query={q}+clase+educativa"
        for p in tf.paragraphs:
            p.font.size = Pt(16)
            p.font.color.rgb = GRIS
            p.space_after = Pt(12)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# =====================================================================
#  INTERFAZ
# =====================================================================
def pantalla_generador():
    st.markdown(
        '<div class="plani-hero"><b>🎓 Crea tu presentación</b>'
        "<p>Escribe un tema y la app reunirá información verificable de varias fuentes, "
        "con imágenes relacionadas y una diapositiva de referencias.</p></div>",
        unsafe_allow_html=True,
    )

    with st.form("form_presentacion"):
        tema = st.text_input("Tema, asignatura o concepto", "La fotosíntesis")
        c1, c2 = st.columns(2)
        with c1:
            num_slides = st.slider("Diapositivas de contenido", 3, 10, 6)
        with c2:
            incluir_yt = st.checkbox("Incluir enlaces de YouTube", value=True)
        generar = st.form_submit_button("🚀 Generar presentación")

    if generar:
        if not tema.strip():
            st.warning("Escribe un tema válido.")
            return
        try:
            with st.spinner("Reuniendo información e imágenes..."):
                resultado = recopilar(tema.strip(), num_slides)
            if not resultado[0]:
                st.error("No encontré información suficiente sobre ese tema. Prueba con otra redacción.")
            else:
                slides, fuentes, portada = resultado
                with st.spinner("Armando el PowerPoint..."):
                    archivo = construir_pptx(tema.strip(), slides, fuentes, portada, incluir_yt)
                st.success(f"Presentación lista: {len(slides)} diapositivas de contenido.")
                st.download_button(
                    "📥 Descargar (.pptx)",
                    data=archivo,
                    file_name=f"{tema.strip().replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
        except Exception as err:
            st.error(f"Ocurrió un error: {err}")


# =====================================================================
#  ENRUTADOR
# =====================================================================
if not st.session_state.get("autenticado"):
    pantalla_login()
    st.stop()

st.markdown(ESTILO_BASE + ESTILO_APP, unsafe_allow_html=True)
banner()
sidebar_usuario()
if st.sidebar.button("Cerrar sesión"):
    for k in ("autenticado", "usuario", "rol", "nombre"):
        st.session_state.pop(k, None)
    st.rerun()
cambiar_mi_clave()

if st.session_state.get("rol") == "admin":
    tab_gen, tab_admin = st.tabs(["🎓 Generar presentación", "👥 Administrar usuarios"])
    with tab_gen:
        pantalla_generador()
    with tab_admin:
        panel_admin()
else:
    pantalla_generador()
