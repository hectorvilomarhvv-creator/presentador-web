"""Presentador: busca un tema y crea una presentación de PowerPoint (.pptx).
Con sistema de Login y Administración de Usuarios.
"""
import hmac
import html
import io
import json
import os
import platform
import re
import tempfile
import time
from collections import defaultdict

import requests
from flask import Flask, jsonify, render_template, request, send_file, redirect, url_for, flash
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10_000

# --- Configuración ---
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'clave_secreta_muy_segura_123')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///usuarios.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-flash-latest")
CONTACTO = os.environ.get("CONTACT", "").strip()
LIMITE_DIARIO = int(os.environ.get("DAILY_LIMIT", "20"))

HEADERS = {"User-Agent": "Presentador/1.0 (herramienta educativa" + (f"; {CONTACTO}" if CONTACTO else "") + ")"}

# --- Base de datos y Login ---
db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(120), nullable=False)
    is_admin = db.Column(db.Boolean, default=False)
    uso_diario = db.Column(db.Integer, default=0)
    fecha_ultimo_uso = db.Column(db.String(10), default="")

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def registrar_uso(usuario):
    """Cuenta una presentación para el usuario logueado."""
    hoy = time.strftime("%Y-%m-%d")
    if usuario.fecha_ultimo_uso != hoy:
        usuario.uso_diario = 0
        usuario.fecha_ultimo_uso = hoy
    
    if usuario.uso_diario >= LIMITE_DIARIO:
        return False
    
    usuario.uso_diario += 1
    db.session.commit()
    return True

TEMAS = {
    "azul": ("1F3A5F", "2E86DE", "F4F7FB"),
    "verde": ("1B4332", "2D9C6B", "F2F8F5"),
    "naranja": ("7A3E00", "E67E22", "FDF6EE"),
    "morado": ("3B2A5E", "7E57C2", "F6F3FB"),
}

# ----------------------------------------------------------------- búsqueda
def buscar_wikipedia(tema, idioma="es"):
    base = f"https://{idioma}.wikipedia.org/w/api.php"
    r = requests.get(base, params={"action": "query", "list": "search", "srsearch": tema, "srlimit": 1, "format": "json"}, headers=HEADERS, timeout=15)
    r.raise_for_status()
    resultados = r.json().get("query", {}).get("search", [])
    if not resultados:
        return None
    titulo = resultados[0]["title"]
    r = requests.get(base, params={"action": "query", "prop": "extracts", "explaintext": 1, "exsectionformat": "plain", "titles": titulo, "format": "json", "redirects": 1}, headers=HEADERS, timeout=15)
    r.raise_for_status()
    paginas = r.json().get("query", {}).get("pages", {})
    texto = next(iter(paginas.values())).get("extract", "")
    url = f"https://{idioma}.wikipedia.org/wiki/{titulo.replace(' ', '_')}"
    return titulo, texto, url

def _limpiar_html(texto):
    return re.sub(r"<[^>]+>", "", html.unescape(texto or "")).strip()

def buscar_imagen(consulta, usadas):
    try:
        r = requests.get("https://commons.wikimedia.org/w/api.php", params={"action": "query", "generator": "search", "gsrsearch": consulta, "gsrnamespace": 6, "gsrlimit": 10, "prop": "imageinfo", "iiprop": "url|mime|size|extmetadata", "iiurlwidth": 900, "format": "json"}, headers=HEADERS, timeout=10)
        r.raise_for_status()
        paginas = r.json().get("query", {}).get("pages", {})
        for p in sorted(paginas.values(), key=lambda p: p.get("index", 0)):
            info = (p.get("imageinfo") or [{}])[0]
            titulo = p.get("title", "")
            if info.get("mime") not in ("image/jpeg", "image/png"):
                continue
            if info.get("width", 0) < 500 or titulo in usadas:
                continue
            if re.search(r"logo|icon|flag|signature", titulo, re.I):
                continue
            img = requests.get(info.get("thumburl") or info["url"], headers=HEADERS, timeout=15)
            img.raise_for_status()
            if len(img.content) > 4_000_000:
                continue
            meta = info.get("extmetadata", {})
            autor = _limpiar_html(meta.get("Artist", {}).get("value"))[:60]
            licencia = _limpiar_html(meta.get("LicenseShortName", {}).get("value"))
            usadas.add(titulo)
            credito = (f'{titulo.replace("File:", "")} — {autor or "autor no indicado"} ({licencia or "ver Wikimedia Commons"})')
            return {"datos": img.content, "credito": credito}
    except Exception:
        return None
    return None

# ------------------------------------------------------- contenido
def oraciones(texto):
    partes = re.split(r"(?<=[.!?])\s+", texto.replace("\n", " "))
    return [p.strip() for p in partes if 40 <= len(p.strip()) <= 220]

def diapositivas_desde_texto(titulo, texto, cantidad):
    texto = re.split(r"\n(?:Véase también|Referencias|Enlaces externos|See also|References)\b", texto)[0]
    oraciones_ok = oraciones(texto)
    if not oraciones_ok:
        return []
    por_slide = 4
    total = min(cantidad, max(1, len(oraciones_ok) // por_slide))
    paso = max(1, len(oraciones_ok) // total)
    slides = []
    for i in range(total):
        trozo = oraciones_ok[i * paso:(i + 1) * paso]
        if trozo:
            puntos = trozo[:por_slide]
            resto = trozo[por_slide:por_slide + 3]
            notas = " ".join(resto) if resto else " ".join(puntos)
            slides.append({"titulo": f"{titulo}: parte {i + 1}", "puntos": puntos, "notas": notas, "imagen": titulo})
    return slides

def diapositivas_con_gemini(tema, texto, cantidad, idioma):
    nombre_idioma = "español" if idioma == "es" else "inglés"
    prompt = (f"Crea el contenido de una presentación en {nombre_idioma} sobre: {tema}.\n"
              f"Usa exactamente {cantidad} diapositivas de contenido (sin portada).\n"
              "Cada diapositiva tiene: un título corto, de 3 a 5 puntos breves (máx. 18 palabras cada uno), "
              "'notas' con 2 a 4 oraciones que el presentador puede decir en voz alta, y "
              "'imagen' con 2 o 3 palabras clave para buscar una foto adecuada.\n"
              "Basa los datos en el texto de referencia; no inventes cifras ni fechas.\n"
              'Responde SOLO JSON con este formato: {"diapositivas":[{"titulo":"...","puntos":["..."],"notas":"...","imagen":"..."}]}\n\n'
              f"Texto de referencia:\n{texto[:12000]}")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    r = requests.post(url, headers={"x-goog-api-key": GEMINI_API_KEY, "Content-Type": "application/json"},
                      json={"contents": [{"parts": [{"text": prompt}]}], "generationConfig": {"responseMimeType": "application/json"}}, timeout=60)
    r.raise_for_status()
    crudo = r.json()["candidates"][0]["content"]["parts"][0]["text"]
    datos = json.loads(crudo)
    return [{"titulo": str(s["titulo"]), "puntos": [str(p) for p in s["puntos"]][:6], "notas": str(s.get("notas", "")), "imagen": str(s.get("imagen", "") or tema)} for s in datos["diapositivas"] if s.get("titulo") and s.get("puntos")]

# -------------------------------------------------------------- archivo .pptx
def _rgb(hex_):
    return RGBColor.from_string(hex_)

def _rect(slide, x, y, w, h, color):
    forma = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    forma.fill.solid()
    forma.fill.fore_color.rgb = _rgb(color)
    forma.line.fill.background()
    return forma

def _texto(slide, x, y, w, h, texto, size, color, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    caja = slide.shapes.add_textbox(x, y, w, h)
    tf = caja.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = texto
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = "Calibri"
    run.font.color.rgb = _rgb(color)
    return caja

def _poner_imagen(slide, datos, x, y, w_max, h_max):
    iw, ih = Image.open(io.BytesIO(datos)).size
    escala = min(w_max / iw, h_max / ih)
    w, h = int(iw * escala), int(ih * escala)
    slide.shapes.add_picture(io.BytesIO(datos), int(x + (w_max - w) / 2), int(y + (h_max - h) / 2), w, h)

def _notas(slide, texto):
    if texto:
        slide.notes_slide.notes_text_frame.text = texto

def crear_pptx(titulo, slides, fuente_url, tema_color, portada=None, con_notas=True):
    oscuro, acento, claro = TEMAS.get(tema_color, TEMAS["azul"])
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    W, H = prs.slide_width, prs.slide_height
    vacio = prs.slide_layouts[6]
    creditos = []

    s = prs.slides.add_slide(vacio)
    _rect(s, 0, 0, W, H, oscuro)
    ancho_titulo = Inches(7.2) if portada else Inches(11.5)
    _rect(s, Inches(0.8), Inches(3.55), Inches(1.6), Emu(60000), acento)
    _texto(s, Inches(0.8), Inches(1.8), ancho_titulo, Inches(1.7), titulo, 44 if portada else 48, "FFFFFF", bold=True, anchor=MSO_ANCHOR.BOTTOM)
    _texto(s, Inches(0.8), Inches(3.8), Inches(6), Inches(0.8), time.strftime("%d/%m/%Y"), 18, "C9D6E8")
    if portada:
        _poner_imagen(s, portada["datos"], Inches(8.4), Inches(1.2), Inches(4.4), Inches(5.1))
        creditos.append(portada["credito"])
    if con_notas:
        _notas(s, f"Presentación sobre {titulo}. Fuente principal: {fuente_url}")

    for n, d in enumerate(slides, start=1):
        s = prs.slides.add_slide(vacio)
        _rect(s, 0, 0, W, H, claro)
        _rect(s, 0, 0, W, Inches(1.35), oscuro)
        _rect(s, 0, Inches(1.35), W, Emu(50000), acento)
        _texto(s, Inches(0.7), 0, Inches(11.9), Inches(1.35), d["titulo"], 30, "FFFFFF", bold=True, anchor=MSO_ANCHOR.MIDDLE)
        imagen = d.get("imagen_datos")
        ancho_texto = Inches(6.9) if imagen else Inches(11.5)
        caja = s.shapes.add_textbox(Inches(0.9), Inches(1.9), ancho_texto, Inches(4.8))
        tf = caja.text_frame
        tf.word_wrap = True
        tamano = (22 if len(d["puntos"]) <= 4 else 19) if imagen else (24 if len(d["puntos"]) <= 4 else 21)
        for i, punto in enumerate(d["puntos"]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(14)
            marca = p.add_run()
            marca.text = "●  "
            marca.font.size = Pt(tamano - 8)
            marca.font.color.rgb = _rgb(acento)
            run = p.add_run()
            run.text = punto
            run.font.size = Pt(tamano)
            run.font.name = "Calibri"
            run.font.color.rgb = _rgb("222222")
        if imagen:
            _poner_imagen(s, imagen, Inches(8.2), Inches(1.9), Inches(4.5), Inches(4.7))
            creditos.append(d["credito"])
        _texto(s, Inches(11.8), Inches(6.9), Inches(1.2), Inches(0.4), str(n), 12, "888888", align=PP_ALIGN.RIGHT)
        if con_notas:
            _notas(s, d.get("notas", ""))

    s = prs.slides.add_slide(vacio)
    _rect(s, 0, 0, W, H, oscuro)
    _texto(s, Inches(0.8), Inches(0.6), Inches(11.5), Inches(0.9), "Fuentes", 36, "FFFFFF", bold=True)
    _texto(s, Inches(0.8), Inches(1.6), Inches(11.5), Inches(0.7), fuente_url or "—", 18, "C9D6E8")
    if creditos:
        _texto(s, Inches(0.8), Inches(2.6), Inches(11.5), Inches(0.5), "Imágenes (Wikimedia Commons)", 18, "FFFFFF", bold=True)
        caja = s.shapes.add_textbox(Inches(0.8), Inches(3.15), Inches(11.8), Inches(4))
        tf = caja.text_frame
        tf.word_wrap = True
        for i, c in enumerate(creditos):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(3)
            run = p.add_run()
            run.text = c
            run.font.size = Pt(11)
            run.font.name = "Calibri"
            run.font.color.rgb = _rgb("C9D6E8")

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

def nombre_seguro(texto):
    limpio = re.sub(r"[^\w\- ]+", "", texto, flags=re.UNICODE).strip().replace(" ", "_")
    return (limpio or "presentacion")[:60]

# ------------------------------------------------------------------ rutas web
@app.get("/")
@login_required
def inicio():
    return render_template("index.html", con_ia=bool(GEMINI_API_KEY), es_windows=platform.system() == "Windows", temas=list(TEMAS), usuario=current_user)

@app.get("/salud")
def salud():
    return "ok"

# ✅ CORREGIDO: Ahora acepta GET y POST
@app.route("/login", methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('inicio'))
    
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = User.query.filter_by(username=username).first()
        
        if user and check_password_hash(user.password_hash, password):
            login_user(user)
            return redirect(url_for('inicio'))
        else:
            flash('Usuario o contraseña incorrectos', 'error')
    
    return render_template('login.html')

@app.get("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route("/admin/crear-usuario", methods=['GET', 'POST'])
@login_required
def crear_usuario():
    if not current_user.is_admin:
        flash('No tienes permisos de administrador', 'error')
        return redirect(url_for('inicio'))
    
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        is_admin = request.form.get('is_admin') == 'on'
        
        if User.query.filter_by(username=username).first():
            flash('El usuario ya existe', 'error')
        else:
            new_user = User(username=username, password_hash=generate_password_hash(password), is_admin=is_admin)
            db.session.add(new_user)
            db.session.commit()
            flash(f'Usuario "{username}" creado exitosamente', 'success')
    return render_template('admin.html')

@app.post("/generar")
@login_required
def generar():
    if not registrar_uso(current_user):
        return jsonify(error=f"Llegaste al límite de {LIMITE_DIARIO} presentaciones por hoy. Vuelve mañana."), 429

    datos = request.get_json(force=True, silent=True) or {}
    tema = (datos.get("tema") or "").strip()[:150]
    try:
        cantidad = max(3, min(int(datos.get("cantidad") or 6), 12))
    except (TypeError, ValueError):
        cantidad = 6

    idioma = datos.get("idioma") if datos.get("idioma") in ("es", "en") else "es"
    color = datos.get("color", "azul")
    abrir = bool(datos.get("abrir"))
    con_imagenes = bool(datos.get("imagenes", True))
    con_notas = bool(datos.get("notas", True))
    
    if not tema:
        return jsonify(error="Escribe un tema."), 400

    try:
        encontrado = buscar_wikipedia(tema, idioma)
    except requests.RequestException as e:
        return jsonify(error=f"No pude conectar con Wikipedia: {e}"), 502
    if not encontrado:
        return jsonify(error="No encontré información sobre ese tema."), 404
    titulo, texto, url = encontrado

    slides, avisos = [], []
    if GEMINI_API_KEY:
        try:
            slides = diapositivas_con_gemini(titulo, texto, cantidad, idioma)
        except Exception as e:
            avisos.append(f"La IA no respondió ({type(e).__name__}); usé solo Wikipedia.")
    if not slides:
        slides = diapositivas_desde_texto(titulo, texto, cantidad)
    if not slides:
        return jsonify(error="El artículo es muy corto para armar diapositivas."), 422

    portada = None
    if con_imagenes:
        usadas = set()
        portada = buscar_imagen(titulo, usadas)
        sin_imagen = 0
        for d in slides:
            img = buscar_imagen(d.get("imagen") or titulo, usadas) or buscar_imagen(titulo, usadas)
            if img:
                d["imagen_datos"], d["credito"] = img["datos"], img["credito"]
            else:
                sin_imagen += 1
        if sin_imagen:
            avisos.append(f"No encontré imagen para {sin_imagen} diapositiva(s).")
    aviso = " ".join(avisos) or None

    archivo = crear_pptx(titulo, slides, url, color, portada=portada, con_notas=con_notas)
    nombre = nombre_seguro(titulo) + ".pptx"

    if abrir and platform.system() == "Windows" and request.remote_addr in ("127.0.0.1", "::1"):
        ruta = os.path.join(tempfile.gettempdir(), nombre)
        with open(ruta, "wb") as f:
            f.write(archivo.getvalue())
        os.startfile(ruta)
        return jsonify(ok=True, abierto=True, archivo=ruta, aviso=aviso, diapositivas=len(slides))

    resp = send_file(archivo, as_attachment=True, download_name=nombre, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    if aviso:
        resp.headers["X-Aviso"] = aviso.encode("latin-1", "ignore").decode("latin-1")
    return resp

# Inicializar base de datos y crear admin automáticamente
with app.app_context():
    db.create_all()
    
    # Crear usuario admin si no existe
    admin_user = User.query.filter_by(username='admin').first()
    if not admin_user:
        admin_user = User(
            username='admin',
            password_hash=generate_password_hash('admin123'),
            is_admin=True
        )
        db.session.add(admin_user)
        db.session.commit()
        print("=" * 50)
        print("✅ Usuario admin creado automáticamente")
        print("👤 Usuario: admin")
        print("🔑 Contraseña: admin123")
        print("️ CAMBIA LA CONTRASEÑA DESPUÉS!")
        print("=" * 50)
    else:
        print("✅ El usuario admin ya existe")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"Abre http://127.0.0.1:{port} en tu navegador")
    app.run(host="0.0.0.0", port=port, debug=False)