"""ALTIVIA HH CONTROL — Streamlit + SQLite local (prototipo sin servidor de BD).
Subir app.py y requirements.txt a GitHub, seleccionar app.py en Streamlit Cloud.
La base de datos creada en Streamlit Cloud es EPHEMERAL: usar respaldos.
"""
import io, re, secrets, unicodedata, sqlite3, tempfile, os, html
from pathlib import Path
from contextlib import closing
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, event, text
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook

st.set_page_config(page_title='ALTIVIA | HH CONTROL',page_icon='⏱️',layout='wide',initial_sidebar_state='expanded',menu_items={'About':'ALTIVIA HH CONTROL · Gestión de horas hombre'})


# -----------------------------------------------------------------------
# CAPA DE DATOS: 100% en este archivo, mediante SQLite (incluido en Python).
# No requiere db.py, DATABASE_URL, Neon ni un servidor de PostgreSQL.
# ATENCION: Streamlit Community Cloud puede eliminar el archivo al reiniciar.
# -----------------------------------------------------------------------
DB_PATH = Path(os.environ.get('ALTIVIA_DB_PATH', str(Path(__file__).resolve().parent / 'altivia_hh.db'))).expanduser().resolve()
DB_PATH.parent.mkdir(parents=True, exist_ok=True)

@st.cache_resource(show_spinner=False)
def engine():
    connection = create_engine('sqlite:///' + DB_PATH.as_posix(), connect_args={'check_same_thread': False, 'timeout': 30})
    @event.listens_for(connection, 'connect')
    def set_pragmas(dbapi_connection, record):
        cur = dbapi_connection.cursor()
        cur.execute('PRAGMA foreign_keys=ON')
        cur.execute('PRAGMA journal_mode=WAL')
        cur.execute('PRAGMA busy_timeout=30000')
        cur.close()
    return connection

def all_rows(sql, args=None):
    with engine().connect() as conn:
        return [dict(row) for row in conn.execute(text(sql), args or {}).mappings().all()]

def one(sql, args=None):
    with engine().connect() as conn:
        row = conn.execute(text(sql), args or {}).mappings().first()
        return dict(row) if row else None

def scalar(sql, args=None):
    with engine().connect() as conn:
        return conn.execute(text(sql), args or {}).scalar()

def exec_sql(sql, args=None):
    with engine().begin() as conn:
        return conn.execute(text(sql), args or {}).rowcount

def log(conn, actor, action, target, tid=None, detail=''):
    conn.execute(text('INSERT INTO audit(actor_id,action,target,target_id,detail,at) VALUES(:actor,:action,:target,:tid,:detail,:at)'),
                 dict(actor=actor, action=action, target=target, tid=tid, detail=str(detail)[:500], at=datetime.utcnow().isoformat()))

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, name TEXT NOT NULL,
    pw TEXT NOT NULL, role TEXT NOT NULL CHECK(role IN ('admin','supervisor','trabajador')),
    active INTEGER NOT NULL DEFAULT 1, must_change_pw INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT UNIQUE NOT NULL, name TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'activo' CHECK(status IN ('activo','desactivado','terminado')),
    supervisor_id INTEGER REFERENCES users(id));
CREATE TABLE IF NOT EXISTS activities (
    id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL REFERENCES projects(id),
    code TEXT NOT NULL, name TEXT NOT NULL,
    budget_quarters INTEGER NOT NULL DEFAULT 0 CHECK(budget_quarters >= 0), active INTEGER NOT NULL DEFAULT 1,
    UNIQUE(project_id,code));
CREATE TABLE IF NOT EXISTS assignments (
    user_id INTEGER REFERENCES users(id), project_id INTEGER REFERENCES projects(id), PRIMARY KEY(user_id,project_id));
CREATE TABLE IF NOT EXISTS activity_assignments (
    user_id INTEGER REFERENCES users(id), activity_id INTEGER REFERENCES activities(id),
    planned_quarters INTEGER NOT NULL DEFAULT 0, PRIMARY KEY(user_id,activity_id));
CREATE TABLE IF NOT EXISTS rates (
    id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER UNIQUE REFERENCES users(id),
    rate_cents INTEGER NOT NULL CHECK(rate_cents >= 0));
CREATE TABLE IF NOT EXISTS entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL REFERENCES users(id),
    activity_id INTEGER NOT NULL REFERENCES activities(id), work_date DATE NOT NULL,
    quarters INTEGER NOT NULL CHECK(quarters BETWEEN 1 AND 48), description TEXT NOT NULL,
    state TEXT NOT NULL DEFAULT 'pendiente' CHECK(state IN ('borrador','pendiente','observado','aprobado','rechazado','liquidado')),
    note TEXT NOT NULL DEFAULT '', created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, approval_by INTEGER REFERENCES users(id),
    approval_at TIMESTAMP, liquidation_id INTEGER, submission_key TEXT UNIQUE);
CREATE TABLE IF NOT EXISTS liquidations (
    id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL REFERENCES users(id),
    period TEXT NOT NULL, cutoff_date DATE NOT NULL, cents INTEGER NOT NULL, quarters INTEGER NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, created_by INTEGER REFERENCES users(id));
CREATE TABLE IF NOT EXISTS liquidation_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT, liquidation_id INTEGER NOT NULL REFERENCES liquidations(id),
    entry_id INTEGER NOT NULL UNIQUE REFERENCES entries(id), quarters INTEGER NOT NULL,
    rate_cents INTEGER NOT NULL, amount_cents INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS audit (
    id INTEGER PRIMARY KEY AUTOINCREMENT, actor_id INTEGER, action TEXT NOT NULL, target TEXT NOT NULL,
    target_id INTEGER, detail TEXT, at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS import_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER REFERENCES projects(id), filename TEXT,
    processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, actor_id INTEGER, notes TEXT);
CREATE INDEX IF NOT EXISTS idx_entries_date ON entries(work_date);
CREATE INDEX IF NOT EXISTS idx_entries_user_state ON entries(user_id,state);
CREATE INDEX IF NOT EXISTS idx_entries_activity ON entries(activity_id);
CREATE INDEX IF NOT EXISTS idx_activities_project ON activities(project_id);
"""

def init_schema():
    with engine().begin() as conn:
        for statement in SCHEMA.split(';'):
            if statement.strip():
                conn.execute(text(statement))

def cents_for_quarters(quarters, rate_cents):
    """Redondeo monetario correcto, por registro, a centimos."""
    return int((Decimal(quarters) * Decimal(rate_cents) / Decimal(4)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))

def sqlite_backup_bytes():
    """Copia SQLite coherente aun con journal WAL. No copia archivos incompletos."""
    buf = io.BytesIO()
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp) / 'altivia_backup.db'
        with closing(sqlite3.connect(str(DB_PATH))) as src, closing(sqlite3.connect(str(target))) as dst:
            src.backup(dst)
        buf.write(target.read_bytes())
    return buf.getvalue()

def restore_sqlite_bytes(payload):
    """Restauracion administrativa. Ejecutar con sesiones de usuarios cerradas."""
    with tempfile.TemporaryDirectory() as tmp:
        restored = Path(tmp) / 'upload.db'
        restored.write_bytes(payload)
        with closing(sqlite3.connect(str(restored))) as src:
            if src.execute('PRAGMA quick_check').fetchone()[0] != 'ok':
                raise ValueError('El archivo SQLite no supera la verificacion de integridad.')
            names = {r[0] for r in src.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            needed = {'users','projects','activities','entries','rates','liquidations','audit'}
            if not needed.issubset(names):
                raise ValueError('No corresponde a un respaldo valido de ALTIVIA HH CONTROL.')
            # Primero extraer una copia de seguridad del contenido actual, antes de restaurar.
            engine().dispose()
            with closing(sqlite3.connect(str(DB_PATH))) as dst:
                src.backup(dst)
    st.cache_resource.clear()

# Diseño ALTIVIA. CSS autocontenido: no requiere archivos externos.
st.markdown("""
<style>
:root { --ink: #142b3b; --muted:#64748b; --line:#e1eaf0; --teal:#087e79; --surface:#ffffff; }
html, body, [class*="css"], [data-testid="stApp"] {font-family: 'Inter', 'Segoe UI', system-ui, sans-serif;}
[data-testid="stAppViewContainer"] {background:linear-gradient(180deg,#eff5f8 0px,#f6f9fb 250px,#f7f9fc 100%);}
.block-container {padding-top:1.65rem; padding-bottom:3.6rem; max-width:1480px;}
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}
h1,h2,h3 {color:#152f40;letter-spacing:-.025em;}
h1 {font-size:clamp(1.65rem,2.7vw,2.55rem)!important;font-weight:780!important;}
h2 {font-size:1.35rem!important;}
h3 {font-size:1.04rem!important;}
[data-testid="stSidebar"] {background:linear-gradient(165deg,#122f41 0%,#0e2637 72%,#0a222f 100%);border-right:1px solid #163b4a;}
[data-testid="stSidebar"] [data-testid="stSidebarContent"]{padding-top:1.25rem;}
[data-testid="stSidebar"] * {color:#f0f6f8;}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p{color:#d6e5eb;}
[data-testid="stSidebar"] [data-testid="stRadio"] label {border-radius:12px;padding:0.15rem .25rem;transition:background .15s;}
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {background:#23485a;}
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {background:#1d5861;border:1px solid #4b9798;}
[data-testid="stSidebar"] [data-testid="stRadio"] label p {font-size:.90rem;font-weight:600;color:#f4fafb!important;}
[data-testid="stSidebar"] [data-testid="stRadio"] [data-testid="stWidgetLabel"]{display:none;}
[data-testid="stSidebar"] [data-testid="stButton"] button {border-color:#54727f;background:#1b4355;color:#f6fafb;}
[data-testid="stSidebar"] [data-testid="stButton"] button:hover {background:#2b6070;border-color:#92c9c7;}
[data-testid="stSidebar"] hr {border-color:#355261;}
[data-testid="stSidebar"] .stCaption {color:#a7bfca!important;}
[data-testid="stMetric"] {background:#fff;border:1px solid #e2ebf0;border-radius:17px;padding:1rem 1.12rem;box-shadow:0 4px 18px rgba(13,45,67,.035);min-height:118px;}
[data-testid="stMetricLabel"] {font-size:.84rem;color:#667788;}
[data-testid="stMetricValue"] {font-size:clamp(1.24rem,2vw,1.86rem);font-weight:780;color:#14394b;}
[data-testid="stMetricDelta"] {font-size:.78rem;}
[data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"] [data-testid="stForm"]) {border-radius:17px;}
[data-testid="stForm"] {background:rgba(255,255,255,.92);border:1px solid #dfebf0;border-radius:17px;padding:1.15rem 1.15rem .75rem;box-shadow:0 4px 18px rgba(10,49,68,.03);}
[data-testid="stDataFrame"] {border:1px solid #e0eaf0;border-radius:14px;overflow:hidden;}
[data-testid="stExpander"] {border:1px solid #e2e9f0;border-radius:14px;background:#fff;overflow:hidden;}
[data-testid="stTabs"] [data-baseweb="tab-list"] {gap:.42rem;border-bottom:1px solid #dbe6ec;}
[data-testid="stTabs"] button[data-baseweb="tab"]{border-radius:10px 10px 0 0;padding:.55rem .9rem;font-weight:650;}
[data-testid="stTabs"] button[aria-selected="true"]{color:#087e79;border-bottom-color:#087e79;}
[data-testid="stSelectbox"], [data-testid="stDateInput"], [data-testid="stTextInput"], [data-testid="stNumberInput"], [data-testid="stTextArea"] {margin-bottom:.35rem;}
[data-testid="stBaseButton-primary"] {background:#067d77;border:1px solid #067d77;border-radius:11px;font-weight:730;box-shadow:0 3px 8px rgba(8,126,120,.13);}
[data-testid="stBaseButton-primary"]:hover {background:#086963;border-color:#086963;}
[data-testid="stBaseButton-secondary"] {border-radius:11px;font-weight:650;}
[data-testid="stDownloadButton"] button {border-radius:11px;font-weight:650;}
[data-testid="stAlert"] {border-radius:12px;}
.alt-hero {border-radius:20px;background:linear-gradient(112deg,#14364a 0%,#145461 65%,#0b7976 100%);padding:1.6rem 1.85rem;margin:.3rem 0 1.45rem;position:relative;overflow:hidden;box-shadow:0 9px 28px rgba(20,55,73,.11);}
.alt-hero:after {content:'';position:absolute;right:-75px;top:-108px;width:330px;height:330px;border-radius:50%;border:50px solid rgba(255,255,255,.045);pointer-events:none;}
.alt-eyebrow {text-transform:uppercase;letter-spacing:.15em;font-weight:800;font-size:.67rem;color:#9cddd6;margin-bottom:.65rem;}
.alt-hero-title {font-size:clamp(1.45rem,2.7vw,2.35rem);font-weight:780;letter-spacing:-.025em;color:#fff;line-height:1.23;position:relative;z-index:1;}
.alt-hero-sub {font-size:.91rem;color:#d5e7ed;margin-top:.6rem;max-width:850px;line-height:1.55;position:relative;z-index:1;}
.alt-section {display:flex;align-items:center;gap:.6rem;margin:1.45rem 0 .75rem;font-size:1.06rem;font-weight:790;color:#1c384a;}
.alt-section:before {content:'';width:4px;height:19px;border-radius:10px;background:#14a49b;display:block;}
.alt-mini {border:1px solid #e0eaf0;background:#fff;border-radius:15px;padding:1rem 1.15rem;margin:.2rem 0 1rem;}
.alt-mini-head {display:flex;justify-content:space-between;align-items:center;gap:.5rem;font-size:.85rem;font-weight:700;color:#243f4f;}
.alt-mini-hint {font-size:.78rem;color:#64778a;margin-top:.32rem;line-height:1.4;}
.alt-bar {height:8px;border-radius:20px;background:#e8eff4;overflow:hidden;margin-top:.7rem;}
.alt-bar-fill {height:100%;border-radius:20px;background:linear-gradient(90deg,#14a49b,#52bdaa);}
.alt-bar-fill.warn {background:linear-gradient(90deg,#d39a32,#e9b960);}
.alt-bar-fill.over {background:linear-gradient(90deg,#e17761,#cf4c4c);}
.alt-value {font-size:1.25rem;font-weight:800;color:#18384a;margin-top:.33rem;}
.alt-tag {font-size:.71rem;border-radius:20px;padding:3px 9px;font-weight:760;background:#e0f4ef;color:#087a6e;}
.alt-tag.wait {background:#fff2d9;color:#a66610;}
.alt-tag.bad {background:#fee7e4;color:#ae4040;}
.alt-tag.info {background:#e6eefb;color:#3f65a0;}
.alt-empty {text-align:center;padding:2.4rem 1rem;background:white;border:1px dashed #d0e0e9;border-radius:18px;margin:.6rem 0 1rem;color:#65798a;}
.alt-empty strong {display:block;color:#29495b;font-size:1.05rem;margin-bottom:.25rem;}
.alt-empty-symbol {font-size:2.1rem;display:block;margin-bottom:.5rem;}
.alt-brand {display:flex;align-items:center;gap:.75rem;padding:4px 0 18px;}
.alt-brand-mark {display:flex;align-items:center;justify-content:center;font-size:1.24rem;border-radius:12px;width:44px;height:44px;background:#1c6870;color:#fff;font-weight:900;border:1px solid #4e9b9e;}
.alt-brand-name {font-size:1.12rem;font-weight:850;letter-spacing:.075em;color:#fff;line-height:1.1;}
.alt-brand-sub {font-size:.69rem;color:#a6c4ca;font-weight:640;margin-top:4px;letter-spacing:.04em;}
.alt-sidebox {border:1px solid #365a69;border-radius:13px;padding:.8rem .85rem;background:rgba(255,255,255,.055);font-size:.78rem;color:#d8e7eb;line-height:1.5;margin:.85rem 0;}
.alt-sideuser {color:#fff;font-size:.86rem;font-weight:730;}
.alt-sidemuted {font-size:.75rem;color:#aec3ca;margin-top:3px;}
.alt-quick {font-size:.8rem;color:#607384;margin:.1rem 0 1rem;}
@media(max-width:900px){.block-container{padding:1rem .75rem 3rem;}.alt-hero{padding:1.25rem 1.2rem;margin-bottom:1.15rem;}[data-testid="stMetric"]{min-height:95px;padding:.75rem;}}
</style>
""",unsafe_allow_html=True)

try:init_schema()
except Exception as exc:
    st.error('No se pudo inicializar la base de datos SQLite.');st.code(str(exc));st.stop()

def norm(v):return ''.join(ch for ch in unicodedata.normalize('NFKD',str(v).casefold()) if not unicodedata.combining(ch)).strip()
def login_from_name(name):
    words=re.findall('[a-z0-9]+',norm(name));return words[-1] if words else 'trabajador'
def hours_to_quarters(v):
    d=Decimal(str(v))
    if d<=0 or d>12 or d*4 != (d*4).to_integral_value():raise ValueError('Ingrese entre 0.25 y 12 HH en incrementos de 0.25.')
    return int(d*4)
def money(v):return f'S/ {v/100:,.2f}'

def safe(v):
    return html.escape(str(v if v is not None else ''),quote=True)

def hero(title, subtitle='', section='CONTROL DE HORAS HOMBRE'):
    st.markdown(f'<div class="alt-hero"><div class="alt-eyebrow">{safe(section)}</div><div class="alt-hero-title">{safe(title)}</div><div class="alt-hero-sub">{safe(subtitle)}</div></div>',unsafe_allow_html=True)

def section(title):
    st.markdown(f'<div class="alt-section">{safe(title)}</div>',unsafe_allow_html=True)

def empty(title, desc='', icon='◌'):
    st.markdown(f'<div class="alt-empty"><span class="alt-empty-symbol">{safe(icon)}</span><strong>{safe(title)}</strong>{safe(desc)}</div>',unsafe_allow_html=True)

def mini(title,value=None,hint=None,percent=None):
    content=f'<div class="alt-mini"><div class="alt-mini-head"><span>{safe(title)}</span>'
    if percent is not None:
        color='over' if percent>100 else 'warn' if percent>=80 else ''
        content+=f'<span class="alt-tag {"bad" if percent>100 else "wait" if percent>=80 else ""}">{percent:,.1f}%</span>'
    content+='</div>'
    if value is not None:content+=f'<div class="alt-value">{safe(value)}</div>'
    if hint:content+=f'<div class="alt-mini-hint">{safe(hint)}</div>'
    if percent is not None:content+=f'<div class="alt-bar"><div class="alt-bar-fill {color}" style="width:{max(0,min(100,percent)):.1f}%"></div></div>'
    st.markdown(content+'</div>',unsafe_allow_html=True)

def export_xlsx(frame,filename,sheet='Datos'):
    if frame.empty:return
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as writer:frame.to_excel(writer,index=False,sheet_name=sheet[:31])
    st.download_button('⬇ Exportar Excel',out.getvalue(),filename,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='dl_'+filename)

def table_data(frame,**kwargs):
    if frame.empty:
        empty('Sin registros para mostrar','Prueba con otros filtros o agrega información.','📋')
    else:
        st.dataframe(frame,use_container_width=True,hide_index=True,**kwargs)

def me():
    uid=st.session_state.get('uid')
    return one('SELECT * FROM users WHERE id=:i AND active=1',{'i':uid}) if uid else None

def require(role):
    if me()['role'] not in role:st.error('Acceso no autorizado.');st.stop()
def options(rows,key='name'):return {f"{r.get('code',r['id'])} · {r[key]}":r['id'] for r in rows}
def choose(title,rows,key='name',empty=False,k=None):
    opts=options(rows,key);names=(['Todos'] if empty else [])+list(opts)
    selected=st.selectbox(title,names,key=k) if names else None
    return None if selected in ('Todos',None) else opts[selected]
def audit_event(c,action,target,tid=None,detail=''):log(c,me()['id'],action,target,tid,detail)

def authenticate():
    if scalar('SELECT COUNT(*) FROM users')==0:
        hero('Configura tu espacio de trabajo','Crea la primera cuenta administradora para empezar a gestionar proyectos, HH y liquidaciones.','PRIMER ACCESO · ALTIVIA')
        st.warning('Este formulario crea el primer administrador. Hazlo antes de compartir el enlace.')
        try:
            setup_key = str(st.secrets.get('ADMIN_SETUP_KEY', os.environ.get('ADMIN_SETUP_KEY','')))
        except Exception:
            setup_key = os.environ.get('ADMIN_SETUP_KEY','')
        with st.form('setup'):
            name=st.text_input('Nombre completo');login=st.text_input('Usuario administrador').lower().strip();pwd=st.text_input('Contraseña (mínimo 12 caracteres)',type='password');confirm=st.text_input('Repetir contraseña',type='password')
            supplied_key=st.text_input('Clave de activación inicial',type='password') if setup_key else ''
            submit=st.form_submit_button('Crear administrador')
        if submit:
            if setup_key and not secrets.compare_digest(supplied_key,setup_key):st.error('Clave de activación incorrecta.')
            elif not name.strip() or not re.fullmatch('[a-z0-9_.-]{3,40}',login) or len(pwd)<12 or pwd!=confirm:st.error('Verifica el nombre, usuario y contraseña (12 caracteres).')
            else:
                try:
                    with engine().begin() as c:
                        result=c.execute(text("INSERT INTO users(username,name,pw,role) SELECT :l,:n,:pw,'admin' WHERE NOT EXISTS(SELECT 1 FROM users)"),dict(l=login,n=name.strip(),pw=generate_password_hash(pwd)))
                        if result.rowcount!=1:raise ValueError('El administrador ya fue creado en otra sesión.')
                    st.success('Administrador creado; inicia sesión.');st.rerun()
                except Exception as exc:st.error(str(exc))
        st.stop()
    if not me():
        hero('Bienvenido a ALTIVIA','Ingresa con tus credenciales para acceder a la gestión de proyectos, presupuestos y horas hombre.','PLATAFORMA INTERNA · ACCESO SEGURO')
        with st.form('login'):
            login=st.text_input('Usuario').strip().lower();pwd=st.text_input('Contraseña',type='password');submitted=st.form_submit_button('Ingresar',type='primary')
        if submitted:
            u=one('SELECT * FROM users WHERE username=:u AND active=1',{'u':login})
            if u and check_password_hash(u['pw'],pwd):
                st.session_state['uid']=u['id'];st.rerun()
            else:st.error('Usuario o contraseña incorrectos.')
        st.stop()
    if me()['must_change_pw']:
        hero('Protege tu cuenta','Por seguridad, reemplaza tu contraseña temporal antes de continuar.','SEGURIDAD DE ACCESO')
        with st.form('change-pwd'):
            pwd=st.text_input('Nueva contraseña (mínimo 12 caracteres)',type='password');again=st.text_input('Repetir nueva contraseña',type='password');ok=st.form_submit_button('Actualizar contraseña')
        if ok:
            if len(pwd)<12 or pwd!=again:st.error('Usa una contraseña de al menos 12 caracteres, idéntica en los dos campos.')
            else:
                with engine().begin() as c:
                    c.execute(text('UPDATE users SET pw=:p,must_change_pw=0 WHERE id=:i'),dict(p=generate_password_hash(pwd),i=me()['id']))
                    audit_event(c,'cambio_contrasena','user',me()['id'])
                st.rerun()
        st.stop()

authenticate()
u=me()
with st.sidebar:
    st.markdown('<div class="alt-brand"><div class="alt-brand-mark">A</div><div><div class="alt-brand-name">ALTIVIA</div><div class="alt-brand-sub">HH CONTROL · WORKSPACE</div></div></div>',unsafe_allow_html=True)
    st.markdown(f'<div class="alt-sidebox"><div class="alt-sideuser">👤 {safe(u["name"])}</div><div class="alt-sidemuted">{safe({"admin":"Administrador","supervisor":"Supervisor","trabajador":"Profesional"}.get(u["role"],u["role"]))}</div></div>',unsafe_allow_html=True)
    st.caption('ESPACIO DE TRABAJO')
menus=['Centro gerencial','Mis HH']
if u['role'] in ('admin','supervisor'):menus+=['HH por profesional','Aprobaciones','Presupuestos EDT']
if u['role']=='admin':menus+=['Proyectos y actividades','Usuarios y tarifas','Importar Excel','Liquidaciones','Auditoría','Respaldo y restauración']
NAV_NAMES={'Centro gerencial':'▦  Centro gerencial','Mis HH':'◷  Mis horas hombre','HH por profesional':'♙  HH por profesional','Aprobaciones':'✓  Aprobaciones','Presupuestos EDT':'▥  Presupuestos EDT','Proyectos y actividades':'▣  Proyectos y actividades','Usuarios y tarifas':'♧  Usuarios y tarifas','Importar Excel':'⇪  Importar Excel','Liquidaciones':'◈  Liquidaciones','Auditoría':'◎  Auditoría','Respaldo y restauración':'⬡  Respaldos'}
if st.session_state.get('nav_page') not in menus:st.session_state['nav_page']=menus[0]
page=st.sidebar.radio('Módulos',menus,key='nav_page',format_func=lambda x:NAV_NAMES.get(x,x),label_visibility='collapsed')
with st.sidebar:
    st.divider()
    st.markdown('<div class="alt-sidebox">🔐 <strong>Datos de prueba / piloto</strong><div class="alt-sidemuted">SQLite local. En Streamlit Cloud debes respaldar antes de cada actualización.</div></div>',unsafe_allow_html=True)
    if st.button('↪  Cerrar sesión',use_container_width=True):
        st.session_state.clear();st.rerun()


def projects_access(include_inactive=True):
    if u['role']=='admin':cond='1=1';args={}
    elif u['role']=='supervisor':cond='supervisor_id=:uid';args={'uid':u['id']}
    else:cond='id IN (SELECT project_id FROM assignments WHERE user_id=:uid)';args={'uid':u['id']}
    if not include_inactive:cond+=" AND status='activo'"
    return all_rows('SELECT * FROM projects WHERE '+cond+' ORDER BY code',args)

def entries_query(pid=None,start=None,end=None,worker=None):
    clauses=[];args={}
    if u['role']=='trabajador':clauses.append('e.user_id=:viewer');args['viewer']=u['id']
    elif u['role']=='supervisor':clauses.append('p.supervisor_id=:viewer');args['viewer']=u['id']
    if pid:clauses.append('p.id=:pid');args['pid']=pid
    if start:clauses.append('e.work_date>=:start');args['start']=start
    if end:clauses.append('e.work_date<=:end');args['end']=end
    if worker:clauses.append('e.user_id=:worker');args['worker']=worker
    return all_rows('''SELECT e.*,us.name worker,p.code project,p.id project_id,a.code activity_code,a.name activity
        FROM entries e JOIN users us ON us.id=e.user_id JOIN activities a ON a.id=e.activity_id
        JOIN projects p ON p.id=a.project_id WHERE '''+(' AND '.join(clauses) if clauses else '1=1')+' ORDER BY e.work_date DESC,e.id DESC',args)

def date_filters():
    c1,c2,c3=st.columns([2,1,1]);projs=projects_access()
    with c1:pid=choose('Proyecto (opcional)',projs,empty=True)
    with c2:since=st.date_input('Desde',value=None,key='from_date')
    with c3:until=st.date_input('Hasta',value=None,key='to_date')
    if since and until and since>until:st.error('La fecha inicial supera la final.');st.stop()
    return pid,since,until

if page=='Centro gerencial':
    hero('Visión general de tus proyectos','Monitorea el presupuesto, el consumo y las horas pendientes de revisión. Todos los indicadores responden a los filtros activos.','RESUMEN EJECUTIVO')
    with st.expander('⚙️ Filtros del tablero',expanded=True):
        pid,start,end=date_filters()
    projs=[p for p in projects_access() if not pid or p['id']==pid]
    entries=[r for r in entries_query(pid,start,end) if r['state'] in ('pendiente','aprobado','liquidado')]
    budgets={r['project_id']:r['qty'] for r in all_rows('SELECT project_id,COALESCE(SUM(budget_quarters),0) qty FROM activities GROUP BY project_id')}
    budget=sum(budgets.get(p['id'],0) for p in projs)/4
    total=sum(e['quarters'] for e in entries)/4
    approved=sum(e['quarters'] for e in entries if e['state'] in ('aprobado','liquidado'))/4
    pending=sum(e['quarters'] for e in entries if e['state']=='pendiente')/4
    section('Indicadores clave')
    c1,c2,c3,c4=st.columns(4)
    c1.metric('Presupuesto asignado',f'{budget:,.2f} HH',help='Presupuesto completo de los proyectos filtrados')
    c2.metric('HH registradas',f'{total:,.2f} HH',help='Solo registros pendientes, aprobados y liquidados del periodo')
    c3.metric('HH validadas',f'{approved:,.2f} HH',help='Registros aprobados + liquidados')
    c4.metric('HH por aprobar',f'{pending:,.2f} HH',help='Registro pendiente de revisión')
    c5,c6,c7=st.columns(3)
    with c5:mini('Avance del presupuesto',f'{(100*total/budget):,.1f}%' if budget else 'Sin presupuesto','Consumo del período frente al presupuesto total.',100*total/budget if budget else 0)
    with c6:mini('Saldo presupuestal',f'{budget-total:,.2f} HH','Presupuesto menos HH consumidas en el filtro.')
    with c7:mini('Proyectos supervisados',str(len(projs)),'Proyectos visibles según tus permisos.')
    rows=[]
    for proj in projs:
        er=[x for x in entries if x['project_id']==proj['id']]
        q=sum(x['quarters'] for x in er)/4;b=budgets.get(proj['id'],0)/4
        rows.append(dict(Proyecto=f"{proj['code']} · {proj['name']}",Estado=proj['status'],Presupuesto_HH=b,HH_cargadas=q,HH_aprobadas=sum(x['quarters'] for x in er if x['state'] in ('aprobado','liquidado'))/4,HH_por_aprobar=sum(x['quarters'] for x in er if x['state']=='pendiente')/4,Diferencia_HH=b-q,Consumo_pct=round(q/b*100,1) if b else 0))
    section('Comparativo por proyecto')
    if rows:
        frame=pd.DataFrame(rows)
        left,right=st.columns([1.65,1],gap='large')
        with left:
            display=frame.rename(columns={'Presupuesto_HH':'Presupuesto (HH)','HH_cargadas':'Consumo (HH)','HH_por_aprobar':'Pendiente (HH)','Consumo_pct':'Avance (%)'})
            table_data(display[['Proyecto','Presupuesto (HH)','Consumo (HH)','Pendiente (HH)','Avance (%)']],column_config={'Avance (%)':st.column_config.NumberColumn(format='%.1f %%')},height=330)
        with right:
            st.caption('Presupuesto vs. consumo · HH')
            chart=frame.set_index('Proyecto')[['Presupuesto_HH','HH_cargadas']].rename(columns={'Presupuesto_HH':'Presupuesto','HH_cargadas':'Consumo'})
            st.bar_chart(chart,horizontal=True,height=320,color=['#143d50','#16a69a'])
        section('Semáforo presupuestal')
        cols=st.columns(3)
        for idx,r in enumerate(rows):
            with cols[idx%3]:mini(r['Proyecto'],f"{r['HH_cargadas']:,.2f} / {r['Presupuesto_HH']:,.2f} HH",f"Saldo: {r['Diferencia_HH']:,.2f} HH",r['Consumo_pct'] if r['Presupuesto_HH'] else None)
    else:empty('Aún no hay proyectos visibles','Crea un proyecto o solicita la asignación correspondiente.','🏗️')
    if entries:
        with st.expander('📥 Descargar detalle del período'):
            export=pd.DataFrame([dict(ID=e['id'],Fecha=e['work_date'],Profesional=e['worker'],Proyecto=e['project'],EDT=e['activity_code'],Actividad=e['activity'],HH=e['quarters']/4,Estado=e['state'],Descripcion=e['description']) for e in entries])
            export_xlsx(export,'altivia_hh_filtradas.xlsx','Horas Hombre')
    if u['role'] in ('admin','supervisor'):
        section('Acciones frecuentes')
        a,b,c=st.columns(3)
        a.button('✓ Revisar aprobaciones',use_container_width=True,on_click=lambda:st.session_state.update(nav_page='Aprobaciones'))
        b.button('▥ Consultar presupuestos',use_container_width=True,on_click=lambda:st.session_state.update(nav_page='Presupuestos EDT'))
        c.button('◷ Registrar mis HH',use_container_width=True,on_click=lambda:st.session_state.update(nav_page='Mis HH'))

elif page=='Mis HH':
    hero('Mis horas hombre','Registra tus tareas, consulta el estado de cada envío y corrige observaciones en un mismo lugar.','REGISTRO PERSONAL')
    history=entries_query(worker=u['id'])[:500]
    totals={s:sum(r['quarters'] for r in history if r['state']==s)/4 for s in ['pendiente','aprobado','liquidado','observado']}
    x1,x2,x3,x4=st.columns(4)
    x1.metric('Pendientes',f"{totals['pendiente']:,.2f} HH")
    x2.metric('Aprobadas',f"{totals['aprobado']:,.2f} HH")
    x3.metric('Liquidadas',f"{totals['liquidado']:,.2f} HH")
    x4.metric('Observadas',str(len([r for r in history if r['state']=='observado'])))
    tab1,tab2,tab3=st.tabs(['＋  Nuevo registro','☷  Mi historial','⚠  Corregir observaciones'])
    with tab1:
        section('Registrar horas trabajadas')
        projs=projects_access(False)
        if not projs:empty('No tienes proyectos activos asignados','Solicita a un administrador acceso a un proyecto y a sus actividades.','📁')
        else:
            pid=choose('1 · Proyecto',projs)
            acts=all_rows('''SELECT a.id,a.code,a.name FROM activities a JOIN activity_assignments aa ON aa.activity_id=a.id
                WHERE aa.user_id=:uid AND a.project_id=:pid AND a.active=1 ORDER BY a.code''',dict(uid=u['id'],pid=pid)) if pid else []
            if not acts:st.info('No tienes actividades autorizadas en este proyecto. Solicita tu asignación.')
            else:
                with st.form('new-entry',clear_on_submit=True):
                    activity_id=choose('2 · Actividad EDT / WBS',acts)
                    c1,c2=st.columns(2)
                    with c1:work_date=st.date_input('3 · Fecha del trabajo',date.today())
                    with c2:hh=st.number_input('4 · Horas hombre (HH)',min_value=0.25,max_value=12.,step=.25,value=1.,help='Hasta 12 HH por día, en incrementos de 15 minutos.')
                    description=st.text_area('5 · Descripción técnica del trabajo',max_chars=1500,placeholder='Ej.: Revisión de resultados de instrumentación geotécnica y elaboración de informe...')
                    st.caption('Antes de enviar, verifica fecha, EDT y horas. El registro irá a la bandeja de aprobación.')
                    sent=st.form_submit_button('✓ Enviar a aprobación',type='primary',use_container_width=True)
                if sent:
                    try:
                        q=hours_to_quarters(hh)
                        if not description.strip():raise ValueError('Describe el trabajo realizado.')
                        with engine().begin() as c:
                            valid=c.execute(text('''SELECT a.id FROM activities a JOIN projects p ON p.id=a.project_id JOIN activity_assignments aa ON aa.activity_id=a.id
                                WHERE a.id=:aid AND p.id=:pid AND p.status='activo' AND a.active=1 AND aa.user_id=:uid'''),dict(aid=activity_id,pid=pid,uid=u['id'])).first()
                            if not valid:raise ValueError('La actividad ya no está disponible para registrar HH.')
                            day_q=c.execute(text("SELECT COALESCE(SUM(quarters),0) FROM entries WHERE user_id=:u AND work_date=:d AND state NOT IN ('rechazado','observado')"),dict(u=u['id'],d=work_date)).scalar()
                            if day_q+q>48:raise ValueError('La jornada no puede superar 12 HH.')
                            eid=c.execute(text('''INSERT INTO entries(user_id,activity_id,work_date,quarters,description,state,submission_key)
                                VALUES(:uid,:aid,:dt,:q,:desc,'pendiente',:key) RETURNING id'''),dict(uid=u['id'],aid=activity_id,dt=work_date,q=q,desc=description.strip(),key=secrets.token_hex(16))).scalar_one()
                            audit_event(c,'registrar_hh','entry',eid)
                        st.success(f'Registro #{eid} enviado correctamente.');st.rerun()
                    except Exception as exc:st.error(str(exc))
    with tab2:
        section('Consulta de registros')
        if history:
            a,b=st.columns([1,1])
            with a:state_filter=st.selectbox('Estado', ['Todos','Pendiente','Aprobado','Liquidado','Observado','Rechazado','Borrador'],key='mine_state')
            with b:search=st.text_input('Buscar proyecto, EDT o actividad',placeholder='Escribe para filtrar...',key='mine_search')
            visible=[r for r in history if (state_filter=='Todos' or r['state']==state_filter.lower()) and (not search or norm(search) in norm(' '.join([r['project'],r['activity_code'],r['activity'],r['description']]))) ]
            table_data(pd.DataFrame([dict(ID=r['id'],Fecha=r['work_date'],Proyecto=r['project'],EDT=r['activity_code'],Actividad=r['activity'],HH=r['quarters']/4,Estado=r['state'].capitalize(),Observacion=r['note']) for r in visible]),height=440)
        else:empty('Todavía no registras HH','Crea tu primer registro desde la pestaña Nuevo registro.','⏱️')
    with tab3:
        observed=[r for r in history if r['state']=='observado']
        if not observed:empty('No tienes observaciones pendientes','Tus registros observados aparecerán aquí para su corrección.','✅')
        else:
            st.info(f'Tienes {len(observed)} registro(s) que requieren atención.')
            eid=st.selectbox('Selecciona el registro observado', [r['id'] for r in observed],format_func=lambda i:f"#{i} · {next(x['activity'] for x in observed if x['id']==i)}")
            r=next(x for x in observed if x['id']==eid)
            if r['note']:st.warning('Motivo de la observación: '+r['note'])
            with st.form('fix'):
                h=st.number_input('HH corregidas',value=r['quarters']/4,min_value=.25,max_value=12.,step=.25)
                desc=st.text_area('Descripción corregida',value=r['description'])
                if st.form_submit_button('↻ Corregir y reenviar',type='primary'):
                    try:
                        q=hours_to_quarters(h)
                        if not desc.strip():raise ValueError('La descripción técnica no puede estar vacía.')
                        with engine().begin() as c:
                            previous=c.execute(text("SELECT work_date FROM entries WHERE id=:id AND user_id=:uid AND state='observado'"),dict(id=eid,uid=u['id'])).scalar()
                            if previous is None:raise ValueError('Registro no disponible.')
                            other_q=c.execute(text("SELECT COALESCE(SUM(quarters),0) FROM entries WHERE user_id=:uid AND work_date=:d AND id<>:id AND state NOT IN ('observado','rechazado')"),dict(uid=u['id'],d=previous,id=eid)).scalar()
                            if other_q+q>48:raise ValueError('La jornada no puede superar 12 HH.')
                            result=c.execute(text('''UPDATE entries SET quarters=:q,description=:d,state='pendiente',note='',approval_by=NULL,approval_at=NULL,updated_at=CURRENT_TIMESTAMP
                                WHERE id=:id AND user_id=:uid AND state='observado' AND EXISTS (SELECT 1 FROM activities a JOIN projects p ON p.id=a.project_id WHERE a.id=entries.activity_id AND p.status='activo' AND a.active=1)'''),dict(q=q,d=desc.strip(),id=eid,uid=u['id']))
                            if result.rowcount!=1:raise ValueError('Registro no disponible para corrección.')
                            audit_event(c,'corregir_hh','entry',eid)
                        st.success('Observación corregida y reenviada.');st.rerun()
                    except Exception as exc:st.error(str(exc))

elif page=='HH por profesional':
    hero('Rendimiento por profesional','Consulta horas hombre por persona, proyecto y período, con un desglose listo para exportar.','ANÁLISIS DE EQUIPOS')
    with st.expander('⚙️ Filtros de consulta',expanded=True):pid,start,end=date_filters()
    entries=entries_query(pid,start,end)
    rows={}
    for e in entries:
        if e['state'] not in ('pendiente','aprobado','liquidado'):continue
        key=(e['worker'],e['project']);r=rows.setdefault(key,dict(Profesional=e['worker'],Proyecto=e['project'],HH_cargadas=0,HH_aprobadas=0,HH_por_aprobar=0,HH_liquidadas=0))
        q=e['quarters']/4;r['HH_cargadas']+=q
        if e['state'] in ('aprobado','liquidado'):r['HH_aprobadas']+=q
        if e['state']=='pendiente':r['HH_por_aprobar']+=q
        if e['state']=='liquidado':r['HH_liquidadas']+=q
    frame=pd.DataFrame(list(rows.values()))
    section('Resumen de personal')
    a,b,c=st.columns(3)
    a.metric('Profesionales',frame['Profesional'].nunique() if not frame.empty else 0)
    b.metric('HH cargadas',f"{frame['HH_cargadas'].sum():,.2f}" if not frame.empty else '0.00')
    c.metric('HH aprobadas',f"{frame['HH_aprobadas'].sum():,.2f}" if not frame.empty else '0.00')
    if not frame.empty:
        left,right=st.columns([1.5,1],gap='large')
        with left:table_data(frame.rename(columns={'HH_cargadas':'Cargadas (HH)','HH_aprobadas':'Aprobadas (HH)','HH_por_aprobar':'Pendientes (HH)','HH_liquidadas':'Liquidadas (HH)'}),height=360)
        with right:
            st.caption('HH registradas por profesional')
            chart=frame.groupby('Profesional',as_index=True)['HH_cargadas'].sum().sort_values(ascending=False).head(12)
            st.bar_chart(chart,horizontal=True,color='#12a398',height=340)
        export_xlsx(frame,'altivia_hh_por_profesional.xlsx','Resumen')
    else:empty('No hay registros para el período','Ajusta las fechas o el proyecto seleccionado.','📊')
    with st.expander('☷ Detalle por actividad y fecha'):
        detail=pd.DataFrame([dict(Profesional=r['worker'],Proyecto=r['project'],EDT=r['activity_code'],Actividad=r['activity'],Fecha=r['work_date'],HH=r['quarters']/4,Estado=r['state']) for r in entries])
        table_data(detail,height=400)
        export_xlsx(detail,'altivia_hh_detalle.xlsx','Detalle')

elif page=='Aprobaciones':
    hero('Bandeja de aprobaciones','Revisa primero el detalle técnico de cada registro y luego aprueba, observa o rechaza.','SUPERVISIÓN · FLUJO DE REVISIÓN')
    pending=[r for r in entries_query() if r['state']=='pendiente' and r['user_id']!=u['id']]
    a,b,c=st.columns(3)
    a.metric('Registros por revisar',len(pending))
    b.metric('HH pendientes',f"{sum(r['quarters'] for r in pending)/4:,.2f}")
    c.metric('Profesionales en espera',len({r['user_id'] for r in pending}))
    if not pending:empty('Bandeja al día','No hay HH pendientes que requieran tu aprobación.','✅')
    else:
        section('Cola de revisión')
        projlist=sorted({r['project'] for r in pending})
        x1,x2=st.columns([1,1.6])
        with x1:project_filter=st.selectbox('Filtrar por proyecto',['Todos']+projlist,key='approval_project')
        with x2:search=st.text_input('Buscar profesional o actividad',placeholder='Nombre o código EDT...',key='approval_search')
        pending_view=[r for r in pending if (project_filter=='Todos' or r['project']==project_filter) and (not search or norm(search) in norm(r['worker']+' '+r['activity']+' '+r['activity_code']))]
        table_data(pd.DataFrame([dict(ID=r['id'],Profesional=r['worker'],Proyecto=r['project'],EDT=r['activity_code'],Fecha=r['work_date'],HH=r['quarters']/4,Actividad=r['activity']) for r in pending_view]),height=305)
        if pending_view:
            section('Resolver un registro')
            with st.form('approve'):
                eid=st.selectbox('Registro para revisar',[r['id'] for r in pending_view],format_func=lambda i:next(f"#{r['id']} · {r['worker']} · {r['quarters']/4:.2f} HH" for r in pending_view if r['id']==i))
                selected=next(r for r in pending_view if r['id']==eid)
                st.markdown('**Proyecto:** '+safe(selected['project'])+'　 · 　**EDT:** '+safe(selected['activity_code'])+'　 · 　**Fecha:** '+safe(selected['work_date']))
                st.info(selected['description'])
                decision=st.radio('Decisión',['Aprobar','Observar','Rechazar'],horizontal=True,help='Observar permite al profesional corregir y reenviar. Rechazar finaliza la revisión.')
                note=st.text_area('Observación / motivo',placeholder='Obligatorio cuando observas o rechazas...')
                if st.form_submit_button('✓ Confirmar decisión',type='primary',use_container_width=True):
                    if decision!='Aprobar' and not note.strip():st.error('Debes ingresar el motivo.')
                    else:
                        state={'Aprobar':'aprobado','Observar':'observado','Rechazar':'rechazado'}[decision]
                        try:
                            with engine().begin() as c:
                                rr=c.execute(text('''UPDATE entries SET state=:s,note=:note,approval_by=:reviewer,approval_at=CURRENT_TIMESTAMP,updated_at=CURRENT_TIMESTAMP
                                    WHERE id=:eid AND state='pendiente' AND user_id<>:reviewer AND EXISTS (
                                        SELECT 1 FROM activities a JOIN projects p ON p.id=a.project_id
                                        WHERE a.id=entries.activity_id AND (:is_admin=1 OR p.supervisor_id=:reviewer))'''),
                                    dict(s=state,note=note,reviewer=u['id'],eid=eid,is_admin=int(u['role']=='admin')))
                                if rr.rowcount!=1:raise ValueError('No puedes aprobar este registro.')
                                audit_event(c,'revision_'+state,'entry',eid,note)
                            st.success('Decisión guardada correctamente.');st.rerun()
                        except Exception as exc:st.error(str(exc))

elif page=='Presupuestos EDT':
    hero('Control de presupuestos EDT','Compara el presupuesto de cada actividad con las HH registradas y detecta desviaciones.','PRESUPUESTOS · EDT/WBS')
    pid=choose('Proyecto a consultar',projects_access())
    rows=all_rows('''SELECT a.code,a.name,a.budget_quarters,COALESCE(SUM(CASE WHEN e.state IN ('pendiente','aprobado','liquidado') THEN e.quarters ELSE 0 END),0) consumed
      FROM activities a LEFT JOIN entries e ON e.activity_id=a.id WHERE a.project_id=:pid GROUP BY a.id ORDER BY a.code''',{'pid':pid}) if pid else []
    if rows:
        b=sum(r['budget_quarters'] for r in rows)/4;t=sum(r['consumed'] for r in rows)/4
        a,c,d=st.columns(3)
        a.metric('Presupuesto total',f'{b:,.2f} HH')
        c.metric('HH consumidas',f'{t:,.2f} HH')
        d.metric('Saldo disponible',f'{b-t:,.2f} HH')
        mini('Utilización del presupuesto',f'{t/b*100:,.1f}%' if b else 'Sin presupuesto','Consumo acumulado, sin filtrar fechas.',100*t/b if b else 0)
        frame=pd.DataFrame([dict(EDT=r['code'],Actividad=r['name'],Presupuesto_HH=r['budget_quarters']/4,HH_cargadas=r['consumed']/4,Saldo_HH=(r['budget_quarters']-r['consumed'])/4,Avance_pct=round(100*r['consumed']/r['budget_quarters'],1) if r['budget_quarters'] else 0) for r in rows])
        section('Detalle por actividad')
        only_over=st.checkbox('Mostrar solo actividades sin saldo / excedidas',key='overruns')
        visible=frame[frame['Saldo_HH']<=0] if only_over else frame
        table_data(visible, height=440, column_config={'Avance_pct':st.column_config.NumberColumn('Avance (%)',format='%.1f %%')})
        export_xlsx(frame,'altivia_presupuestos_edt.xlsx','Presupuestos')
    else:empty('Este proyecto no tiene actividades','Agrégalas desde Proyectos y actividades o mediante Importar Excel.','▥')

elif page=='Proyectos y actividades':
    hero('Proyectos y actividades','Crea proyectos, designa responsables y administra el presupuesto de cada EDT.','CONFIGURACIÓN DE PROYECTOS')
    section('Alta y edición de proyectos')
    with st.expander('Nuevo proyecto'):
        with st.form('newproj'):
            code=st.text_input('Código (único)');name=st.text_input('Nombre');supervisors=all_rows("SELECT id,name FROM users WHERE role IN ('supervisor','admin') AND active=1 ORDER BY name")
            sup=choose('Supervisor responsable',supervisors,empty=True);save=st.form_submit_button('Crear proyecto')
        if save:
            try:
                with engine().begin() as c:
                    iid=c.execute(text("INSERT INTO projects(code,name,supervisor_id) VALUES(:code,:name,:sup) RETURNING id"),dict(code=code.strip(),name=name.strip(),sup=sup)).scalar_one()
                    audit_event(c,'crear_proyecto','project',iid)
                st.rerun()
            except Exception as e:st.error(str(e))
    projs=projects_access();pid=choose('Selecciona proyecto',projs)
    if pid:
        p=next(x for x in projs if x['id']==pid)
        details_tab,activities_tab=st.tabs(['▣  Datos del proyecto','▥  Actividades EDT'])
        with details_tab:
            with st.form('change-project'):
                code=st.text_input('Código',value=p['code']);name=st.text_input('Nombre',value=p['name']);status=st.selectbox('Estado',['activo','desactivado','terminado'],index=['activo','desactivado','terminado'].index(p['status']))
                sup_opts={**{'Sin asignar':None}, **options(all_rows("SELECT id,name FROM users WHERE role IN ('supervisor','admin') AND active=1"))}
                previous=next((label for label,value in sup_opts.items() if value==p['supervisor_id']), 'Sin asignar')
                sup_label=st.selectbox('Supervisor',list(sup_opts),index=list(sup_opts).index(previous))
                sup=sup_opts[sup_label]
                if st.form_submit_button('Guardar proyecto'):
                    with engine().begin() as c:
                        c.execute(text('UPDATE projects SET code=:cd,name=:n,status=:s,supervisor_id=:sup WHERE id=:pid'),dict(cd=code.strip(),n=name.strip(),s=status,sup=sup,pid=pid))
                        audit_event(c,'editar_proyecto','project',pid,'Estado '+status)
                    st.rerun()
        with activities_tab:
            st.info('Los proyectos desactivados o terminados bloquean la carga y corrección de HH, pero conservan el historial.')
            acts=all_rows('SELECT id,code,name,budget_quarters,active FROM activities WHERE project_id=:pid ORDER BY code',{'pid':pid})
            section('Actividades EDT · edición rápida')
            if acts:
                df=pd.DataFrame([dict(ID=r['id'],EDT=r['code'],Actividad=r['name'],Presupuesto_HH=r['budget_quarters']/4,Activa=bool(r['active'])) for r in acts])
                edited=st.data_editor(df,hide_index=True,use_container_width=True,disabled=['ID'],num_rows='fixed',key=f'activities_edit_{pid}')
                if st.button('Guardar modificaciones de actividades',type='primary'):
                    try:
                        with engine().begin() as c:
                            for _,r in edited.iterrows():
                                b=Decimal(str(r['Presupuesto_HH']))
                                if b<0 or b*4!=(b*4).to_integral_value():raise ValueError('El presupuesto debe ser múltiplo de 0.25 HH.')
                                c.execute(text('UPDATE activities SET code=:code,name=:name,budget_quarters=:q,active=:active WHERE id=:id AND project_id=:pid'),dict(code=str(r['EDT']).strip(),name=str(r['Actividad']).strip(),q=int(b*4),active=int(bool(r['Activa'])),id=int(r['ID']),pid=pid))
                            audit_event(c,'edicion_masiva_actividades','project',pid)
                        st.success('Actividades actualizadas.');st.rerun()
                    except Exception as exc:st.error(str(exc))
            with st.expander('Agregar actividad'):
                with st.form('newact'):
                    cd=st.text_input('EDT/WBS');nm=st.text_input('Actividad');budget=st.number_input('HH presupuesto',min_value=0.,step=.25)
                    if st.form_submit_button('Agregar'):
                        try:
                            with engine().begin() as c:
                                aid=c.execute(text('INSERT INTO activities(project_id,code,name,budget_quarters) VALUES(:pid,:code,:name,:b) RETURNING id'),dict(pid=pid,code=cd.strip(),name=nm.strip(),b=int(Decimal(str(budget))*4))).scalar_one()
                                audit_event(c,'crear_actividad','activity',aid)
                            st.rerun()
                        except Exception as exc:st.error(str(exc))

elif page=='Usuarios y tarifas':
    hero('Equipo, roles y tarifas','Administra accesos, define tarifas y asigna profesionales a los proyectos.','ADMINISTRACIÓN DEL EQUIPO')
    team_tab,rates_tab,assignment_tab=st.tabs(['♧  Directorio','S/  Tarifas','▦  Asignaciones'])
    with team_tab:
        section('Directorio de usuarios')
        st.caption('Edita el directorio con cuidado: los roles determinan los módulos accesibles. Las tarifas de liquidaciones ya cerradas no cambian.')
        users=all_rows('SELECT id,username,name,role,active,must_change_pw FROM users ORDER BY name')
        df=pd.DataFrame([dict(ID=x['id'],Usuario=x['username'],Nombre=x['name'],Rol=x['role'],Activo=bool(x['active'])) for x in users])
        edited=st.data_editor(df,hide_index=True,use_container_width=True,disabled=['ID'],column_config={'Rol':st.column_config.SelectboxColumn(options=['admin','supervisor','trabajador'])},key='useredit')
        if st.button('Guardar cambios de usuarios',type='primary'):
            try:
                with engine().begin() as c:
                    active_admins=0
                    for _,row in edited.iterrows():
                        if row['Rol']=='admin' and row['Activo']:active_admins+=1
                    if active_admins<1:raise ValueError('Debe existir al menos un administrador activo.')
                    for _,row in edited.iterrows():
                        c.execute(text('UPDATE users SET username=:login,name=:name,role=:role,active=:active WHERE id=:id'),dict(login=str(row['Usuario']).lower().strip(),name=str(row['Nombre']).strip(),role=row['Rol'],active=int(row['Activo']),id=int(row['ID'])))
                    audit_event(c,'edicion_usuarios','users')
                st.success('Usuarios actualizados.');st.rerun()
            except Exception as exc:st.error(str(exc))
        with st.expander('Crear usuario manualmente'):
            with st.form('newusr'):
                name=st.text_input('Nombre completo');login=st.text_input('Usuario (opcional)');role=st.selectbox('Rol',['trabajador','supervisor','admin']);pwd=st.text_input('Contraseña temporal (mín. 12 caracteres)',type='password')
                if st.form_submit_button('Crear'):
                    if not name.strip():st.error('Ingresa el nombre completo.')
                    elif len(pwd)<12:st.error('La contraseña temporal debe tener al menos 12 caracteres.')
                    else:
                        try:
                            with engine().begin() as c:
                                uid=c.execute(text('INSERT INTO users(username,name,pw,role,must_change_pw) VALUES(:login,:name,:pw,:role,1) RETURNING id'),dict(login=(login.strip().lower() or login_from_name(name)),name=name.strip(),pw=generate_password_hash(pwd),role=role)).scalar_one()
                                audit_event(c,'crear_usuario','user',uid)
                            st.rerun()
                        except Exception as exc:st.error(str(exc))
    with rates_tab:
        section('Tarifas estándar por HH')
        rates=all_rows('SELECT u.id,u.name,r.rate_cents FROM users u LEFT JOIN rates r ON r.user_id=u.id WHERE u.active=1 ORDER BY u.name')
        rdf=pd.DataFrame([dict(ID=r['id'],Trabajador=r['name'],Tarifa_Soles=round((r['rate_cents'] or 0)/100,2)) for r in rates])
        edited_rates=st.data_editor(rdf,hide_index=True,use_container_width=True,disabled=['ID','Trabajador'],key='rates_edit')
        if st.button('Guardar tarifas',type='primary'):
            try:
                with engine().begin() as c:
                    for _,r in edited_rates.iterrows():
                        cents=int((Decimal(str(r['Tarifa_Soles']))*100).quantize(Decimal('1')))
                        if cents<0:raise ValueError('No se aceptan tarifas negativas.')
                        c.execute(text('INSERT INTO rates(user_id,rate_cents) VALUES(:u,:rate) ON CONFLICT (user_id) DO UPDATE SET rate_cents=EXCLUDED.rate_cents'),dict(u=int(r['ID']),rate=cents))
                    audit_event(c,'actualizar_tarifas','rates')
                st.success('Tarifas actualizadas.');st.rerun()
            except Exception as e:st.error(str(e))
    with assignment_tab:
        section('Asignar profesionales a actividades')
        with st.expander('Asignar trabajador a proyecto o actividad'):
            people=all_rows('SELECT id,name FROM users WHERE active=1 ORDER BY name');person=choose('Trabajador',people)
            pid=choose('Proyecto',projects_access());aa=all_rows('SELECT id,code,name FROM activities WHERE project_id=:pid ORDER BY code',{'pid':pid}) if pid else []
            selected=st.multiselect('Actividades autorizadas',list(options(aa)),key='assign_act')
            if st.button('Agregar asignaciones'):
                with engine().begin() as c:
                    c.execute(text('INSERT INTO assignments(user_id,project_id) VALUES(:u,:p) ON CONFLICT DO NOTHING'),dict(u=person,p=pid))
                    for label in selected:c.execute(text('INSERT INTO activity_assignments(user_id,activity_id) VALUES(:u,:a) ON CONFLICT DO NOTHING'),dict(u=person,a=options(aa)[label]))
                    audit_event(c,'asignar_actividades','user',person)
                st.success('Asignaciones registradas.')

elif page=='Importar Excel':
    hero('Importar matriz de presupuestos','Carga tu matriz de estimación de costos y revisa las actividades detectadas antes de confirmar.','IMPORTACIÓN · EXCEL')
    section('Paso 1 · Proyecto y archivo')
    st.info('Formato esperado: hoja «Estimado de Costos» · EDT en columna B · actividad en C · profesionales en fila 5, columnas J:T · HH positivas por cruce.')
    pid=choose('Proyecto destino',projects_access())
    file=st.file_uploader('Seleccionar libro .xlsx',type=['xlsx'])
    if file and pid:
        try:
            wb=load_workbook(io.BytesIO(file.getvalue()),read_only=True,data_only=True)
            if 'Estimado de Costos' not in wb.sheetnames:raise ValueError('No existe la hoja Estimado de Costos.')
            ws=wb['Estimado de Costos'];persons=[(col,str(ws.cell(5,col).value).strip()) for col in range(10,21) if ws.cell(5,col).value and norm(ws.cell(5,col).value) not in ('tbd','por definir')]
            preview=[]
            for row in ws.iter_rows(min_row=10,values_only=True):
                code=row[1] if len(row)>2 else None;title=row[2] if len(row)>2 else None
                if isinstance(code,bool) or not isinstance(code,(int,float)) or not title:continue
                code=str(int(code)) if float(code).is_integer() else str(code)
                if code.endswith('000'):continue
                cross=[]
                for col,fullname in persons:
                    val=row[col-1] if len(row)>=col else None
                    if isinstance(val,(int,float)) and not isinstance(val,bool) and val>0:
                        q=Decimal(str(val))*4
                        if q==q.to_integral_value():cross.append((fullname,int(q)))
                if cross:preview.append(dict(code=code,name=str(title).strip(),cross=cross,budget=sum(q for _,q in cross)))
            st.success(f'Vista previa preparada: {len(persons)} profesionales y {len(preview)} actividades identificadas.')
            section('Paso 2 · Validar datos detectados')
            table_data(pd.DataFrame([dict(EDT=x['code'],Actividad=x['name'],Presupuesto_HH=x['budget']/4,Profesionales=', '.join(n for n,_ in x['cross'])) for x in preview]))
            st.warning('Antes de confirmar: la importación reemplazará las asignaciones de las actividades incluidas. No sustituirá la base completa. Comprueba el proyecto de destino.')
            section('Paso 3 · Confirmar importación')
            if st.button('Confirmar importación de matriz',type='primary'):
                with engine().begin() as c:
                    mapping={}
                    for _,fullname in persons:
                        existing=c.execute(text('SELECT id FROM users WHERE lower(name)=lower(:name)'),{'name':fullname}).scalar()
                        if existing:mapping[fullname]=existing;continue
                        root=login_from_name(fullname);login=root;seq=2
                        while c.execute(text('SELECT 1 FROM users WHERE username=:login'),{'login':login}).first():login=f'{root}{seq}';seq+=1
                        # Random invitation password prevents predictable name-based credentials in public deployment.
                        initial=secrets.token_urlsafe(18)
                        uid=c.execute(text('INSERT INTO users(username,name,pw,role,must_change_pw) VALUES(:login,:name,:pw,\'trabajador\',1) RETURNING id'),dict(login=login,name=fullname,pw=generate_password_hash(initial))).scalar_one()
                        mapping[fullname]=uid
                        st.session_state.setdefault('new_import_credentials',[]).append((fullname,login,initial))
                    for item in preview:
                        existing=c.execute(text('SELECT id FROM activities WHERE project_id=:pid AND code=:code'),dict(pid=pid,code=item['code'])).scalar()
                        if existing:
                            aid=existing
                            if c.execute(text("SELECT 1 FROM entries WHERE activity_id=:a AND state='liquidado' LIMIT 1"),{'a':aid}).first():raise ValueError(f"No se puede reemplazar actividad {item['code']} con HH liquidadas.")
                            c.execute(text('UPDATE activities SET name=:name,budget_quarters=:b,active=1 WHERE id=:id'),dict(name=item['name'],b=item['budget'],id=aid))
                            c.execute(text('DELETE FROM activity_assignments WHERE activity_id=:aid'),{'aid':aid})
                        else:aid=c.execute(text('INSERT INTO activities(project_id,code,name,budget_quarters) VALUES(:pid,:code,:name,:b) RETURNING id'),dict(pid=pid,code=item['code'],name=item['name'],b=item['budget'])).scalar_one()
                        for fullname,quarters in item['cross']:
                            uid=mapping[fullname]
                            c.execute(text('INSERT INTO assignments(user_id,project_id) VALUES(:uid,:pid) ON CONFLICT DO NOTHING'),dict(uid=uid,pid=pid))
                            c.execute(text('INSERT INTO activity_assignments(user_id,activity_id,planned_quarters) VALUES(:uid,:aid,:q) ON CONFLICT (user_id,activity_id) DO UPDATE SET planned_quarters=EXCLUDED.planned_quarters'),dict(uid=uid,aid=aid,q=quarters))
                    c.execute(text('INSERT INTO import_log(project_id,filename,actor_id,notes) VALUES(:pid,:file,:actor,:notes)'),dict(pid=pid,file=file.name,actor=u['id'],notes=f'{len(preview)} actividades'))
                    audit_event(c,'importar_excel','project',pid)
                st.success('Importación completada.');st.rerun()
            credentials=st.session_state.get('new_import_credentials',[])
            if credentials:
                st.warning('Credenciales iniciales aleatorias, mostradas una sola vez. Entrega cada contraseña privadamente.')
                st.dataframe(pd.DataFrame(credentials,columns=['Nombre','Usuario','Contraseña temporal']),hide_index=True)
                if st.button('Ocultar credenciales'):st.session_state.pop('new_import_credentials',None);st.rerun()
        except Exception as e:st.error(str(e))

elif page=='Liquidaciones':
    hero('Liquidaciones de HH','Prepara cierres con fecha de corte, revisa los importes y consulta el historial de pagos calculados.','CONTROL DE LIQUIDACIONES')
    st.info('Las HH deben estar aprobadas antes o en la fecha de corte. Cada registro se liquida una sola vez y la tarifa aplicada queda congelada.')
    cutoff=st.date_input('📅 Fecha de corte',date.today(),help='Solo se incluyen trabajos realizados y aprobados hasta esta fecha.',key='liquidation_cutoff')
    candidates=all_rows('''SELECT us.id,us.name,COUNT(e.id) records,COALESCE(SUM(e.quarters),0) quarters,r.rate_cents
        FROM users us JOIN entries e ON e.user_id=us.id LEFT JOIN rates r ON r.user_id=us.id
        WHERE e.state='aprobado' AND e.liquidation_id IS NULL AND e.work_date<=:cutoff AND date(e.approval_at)<=:cutoff
        GROUP BY us.id,r.rate_cents ORDER BY us.name''',{'cutoff':cutoff})
    closed=all_rows('''SELECT l.id,us.name,l.period,l.cutoff_date,l.quarters,l.cents,l.created_at
        FROM liquidations l JOIN users us ON us.id=l.user_id ORDER BY l.id DESC LIMIT 500''')
    x1,x2,x3,x4=st.columns(4)
    x1.metric('Profesionales listos',len(candidates))
    x2.metric('HH listas',f"{sum(r['quarters'] for r in candidates)/4:,.2f}")
    x3.metric('Sin tarifa',len([r for r in candidates if r['rate_cents'] is None]))
    x4.metric('Liquidaciones cerradas',len(closed))
    ready_tab,history_tab=st.tabs(['◈  Preparar liquidación','☷  Historial y exportaciones'])
    with ready_tab:
        section('Registros elegibles al corte')
        if not candidates:empty('No hay HH aprobadas por liquidar','Puedes cambiar la fecha de corte para consultar otros registros.','✅')
        else:
            # Cálculo por cada registro, idéntico al utilizado al guardar las líneas de liquidación.
            eligible_parts=all_rows('''SELECT user_id,quarters FROM entries WHERE state='aprobado' AND liquidation_id IS NULL
                AND work_date<=:cutoff AND date(approval_at)<=:cutoff ORDER BY id''',{'cutoff':cutoff})
            parts_by_user={}
            for part in eligible_parts:parts_by_user.setdefault(part['user_id'],[]).append(part)
            estimates=[]
            for x in candidates:
                parts=parts_by_user.get(x['id'],[])
                exact=sum(cents_for_quarters(r['quarters'],x['rate_cents']) for r in parts) if x['rate_cents'] is not None else None
                estimates.append(dict(Profesional=x['name'],Registros=x['records'],HH=x['quarters']/4,Tarifa=money(x['rate_cents']) if x['rate_cents'] is not None else 'SIN TARIFA',Importe=money(exact) if exact is not None else 'Configurar tarifa'))
            table_data(pd.DataFrame(estimates),height=325)
            section('Seleccionar profesional')
            uid=choose('Trabajador a liquidar',candidates)
            chosen=next(x for x in candidates if x['id']==uid)
            records_preview=all_rows('''SELECT e.id,e.work_date,e.quarters,e.description,p.code project,a.code activity
                FROM entries e JOIN activities a ON a.id=e.activity_id JOIN projects p ON p.id=a.project_id
                WHERE e.user_id=:uid AND e.state='aprobado' AND e.liquidation_id IS NULL
                  AND e.work_date<=:cutoff AND date(e.approval_at)<=:cutoff ORDER BY e.work_date,e.id''',dict(uid=uid,cutoff=cutoff))
            if chosen['rate_cents'] is None:st.error('Este trabajador no tiene tarifa configurada. Ve a Usuarios y tarifas antes de liquidar.')
            else:
                total_cents=sum(cents_for_quarters(r['quarters'],chosen['rate_cents']) for r in records_preview)
                a,b,c=st.columns(3)
                a.metric('Registros del cierre',len(records_preview))
                b.metric('HH del cierre',f"{sum(r['quarters'] for r in records_preview)/4:,.2f}")
                c.metric('Importe a liquidar',money(total_cents))
                with st.expander('☷ Ver las HH incluidas en este cierre',expanded=True):
                    table_data(pd.DataFrame([dict(ID=r['id'],Fecha=r['work_date'],Proyecto=r['project'],EDT=r['activity'],HH=r['quarters']/4,Descripcion=r['description']) for r in records_preview]),height=260)
                with st.form('confirm_liquidation'):
                    st.warning('El cierre es definitivo desde esta interfaz. Los registros quedan marcados como liquidados y no volverán a incluirse en otro cierre.')
                    consent=st.checkbox(f"Confirmo el corte {cutoff:%d/%m/%Y} para {chosen['name']} por {money(total_cents)}")
                    submit=st.form_submit_button('✓ Cerrar liquidación',type='primary',disabled=not bool(records_preview),use_container_width=True)
                if submit:
                    if not consent:st.error('Debes confirmar el trabajador, la fecha y el importe antes de cerrar.')
                    else:
                        try:
                            with engine().begin() as c:
                                rate=c.execute(text('SELECT rate_cents FROM rates WHERE user_id=:uid'),{'uid':uid}).scalar()
                                if rate is None:raise ValueError('Debes configurar la tarifa del trabajador.')
                                records=c.execute(text('''SELECT id,quarters FROM entries WHERE user_id=:uid AND state='aprobado' AND liquidation_id IS NULL
                                  AND work_date<=:cutoff AND date(approval_at)<=:cutoff'''),dict(uid=uid,cutoff=cutoff)).mappings().all()
                                if not records:raise ValueError('No hay HH pendientes para este corte.')
                                total_q=sum(r['quarters'] for r in records);amount=sum(cents_for_quarters(r['quarters'], rate) for r in records)
                                lid=c.execute(text('''INSERT INTO liquidations(user_id,period,cutoff_date,cents,quarters,created_by)
                                    VALUES(:uid,:p,:cutoff,:amount,:q,:actor) RETURNING id'''),dict(uid=uid,p=cutoff.strftime('%Y-%m'),cutoff=cutoff,amount=amount,q=total_q,actor=u['id'])).scalar_one()
                                for record in records:
                                    cents=cents_for_quarters(record['quarters'],rate)
                                    c.execute(text('INSERT INTO liquidation_lines(liquidation_id,entry_id,quarters,rate_cents,amount_cents) VALUES(:lid,:eid,:q,:rate,:cents)'),dict(lid=lid,eid=record['id'],q=record['quarters'],rate=rate,cents=cents))
                                    c.execute(text("UPDATE entries SET state='liquidado',liquidation_id=:lid WHERE id=:id"),dict(lid=lid,id=record['id']))
                                audit_event(c,'cerrar_liquidacion','liquidation',lid,f'Corte {cutoff}')
                            st.success(f'Liquidación #{lid} cerrada: {money(amount)}');st.rerun()
                        except Exception as exc:st.error(str(exc))
    with history_tab:
        section('Liquidaciones realizadas')
        df=pd.DataFrame([dict(ID=r['id'],Trabajador=r['name'],Periodo=r['period'],Corte=r['cutoff_date'],HH=r['quarters']/4,Total_Soles=r['cents']/100,Creada=r['created_at']) for r in closed])
        if df.empty:empty('Todavía no hay liquidaciones cerradas','Los cierres confirmados aparecerán aquí.','🧾')
        else:
            c1,c2=st.columns(2)
            with c1:person_search=st.text_input('Buscar profesional',key='liquidation_search')
            with c2:period_filter=st.selectbox('Período',['Todos']+sorted(df['Periodo'].dropna().unique().tolist(),reverse=True),key='liquidation_period')
            filt=df.copy()
            if person_search:filt=filt[filt['Trabajador'].str.contains(person_search,case=False,regex=False,na=False)]
            if period_filter!='Todos':filt=filt[filt['Periodo']==period_filter]
            table_data(filt,height=380)
            if not filt.empty:
                export_xlsx(filt,'altivia_liquidaciones.xlsx','Liquidaciones')
                from reportlab.pdfgen import canvas
                from reportlab.lib.pagesizes import A4
                pdf=io.BytesIO();cv=canvas.Canvas(pdf,pagesize=A4)
                cv.setFont('Helvetica-Bold',14);cv.drawString(42,800,'ALTIVIA - Liquidaciones cerradas');cv.setFont('Helvetica',9)
                y=773
                for _,r in filt.iterrows():
                    if y<50:cv.showPage();cv.setFont('Helvetica',9);y=800
                    line=f"#{r['ID']}  {str(r['Trabajador'])[:30]}   {r['HH']:.2f} HH    S/ {r['Total_Soles']:,.2f}"
                    cv.drawString(42,y,line.encode('latin-1','replace').decode('latin-1'));y-=13
                cv.save()
                st.download_button('⬇ Exportar resumen PDF',data=pdf.getvalue(),file_name='altivia_liquidaciones.pdf',mime='application/pdf')

elif page=='Auditoría':
    hero('Registro de auditoría','Historial de operaciones para facilitar el seguimiento y la trazabilidad.','SEGURIDAD · TRAZABILIDAD')
    rows=all_rows('''SELECT a.id,a.actor_id,COALESCE(u.name,'Sistema / eliminado') actor,a.action,a.target,a.target_id,a.detail,a.at
        FROM audit a LEFT JOIN users u ON u.id=a.actor_id ORDER BY a.id DESC LIMIT 1000''')
    section('Eventos recientes')
    a,b=st.columns(2)
    a.metric('Eventos consultados',len(rows))
    b.metric('Actores distintos',len({r['actor_id'] for r in rows}))
    if rows:
        x,y=st.columns([1,1.4])
        with x:action=st.selectbox('Tipo de acción',['Todas']+sorted({r['action'] for r in rows}),key='audit_action')
        with y:who=st.text_input('Buscar actor o detalle',key='audit_search')
        records=[r for r in rows if (action=='Todas' or r['action']==action) and (not who or norm(who) in norm(str(r['actor'])+' '+str(r['detail'])))]
        df=pd.DataFrame([dict(ID=r['id'],Fecha=r['at'],Usuario=r['actor'],Accion=r['action'],Entidad=r['target'],ID_Entidad=r['target_id'],Detalle=r['detail']) for r in records])
        table_data(df,height=500)
        export_xlsx(df,'altivia_auditoria.xlsx','Auditoria')
    else:empty('Sin actividad registrada','Las operaciones auditadas aparecerán aquí.','◎')

elif page=='Respaldo y restauración':
    hero('Protección de tus datos','Descarga copias de seguridad y restáuralas cuando sea necesario. Solo administradores.','ADMINISTRACIÓN · DATOS')
    st.warning('Streamlit Community Cloud utiliza almacenamiento temporal. Si se reinicia o reconstruye el servidor, podrías perder la base SQLite. Este módulo no sustituye una base de datos persistente.')
    backup_tab,restore_tab=st.tabs(['⬇  Crear respaldo','↻  Restaurar respaldo'])
    with backup_tab:
        section('Copia completa de la plataforma')
        mini('Contenido incluido','SQLite · base completa','Usuarios, proyectos, asignaciones, HH, tarifas, aprobaciones, liquidaciones y auditoría. El archivo contiene información confidencial.')
        st.download_button('⬇ Descargar respaldo (.db)',data=sqlite_backup_bytes(),file_name=f"altivia_hh_backup_{datetime.now():%Y%m%d_%H%M%S}.db",mime='application/octet-stream',type='primary')
        st.caption('Guarda el respaldo en una ubicación privada. Nunca lo subas a un repositorio público de GitHub.')
    with restore_tab:
        section('Restauración controlada')
        st.error('Restaurar sobrescribe toda la base actual, incluso usuarios y contraseñas. Primero descarga un respaldo del estado actual y asegúrate de que los demás usuarios hayan cerrado sesión.')
        upload=st.file_uploader('Selecciona un respaldo .db / .sqlite',type=['db','sqlite','sqlite3'])
        confirmation=st.checkbox('Entiendo que se reemplazarán TODOS los registros y cuentas actuales.')
        if st.button('↻ Restaurar respaldo',disabled=not (upload and confirmation),type='primary'):
            try:
                if upload.size>50*1024*1024:raise ValueError('El respaldo supera 50 MB; restaura localmente o contacta soporte.')
                restore_sqlite_bytes(upload.getvalue())
                st.session_state.clear()
                st.success('Respaldo restaurado. Inicia sesión con las credenciales de la copia restaurada.')
                st.rerun()
            except Exception as exc:st.error(f'No se pudo restaurar: {exc}')
