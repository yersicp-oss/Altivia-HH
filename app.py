"""ALTIVIA HH CONTROL — Streamlit + SQLite local (prototipo sin servidor de BD).
Subir app.py y requirements.txt a GitHub, seleccionar app.py en Streamlit Cloud.
La base de datos creada en Streamlit Cloud es EPHEMERAL: usar respaldos.
"""
import io, re, secrets, unicodedata, sqlite3, tempfile, os
from pathlib import Path
from contextlib import closing
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, event, text
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook

st.set_page_config(page_title='ALTIVIA HH CONTROL',page_icon='📊',layout='wide',initial_sidebar_state='expanded')


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

st.markdown('''<style>.block-container{padding-top:1.4rem;max-width:1450px}div[data-testid="stMetric"]{background:#fff;padding:16px;border:1px solid #e2ebef;border-radius:14px}h1,h2,h3{color:#11384b}section[data-testid="stSidebar"]{background:#11384b}section[data-testid="stSidebar"] *{color:#f5f9fc!important}section[data-testid="stSidebar"] button{color:#17364a!important}</style>''',unsafe_allow_html=True)
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
        st.title('Configuración inicial · ALTIVIA')
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
        st.title('ALTIVIA HH CONTROL')
        st.caption('Acceso privado · gestión de horas hombre')
        with st.form('login'):
            login=st.text_input('Usuario').strip().lower();pwd=st.text_input('Contraseña',type='password');submitted=st.form_submit_button('Ingresar',type='primary')
        if submitted:
            u=one('SELECT * FROM users WHERE username=:u AND active=1',{'u':login})
            if u and check_password_hash(u['pw'],pwd):
                st.session_state['uid']=u['id'];st.rerun()
            else:st.error('Usuario o contraseña incorrectos.')
        st.stop()
    if me()['must_change_pw']:
        st.title('Cambiar contraseña temporal')
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
u=me();st.sidebar.title('ALTIVIA · HH')
st.sidebar.caption(f"{u['name']} · {u['role'].capitalize()}")
st.sidebar.caption('Motor: Python + SQLite')
st.sidebar.caption('⚠️ En Streamlit Cloud: respaldar datos antes de reinicios')
menus=['Centro gerencial','Mis HH']
if u['role'] in ('admin','supervisor'):menus+=['HH por profesional','Aprobaciones','Presupuestos EDT']
if u['role']=='admin':menus+=['Proyectos y actividades','Usuarios y tarifas','Importar Excel','Liquidaciones','Auditoría','Respaldo y restauración']
page=st.sidebar.radio('Navegación',menus)
if st.sidebar.button('Cerrar sesión'):
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
    st.title('Centro gerencial')
    st.caption('Filtros por proyecto y fecha. El presupuesto se muestra completo; el consumo corresponde al período seleccionado.')
    pid,start,end=date_filters()
    projs=[p for p in projects_access() if not pid or p['id']==pid]
    entries=[r for r in entries_query(pid,start,end) if r['state'] in ('pendiente','aprobado','liquidado')]
    budgets={r['project_id']:r['qty'] for r in all_rows('SELECT project_id,COALESCE(SUM(budget_quarters),0) qty FROM activities GROUP BY project_id')}
    budget=sum(budgets.get(p['id'],0) for p in projs)/4
    total=sum(e['quarters'] for e in entries)/4
    approved=sum(e['quarters'] for e in entries if e['state'] in ('aprobado','liquidado'))/4
    pending=sum(e['quarters'] for e in entries if e['state']=='pendiente')/4
    k=st.columns(5)
    for col,label,val in zip(k,['Presupuesto total','HH cargadas (filtro)','HH aprobadas (filtro)','HH por aprobar','Diferencia presupuesto/filtro'],[budget,total,approved,pending,budget-total]):col.metric(label,f'{val:,.2f} HH')
    rows=[]
    for p in projs:
        e=[x for x in entries if x['project_id']==p['id']]
        q=sum(x['quarters'] for x in e)/4;b=budgets.get(p['id'],0)/4
        rows.append(dict(Proyecto=f"{p['code']} · {p['name']}",Estado=p['status'],Presupuesto_HH=b,HH_cargadas=q,HH_aprobadas=sum(x['quarters'] for x in e if x['state'] in ('aprobado','liquidado'))/4,HH_por_aprobar=sum(x['quarters'] for x in e if x['state']=='pendiente')/4,Diferencia_HH=b-q,Consumo_pct=round(q/b*100,1) if b else 0))
    st.subheader('Presupuesto vs. consumo por proyecto')
    st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
    if rows:st.bar_chart(pd.DataFrame(rows).set_index('Proyecto')[['Presupuesto_HH','HH_cargadas']])
    if entries:
        export=pd.DataFrame([dict(ID=e['id'],Fecha=e['work_date'],Profesional=e['worker'],Proyecto=e['project'],EDT=e['activity_code'],Actividad=e['activity'],HH=e['quarters']/4,Estado=e['state'],Descripcion=e['description']) for e in entries])
        b=io.BytesIO()
        with pd.ExcelWriter(b,engine='openpyxl') as wr:export.to_excel(wr,index=False,sheet_name='HH')
        st.download_button('Exportar detalle HH filtradas a Excel',b.getvalue(),'altivia_hh_filtradas.xlsx')

elif page=='Mis HH':
    st.title('Mis horas hombre')
    projs=projects_access(False)
    if not projs:st.info('No tienes proyectos activos asignados.')
    else:
        pid=choose('1. Selecciona el proyecto',projs)
        acts=all_rows('''SELECT a.id,a.code,a.name FROM activities a JOIN activity_assignments aa ON aa.activity_id=a.id
            WHERE aa.user_id=:uid AND a.project_id=:pid AND a.active=1 ORDER BY a.code''',dict(uid=u['id'],pid=pid)) if pid else []
        if not acts:st.warning('No tienes actividades autorizadas en este proyecto.')
        else:
            with st.form('new-entry',clear_on_submit=True):
                activity_id=choose('2. Selecciona actividad',acts)
                c1,c2=st.columns(2)
                with c1:work_date=st.date_input('Fecha de trabajo',date.today())
                with c2:hh=st.number_input('Horas hombre',min_value=0.25,max_value=12.,step=.25,value=1.)
                description=st.text_area('Descripción técnica',max_chars=1500)
                sent=st.form_submit_button('Enviar para aprobación',type='primary')
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
                    st.success('HH registradas y enviadas para aprobación.');st.rerun()
                except Exception as exc:st.error(str(exc))
    st.subheader('Historial personal')
    history=entries_query(worker=u['id'])[:250]
    st.dataframe(pd.DataFrame([dict(ID=r['id'],Fecha=r['work_date'],Proyecto=r['project'],EDT=r['activity_code'],Actividad=r['activity'],HH=r['quarters']/4,Estado=r['state'],Observacion=r['note']) for r in history]),hide_index=True,use_container_width=True)
    observed=[r for r in history if r['state']=='observado']
    if observed:
        st.subheader('Corregir registros observados')
        eid=st.selectbox('Registro observado', [r['id'] for r in observed],format_func=lambda i:f"#{i} · {next(x['activity'] for x in observed if x['id']==i)}")
        r=next(x for x in observed if x['id']==eid)
        with st.form('fix'):
            h=st.number_input('HH corregidas',value=r['quarters']/4,min_value=.25,max_value=12.,step=.25)
            desc=st.text_area('Descripción corregida',value=r['description'])
            if st.form_submit_button('Reenviar'):
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
    st.title('HH por proyecto y profesional')
    pid,start,end=date_filters();entries=entries_query(pid,start,end)
    rows={}
    for e in entries:
        if e['state'] not in ('pendiente','aprobado','liquidado'):continue
        key=(e['worker'],e['project']);r=rows.setdefault(key,dict(Profesional=e['worker'],Proyecto=e['project'],HH_cargadas=0,HH_aprobadas=0,HH_por_aprobar=0,HH_liquidadas=0))
        q=e['quarters']/4;r['HH_cargadas']+=q
        if e['state'] in ('aprobado','liquidado'):r['HH_aprobadas']+=q
        if e['state']=='pendiente':r['HH_por_aprobar']+=q
        if e['state']=='liquidado':r['HH_liquidadas']+=q
    st.dataframe(pd.DataFrame(list(rows.values())),hide_index=True,use_container_width=True)
    with st.expander('Ver detalle por actividad'):
        st.dataframe(pd.DataFrame([dict(Profesional=r['worker'],Proyecto=r['project'],EDT=r['activity_code'],Actividad=r['activity'],Fecha=r['work_date'],HH=r['quarters']/4,Estado=r['state']) for r in entries]),use_container_width=True,hide_index=True)

elif page=='Aprobaciones':
    st.title('Aprobación de HH')
    pending=[r for r in entries_query() if r['state']=='pendiente' and r['user_id']!=u['id']]
    st.dataframe(pd.DataFrame([dict(ID=r['id'],Trabajador=r['worker'],Proyecto=r['project'],Actividad=r['activity'],Fecha=r['work_date'],HH=r['quarters']/4,Descripción=r['description']) for r in pending]),hide_index=True,use_container_width=True)
    if pending:
        with st.form('approve'):
            eid=st.selectbox('Registro', [r['id'] for r in pending])
            decision=st.radio('Decisión',['Aprobar','Observar','Rechazar'],horizontal=True)
            note=st.text_area('Observación / motivo')
            if st.form_submit_button('Registrar decisión'):
                if decision!='Aprobar' and not note.strip():st.error('Debes ingresar el motivo.')
                else:
                    state={'Aprobar':'aprobado','Observar':'observado','Rechazar':'rechazado'}[decision]
                    with engine().begin() as c:
                        rr=c.execute(text('''UPDATE entries SET state=:s,note=:note,approval_by=:reviewer,approval_at=CURRENT_TIMESTAMP,updated_at=CURRENT_TIMESTAMP
                            WHERE id=:eid AND state='pendiente' AND user_id<>:reviewer AND EXISTS (
                                SELECT 1 FROM activities a JOIN projects p ON p.id=a.project_id
                                WHERE a.id=entries.activity_id AND (:is_admin=1 OR p.supervisor_id=:reviewer))'''),
                            dict(s=state,note=note,reviewer=u['id'],eid=eid,is_admin=int(u['role']=='admin')))
                        if rr.rowcount!=1:raise ValueError('No puedes aprobar este registro.')
                        audit_event(c,'revision_'+state,'entry',eid,note)
                    st.success('Revisión registrada.');st.rerun()

elif page=='Presupuestos EDT':
    st.title('Presupuestos por actividad')
    pid=choose('Proyecto',projects_access())
    rows=all_rows('''SELECT a.code,a.name,a.budget_quarters,COALESCE(SUM(CASE WHEN e.state IN ('pendiente','aprobado','liquidado') THEN e.quarters ELSE 0 END),0) consumed
      FROM activities a LEFT JOIN entries e ON e.activity_id=a.id WHERE a.project_id=:pid GROUP BY a.id ORDER BY a.code''',{'pid':pid}) if pid else []
    st.dataframe(pd.DataFrame([dict(EDT=r['code'],Actividad=r['name'],Presupuesto_HH=r['budget_quarters']/4,HH_cargadas=r['consumed']/4,Saldo_HH=(r['budget_quarters']-r['consumed'])/4) for r in rows]),hide_index=True,use_container_width=True)
elif page=='Proyectos y actividades':
    st.title('Proyectos y actividades')
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
        st.info('Los proyectos desactivados o terminados bloquean la carga y corrección de HH, pero conservan el historial.')
        acts=all_rows('SELECT id,code,name,budget_quarters,active FROM activities WHERE project_id=:pid ORDER BY code',{'pid':pid})
        st.subheader('Actividades · edición en tabla')
        if acts:
            df=pd.DataFrame([dict(ID=r['id'],EDT=r['code'],Actividad=r['name'],Presupuesto_HH=r['budget_quarters']/4,Activa=bool(r['active'])) for r in acts])
            edited=st.data_editor(df,hide_index=True,use_container_width=True,disabled=['ID'],num_rows='fixed',key='activities_edit')
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
    st.title('Usuarios y tarifas estándar')
    st.caption('Las tarifas no tienen fecha de vigencia. Los cierres anteriores conservan la tarifa congelada.')
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
    st.subheader('Tarifas estándar por HH')
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
    st.title('Importación EDT/WBS desde Excel')
    st.caption('Hoja Estimado de Costos: EDT en B; actividad C; profesionales fila 5, columnas J:T; HH positivas por cruce.')
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
            st.write(f'{len(persons)} profesionales identificados · {len(preview)} actividades asignables.')
            st.dataframe(pd.DataFrame([dict(EDT=x['code'],Actividad=x['name'],Presupuesto_HH=x['budget']/4,Profesionales=', '.join(n for n,_ in x['cross'])) for x in preview]),use_container_width=True,hide_index=True)
            st.warning('La importación sustituye asignaciones de las actividades incluidas. No sustituye la base completa. Verifica el proyecto.')
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
    st.title('Liquidaciones con fecha de corte')
    st.caption('Se incluyen HH aprobadas por fecha de trabajo y aprobadas hasta el cierre; no se vuelven a liquidar.')
    cutoff=st.date_input('Fecha de corte',date.today());workers=all_rows("SELECT id,name FROM users WHERE role IN ('trabajador','supervisor','admin') AND active=1 ORDER BY name")
    candidates=all_rows('''SELECT us.id,us.name,COUNT(e.id) records,COALESCE(SUM(e.quarters),0) quarters,r.rate_cents
        FROM users us JOIN entries e ON e.user_id=us.id LEFT JOIN rates r ON r.user_id=us.id
        WHERE e.state='aprobado' AND e.liquidation_id IS NULL AND e.work_date<=:cutoff AND date(e.approval_at)<=:cutoff
        GROUP BY us.id,r.rate_cents ORDER BY us.name''',{'cutoff':cutoff})
    st.subheader('HH listas para liquidar')
    st.dataframe(pd.DataFrame([dict(Trabajador=x['name'],Registros=x['records'],HH=x['quarters']/4,Tarifa=money(x['rate_cents']) if x['rate_cents'] is not None else 'SIN TARIFA',Importe=money(cents_for_quarters(x['quarters'],x['rate_cents'])) if x['rate_cents'] is not None else '-') for x in candidates]),hide_index=True,use_container_width=True)
    if candidates:
        uid=choose('Trabajador a liquidar',candidates)
        if st.button('Cerrar liquidación de este trabajador',type='primary'):
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
    st.subheader('Liquidaciones cerradas')
    closed=all_rows('''SELECT l.id,us.name,l.period,l.cutoff_date,l.quarters,l.cents,l.created_at
        FROM liquidations l JOIN users us ON us.id=l.user_id ORDER BY l.id DESC LIMIT 500''')
    df=pd.DataFrame([dict(ID=r['id'],Trabajador=r['name'],Periodo=r['period'],Corte=r['cutoff_date'],HH=r['quarters']/4,Total_Soles=r['cents']/100,Creada=r['created_at']) for r in closed])
    st.dataframe(df,hide_index=True,use_container_width=True)
    if not df.empty:
        buf=io.BytesIO()
        with pd.ExcelWriter(buf,engine='openpyxl') as writer:df.to_excel(writer,index=False,sheet_name='Liquidaciones')
        st.download_button('Exportar liquidaciones a Excel',data=buf.getvalue(),file_name='altivia_liquidaciones.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        pdf=io.BytesIO();cv=canvas.Canvas(pdf,pagesize=A4)
        cv.setFont('Helvetica-Bold',14);cv.drawString(42,800,'ALTIVIA - Liquidaciones cerradas');cv.setFont('Helvetica',9)
        y=773
        for _,r in df.head(55).iterrows():
            cv.drawString(42,y,f"#{r['ID']}  {str(r['Trabajador'])[:30]}   {r['HH']:.2f} HH    S/ {r['Total_Soles']:,.2f}")
            y-=13
            if y<45:cv.showPage();y=800
        cv.save()
        st.download_button('Exportar resumen a PDF',data=pdf.getvalue(),file_name='altivia_liquidaciones.pdf',mime='application/pdf')

elif page=='Auditoría':
    st.title('Auditoría')
    rows=all_rows('SELECT * FROM audit ORDER BY id DESC LIMIT 500')
    st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)


elif page=='Respaldo y restauración':
    st.title('Respaldo y restauración de datos')
    st.warning('Streamlit Community Cloud NO asegura persistencia de la base SQLite local. Antes de una suspensión, reinicio o actualización, descarga una copia. Para trabajo empresarial permanente, se necesita almacenamiento persistente externo.')
    st.subheader('Crear respaldo completo')
    st.caption('La base .db contiene usuarios, contraseñas cifradas mediante hash, proyectos, HH, aprobaciones, tarifas, liquidaciones y auditoría. Guárdala en un lugar seguro.')
    st.download_button('Descargar respaldo SQLite (.db)',data=sqlite_backup_bytes(),file_name=f"altivia_hh_backup_{datetime.now():%Y%m%d_%H%M%S}.db",mime='application/octet-stream',type='primary')
    st.subheader('Restaurar una copia')
    st.error('IMPORTANTE: restaurar SUSTITUYE los datos actuales. Realiza esta operación solamente con los demás usuarios desconectados.')
    upload=st.file_uploader('Archivo de respaldo .db',type=['db','sqlite','sqlite3'])
    confirmation=st.checkbox('Entiendo que se reemplazarán todos los registros y las cuentas actuales.')
    if st.button('Restaurar respaldo',disabled=not (upload and confirmation)):
        try:
            if upload.size > 50 * 1024 * 1024:
                raise ValueError('El respaldo supera 50 MB; restaura localmente o contacta soporte.')
            # Los campos de autenticación pertenecen a la base recuperada.
            restore_sqlite_bytes(upload.getvalue())
            st.session_state.clear()
            st.success('Copia restaurada. Inicia sesión con las credenciales guardadas en el respaldo.')
            st.rerun()
        except Exception as exc:
            st.error(f'No se pudo restaurar: {exc}')
