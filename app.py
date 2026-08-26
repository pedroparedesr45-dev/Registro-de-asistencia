import streamlit as st
import streamlit.components.v1 as components
from supabase import create_client
import pandas as pd
from datetime import date

# --- 1. CONFIGURACIÓN DE PÁGINA E INYECCIÓN NATIVA PWA ---
st.set_page_config(page_title="Sistema PF", layout="wide", page_icon="📱")

# Inyección dinámica para reconocimiento PWA en móviles (Android/iOS)
pwa_code = """
<script>
    const linkManifest = parent.document.createElement('link');
    linkManifest.rel = 'manifest';
    linkManifest.href = 'https://raw.githubusercontent.com/pedroparedesr45-dev/Registro-de-asistencia/main/manifest.json';
    parent.document.head.appendChild(linkManifest);

    const metaViewport = parent.document.createElement('meta');
    metaViewport.name = 'viewport';
    metaViewport.content = 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no';
    parent.document.head.appendChild(metaViewport);

    const metaTheme = parent.document.createElement('meta');
    metaTheme.name = 'theme-color';
    metaTheme.content = '#0E1117';
    parent.document.head.appendChild(metaTheme);
</script>
"""
components.html(pwa_code, height=0, width=0)

# --- 2. CONEXIÓN A SUPABASE ---
@st.cache_resource
def init_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_supabase()

# --- 3. GESTIÓN DE SESIÓN ---
if "user" not in st.session_state:
    st.session_state.user = None

def login(email, password):
    res = supabase.table("usuarios").select("*").eq("email", email).eq("password_hash", password).execute()
    if res.data:
        st.session_state.user = res.data[0]
        st.rerun()
    else:
        st.error("Credenciales incorrectas")

# Pantalla de Login
if not st.session_state.user:
    st.title("🔑 Acceso al Sistema PF")
    with st.form("login_form"):
        email = st.text_input("Correo Electrónico")
        password = st.text_input("Contraseña", type="password")
        if st.form_submit_button("Ingresar", use_container_width=True):
            login(email, password)
    st.stop()

# Usuario Autenticado
user = st.session_state.user
st.sidebar.write(f"👤 **{user['email']}**")
st.sidebar.caption(f"Rol: {user['rol']}")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.user = None
    st.rerun()

# --- 4. VISTA SUPER ADMIN ---
if user["rol"] == "SUPER_ADMIN":
    st.title("🛡️ Panel Super Admin")
    tab1, tab2 = st.tabs(["➕ Empresas y Admins", "📊 Consolidado General"])

    with tab1:
        st.subheader("Registrar Nueva Empresa")
        with st.form("form_empresa"):
            nombre_emp = st.text_input("Nombre de la Empresa")
            admin_email = st.text_input("Correo del Admin")
            admin_pass = st.text_input("Contraseña del Admin", type="password")
            if st.form_submit_button("Crear Empresa"):
                if nombre_emp and admin_email and admin_pass:
                    emp_res = supabase.table("empresas").insert({"nombre": nombre_emp}).execute()
                    emp_id = emp_res.data[0]["id"]
                    supabase.table("usuarios").insert({
                        "email": admin_email,
                        "password_hash": admin_pass,
                        "rol": "ADMIN_EMPRESA",
                        "empresa_id": emp_id
                    }).execute()
                    st.success(f"Empresa '{nombre_emp}' creada correctamente.")

    with tab2:
        st.subheader("Reportes Globales")
        reportes = supabase.table("reportes").select("*, empresas(nombre)").execute()
        if reportes.data:
            st.dataframe(pd.DataFrame(reportes.data), use_container_width=True)

# --- 5. VISTA ADMIN EMPRESA INDEPENDIENTE ---
elif user["rol"] == "ADMIN_EMPRESA":
    emp_data = supabase.table("empresas").select("nombre").eq("id", user["empresa_id"]).execute()
    st.title(f"🏢 Panel: {emp_data.data[0]['nombre']}")
    
    tab1, tab2 = st.tabs(["📝 Registrar Asistencia", "📈 Mis Reportes"])
    
    with tab1:
        with st.form("form_asistencia"):
            empleado = st.text_input("Nombre del Empleado")
            horas = st.number_input("Horas Trabajadas", min_value=0.0, step=0.5)
            monto = st.number_input("Monto Nómina ($)", min_value=0.0, step=10.0)
            if st.form_submit_button("Guardar Registro"):
                if empleado:
                    supabase.table("reportes").insert({
                        "empresa_id": user["empresa_id"],
                        "empleado_nombre": empleado,
                        "fecha": str(date.today()),
                        "horas_trabajadas": horas,
                        "monto_nomina": monto
                    }).execute()
                    st.success("Registro guardado exitosamente.")
                    st.rerun()

    with tab2:
        data = supabase.table("reportes").select("*").eq("empresa_id", user["empresa_id"]).execute()
        if data.data:
            df = pd.DataFrame(data.data)
            st.metric("Total Nómina", f"${df['monto_nomina'].sum():,.2f}")
            st.dataframe(df[['empleado_nombre', 'fecha', 'horas_trabajadas', 'monto_nomina']], use_container_width=True)
