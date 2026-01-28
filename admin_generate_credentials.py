"""
Admin Tool para generar credenciales de Streamlit Auth
Ejecutar: streamlit run admin_generate_credentials.py
"""

import streamlit as st
import bcrypt
from datetime import datetime
import os

st.set_page_config(page_title="Admin - Generar Credenciales", layout="wide", initial_sidebar_state="collapsed")

st.title("🔑 Generador de Credenciales Streamlit Auth")
st.markdown("Crea credenciales para tus compañeros de forma segura")

# Inicializar sesión
if "credentials_list" not in st.session_state:
    st.session_state.credentials_list = []

if "toml_output" not in st.session_state:
    st.session_state.toml_output = ""

# ============================================================================
# SECCIÓN 1: CREAR UN NUEVO USUARIO
# ============================================================================
st.subheader("📝 Agregar Nuevo Usuario")

col1, col2 = st.columns(2)

with col1:
    username = st.text_input("👤 Nombre de usuario", placeholder="juan.garcia", key="username")
    name = st.text_input("📛 Nombre completo", placeholder="Juan García", key="name")

with col2:
    email = st.text_input("✉️ Email", placeholder="juan@company.com", key="email")
    password = st.text_input("🔐 Contraseña", type="password", placeholder="MiContraseña123!", key="password")

col_btn1, col_btn2 = st.columns([1, 3])

with col_btn1:
    if st.button("✅ Agregar Usuario", use_container_width=True):
        if not username or not name or not email or not password:
            st.error("❌ Todos los campos son requeridos")
        elif any(u["username"] == username for u in st.session_state.credentials_list):
            st.error(f"❌ El usuario '{username}' ya existe")
        elif len(password) < 6:
            st.error("❌ La contraseña debe tener al menos 6 caracteres")
        else:
            # Generar hash bcrypt
            password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
            
            st.session_state.credentials_list.append({
                "username": username,
                "name": name,
                "email": email,
                "password_hash": password_hash
            })
            
            st.success(f"✅ Usuario '{username}' agregado correctamente")
            st.balloons()

# ============================================================================
# SECCIÓN 2: LISTA DE USUARIOS
# ============================================================================
if st.session_state.credentials_list:
    st.divider()
    st.subheader("👥 Usuarios Agregados")
    
    # Tabla de usuarios
    for idx, cred in enumerate(st.session_state.credentials_list):
        col1, col2, col3, col4, col_del = st.columns([2, 2, 2, 2.5, 0.8])
        
        with col1:
            st.caption(f"👤 {cred['username']}")
        with col2:
            st.caption(f"📛 {cred['name']}")
        with col3:
            st.caption(f"✉️ {cred['email']}")
        with col4:
            st.caption(f"🔐 Hash: {cred['password_hash'][:20]}...")
        with col_del:
            if st.button("🗑️", key=f"del_{idx}", help="Eliminar usuario"):
                st.session_state.credentials_list.pop(idx)
                st.rerun()
    
    # ========================================================================
    # SECCIÓN 3: GENERAR CONFIGURACIÓN
    # ========================================================================
    st.divider()
    st.subheader("📋 Configuración para secrets.toml")
    
    # Generar TOML
    toml_lines = ["[credentials]"]
    for cred in st.session_state.credentials_list:
        toml_lines.append(f'usernames.{cred["username"]}.email = "{cred["email"]}"')
        toml_lines.append(f'usernames.{cred["username"]}.name = "{cred["name"]}"')
        toml_lines.append(f'usernames.{cred["username"]}.password = "{cred["password_hash"]}"')
        toml_lines.append("")  # Línea en blanco
    
    toml_output = "\n".join(toml_lines)
    
    st.info("""
    **📌 Pasos:**
    1. Copia todo el texto de abajo
    2. Pégalo en `.streamlit/secrets.toml` (al final del archivo, después de `DATALAB_API_KEY`)
    3. Guarda el archivo
    """)
    
    # Text area copiable
    st.code(toml_output, language="toml")
    
    # Botón para copiar al portapapeles
    col_copy, col_clear = st.columns([1, 1])
    
    with col_copy:
        st.info("📋 Selecciona el texto del recuadro arriba y presiona `Ctrl+C` para copiar")
    
    with col_clear:
        if st.button("🗑️ Limpiar Todo", use_container_width=True):
            st.session_state.credentials_list = []
            st.rerun()
    
    # ========================================================================
    # SECCIÓN 4: INSTRUCCIONES
    # ========================================================================
    st.divider()
    st.subheader("🚀 Próximos Pasos")
    
    with st.expander("Paso 1: Guardar en secrets.toml (local)", expanded=False):
        st.markdown("""
        1. Abre `.streamlit/secrets.toml` en tu editor de código
        2. Ve al final del archivo (después de `DATALAB_API_KEY`)
        3. Pega el contenido que copiaste arriba
        4. Guarda el archivo
        
        Debería verse así:
        """)
        st.code("""DATALAB_API_KEY = "tu_api_key_aqui"

[credentials]
usernames.juan.garcia.email = "juan@company.com"
usernames.juan.garcia.name = "Juan García"
usernames.juan.garcia.password = "$2b$12$..."
# ... más usuarios aquí""", language="toml")
    
    with st.expander("Paso 2: Configurar en Streamlit Cloud", expanded=False):
        st.markdown("""
        1. Ve a https://share.streamlit.io/
        2. Click en tu app: **big-pdf-to-excel-converter**
        3. Menú **⋮** (arriba a la derecha)
        4. Click en **Settings**
        5. Selecciona tab **Secrets**
        6. Pega **TODO el contenido** de `.streamlit/secrets.toml` (incluye `DATALAB_API_KEY` + `[credentials]`)
        7. Click en **Save**
        8. Espera 30 segundos a que reinicie
        """)
    
    with st.expander("Paso 3: Verificar que funciona", expanded=False):
        st.markdown("""
        1. Abre https://big-pdf-to-excel-converter.streamlit.app
        2. Verás la pantalla de login
        3. Intenta entrar con uno de los usuarios creados
        4. Deberías ver la app de conversión
        5. Tu nombre aparece en la sidebar (arriba a la izquierda)
        """)

else:
    st.info("👆 Agrega usuarios arriba para comenzar")

# ============================================================================
# FOOTER
# ============================================================================
st.divider()
st.markdown("""
<hr style='opacity: 0.3;'>

**⚠️ IMPORTANTE:**
- Cada usuario debe tener una **contraseña única**
- Las contraseñas se **hashean con bcrypt** (no se pueden recuperar)
- **NUNCA** compartas contraseñas por email sin encriptar
- El archivo `secrets.toml` **NO se sube a GitHub** (está en .gitignore)

**¿Necesitas ayuda?** Ver [AUTH_SETUP.md](AUTH_SETUP.md)
""", unsafe_allow_html=True)
