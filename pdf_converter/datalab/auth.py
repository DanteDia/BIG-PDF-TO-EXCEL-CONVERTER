"""
Módulo de autenticación simple para la app Streamlit
Maneja login y acceso de usuarios autorizados (sin streamlit-authenticator)
"""

import streamlit as st
import bcrypt

def load_credentials():
    """Carga las credenciales desde secrets.toml"""
    try:
        if "credentials" in st.secrets and "usernames" in st.secrets["credentials"]:
            return st.secrets["credentials"]["usernames"]
    except Exception as e:
        st.error(f"Error cargando credenciales: {e}")
    
    return None

def verify_password(password, hashed_password):
    """Verifica si la contraseña coincide con el hash bcrypt"""
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))
    except Exception:
        return False

def login_page():
    """Muestra la página de login"""
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("---")
        st.title("🔐 Acceso Restringido")
        st.markdown("""
        ### PDF to Excel Converter
        Convierte tus PDFs de resumen impositivo a Excel automáticamente
        """)
        st.markdown("---")
        
        # Cargar credenciales
        credentials = load_credentials()
        
        if not credentials:
            st.error("❌ No se encontraron credenciales configuradas")
            st.stop()
        
        # Formulario de login
        with st.form("login_form"):
            username = st.text_input("👤 Usuario")
            password = st.text_input("🔐 Contraseña", type="password")
            submit = st.form_submit_button("Iniciar Sesión", use_container_width=True)
            
            if submit:
                # Verificar si el usuario existe
                if username in credentials:
                    user_data = credentials[username]
                    stored_password = user_data["password"]
                    
                    # Verificar contraseña
                    if verify_password(password, stored_password):
                        st.session_state.authenticated = True
                        st.session_state.username = user_data["name"]
                        st.session_state.user_email = user_data["email"]
                        st.success(f"✅ ¡Bienvenido {user_data['name']}!")
                        st.balloons()
                        st.rerun()
                    else:
                        st.error("❌ Usuario o contraseña incorrectos")
                else:
                    st.error("❌ Usuario o contraseña incorrectos")
        
        # Información de ayuda
        st.markdown("---")
        st.markdown("""
        **¿Necesitas ayuda?**
        - Contacta al administrador para obtener tus credenciales
        - Asegúrate de escribir correctamente el usuario y contraseña
        """)

def check_authentication():
    """Verifica si el usuario está autenticado"""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    
    return st.session_state.authenticated

def require_login():
    """Requiere login para acceder a la app"""
    if not check_authentication():
        login_page()
        st.stop()

def logout_button():
    """Muestra botón de logout en la sidebar"""
    with st.sidebar:
        st.markdown("---")
        st.write(f"👤 {st.session_state.get('username', 'Usuario')}")
        
        if st.button("🚪 Cerrar Sesión"):
            st.session_state.authenticated = False
            st.session_state.username = None
            st.session_state.user_email = None
            st.rerun()
