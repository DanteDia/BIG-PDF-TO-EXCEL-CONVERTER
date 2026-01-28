"""
Módulo de autenticación para la app Streamlit
Maneja login y acceso de usuarios autorizados
"""

import streamlit as st
import streamlit_authenticator as stauth
import yaml
import os
from pathlib import Path

def load_credentials():
    """Carga las credenciales desde secrets.toml"""
    try:
        # En Streamlit Cloud, lee desde st.secrets
        credentials = st.secrets.get("credentials", {})
        if credentials:
            return credentials
        
        # En local, intenta cargar desde archivo YAML
        config_path = Path(".streamlit/auth_config.yaml")
        if config_path.exists():
            with open(config_path) as f:
                config = yaml.safe_load(f)
                return config.get("credentials", {})
    except Exception as e:
        st.error(f"Error cargando credenciales: {e}")
    
    return {}

def initialize_authenticator():
    """Inicializa el autenticador de Streamlit"""
    credentials = load_credentials()
    
    if not credentials:
        st.error("❌ No se encontraron credenciales configuradas")
        st.stop()
    
    authenticator = stauth.Authenticate(
        credentials,
        "big_pdf_converter",  # cookie name
        "your_secret_key_here",  # cookie key (cambiar en producción)
        cookie_expiry_days=30,
        preauthorized=[]
    )
    
    return authenticator

def login_page(authenticator):
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
        
        # Mostrar autenticador
        try:
            name, authentication_status, username = authenticator.login("main")
            
            if authentication_status:
                st.session_state.authenticated = True
                st.session_state.username = name
                st.success(f"✅ ¡Bienvenido {name}!")
                st.balloons()
                st.rerun()
                
            elif authentication_status is False:
                st.error("❌ Usuario o contraseña incorrectos")
                
            elif authentication_status is None:
                st.warning("⚠️ Por favor ingresa tu usuario y contraseña")
                
                # Información de ayuda
                st.markdown("---")
                st.markdown("""
                **¿Necesitas ayuda?**
                - Contacta al administrador para obtener tus credenciales
                - Asegúrate de escribir correctamente el usuario y contraseña
                """)
                
        except Exception as e:
            st.error(f"Error en autenticación: {e}")
            st.stop()

def check_authentication():
    """Verifica si el usuario está autenticado"""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    
    return st.session_state.authenticated

def require_login():
    """Requiere login para acceder a la app"""
    if not check_authentication():
        authenticator = initialize_authenticator()
        login_page(authenticator)
        st.stop()

def logout_button(authenticator):
    """Muestra botón de logout en la sidebar"""
    with st.sidebar:
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.write(f"👤 {st.session_state.get('username', 'Usuario')}")
        
        with col2:
            try:
                authenticator.logout("Logout")
                if not st.session_state.get("login_widget", True):
                    st.session_state.authenticated = False
                    st.rerun()
            except:
                pass
