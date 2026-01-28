"""
Streamlit app for PDF to Excel conversion.
Supports both Gallo and Visual format financial reports.
"""

import streamlit as st
import pandas as pd
import tempfile
import os
import io
import re
import sys
from pathlib import Path
from datetime import datetime

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from dotenv import load_dotenv

# Load environment variables
load_dotenv(Path(__file__).parent.parent / "pdf_converter" / ".env")

# Import authentication
from pdf_converter.datalab.auth import require_login, logout_button

# ==================== AUTENTICACIÓN ====================
# Requiere login antes de mostrar la app
require_login()

# ==================== APP PRINCIPAL ====================

# Import our converter
from pdf_converter.datalab import DatalabClient
from pdf_converter.datalab.md_to_excel import convert_markdown_to_excel
from pdf_converter.datalab.postprocess import postprocess_gallo_workbook, postprocess_visual_workbook
from openpyxl import load_workbook

# Page config
st.set_page_config(
    page_title="Procesador Financiero - Gallo + Visual",
    page_icon="💼",
    layout="wide"
)

# ==================== SIDEBAR ====================
with st.sidebar:
    st.markdown("---")
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Obtener nombre del usuario autenticado
        username = st.session_state.get('username', 'Usuario')
        st.markdown(f"👤 **{username}**")
    
    # En producción, agregar botón de logout aquí
    st.markdown("---")

# Initialize session state
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = None

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #003366;
        font-weight: bold;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">💼 Procesador Financiero Gallo + Visual</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Sistema Automatizado de Conversión de Reportes</div>', unsafe_allow_html=True)

# Info box
st.markdown("""
<div class="info-box">
    <h4>🎯 Cómo funciona:</h4>
    <ol>
        <li><strong>Carga los PDFs</strong>: Sube el reporte Gallo y/o Visual</li>
        <li><strong>Procesamiento</strong>: Extracción y estructuración automática</li>
        <li><strong>Excel estructurado</strong>: Con todas las columnas necesarias</li>
    </ol>
    <p><strong>✨ Hojas generadas:</strong> Boletos | Resultado Ventas ARS/USD | Rentas Dividendos ARS/USD | Resumen | Posición Títulos</p>
</div>
""", unsafe_allow_html=True)


def convert_pdf_to_excel_streamlit(pdf_bytes: bytes, pdf_name: str, format_type: str, temp_dir: str, progress_callback=None) -> tuple:
    """
    Convert PDF to Excel using Datalab API.
    
    Args:
        pdf_bytes: PDF file content
        pdf_name: Original filename
        format_type: 'gallo' or 'visual'
        temp_dir: Temporary directory path
        progress_callback: Optional callback for progress updates
    
    Returns:
        Tuple of (excel_path, comitente_number, comitente_name)
    """
    from pdf_converter.datalab.md_to_excel import extract_comitente_info
    
    api_key = os.environ.get("DATALAB_API_KEY", "").strip()
    if not api_key:
        raise ValueError("DATALAB_API_KEY no encontrada. Configure la variable de entorno.")
    
    # Save PDF to temp
    pdf_path = os.path.join(temp_dir, f"{format_type}.pdf")
    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)
    
    if progress_callback:
        progress_callback(f"Procesando {format_type.upper()} con OCR...")
    
    # Convert PDF to Markdown using Datalab
    with DatalabClient(api_key=api_key, mode="accurate") as client:
        result = client.convert_pdf(pdf_path, paginate=True)
        
        if not result.success:
            raise RuntimeError(f"Error en OCR: {result.error}")
    
    markdown_content = result.markdown or ""
    
    # Extract comitente info from markdown
    comitente_number, comitente_name = extract_comitente_info(markdown_content)
    
    # Save markdown
    md_path = os.path.join(temp_dir, f"{format_type}.datalab.md")
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(markdown_content)
    
    if progress_callback:
        progress_callback(f"Creando Excel de {format_type.upper()}...")
    
    # Convert Markdown to Excel
    excel_path = os.path.join(temp_dir, f"{format_type}_estructurado.xlsx")
    convert_markdown_to_excel(md_path, excel_path, apply_postprocess=True)
    
    return excel_path, comitente_number, comitente_name


# File uploaders
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📄 PDF Gallo")
    gallo_file = st.file_uploader(
        "Selecciona el PDF del sistema Gallo",
        type=['pdf'],
        key='gallo',
        help="Reporte formato Gallo (antes de Jun 2025)"
    )
    if gallo_file:
        st.success(f"✅ {gallo_file.name} ({gallo_file.size / 1024:.1f} KB)")

with col2:
    st.markdown("### 📄 PDF Visual")
    visual_file = st.file_uploader(
        "Selecciona el PDF del sistema Visual",
        type=['pdf'],
        key='visual',
        help="Reporte formato Visual (post Jun 2025)"
    )
    if visual_file:
        st.success(f"✅ {visual_file.name} ({visual_file.size / 1024:.1f} KB)")

# Process button
if st.button("🚀 Procesar Reportes", type="primary", use_container_width=True):
    if not gallo_file and not visual_file:
        st.error("⚠️ Por favor carga al menos un archivo PDF")
    else:
        api_key = os.environ.get("DATALAB_API_KEY", "").strip()
        if not api_key:
            st.error("⚠️ DATALAB_API_KEY no configurada. Agregue la API key en el archivo .env")
        else:
            try:
                with st.spinner("🔄 Procesando reportes..."):
                    temp_dir = tempfile.mkdtemp()
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    results = {}
                    step = 0
                    total_steps = (1 if gallo_file else 0) + (1 if visual_file else 0)
                    
                    # Process Gallo if provided
                    if gallo_file:
                        status_text.text("📊 Procesando reporte Gallo...")
                        step += 1
                        progress_bar.progress(int(step / (total_steps + 1) * 100))
                        
                        gallo_excel, gallo_comitente_num, gallo_comitente_name = convert_pdf_to_excel_streamlit(
                            gallo_file.getvalue(),
                            gallo_file.name,
                            "gallo",
                            temp_dir,
                            lambda msg: status_text.text(msg)
                        )
                        
                        with open(gallo_excel, "rb") as f:
                            results['gallo'] = f.read()
                        results['gallo_comitente_num'] = gallo_comitente_num
                        results['gallo_comitente_name'] = gallo_comitente_name
                    
                    # Process Visual if provided
                    if visual_file:
                        status_text.text("📊 Procesando reporte Visual...")
                        step += 1
                        progress_bar.progress(int(step / (total_steps + 1) * 100))
                        
                        visual_excel, visual_comitente_num, visual_comitente_name = convert_pdf_to_excel_streamlit(
                            visual_file.getvalue(),
                            visual_file.name,
                            "visual",
                            temp_dir,
                            lambda msg: status_text.text(msg)
                        )
                        
                        with open(visual_excel, "rb") as f:
                            results['visual'] = f.read()
                        results['visual_comitente_num'] = visual_comitente_num
                        results['visual_comitente_name'] = visual_comitente_name
                    
                    progress_bar.progress(100)
                    status_text.text("✅ Procesamiento completado!")
                    
                    # Store results
                    st.session_state.processed_files = {
                        **results,
                        'timestamp': datetime.now().strftime('%Y%m%d_%H%M')
                    }
                    
                    st.markdown("""
                    <div class="success-box">
                        <h3>✅ Procesamiento Exitoso!</h3>
                        <p>Los archivos Excel estructurados están listos para descargar.</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
            except Exception as e:
                st.error(f"❌ Error durante el procesamiento: {str(e)}")
                st.exception(e)

# Download buttons
if st.session_state.processed_files is not None:
    st.markdown("### 📥 Descargar Resultados")
    
    timestamp = st.session_state.processed_files['timestamp']
    
    cols = st.columns(2)
    col_idx = 0
    
    if 'gallo' in st.session_state.processed_files:
        with cols[col_idx]:
            # Build filename with comitente info
            gallo_num = st.session_state.processed_files.get('gallo_comitente_num', '')
            gallo_name = st.session_state.processed_files.get('gallo_comitente_name', '')
            if gallo_num and gallo_name:
                # Clean name for filename (remove special chars)
                clean_name = re.sub(r'[^\w\s]', '', gallo_name).strip().replace(' ', '_')[:30]
                gallo_filename = f"{gallo_num}_{clean_name}_Gallo_{timestamp}.xlsx"
            else:
                gallo_filename = f"Gallo_Estructurado_{timestamp}.xlsx"
            
            st.download_button(
                label="📊 Excel Gallo",
                data=st.session_state.processed_files['gallo'],
                file_name=gallo_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key='download_gallo'
            )
        col_idx = 1
    
    if 'visual' in st.session_state.processed_files:
        with cols[col_idx]:
            # Build filename with comitente info
            visual_num = st.session_state.processed_files.get('visual_comitente_num', '')
            visual_name = st.session_state.processed_files.get('visual_comitente_name', '')
            if visual_num and visual_name:
                # Clean name for filename (remove special chars)
                clean_name = re.sub(r'[^\w\s]', '', visual_name).strip().replace(' ', '_')[:30]
                visual_filename = f"{visual_num}_{clean_name}_Visual_{timestamp}.xlsx"
            else:
                visual_filename = f"Visual_Estructurado_{timestamp}.xlsx"
            
            st.download_button(
                label="📊 Excel Visual",
                data=st.session_state.processed_files['visual'],
                file_name=visual_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key='download_visual'
            )
    
    # Preview
    st.markdown("### 👁️ Vista Previa")
    
    preview_options = []
    if 'gallo' in st.session_state.processed_files:
        preview_options.append("Gallo")
    if 'visual' in st.session_state.processed_files:
        preview_options.append("Visual")
    
    if preview_options:
        selected = st.selectbox("Seleccionar archivo para vista previa:", preview_options)
        
        data_key = selected.lower()
        xl = pd.ExcelFile(io.BytesIO(st.session_state.processed_files[data_key]))
        
        tabs = st.tabs(xl.sheet_names)
        
        for i, sheet_name in enumerate(xl.sheet_names):
            with tabs[i]:
                df = pd.read_excel(io.BytesIO(st.session_state.processed_files[data_key]), sheet_name=sheet_name)
                st.dataframe(df, use_container_width=True, hide_index=True)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 1rem;">
    <p>💼 <strong>Procesador Financiero Gallo + Visual</strong></p>
    <p style="font-size: 0.9rem;">Conversión automática de reportes</p>
</div>
""", unsafe_allow_html=True)
