import streamlit as st
import pandas as pd
import tempfile
import os
import io
import time
from pathlib import Path
from datetime import datetime
from gemini_converter import GeminiPDFConverter
from excel_merger import ExcelMerger
from fix_especie_column import split_especie
from validation_module import validate_visual, validate_gallo, ValidationReport

# Page config
st.set_page_config(
    page_title="Procesador Financiero - Gallo + Visual",
    page_icon="💼",
    layout="wide"
)

# Initialize session state for file storage
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
    .validation-pass {
        color: #28a745;
        font-weight: bold;
    }
    .validation-fail {
        color: #dc3545;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">💼 Procesador Financiero Gallo + Visual</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Sistema Automatizado de Conversión y Unificación de Reportes</div>', unsafe_allow_html=True)

# Info box
st.markdown("""
<div class="info-box">
    <h4>🎯 Cómo funciona:</h4>
    <ol>
        <li><strong>Carga los PDFs</strong>: Sube el reporte de Gallo (pre-Jun 2025) y Visual (post-Jun 2025)</li>
        <li><strong>Procesamiento Automático</strong>: El sistema extrae y estructura todos los datos de ambos reportes</li>
        <li><strong>Tres outputs</strong>:
            <ul>
                <li>📊 Excel Gallo estructurado (7 hojas)</li>
                <li>📊 Excel Visual estructurado (7 hojas)</li>
                <li>📊 Excel Unificado con ambos periodos</li>
            </ul>
        </li>
    </ol>
    <p><strong>✨ Hojas generadas:</strong> Boletos | Resultado Ventas ARS/USD | Rentas Dividendos ARS/USD | Resumen | Posición Títulos</p>
</div>
""", unsafe_allow_html=True)

# File uploaders
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📄 PDF Gallo (Enero - Mayo 2025)")
    gallo_file = st.file_uploader(
        "Selecciona el PDF del sistema Gallo",
        type=['pdf'],
        key='gallo',
        help="Reporte del período enero a mayo 2025"
    )
    if gallo_file:
        st.success(f"✅ Archivo cargado: {gallo_file.name} ({gallo_file.size / 1024:.1f} KB)")

with col2:
    st.markdown("### 📄 PDF Visual (Junio - Nov 2025)")
    visual_file = st.file_uploader(
        "Selecciona el PDF del sistema Visual",
        type=['pdf'],
        key='visual',
        help="Reporte del período junio a noviembre 2025"
    )
    if visual_file:
        st.success(f"✅ Archivo cargado: {visual_file.name} ({visual_file.size / 1024:.1f} KB)")

# Process button
if st.button("🚀 Procesar Reportes", type="primary", use_container_width=True):
    if not gallo_file or not visual_file:
        st.error("⚠️ Por favor carga ambos archivos PDF")
    else:
        # Get API key from environment variable
        api_key = os.getenv('GOOGLE_API_KEY')
        if not api_key:
            st.error("⚠️ Error de configuración: Google API Key no encontrada. Contacte al administrador.")
        else:
            try:
                with st.spinner("🔄 Procesando reportes..."):
                    # Create temp directory
                    temp_dir = tempfile.mkdtemp()
                    
                    # Save uploaded files
                    gallo_path = os.path.join(temp_dir, "gallo.pdf")
                    visual_path = os.path.join(temp_dir, "visual.pdf")
                    
                    with open(gallo_path, "wb") as f:
                        f.write(gallo_file.getbuffer())
                    with open(visual_path, "wb") as f:
                        f.write(visual_file.getbuffer())
                    
                    # Initialize converter (Gemini - faster with higher limits)
                    converter = GeminiPDFConverter(api_key=api_key)
                    
                    # Progress tracking
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Step 1: Process Gallo
                    status_text.text("📊 Paso 1/5: Procesando reporte Gallo...")
                    progress_bar.progress(10)
                    gallo_excel = os.path.join(temp_dir, "gallo_estructurado.xlsx")
                    
                    try:
                        converter.convert_pdf_to_excel(gallo_path, gallo_excel, "gallo")
                        if not os.path.exists(gallo_excel):
                            raise FileNotFoundError(f"Gallo Excel no fue creado: {gallo_excel}")
                        st.info(f"✅ Gallo procesado exitosamente: {os.path.getsize(gallo_excel)} bytes")
                    except Exception as e:
                        st.error(f"❌ Error procesando Gallo: {str(e)}")
                        raise
                    
                    progress_bar.progress(35)
                    
                    # Step 2: Process Visual
                    status_text.text("📊 Paso 2/5: Procesando reporte Visual...")
                    visual_excel_temp = os.path.join(temp_dir, "visual_temp.xlsx")
                    converter.convert_pdf_to_excel(visual_path, visual_excel_temp, "visual")
                    progress_bar.progress(55)
                    
                    # Fix especie column for Visual
                    status_text.text("🔧 Paso 3/5: Estructurando Excel de Visual...")
                    visual_excel = os.path.join(temp_dir, "visual_estructurado.xlsx")
                    split_especie(visual_excel_temp, visual_excel)
                    progress_bar.progress(70)
                
                    
                    # Step 4: Validate generated files
                    status_text.text("✅ Paso 4/5: Validando integridad matemática...")
                    
                    # Validate Gallo
                    gallo_validation = validate_gallo(gallo_excel)
                    
                    # Validate Visual  
                    visual_validation = validate_visual(visual_excel)
                    
                    progress_bar.progress(80)
                    
                    # Step 5: Merge
                    status_text.text("🔗 Paso 5/5: Unificando reportes...")
                    progress_bar.progress(85)
                    merged_excel = os.path.join(temp_dir, "reporte_unificado.xlsx")
                    merger = ExcelMerger(gallo_excel, visual_excel)
                    merger.merge(merged_excel)
                    
                    progress_bar.progress(100)
                    status_text.text("✅ Procesamiento completado!")
                    
                    # Read files into memory and store in session state
                    with open(gallo_excel, "rb") as f:
                        gallo_data = f.read()
                    with open(visual_excel, "rb") as f:
                        visual_data = f.read()
                    with open(merged_excel, "rb") as f:
                        merged_data = f.read()
                    
                    # Store in session state
                    st.session_state.processed_files = {
                        'gallo': gallo_data,
                        'visual': visual_data,
                        'merged': merged_data,
                        'timestamp': datetime.now().strftime('%Y%m%d_%H%M'),
                        'merged_excel_path': merged_excel,
                        'gallo_validation': gallo_validation,
                        'visual_validation': visual_validation
                    }
                    
                    # Success message with validation results
                    gallo_status = "✅" if gallo_validation.all_passed else "⚠️"
                    visual_status = "✅" if visual_validation.all_passed else "⚠️"
                    
                    if gallo_validation.all_passed and visual_validation.all_passed:
                        st.markdown("""
                        <div class="success-box">
                            <h3>✅ Procesamiento Exitoso!</h3>
                            <p>Se generaron 3 archivos Excel estructurados con todas las transacciones.</p>
                            <p><strong>Validación matemática:</strong> Todos los cálculos verificados correctamente.</p>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div class="warning-box">
                            <h3>⚠️ Procesamiento Completo con Advertencias</h3>
                            <p>Los archivos se generaron pero hay diferencias en la validación:</p>
                            <ul>
                                <li>Gallo: {gallo_validation.passed_count}/{len(gallo_validation.results)} validaciones pasaron</li>
                                <li>Visual: {visual_validation.passed_count}/{len(visual_validation.results)} validaciones pasaron</li>
                            </ul>
                            <p>Revise los detalles de validación abajo.</p>
                        </div>
                        """, unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"❌ Error durante el procesamiento: {str(e)}")
                st.exception(e)

# Download buttons (outside the processing block to prevent reset)
if st.session_state.processed_files is not None:
    st.markdown("### 📥 Descargar Resultados")
    
    col1, col2, col3 = st.columns(3)
    
    timestamp = st.session_state.processed_files['timestamp']
    
    with col1:
        st.download_button(
            label="📊 Excel Gallo",
            data=st.session_state.processed_files['gallo'],
            file_name=f"Gallo_Estructurado_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key='download_gallo'
        )
    
    with col2:
        st.download_button(
            label="📊 Excel Visual",
            data=st.session_state.processed_files['visual'],
            file_name=f"Visual_Estructurado_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key='download_visual'
        )
    
    with col3:
        st.download_button(
            label="📊 Excel Unificado",
            data=st.session_state.processed_files['merged'],
            file_name=f"Reporte_Unificado_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key='download_merged'
        )
                
    # Preview (also outside processing block)
    st.markdown("### 👁️ Vista Previa del Reporte Unificado")
    
    # Show summary stats
    xl = pd.ExcelFile(io.BytesIO(st.session_state.processed_files['merged']))
    
    tabs = st.tabs(["✅ Validación", "📊 Estadísticas", "💰 Resumen", "📈 Posición"])
    
    with tabs[0]:
        st.markdown("#### Verificación de Integridad Matemática")
        st.markdown("Se verifica que los totales en las hojas de resumen coincidan con la suma de las transacciones individuales.")
        
        col_v1, col_v2 = st.columns(2)
        
        with col_v1:
            gallo_val = st.session_state.processed_files.get('gallo_validation')
            if gallo_val:
                status_class = "validation-pass" if gallo_val.all_passed else "validation-fail"
                status_icon = "✅" if gallo_val.all_passed else "⚠️"
                st.markdown(f"##### {status_icon} Gallo: <span class='{status_class}'>{gallo_val.passed_count}/{len(gallo_val.results)}</span>", unsafe_allow_html=True)
                
                # Create validation table
                val_data = []
                for r in gallo_val.results:
                    val_data.append({
                        "Campo": r.field,
                        "Calculado": f"{r.calculated:,.2f}",
                        "Esperado": f"{r.expected:,.2f}",
                        "Estado": "✓" if r.match else "✗"
                    })
                st.dataframe(pd.DataFrame(val_data), use_container_width=True, hide_index=True)
            else:
                st.info("Sin datos de validación de Gallo")
        
        with col_v2:
            visual_val = st.session_state.processed_files.get('visual_validation')
            if visual_val:
                status_class = "validation-pass" if visual_val.all_passed else "validation-fail"
                status_icon = "✅" if visual_val.all_passed else "⚠️"
                st.markdown(f"##### {status_icon} Visual: <span class='{status_class}'>{visual_val.passed_count}/{len(visual_val.results)}</span>", unsafe_allow_html=True)
                
                # Create validation table
                val_data = []
                for r in visual_val.results:
                    val_data.append({
                        "Campo": r.field,
                        "Calculado": f"{r.calculated:,.2f}",
                        "Esperado": f"{r.expected:,.2f}",
                        "Estado": "✓" if r.match else "✗"
                    })
                st.dataframe(pd.DataFrame(val_data), use_container_width=True, hide_index=True)
            else:
                st.info("Sin datos de validación de Visual")
    
    with tabs[1]:
        stats = []
        for sheet in xl.sheet_names:
            df = pd.read_excel(io.BytesIO(st.session_state.processed_files['merged']), sheet_name=sheet)
            stats.append({
                "Hoja": sheet,
                "Filas": len(df),
                "Columnas": len(df.columns)
            })
        st.dataframe(pd.DataFrame(stats), use_container_width=True, hide_index=True)
    
    with tabs[2]:
        if "Resumen" in xl.sheet_names:
            df_resumen = pd.read_excel(io.BytesIO(st.session_state.processed_files['merged']), sheet_name="Resumen")
            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
        else:
            st.info("No hay hoja de Resumen disponible")
    
    with tabs[3]:
        if "Posicion Titulos" in xl.sheet_names:
            df_pos = pd.read_excel(io.BytesIO(st.session_state.processed_files['merged']), sheet_name="Posicion Titulos")
            # Convert mixed-type columns to string to avoid Arrow serialization issues
            for col in df_pos.columns:
                if df_pos[col].dtype == 'object':
                    df_pos[col] = df_pos[col].astype(str)
            st.dataframe(df_pos, use_container_width=True, hide_index=True)
        else:
            st.info("No hay hoja de Posición disponible")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 1rem;">
    <p>💼 <strong>Sistema de Procesamiento Financiero Automatizado</strong></p>
    <p style="font-size: 0.9rem;">Gallo + Visual • Unificación de Reportes</p>
    <p style="font-size: 0.8rem;">⚠️ Sistema diseñado para procesamiento consistente de miles de clientes</p>
</div>
""", unsafe_allow_html=True)
