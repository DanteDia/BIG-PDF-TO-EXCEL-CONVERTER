# 📊 Resumen Impositivo - PDF to Excel Converter

Herramienta automatizada para convertir informes de resumen impositivo en formato PDF a archivos Excel estructurados.

**🌐 [Pruébalo online aquí](https://big-pdf-to-excel-converter.streamlit.app)** (sin instalación requerida)

## 🚀 Características

- **Conversión automática**: Procesa PDFs de Datalab y genera Excel con múltiples hojas
- **Dos formatos soportados**: 
  - **Gallo**: Formato transaccional detallado
  - **Visual**: Formato resumen consolidado
- **Post-procesamiento inteligente**:
  - Agrupa operaciones por tipo de instrumento
  - Detecta y separa secciones (Cauciones, Rentas, Dividendos)
  - Maneja Posición Inicial y Posición Final
  - Agrega columna de fecha automáticamente
  - Formatea números y monedas
- **Interfaz web**: Aplicación Streamlit para fácil uso

## 📋 Requisitos

- Python 3.13+
- Cuenta en [Datalab.to](https://datalab.to) con API key

## 🔧 Instalación

1. **Clonar el repositorio**:
```bash
git clone <repository-url>
cd "Resumen Impositivo- Branch dots.OCR"
```

2. **Crear entorno virtual**:
```bash
python -m venv .venv
.\.venv\Scripts\Activate.ps1  # Windows PowerShell
# o
source .venv/bin/activate     # Linux/Mac
```

3. **Instalar dependencias**:
```bash
pip install -r requirements.txt
```

4. **Configurar API key de Datalab**:
```bash
# Windows PowerShell
$env:DATALAB_API_KEY="tu_api_key_aqui"

# Linux/Mac
export DATALAB_API_KEY="tu_api_key_aqui"
```

## 🎯 Uso

### 🌐 Online (Sin Instalación)

La forma más fácil: **[Abre la app aquí](https://big-pdf-to-excel-converter.streamlit.app)**

1. Sube tu PDF de resumen impositivo
2. Selecciona modo "accurate" para mejor OCR
3. Espera procesamiento (1-2 minutos)
4. Descarga tu Excel

Ver [STREAMLIT_DEPLOYMENT.md](STREAMLIT_DEPLOYMENT.md) para más detalles.

### 💻 Localmente (Instalación Requerida)

#### Interfaz Web (Recomendado)

```bash
streamlit run export_validation\app_datalab.py
```

Luego abre tu navegador en `http://localhost:8501`

**Pasos en la interfaz**:
1. Sube un PDF de resumen impositivo
2. Selecciona el modo de OCR (accurate recomendado)
3. Espera el procesamiento
4. Descarga el Excel generado

#### Línea de Comandos

```python
from pdf_converter.datalab.md_to_excel import convert_markdown_to_excel

# Si ya tienes el markdown de Datalab
convert_markdown_to_excel(
    'archivo.datalab.md',
    'salida.xlsx',
    apply_postprocess=True
)
```

## 📁 Estructura del Proyecto

```
.
├── pdf_converter/
│   └── datalab/
│       ├── md_to_excel.py      # Parser de markdown a Excel
│       ├── postprocess.py       # Post-procesamiento de hojas
│       └── datalab_client.py    # Cliente API Datalab
├── export_validation/
│   └── app_datalab.py          # Aplicación Streamlit
├── requirements.txt            # Dependencias
├── .gitignore                  # Archivos ignorados
└── README.md                   # Esta documentación
```

## 🔒 Seguridad

- **No incluyas API keys en el código**: Usa variables de entorno
- **Archivos sensibles**: Ya están en `.gitignore` (PDFs, Excel, backups)
- **Datos privados**: Los PDFs y Excel no se suben al repositorio

## 📝 Formatos Soportados

### Formato Gallo
Hojas generadas:
- Resultado Totales
- Títulos Privados (Exentos, Exterior, etc.)
- Renta Fija (Pesos, Dólares)
- Cauciones (Pesos, Dólares)
- **Posición Inicial** (con fecha)
- **Posición Final** (con fecha)

### Formato Visual
Hojas generadas:
- Boletos
- Resultado Ventas (ARS/USD)
- Rentas Dividendos (ARS/USD)
- Cauciones (ARS/USD)
- Resumen
- Posición Títulos

## 🐛 Solución de Problemas

### "DATALAB_API_KEY not found"
Asegúrate de configurar la variable de entorno antes de ejecutar la aplicación.

### "Markdown file not found"
Verifica que el archivo `.datalab.md` existe en el directorio actual.

### Errores de formato en Excel
Revisa que el PDF sea de resumen impositivo válido de Datalab.

## 🤝 Contribuir

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto es de uso interno. Consulta con el equipo antes de compartir externamente.

## 👥 Autores

Equipo de desarrollo - Resumen Impositivo

## 🙏 Agradecimientos

- [Datalab.to](https://datalab.to) - API de OCR
- [Streamlit](https://streamlit.io) - Framework de interfaz web
- [openpyxl](https://openpyxl.readthedocs.io) - Manipulación de Excel
