# 📥 Guía de Instalación

Pasos detallados para instalar y ejecutar el BIG PDF to Excel Converter.

## ⚙️ Requisitos Previos

- **Python 3.13 o superior**: [Descargar](https://www.python.org/downloads/)
- **Git**: [Descargar](https://git-scm.com/)
- **Cuenta en Datalab**: [Crear cuenta](https://datalab.to)
- **API Key de Datalab**: Solicita en tu cuenta Datalab

## 🖥️ Windows

### 1. Clonar el repositorio

Abre PowerShell o CMD en la carpeta donde quieras el proyecto:

```powershell
git clone https://github.com/DanteDia/BIG-PDF-TO-EXCEL-CONVERTER.git
cd BIG-PDF-TO-EXCEL-CONVERTER
```

### 2. Crear entorno virtual

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Si te da error de permisos, ejecuta PowerShell como administrador.

### 3. Instalar dependencias

```powershell
pip install -r requirements.txt
```

### 4. Configurar API Key

**Opción A: Variable de entorno (Recomendado)**

```powershell
# Temporal (solo esta sesión)
$env:DATALAB_API_KEY="tu_api_key_aqui"

# Permanente (requiere editar variables de entorno de Windows)
# Sistema > Configuración avanzada > Variables de entorno > Nueva
```

**Opción B: Archivo .env local**

Copia `.env.example` a `.env` y edita:

```
DATALAB_API_KEY=tu_api_key_aqui
```

### 5. Ejecutar la aplicación

```powershell
streamlit run export_validation\app_datalab.py
```

Abrirá automáticamente en `http://localhost:8501`

---

## 🍎 macOS

### 1. Clonar el repositorio

```bash
git clone https://github.com/DanteDia/BIG-PDF-TO-EXCEL-CONVERTER.git
cd BIG-PDF-TO-EXCEL-CONVERTER
```

### 2. Crear entorno virtual

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 4. Configurar API Key

```bash
# Temporal
export DATALAB_API_KEY="tu_api_key_aqui"

# Permanente (agregar al ~/.zshrc o ~/.bash_profile)
echo 'export DATALAB_API_KEY="tu_api_key_aqui"' >> ~/.zshrc
source ~/.zshrc
```

### 5. Ejecutar la aplicación

```bash
streamlit run export_validation/app_datalab.py
```

---

## 🐧 Linux

### 1. Clonar el repositorio

```bash
git clone https://github.com/DanteDia/BIG-PDF-TO-EXCEL-CONVERTER.git
cd BIG-PDF-TO-EXCEL-CONVERTER
```

### 2. Crear entorno virtual

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 4. Configurar API Key

```bash
# Temporal
export DATALAB_API_KEY="tu_api_key_aqui"

# Permanente
echo 'export DATALAB_API_KEY="tu_api_key_aqui"' >> ~/.bashrc
source ~/.bashrc
```

### 5. Ejecutar la aplicación

```bash
streamlit run export_validation/app_datalab.py
```

---

## 📱 Cómo Obtener tu API Key de Datalab

1. Inicia sesión en [Datalab.to](https://datalab.to)
2. Ve a **Settings** → **API Keys**
3. Copia tu API Key (o crea una nueva)
4. Configúrala en tu sistema (ver pasos arriba)

---

## 🚀 Verificar Instalación

```bash
# Ver versión de Python
python --version  # Debe ser 3.13+

# Ver versiones instaladas
pip list

# Verificar API Key configurada
python -c "import os; print('API Key:', 'Configurada ✓' if os.environ.get('DATALAB_API_KEY') else 'No configurada ✗')"
```

---

## ✅ Tu primera conversión

1. Abre http://localhost:8501 en tu navegador
2. Haz clic en "Browse files"
3. Selecciona un PDF de resumen impositivo
4. Selecciona modo "accurate" para mejor OCR
5. Espera a que procese
6. Descarga tu Excel

**¡Listo!** 🎉

---

## 🆘 Solución de Problemas

### Error: "DATALAB_API_KEY not found"

**Solución**: Asegúrate de configurar la variable de entorno correctamente:

```powershell
# Windows - Verificar
$env:DATALAB_API_KEY

# Linux/Mac - Verificar
echo $DATALAB_API_KEY
```

### Error: "ModuleNotFoundError: No module named 'streamlit'"

**Solución**: Reinstala las dependencias:

```bash
pip install --upgrade -r requirements.txt
```

### Error: "Python 3.13 required"

**Solución**: Actualiza Python desde [python.org](https://www.python.org/downloads/) o usa un manager como `pyenv`:

```bash
# macOS
brew install pyenv
pyenv install 3.13.0
pyenv local 3.13.0

# Linux
pyenv install 3.13.0
pyenv local 3.13.0
```

### Error: "Permission denied" en Linux/Mac

**Solución**: Usa `chmod` para ejecutar scripts:

```bash
chmod +x .venv/bin/activate
source .venv/bin/activate
```

---

## 🔄 Actualizar a la última versión

```bash
git pull origin main
pip install --upgrade -r requirements.txt
```

---

## 💬 ¿Necesitas ayuda?

- Abre un [Issue](https://github.com/DanteDia/BIG-PDF-TO-EXCEL-CONVERTER/issues)
- Revisa [CONTRIBUTING.md](CONTRIBUTING.md)
- Lee el [README](README.md)

¡Feliz conversión! 🚀
