# 🔧 Configuración Final en Streamlit Cloud

Pasos para completar el deploy con autenticación.

## 1️⃣ Generar Credenciales Localmente

Primero, genera las credenciales de tus compañeros:

```bash
python generate_credentials.py
```

Te pedirá ingresar usuario, nombre, email y contraseña para cada compañero. 

**Ejemplo:**
```
Nombre de usuario: juan.garcia
Nombre completo: Juan García
Email: juan@company.com
Contraseña: MiContraseñaSegura123!
```

## 2️⃣ Copiar Credenciales a secrets.toml

El script genera un bloque de código. Cópialo y pégalo en `.streamlit/secrets.toml`:

```toml
[credentials]
usernames.juan.garcia.email = "juan@company.com"
usernames.juan.garcia.name = "Juan García"
usernames.juan.garcia.password = "$2b$12$pGXgPqCJKqMy02fH9Y1Wh..."
```

## 3️⃣ Configurar en Streamlit Cloud

### Acceso a Settings > Secrets

1. Ve a https://share.streamlit.io/
2. Haz clic en tu app
3. Menú ⋮ (arriba a la derecha)
4. Click en **Settings**
5. Selecciona tab **Secrets**

### Agregar Secrets

En el editor de texto, pega TODO el contenido de `.streamlit/secrets.toml`:

```toml
DATALAB_API_KEY = "tu_api_key_real_aqui"

[credentials]
usernames.juan.garcia.email = "juan@company.com"
usernames.juan.garcia.name = "Juan García"
usernames.juan.garcia.password = "$2b$12$..."
usernames.maria.lopez.email = "maria@company.com"
usernames.maria.lopez.name = "María López"
usernames.maria.lopez.password = "$2b$12$..."
# ... más usuarios
```

### Guardar y Reiniciar

1. Haz clic en **"Save"**
2. Streamlit Cloud reinicia automáticamente
3. Espera 30 segundos

## 4️⃣ Verificar que Funciona

1. Abre https://big-pdf-to-excel-converter.streamlit.app
2. Verás la pantalla de login
3. Intenta con una de las credenciales creadas
4. Deberías poder acceder a la app

## ✅ Checklist Final

- [ ] Credenciales generadas con `generate_credentials.py`
- [ ] `secrets.toml` contiene `DATALAB_API_KEY`
- [ ] `secrets.toml` contiene bloque `[credentials]` con usuarios
- [ ] Secrets copiados a Streamlit Cloud
- [ ] App reiniciada después de agregar secrets
- [ ] Login page aparece al abrir la app
- [ ] Puedo iniciar sesión con mis credenciales
- [ ] Puedo usar la app después del login
- [ ] El username aparece en la sidebar

## 📚 Archivos Clave

| Archivo | Propósito |
|---------|-----------|
| `generate_credentials.py` | Script para crear credenciales |
| `.streamlit/secrets.toml` | Archivo local con credenciales (NO SUBIR) |
| `pdf_converter/datalab/auth.py` | Módulo de autenticación |
| `export_validation/app_datalab.py` | App principal con login |
| `AUTH_SETUP.md` | Guía de configuración |

## 🚀 Flujo Completo

```
Tu PC (local)
    ↓
generate_credentials.py
    ↓
.streamlit/secrets.toml
    ↓
Copia a Streamlit Cloud Settings > Secrets
    ↓
App redeploya
    ↓
Login page visible
    ↓
Compañeros pueden usar
```

## ⚠️ Importante

- **NUNCA** subas `secrets.toml` a GitHub (está en `.gitignore`)
- **NUNCA** compartas las credenciales por email sin encriptar
- **SIEMPRE** usa contraseñas fuertes
- **SIEMPRE** crea un usuario único por compañero

## 🆘 Problemas Comunes

### "No se encontraron credenciales" 

Falta agregar el bloque `[credentials]` en Streamlit Cloud Secrets.

**Solución**: Revisa que copiaste TODO el contenido de `secrets.toml`, incluyendo la sección `[credentials]`.

### "Usuario o contraseña incorrectos"

Las credenciales no coinciden.

**Solución**: Ejecuta `generate_credentials.py` nuevamente y copia exactamente el hash generado.

### "DATALAB_API_KEY not found"

Falta la API key en Streamlit Cloud Secrets.

**Solución**: Asegúrate de que `DATALAB_API_KEY = "..."` esté en Secrets.

### Login page aparece pero no puedo entrar

El hash bcrypt puede ser incorrecto.

**Solución**: 
1. Genera nuevas credenciales: `python generate_credentials.py`
2. Reemplaza en Streamlit Cloud Secrets
3. Espera reboot (30s)

## 📖 Documentos de Referencia

- [USERGUIDE.md](USERGUIDE.md) - Guía para compañeros
- [AUTH_SETUP.md](AUTH_SETUP.md) - Detalles técnicos
- [STREAMLIT_DEPLOYMENT.md](STREAMLIT_DEPLOYMENT.md) - Deploy general

---

**Una vez hayas completado estos pasos, tus compañeros podrán usar la app con solo el link y sus credenciales.** ✅
