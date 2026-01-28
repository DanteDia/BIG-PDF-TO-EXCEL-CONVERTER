# 🚀 Deploy a Streamlit Cloud

Guía para desplegar la aplicación en Streamlit Cloud y obtener una URL pública.

## ✅ Requisitos

- ✓ Código en GitHub (ya lo está en https://github.com/DanteDia/BIG-PDF-TO-EXCEL-CONVERTER)
- ✓ Cuenta de GitHub
- ✓ API Key de Datalab

## 📋 Pasos para el Deploy

### Paso 1: Crear cuenta en Streamlit Cloud

1. Ve a https://share.streamlit.io/
2. Haz clic en **"Sign up with GitHub"**
3. Autoriza Streamlit Cloud a acceder a tus repositorios
4. Completa tu perfil

### Paso 2: Crear nueva aplicación

1. En Streamlit Cloud, haz clic en **"Create app"**
2. Completa los campos:
   - **Repository**: `DanteDia/BIG-PDF-TO-EXCEL-CONVERTER`
   - **Branch**: `main`
   - **Main file path**: `export_validation/app_datalab.py`

3. Haz clic en **"Deploy!"**

Streamlit Cloud comenzará el deploy automáticamente. Espera 2-5 minutos mientras construye e inicia la app.

### Paso 3: Configurar la API Key (IMPORTANTE)

Una vez desplegada, **debes configurar tu API Key de Datalab**:

1. En Streamlit Cloud, ve a tu app
2. Haz clic en el menú **⋮** (tres puntos) en la esquina superior derecha
3. Selecciona **"Settings"**
4. Ve a la pestaña **"Secrets"**
5. En el editor de texto, agrega:

```toml
DATALAB_API_KEY = "tu_api_key_aqui"
```

**⚠️ Importante**: Reemplaza `tu_api_key_aqui` con tu verdadera API Key de Datalab.

6. Haz clic en **"Save"**

La app se reiniciará automáticamente con los secrets configurados.

## 🌐 Acceder a tu App

Tu app estará disponible en: **`https://big-pdf-to-excel-converter.streamlit.app`**

(O la URL personalizada que Streamlit Cloud haya generado)

## 📊 Monitorear tu Deployment

### Dashboard de Streamlit Cloud

- Ver logs en tiempo real
- Monitorear uso de recursos
- Ver estado de la app
- Redeployar cambios automáticamente

### Auto-Deploy desde GitHub

Cuando hagas push a `main`, Streamlit Cloud **automáticamente**:
1. Detecta los cambios
2. Reconstruye la app
3. Inicia la nueva versión

**No necesitas hacer nada más después del primer deploy.**

## 🔧 Troubleshooting

### "DATALAB_API_KEY not found"

**Causa**: No configuraste el secret  
**Solución**:
1. Ve a Settings > Secrets
2. Agrega `DATALAB_API_KEY = "tu_key"`
3. Espera a que la app se reinicie

### "App is not loading"

**Causa**: Error durante el deploy  
**Solución**:
1. Ve a "Settings" → "Logs"
2. Revisa qué error aparece
3. Verifica requirements.txt
4. Si es necesario, haz push de cambios a GitHub
5. Streamlit Cloud redesplegará automáticamente

### "Timeout o carga lenta"

**Causa**: El servidor de Datalab está congestionado  
**Solución**:
1. Espera unos minutos e intenta de nuevo
2. Usa modo "standard" en lugar de "accurate" en OCR
3. Reinicia la app desde Settings > Reboot

## 📈 Compartir tu App

**URL para compartir**: `https://big-pdf-to-excel-converter.streamlit.app`

Puedes compartir directamente con tus compañeros. Ellos solo necesitarán:
- La URL
- Un PDF de resumen impositivo

## 🔐 Seguridad en Streamlit Cloud

✅ **Lo que Streamlit Cloud protege**:
- Tu API Key está encriptada
- Los secrets no se muestran en los logs
- No se exponen en GitHub
- Comunicación HTTPS

⚠️ **Ten en cuenta**:
- Los PDFs se procesan en tiempo real
- Asegúrate de que los usuarios confíen en la plataforma
- Los archivos generados son descargables (no almacenados)

## 🔄 Actualizar la App

Cuando hagas cambios en el código:

```bash
git add .
git commit -m "Descripción del cambio"
git push origin main
```

Streamlit Cloud **automáticamente**:
1. Detecta el cambio
2. Reconstruye
3. Redeploya en 1-2 minutos

No necesitas hacer nada más.

## 📊 Límites de Streamlit Cloud (Free Plan)

| Límite | Free |
|--------|------|
| Apps | Ilimitadas |
| Uptime | ~99% |
| Duración de sesión | 48 horas |
| Memoria RAM | 1 GB |
| CPU | Compartida |
| Procesamiento OCR | Limitado por Datalab API |

Para producción con mayor carga, considera upgrading a [Streamlit for Teams](https://streamlit.io/cloud).

## ✨ Próximas Mejoras

- [ ] Agregar caché para conversiones recientes
- [ ] Mostrar estadísticas de uso
- [ ] Enviar Excel por email
- [ ] Soporte para batch processing

## 📞 Soporte

- **Documentación Streamlit**: https://docs.streamlit.io/
- **GitHub Issues**: https://github.com/DanteDia/BIG-PDF-TO-EXCEL-CONVERTER/issues
- **Datalab Help**: https://datalab.to/help

---

**¡Tu app está lista para que todos la usen! 🎉**
