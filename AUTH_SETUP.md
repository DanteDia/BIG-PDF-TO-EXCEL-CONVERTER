# 🔐 Configuración de Autenticación

Este proyecto usa autenticación para proteger la API key de acceso no autorizado.

## ⚙️ Configuración Inicial

### Paso 1: Generar credenciales

Ejecuta el generador de usuarios:

```bash
python generate_credentials.py
```

El script te pedirá:
- Nombre de usuario
- Nombre completo
- Email
- Contraseña

Ejemplo:
```
Nombre de usuario: juan
Nombre completo: Juan García
Email: juan@company.com
Contraseña: MiContraseñaSegura123
```

### Paso 2: Copiar credenciales a secrets.toml

El script genera un bloque de código que debes copiar a `.streamlit/secrets.toml`:

```toml
[credentials]
usernames.juan.email = "juan@company.com"
usernames.juan.name = "Juan García"
usernames.juan.password = "$2b$12$pGXgPqCJKqMy02fH9Y1Wh..."
```

### Paso 3: Configurar en Streamlit Cloud

1. Ve a tu app en Streamlit Cloud
2. Settings > Secrets
3. Copia todo el contenido de `.streamlit/secrets.toml`
4. Pega en el editor de Secrets de Streamlit Cloud

## 🔑 Variables de Entorno Requeridas

En `.streamlit/secrets.toml` debe haber:

```toml
DATALAB_API_KEY = "tu_api_key_real"

[credentials]
usernames.usuario1.email = "user1@mail.com"
usernames.usuario1.name = "Usuario 1"
usernames.usuario1.password = "hash_bcrypt_aqui"
# ... más usuarios
```

## 👥 Agregar más usuarios

1. Ejecuta `python generate_credentials.py` nuevamente
2. Agrega los nuevos usuarios
3. Actualiza `secrets.toml` en Streamlit Cloud

## 🔒 Seguridad

- ✅ Las contraseñas se hashean con bcrypt (no se guardan en texto plano)
- ✅ `.streamlit/secrets.toml` está en `.gitignore` (no se sube a GitHub)
- ✅ La API key no se expone en el código
- ✅ Solo usuarios autenticados pueden acceder

## 🐛 Troubleshooting

### "No se encontraron credenciales"

**Causa**: Las credenciales no están en `secrets.toml`  
**Solución**: Ejecuta `generate_credentials.py` y actualiza `secrets.toml`

### "Usuario o contraseña incorrectos"

**Causa**: Credenciales mal escritas o hash incorrecto  
**Solución**: Regenera las credenciales con `generate_credentials.py`

### App aún requiere login después de cambiar secrets

**Causa**: Streamlit Cloud no recargó los secrets  
**Solución**: Ve a Settings > Reboot app

## 📖 Más información

- [streamlit-authenticator docs](https://github.com/mokerson/streamlit_authenticator)
- [bcrypt documentation](https://github.com/pyca/bcrypt)
