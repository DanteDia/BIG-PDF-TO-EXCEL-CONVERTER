# Contribuyendo a BIG PDF to Excel Converter

¡Gracias por tu interés en contribuir! Este documento te guiará en el proceso.

## 🐛 Reportar Bugs

Si encuentras un bug, abre un **Issue** con:
- Título claro y descriptivo
- Descripción detallada del problema
- Pasos para reproducir
- Comportamiento esperado vs actual
- Versión de Python y sistema operativo

**Ejemplo:**
```
Título: Error en conversión de Posición Final - filas duplicadas
Descripción: Al procesar el archivo Aguiar_Gallo.pdf, la hoja "Posición Final" 
muestra filas duplicadas...
```

## 💡 Sugerir Mejoras

Abre un **Issue** con:
- Descripción de la mejora propuesta
- Por qué sería útil
- Ejemplos de uso

## 🔧 Código

### Configuración de Desarrollo

1. Fork el proyecto
2. Clona tu fork:
```bash
git clone https://github.com/tu-usuario/BIG-PDF-TO-EXCEL-CONVERTER.git
cd BIG-PDF-TO-EXCEL-CONVERTER
```

3. Crea una rama para tu feature:
```bash
git checkout -b feature/mi-mejora
# o para bugs:
git checkout -b fix/correccion-importante
```

4. Instala dependencias de desarrollo:
```bash
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install pytest black pylint  # Testing y linting
```

### Estándares de Código

- **Python**: Sigue [PEP 8](https://pep8.org/)
- **Formato**: Usa `black` antes de commitear
- **Imports**: Organiza según `isort` (imports estándar, terceros, locales)
- **Docstrings**: Usa formato Google

**Ejemplo de función bien documentada:**
```python
def convert_markdown_to_excel(md_path: str, output_path: str, apply_postprocess: bool = True) -> None:
    """
    Convierte archivo markdown de Datalab a Excel estructurado.
    
    Args:
        md_path: Ruta al archivo .datalab.md
        output_path: Ruta donde guardar el Excel
        apply_postprocess: Si aplica post-procesamiento (default: True)
    
    Raises:
        FileNotFoundError: Si el archivo markdown no existe
        ValueError: Si el formato del markdown no es válido
    """
```

### Antes de Commitear

```bash
# Formatea el código
black pdf_converter/ export_validation/

# Revisa errores
pylint pdf_converter/

# Ejecuta tests (si aplica)
pytest
```

### Estructura de Commits

```bash
git commit -m "feature: agregar soporte para formato XYZ

- Descripción detallada de los cambios
- Menciona qué archivos se modificaron
- Sé específico sobre la lógica implementada"
```

**Tipos de commits:**
- `feature`: Nueva funcionalidad
- `fix`: Corrección de bug
- `docs`: Cambios en documentación
- `refactor`: Reorganización de código sin cambiar funcionalidad
- `test`: Agregar o mejorar tests

## 📝 Pull Request

1. **Asegúrate de que tu rama está actualizada:**
```bash
git fetch origin
git rebase origin/main
```

2. **Pushea tu rama:**
```bash
git push origin feature/mi-mejora
```

3. **Abre un Pull Request en GitHub con:**
   - Título claro
   - Descripción de los cambios
   - Referencia a Issues relacionados (cierra #123)
   - Screenshots o ejemplos si es visual

**Descripción de PR útil:**
```
## Descripción
Agrega validación automática de archivos antes de procesarlos para evitar errores

## Tipo de cambio
- [x] Bug fix
- [x] Nueva funcionalidad
- [ ] Breaking change

## Testing
- [x] Testeado localmente
- [x] Testeado con archivos Gallo
- [x] Testeado con archivos Visual

Cierra #456
```

## ✅ Checklist para Contribuidores

Antes de subir tu PR:

- [ ] Mi código sigue los estándares de estilo (PEP 8)
- [ ] He actualizado la documentación necesaria
- [ ] He probado mis cambios localmente
- [ ] Mi rama está basada en `main` actualizado
- [ ] Los commits tienen mensajes descriptivos
- [ ] He testeado con archivos reales de Datalab
- [ ] No introduzco dependencias innecesarias

## 🤝 Cultura de Colaboración

- Sé respetuoso con otros contribuidores
- Proporciona feedback constructivo
- Si tienes dudas, pregunta en los Issues
- Lee el README antes de empezar
- Si es tu primer PR, ¡no dudes en pedir ayuda!

## 📧 Contacto

Si tienes preguntas, abre un Issue o discute en la sección de Discussions.

¡Gracias por contribuir! 🚀
