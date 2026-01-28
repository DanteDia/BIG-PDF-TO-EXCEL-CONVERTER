# Changelog

Todos los cambios notables a este proyecto serán documentados en este archivo.

## [1.0.0] - 2026-01-28

### ✨ Características
- **Conversión Gallo Format**: Soporte completo para formato transaccional
  - Resultado Totales
  - Títulos Privados (Exentos, Exterior, etc.)
  - Renta Fija (Pesos, Dólares)
  - Cauciones (Pesos, Dólares)
  - Posición Inicial y Final
  
- **Conversión Visual Format**: Soporte para formato resumen
  - Boletos
  - Resultado Ventas (ARS/USD)
  - Rentas Dividendos (ARS/USD)
  - Cauciones (ARS/USD)
  - Resumen consolidado
  - Posición Títulos

- **Post-procesamiento Inteligente**:
  - Agrupación automática por tipo de instrumento
  - Detección de secciones (Cauciones, Rentas, Dividendos)
  - Manejo de Posición Inicial y Final con subtotales
  - Extracción y agregación de fecha en columna

- **Interfaz Streamlit**:
  - Upload de PDFs
  - Selección de modo OCR (accurate/standard)
  - Descarga de Excel procesado
  - Validación de datos en tiempo real

- **Seguridad**:
  - API keys via variables de entorno
  - No almacena credenciales
  - Validación de inputs

### 🔧 Técnico
- Parser markdown de Datalab OCR
- Post-procesamiento con openpyxl
- Metadata propagation (fecha) a través del pipeline
- Manejo robusto de OCR split rows
- Detección automática de formato (Gallo/Visual)

### 🐛 Correcciones
- Validación de filas divididas por OCR
- Manejo de valores vacíos en merges
- Detección de posiciones múltiples (Inicial/Final)
- Separación correcta de Cauciones

### 📝 Documentación
- README completo con instrucciones
- .env.example para variables de entorno
- CONTRIBUTING.md para colaboradores
- Docstrings en código principal

---

## Notas de Release

**Breaking Changes**: Ninguno en v1.0.0

**Migración**: N/A (primera versión)

**Dependencias nuevas**:
- streamlit>=1.28
- openpyxl>=3.10
- requests>=2.31

---

## Próximas Mejoras (Roadmap)

### v1.1.0
- [ ] Validación de Excel generado
- [ ] Exportación a PDF
- [ ] Batch processing de múltiples archivos
- [ ] Caché de conversiones

### v1.2.0
- [ ] Soporte para otros formatos de Datalab
- [ ] API REST para integración
- [ ] Tests automatizados
- [ ] Comparación visual vs actual

### v2.0.0
- [ ] Soporte para otros proveedores de OCR
- [ ] Machine learning para detección de errores
- [ ] Dashboard de estadísticas
- [ ] Exportación a múltiples formatos

---

**Para reportar bugs o sugerir mejoras:** Abre un [Issue](https://github.com/DanteDia/BIG-PDF-TO-EXCEL-CONVERTER/issues)
