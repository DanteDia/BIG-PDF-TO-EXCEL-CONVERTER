# Módulo de Validación y UI para Reportes Financieros Gallo/Visual

Este módulo proporciona verificación automática de integridad matemática para archivos Excel generados a partir de PDFs financieros de los sistemas Gallo y Visual.

## Archivos Incluidos

| Archivo | Descripción |
|---------|-------------|
| `validation_module.py` | Lógica de validación matemática para Gallo y Visual |
| `app.py` | UI de Streamlit para procesamiento y visualización |
| `excel_merger.py` | Unificador de reportes Gallo + Visual |
| `fix_especie_column.py` | Utilidad para estructurar columna especie |

## Uso del Módulo de Validación

### Importación

```python
from validation_module import validate_visual, validate_gallo, ValidationReport
```

### Validar archivo Visual

```python
report = validate_visual('archivo_visual.xlsx')

print(f"Pasaron: {report.passed_count}/{len(report.results)}")
print(f"Todas pasaron: {report.all_passed}")

# Ver detalles
for r in report.results:
    status = '✓' if r.match else '✗'
    print(f"{r.field}: calc={r.calculated:.2f}, expected={r.expected:.2f} {status}")
```

### Validar archivo Gallo

```python
report = validate_gallo('archivo_gallo.xlsx')
report.print_report()  # Imprime reporte formateado
```

### Detección automática de tipo

```python
from validation_module import run_full_validation

# Detecta automáticamente si es Visual o Gallo
report = run_full_validation('archivo.xlsx')
```

## Relaciones Matemáticas Verificadas

### Visual

| Celda Resumen | Fórmula |
|---------------|---------|
| B2 (ARS ventas) | `=SUM('Resultado Ventas ARS'!resultado)` |
| B3 (USD ventas) | `=SUM('Resultado Ventas USD'!resultado)` |
| E2 (ARS rentas) | `=SUMIF(Rentas Div ARS, categoria="RENTAS", importe)` |
| E3 (USD rentas) | `=SUMIF(Rentas Div USD, categoria="Rentas", importe)` |
| F2 (ARS dividendos) | `=SUMIF(Rentas Div ARS, categoria="DIVIDENDOS", importe)` |
| F3 (USD dividendos) | `=SUMIF(Rentas Div USD, categoria="Dividendos", importe)` |
| L2/L3 (total) | `=SUM(B:K)` por fila |

### Gallo

La validación de Gallo es **dinámica**: lee las categorías de la hoja "Resultado Totales" y verifica cada una contra su hoja de detalle correspondiente.

Para cada categoría (ej: "RENTA FIJA EN DOLARES (Enajenacion)"):
1. Identifica la hoja de detalle (ej: "Renta Fija Dolares")
2. Suma las **transacciones individuales** (excluyendo filas de Total)
3. Filtra por tipo de operación:
   - **Enajenación**: compra, venta, amortización, cpra cable, ret ajuste
   - **Renta**: renta, dividendo
4. Compara con el valor esperado en Resultado Totales

**Mapeo de categorías a hojas:**

| Categoría | Hoja |
|-----------|------|
| TIT.PRIVADOS EXENTOS | Tit.Privados Exentos |
| TIT.PRIVADOS DEL EXTERIOR | Tit.Privados Exterior |
| RENTA FIJA EN PESOS | Renta Fija Pesos |
| RENTA FIJA EN DOLARES | Renta Fija Dolares |
| CAUCIONES EN PESOS | Cauciones Pesos |
| CAUCIONES EN DOLARES | Cauciones Dolares |

## Estructura Excel Esperada

### Gallo

- **Resultado Totales**: Resumen con columnas `categoria`, `valor_pesos`, `valor_usd`
- **Hojas de detalle**: Columnas incluyen `tipo_fila`, `especie`, `fecha`, `operacion`, `resultado_pesos`, `resultado_usd`

### Visual

- **Resumen**: Columnas `moneda`, `ventas`, `rentas`, `dividendos`, `total`, etc.
- **Resultado Ventas ARS/USD**: Columna `resultado`
- **Rentas Dividendos ARS/USD**: Columnas `categoria`, `importe`

## Integración con tu Conversor

```python
# Tu conversor genera el Excel
mi_conversor.convert_pdf_to_excel('entrada.pdf', 'salida.xlsx', tipo='gallo')

# Validar el resultado
from validation_module import validate_gallo

report = validate_gallo('salida.xlsx')

if report.all_passed:
    print("✅ Todas las validaciones pasaron")
else:
    print(f"⚠️ {report.failed_count} validaciones fallaron")
    for r in report.results:
        if not r.match:
            print(f"  - {r.field}: calc={r.calculated:.2f} vs expected={r.expected:.2f}")
```

## Ejecutar la UI

```bash
pip install streamlit pandas openpyxl
streamlit run app.py
```

## Dependencias

```
pandas
openpyxl
streamlit  # solo para la UI
```

## Notas Importantes

1. **Sin tolerancia**: Las validaciones son exactas. Cualquier diferencia indica un problema en la extracción.

2. **Filas de Total**: El sistema excluye las filas marcadas con `tipo_fila` conteniendo "total" para evitar contar doble.

3. **Limpieza de números**: Maneja automáticamente:
   - Comas de miles: "1,234.56" → 1234.56
   - Signos negativos al final: "538.62-" → -538.62

4. **Escalabilidad**: Diseñado para procesar miles de clientes de forma consistente.
