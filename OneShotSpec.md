# One-Shot Spec: Conversor PDF→Excel Dinámico (Gallo + Visual)

## Resumen Ejecutivo

Aplicación web que convierte reportes financieros en PDF (formatos "Gallo" y "Visual") a Excel estructurado con validación automática. Detecta dinámicamente las secciones presentes en cada PDF y genera hojas correspondientes.

**Inputs:** PDFs de reportes financieros (Gallo o Visual)  
**Outputs:** Excel con múltiples hojas, validación de integridad matemática, reporte de errores

---

## 1. Flujo Principal

```
┌─────────────┐     ┌──────────────┐     ┌─────────────┐     ┌──────────────┐
│  Upload PDF │ ──▶ │ Detectar     │ ──▶ │ Extraer     │ ──▶ │ Post-proceso │
│  + Tipo     │     │ Secciones    │     │ por Sección │     │ + Validar    │
└─────────────┘     └──────────────┘     └─────────────┘     └──────────────┘
                                                                     │
                                                                     ▼
                                                            ┌──────────────┐
                                                            │ Generar XLSX │
                                                            │ + Descargar  │
                                                            └──────────────┘
```

---

## 2. Tipos de Reporte

### 2.1 Visual
Reporte de broker con operaciones de trading, rentas y dividendos.

**Secciones fijas:**
| Sección | Hoja Excel | Descripción |
|---------|------------|-------------|
| boletos | Boletos | Comprobantes de operaciones |
| resultado_ventas_ars | Resultado Ventas ARS | Operaciones en pesos |
| resultado_ventas_usd | Resultado Ventas USD | Operaciones en dólares |
| rentas_dividendos_ars | Rentas Dividendos ARS | Rentas/dividendos en pesos |
| rentas_dividendos_usd | Rentas Dividendos USD | Rentas/dividendos en dólares |
| resumen | Resumen | Totales consolidados |
| posicion_titulos | Posicion Titulos | Tenencias actuales |

### 2.2 Gallo
Reporte impositivo con resultados por categoría de instrumento.

**Secciones dinámicas** (detectadas de "Resultado Totales"):
| Categoría en PDF | Sección Key | Hoja Excel |
|------------------|-------------|------------|
| TIT.PRIVADOS EXENTOS | tit_privados_exentos | Tit.Privados Exentos |
| TIT.PRIVADOS DEL EXTERIOR | tit_privados_exterior | Tit.Privados Exterior |
| RENTA FIJA EN PESOS | renta_fija_pesos | Renta Fija Pesos |
| RENTA FIJA EN DOLARES | renta_fija_dolares | Renta Fija Dolares |
| CAUCIONES EN PESOS | cauciones_pesos | Cauciones Pesos |
| CAUCIONES EN DOLARES | cauciones_dolares | Cauciones Dolares |
| FCI | fci | FCI |
| OPCIONES | opciones | Opciones |
| FUTUROS | futuros | Futuros |

**Secciones fijas adicionales:**
- resultado_totales → Resultado Totales (siempre primero)
- posicion_inicial → Posicion Inicial (siempre al final)
- posicion_final → Posicion Final (siempre al final)

---

## 3. Detección Dinámica de Secciones (Gallo)

### Algoritmo
1. Extraer primero la tabla "Resultado Totales"
2. Leer cada fila de la columna `categoria`
3. Quitar el sufijo entre paréntesis: `"TIT.PRIVADOS EXENTOS (Enajenacion)"` → `"TIT.PRIVADOS EXENTOS"`
4. Mapear a sección usando tabla 2.2 (case-insensitive, tolerante a acentos)
5. Extraer solo las secciones detectadas + posiciones

### Ejemplo
Si Resultado Totales contiene:
```
- TIT.PRIVADOS EXENTOS (Enajenacion)
- TIT.PRIVADOS EXENTOS (Renta)
- CAUCIONES EN DOLARES (Enajenacion)
```
→ Extraer: `tit_privados_exentos`, `cauciones_dolares`, `posicion_inicial`, `posicion_final`

---

## 4. Schemas de Columnas por Hoja

### 4.1 Visual - Boletos
```json
["tipo_instrumento", "concertacion", "liquidacion", "nro_boleto", "moneda", 
 "tipo_operacion", "cod_instrumento", "instrumento", "cantidad", "precio", 
 "tipo_cambio", "bruto", "interes", "gastos", "neto"]
```

### 4.2 Visual - Resultado Ventas (ARS/USD)
```json
["tipo_instrumento", "instrumento", "cod_instrumento", "concertacion", 
 "liquidacion", "moneda", "tipo_operacion", "cantidad", "precio", "bruto", 
 "interes", "tipo_cambio", "gastos", "iva", "resultado"]
```

### 4.3 Visual - Rentas Dividendos (ARS/USD)
```json
["instrumento", "cod_instrumento", "categoria", "tipo_instrumento", 
 "concertacion", "liquidacion", "nro_operacion", "tipo_operacion", 
 "cantidad", "moneda", "tipo_cambio", "gastos", "importe"]
```

### 4.4 Visual - Resumen
```json
["moneda", "ventas", "fci", "opciones", "rentas", "dividendos", 
 "ef_cpd", "pagares", "futuros", "cau_int", "cau_cf", "total"]
```
- Filas: ARS, USD

### 4.5 Visual - Posicion Titulos
```json
["instrumento", "codigo", "ticker", "cantidad", "importe", "moneda"]
```

### 4.6 Gallo - Resultado Totales
```json
["categoria", "valor_pesos", "valor_usd"]
```

### 4.7 Gallo - Transacciones (Tit.*, Renta Fija *)
```json
["tipo_fila", "cod_especie", "especie", "fecha", "operacion", "numero", 
 "cantidad", "precio", "importe", "costo", "resultado_pesos", "resultado_usd", 
 "gastos_pesos", "gastos_usd"]
```
- `tipo_fila`: vacío para transacciones normales, `"Total Enajenacion"` o `"Total Renta"` para totales

### 4.8 Gallo - Cauciones (Pesos/Dolares)
```json
["tipo_fila", "cod_especie", "especie", "fecha", "vencimiento", "operacion", 
 "numero", "colocado", "al_vencimiento", "interes_pesos", "interes_usd", 
 "gastos_pesos", "gastos_usd"]
```

### 4.9 Gallo - Posición (Inicial/Final)
```json
["tipo_especie", "especie", "detalle", "custodia", "cantidad", "precio", 
 "importe_pesos", "pct_cartera_pesos", "importe_dolares", "pct_cartera_dolares"]
```

---

## 5. Reglas de Extracción (Instrucciones para el LLM)

### 5.1 Reglas Globales
```
REGLAS CRÍTICAS:
1. Devolver SOLO JSON válido. Sin texto adicional, sin markdown.
2. NUNCA inventar valores. Si una celda está vacía, usar 0 (cero). NUNCA null.
3. Extraer TODAS las páginas de la sección. No detenerse.
4. Campos desconocidos: ignorar (no agregar propiedades extra).
```

### 5.2 Formato Numérico
```
NÚMEROS:
- Formato PDF: punto = miles, coma = decimales (europeo: "1.234,56")
- Formato JSON: usar punto decimal americano ("1234.56")
- Conversión: "4.000,00" → 4000.0
```

### 5.3 Números Negativos

**Visual (paréntesis):**
```
- Valores entre paréntesis son NEGATIVOS
- "(10.000,00)" → -10000.00
- "(42.750,09)" → -42750.09
```

**Gallo (signo al final):**
```
- Signo menos (-) AL FINAL del número = NEGATIVO
- "5,212,573.58-" → -5212573.58
- MOVER el signo al inicio en el JSON
```

### 5.4 Continuidad de Especie (Gallo)
```
CONTINUIDAD:
1. El nombre de especie aparece solo en la primera fila de cada grupo
2. REPETIR el mismo valor para TODAS las filas siguientes hasta nuevo nombre
3. Si una página comienza sin encabezado de especie, usar la ÚLTIMA especie 
   de la página anterior
4. Incluir filas "Total Enajenacion" y "Total Renta" marcándolas en tipo_fila
```

### 5.5 Limpieza de Instrumentos (Visual)
```
INSTRUMENTOS:
- Quitar sufijo de moneda del nombre: " - Pesos", " - Dolar MEP", " - Dolar Cable"
- "CEDEAR APPLE INC. - Pesos" → "CEDEAR APPLE INC."
```

---

## 6. Post-Procesamiento

### 6.1 Parseo Numérico
```python
def parse_number(value):
    if value is None or value == "":
        return 0.0
    s = str(value).strip()
    # Quitar puntos de miles, coma a punto decimal
    s = s.replace(".", "").replace(",", ".")
    return float(s)
```

### 6.2 Paréntesis a Negativos (Visual)
```python
def convert_parenthesis_to_negative(value):
    s = str(value).strip()
    match = re.match(r'^\(([\d.,]+)\)$', s)
    if match:
        num = parse_number(match.group(1))
        return -num
    return parse_number(s)
```

### 6.3 Corrección Decimales x100 (Visual Resumen)

**Problema:** A veces el LLM extrae valores multiplicados por 100 (ignora decimales).

**Detección y corrección:**
```python
def fix_resumen_decimals(resumen_df, detail_data):
    """
    Compara valores del Resumen con sumas de hojas de detalle.
    Si el ratio es ~100, divide por 100.
    """
    # Calcular valores esperados de detalle
    calc_ventas_ars = sum(detail_data['resultado_ventas_ars']['resultado'])
    calc_dividendos_ars = sum(
        row['importe'] for row in detail_data['rentas_dividendos_ars'] 
        if row['categoria'].upper() == 'DIVIDENDOS'
    )
    
    # Para cada campo en Resumen
    for campo in ['ventas', 'rentas', 'dividendos']:
        valor_resumen = resumen_df[campo]
        valor_calculado = calc_values[campo]
        
        if valor_calculado != 0:
            ratio = abs(valor_resumen / valor_calculado)
            if 95 < ratio < 105:  # ~100x
                resumen_df[campo] = valor_resumen / 100
                print(f"[FIX] {campo}: {valor_resumen} → {valor_resumen/100}")
    
    # Recalcular total
    resumen_df['total'] = sum(campos_numericos)
    return resumen_df
```

### 6.4 De-duplicación
```python
# Eliminar duplicados por clave compuesta
df.drop_duplicates(
    subset=['instrumento', 'cod_instrumento', 'concertacion', 'tipo_operacion', 'cantidad'],
    keep='first'
)
```

---

## 7. Validación Matemática

### 7.1 Visual - Reglas de Validación

| Campo Resumen | Fórmula de Validación |
|---------------|----------------------|
| ventas ARS | `= SUM(Resultado Ventas ARS.resultado)` |
| ventas USD | `= SUM(Resultado Ventas USD.resultado)` |
| rentas ARS | `= SUMIF(Rentas Div ARS, categoria="RENTAS", importe)` |
| rentas USD | `= SUMIF(Rentas Div USD, categoria="Rentas", importe)` |
| dividendos ARS | `= SUMIF(Rentas Div ARS, categoria="DIVIDENDOS", importe)` |
| dividendos USD | `= SUMIF(Rentas Div USD, categoria="Dividendos", importe)` |
| total (por fila) | `= SUM(ventas, fci, opciones, rentas, dividendos, ef_cpd, pagares, futuros, cau_int, cau_cf)` |

### 7.2 Gallo - Reglas de Validación (Dinámicas)

Para cada fila en Resultado Totales (excepto TOTAL GENERAL):

```python
def validate_gallo_row(categoria, valor_pesos, valor_usd):
    # Extraer tipo: "TIT.PRIVADOS EXENTOS (Enajenacion)" → tipo="enajenacion"
    match = re.match(r'(.+)\s*\((.+)\)', categoria)
    cat_base = match.group(1)  # "TIT.PRIVADOS EXENTOS"
    tipo = match.group(2)       # "Enajenacion"
    
    sheet = get_sheet_for_categoria(cat_base)
    
    if "caucion" in cat_base.lower():
        # Cauciones: sumar intereses (excluyendo filas de total)
        calc = sum(row['interes_pesos'] for row in sheet if 'total' not in row['tipo_fila'].lower())
    else:
        # Otras: sumar filas donde tipo_fila contiene "Total" + tipo
        calc = sum(row['resultado_pesos'] for row in sheet 
                   if 'total' in row['tipo_fila'].lower() and tipo.lower() in row['tipo_fila'].lower())
    
    return abs(calc - valor_pesos) < TOLERANCE
```

### 7.3 Tolerancia
```python
TOLERANCE = 0.01  # Diferencia máxima aceptable
```

### 7.4 Reporte de Validación
```json
{
  "file": "archivo.xlsx",
  "type": "gallo",
  "passed": 7,
  "failed": 0,
  "results": [
    {"field": "TIT.PRIVADOS EXENTOS (Enajenacion) pesos", "calculated": 1076527.46, "expected": 1076527.46, "match": true},
    {"field": "CAUCIONES EN DOLARES (Enajenacion) usd", "calculated": 150.32, "expected": 150.32, "match": true}
  ]
}
```

---

## 8. Interfaz de Usuario

### 8.1 Layout
```
┌──────────────────────────────────────────────────────────────┐
│  📄 Conversor PDF → Excel (Gallo + Visual)                   │
├──────────────────────────────────────────────────────────────┤
│                                                              │
│  [Upload PDF Gallo]          [Upload PDF Visual]             │
│                                                              │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │ 🔄 Progreso: [████████████░░░░░░░░] 60%                 │ │
│  │    Procesando: Renta Fija Dolares...                    │ │
│  └─────────────────────────────────────────────────────────┘ │
│                                                              │
│  ✅ Validación: 7/7 Gallo | 6/6 Visual                       │
│                                                              │
│  [📥 Descargar Gallo.xlsx]  [📥 Descargar Visual.xlsx]       │
│                             [📥 Descargar Unificado.xlsx]    │
│                                                              │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │ Preview: [Resultado Totales ▼]                          │ │
│  │ ┌─────────────┬────────────┬───────────┐                │ │
│  │ │ categoria   │ valor_pesos│ valor_usd │                │ │
│  │ ├─────────────┼────────────┼───────────┤                │ │
│  │ │ TIT.PRIV... │ 1076527.46 │ 0.00      │                │ │
│  │ └─────────────┴────────────┴───────────┘                │ │
│  └─────────────────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────┘
```

### 8.2 Estados y Mensajes

| Estado | Mensaje |
|--------|---------|
| Inicial | "Sube un PDF para comenzar" |
| Procesando | "Procesando sección: {nombre}..." |
| Validando | "Validando integridad matemática..." |
| Éxito | "✅ Conversión completada. {N} hojas generadas." |
| Advertencia | "⚠️ Validación con diferencias: {campos}" |
| Error | "❌ Error en {sección}: {mensaje}" |

### 8.3 Tabla de Validación
```
┌─────────────────────────────────────┬────────────┬────────────┬───────┐
│ Campo                               │ Calculado  │ Esperado   │ Match │
├─────────────────────────────────────┼────────────┼────────────┼───────┤
│ TIT.PRIVADOS EXENTOS (Enajenacion)  │ 1076527.46 │ 1076527.46 │  ✅   │
│ RENTA FIJA EN DOLARES (Enajenacion) │ -5212573.58│ -5212573.58│  ✅   │
│ CAUCIONES EN DOLARES (Enajenacion)  │ 150.32     │ 150.32     │  ✅   │
└─────────────────────────────────────┴────────────┴────────────┴───────┘
```

---

## 9. Manejo de Errores

### 9.1 Errores Recuperables
| Error | Acción |
|-------|--------|
| JSON malformado | Intentar reparar (balancear llaves/corchetes) |
| Sección vacía | Crear hoja vacía, continuar con otras |
| Columna faltante | Usar 0 para valores numéricos, "" para texto |

### 9.2 Errores Fatales
| Error | Mensaje al Usuario |
|-------|-------------------|
| PDF corrupto | "El archivo PDF no se puede leer" |
| Tipo no reconocido | "No se detectó formato Gallo ni Visual" |
| Sin secciones | "No se encontraron datos para extraer" |

### 9.3 Logging
```
[INFO] Procesando: resultado_totales
[INFO]   → 8 registros extraídos
[INFO] Secciones detectadas: ['tit_privados_exentos', 'cauciones_dolares']
[WARN] Sección 'fci' no encontrada en PDF
[FIX] ARS dividendos: -4275009.00 → -42750.09 (corregido x100)
[OK] Validación: 7/7 passed
```

---

## 10. Configuración

```yaml
# config.yaml
validation:
  tolerance: 0.01              # Diferencia máxima aceptable
  decimal_fix_ratio_min: 95    # Ratio mínimo para detectar x100
  decimal_fix_ratio_max: 105   # Ratio máximo para detectar x100

excel:
  rounding_decimals: 2
  date_format: "dd/mm/yyyy"

extraction:
  max_retries: 3
  temperature: 0.0             # Determinístico

ui:
  show_preview: true
  max_preview_rows: 100
```

---

## 11. Casos de Prueba

### 11.1 Casos Básicos
| Caso | Input | Validación Esperada |
|------|-------|---------------------|
| Visual simple | PDF con 2 instrumentos | 6/6 validaciones pasan |
| Gallo simple | PDF con 3 secciones | Detecta 3 secciones, valida cada una |

### 11.2 Casos Edge
| Caso | Input | Comportamiento Esperado |
|------|-------|------------------------|
| Decimales x100 | Visual con dividendos mal | Auto-corrige y recalcula total |
| Sección nueva | Gallo con "cauciones dolares" | Detecta, extrae, valida dinámicamente |
| Múltiples páginas | Sección que cruza 5 páginas | Continuidad de especie preservada |
| Números negativos | "(42.750,09)" o "5.212,58-" | Convierte correctamente a -42750.09 |

### 11.3 Golden Files
Mantener archivos XLSX de referencia para comparación:
- `tests/golden/Visual_OK.xlsx`
- `tests/golden/Gallo_OK.xlsx`
- `tests/golden/Gallo_con_cauciones_dolares.xlsx`

---

## 12. Estructura de Archivos del Proyecto

```
/
├── app.py                    # Entrada principal (UI)
├── extractor/
│   ├── gallo.py              # Extracción Gallo (dinámico)
│   ├── visual.py             # Extracción Visual
│   └── prompts/              # Templates de prompts por sección
├── postprocess/
│   ├── numbers.py            # Parseo numérico
│   ├── decimals_fix.py       # Corrección x100
│   └── cleanup.py            # Limpieza y de-dup
├── validation/
│   ├── visual.py             # Validación Visual
│   └── gallo.py              # Validación Gallo (dinámica)
├── export/
│   └── excel_writer.py       # Generador XLSX
├── config.yaml               # Configuración
└── tests/
    ├── data/                 # PDFs de prueba
    └── golden/               # XLSX de referencia
```

---

## Apéndice A: Mapeo Completo Categoría → Hoja (Gallo)

```python
CATEGORIA_TO_SHEET = {
    'tit.privados exentos': 'Tit.Privados Exentos',
    'tit privados exentos': 'Tit.Privados Exentos',
    'titulos privados exentos': 'Tit.Privados Exentos',
    'tit.privados del exterior': 'Tit.Privados Exterior',
    'tit privados del exterior': 'Tit.Privados Exterior',
    'renta fija en pesos': 'Renta Fija Pesos',
    'renta fija pesos': 'Renta Fija Pesos',
    'renta fija en dolares': 'Renta Fija Dolares',
    'renta fija dolares': 'Renta Fija Dolares',
    'renta fija en dólares': 'Renta Fija Dolares',
    'cauciones en pesos': 'Cauciones Pesos',
    'cauciones pesos': 'Cauciones Pesos',
    'cauciones en dolares': 'Cauciones Dolares',
    'cauciones dolares': 'Cauciones Dolares',
    'cauciones en dólares': 'Cauciones Dolares',
    'fci': 'FCI',
    'fondos comunes de inversion': 'FCI',
    'opciones': 'Opciones',
    'futuros': 'Futuros',
}
```

## Apéndice B: Fórmulas Visual Resumen → Detalle

```python
VISUAL_VALIDATIONS = {
    'ventas': {
        'ARS': {'sheet': 'Resultado Ventas ARS', 'formula': 'SUM(resultado)'},
        'USD': {'sheet': 'Resultado Ventas USD', 'formula': 'SUM(resultado)'},
    },
    'rentas': {
        'ARS': {'sheet': 'Rentas Dividendos ARS', 'formula': 'SUMIF(categoria="RENTAS", importe)'},
        'USD': {'sheet': 'Rentas Dividendos USD', 'formula': 'SUMIF(categoria="Rentas", importe)'},
    },
    'dividendos': {
        'ARS': {'sheet': 'Rentas Dividendos ARS', 'formula': 'SUMIF(categoria="DIVIDENDOS", importe)'},
        'USD': {'sheet': 'Rentas Dividendos USD', 'formula': 'SUMIF(categoria="Dividendos", importe)'},
    },
    'total': {
        'formula': 'SUM(ventas, fci, opciones, rentas, dividendos, ef_cpd, pagares, futuros, cau_int, cau_cf)'
    }
}
```

---

## Apéndice C: Prompts de Extracción por Sección

### C.1 Gallo - Resultado Totales
```
Extrae SOLO la tabla "RESULTADO TOTALES" del PDF de Gallo.
Es la tabla resumen al inicio con categorias y sus valores en pesos/USD.

MUY IMPORTANTE - NUMEROS NEGATIVOS:
Si un numero tiene el signo menos (-) AL FINAL del numero, es un numero NEGATIVO.
Ejemplo: "5,212,573.58-" en el PDF significa -5212573.58 en el JSON
Debes MOVER el signo menos al INICIO del numero en tu respuesta.

Devuelve JSON:
{"resultado_totales": [
    {"categoria": "TIT.PRIVADOS EXENTOS (Enajenacion)", "valor_pesos": -5212573.58, "valor_usd": 0}
]}
Usa 0 para valores vacios, NUNCA null.
```

### C.2 Gallo - Secciones de Transacciones
```
Extrae COMPLETAMENTE la seccion "[NOMBRE_SECCION]" del PDF.
Esta seccion puede abarcar MULTIPLES PAGINAS - revisa TODAS hasta encontrar el siguiente encabezado.

ESTRUCTURA:
- Cada especie tiene un encabezado con codigo y nombre (ej: "00007 ALUA ALUAR")
- Debajo del encabezado hay multiples filas de transacciones
- Al final de cada especie puede haber filas "Total Enajenacion" y/o "Total Renta"

COLUMNAS (14 total):
cod_especie, especie, fecha, operacion, numero, cantidad, precio, importe, costo, 
resultado_pesos, resultado_usd, gastos_pesos, gastos_usd, tipo_fila

tipo_fila: vacio para transacciones normales, "Total Enajenacion" o "Total Renta" para totales.

REGLAS CRITICAS:
1. CONTINUIDAD DE ESPECIE: El nombre de especie aparece solo en la primera fila. 
   REPITE el mismo valor para TODAS las filas siguientes hasta nuevo nombre.
2. CONTINUIDAD ENTRE PAGINAS: Si una pagina comienza sin encabezado de especie, 
   pertenecen a la ULTIMA especie de la pagina anterior.
3. NUMEROS NEGATIVOS: Si un numero tiene el signo - AL FINAL, es NEGATIVO.
4. CELDAS VACIAS = 0: Si una celda esta vacia, usa 0 (cero). NUNCA null.
5. EXTRAE TODO: Recorre TODAS las paginas de la seccion. No te detengas.
6. NO INVENTES: Solo extrae datos que existen en el PDF.

Devuelve: {"[nombre_seccion]": [...]}
```

### C.3 Visual - Resumen
```
Extrae la tabla "RESUMEN" del PDF Visual.
Es la tabla al inicio con totales por moneda (ARS, USD) y categoría.

COLUMNAS: moneda, ventas, fci, opciones, rentas, dividendos, ef_cpd, pagares, futuros, cau_int, cau_cf, total

NUMEROS NEGATIVOS: Valores entre parentesis son NEGATIVOS.
"(42.750,09)" → -42750.09

FORMATO: punto=miles, coma=decimales en PDF. Usar punto decimal en JSON.

Devuelve: {"resumen": [
    {"moneda": "ARS", "ventas": 0, "fci": 0, ..., "total": -42750.09},
    {"moneda": "USD", "ventas": -14670053.15, ..., "total": -14667979.69}
]}
```

### C.4 Visual - Resultado Ventas
```
Extrae TODA la seccion "RESULTADO DE VENTAS [ARS/USD]" del PDF Visual.

COLUMNAS: tipo_instrumento, instrumento, cod_instrumento, concertacion, liquidacion, 
moneda, tipo_operacion, cantidad, precio, bruto, interes, tipo_cambio, gastos, iva, resultado

LIMPIEZA: Quitar sufijo de moneda del instrumento: " - Pesos", " - Dolar MEP", " - Dolar Cable"

NUMEROS NEGATIVOS: Valores entre parentesis son NEGATIVOS.

Devuelve: {"resultado_ventas_[ars/usd]": [...]}
```

---

## Apéndice D: Ejemplos de PDFs de Entrada

### D.1 Gallo - Estructura típica
```
RESULTADO TOTALES
┌─────────────────────────────────────┬────────────┬────────────┐
│ categoria                           │ valor_pesos│ valor_usd  │
├─────────────────────────────────────┼────────────┼────────────┤
│ TIT.PRIVADOS EXENTOS (Enajenacion)  │ 1076527.46 │ 0.00       │
│ TIT.PRIVADOS EXENTOS (Renta)        │ 785.78     │ 0.00       │
│ RENTA FIJA EN DOLARES (Enajenacion) │ 0.00       │ 5212573.58-│  ← Negativo!
│ CAUCIONES EN DOLARES (Enajenacion)  │ 0.00       │ 150.32     │  ← Nueva sección
│ TOTAL GENERAL                       │ 1114800.56 │ 5212502.63-│
└─────────────────────────────────────┴────────────┴────────────┘

TIT.PRIVADOS EXENTOS
00007 ALUA ALUAR
  01/03/2025  VENTA     123    1000   15.50   15500.00  14000.00  1500.00  0.00  50.00  0.00
  15/03/2025  VENTA     124    500    16.00   8000.00   7200.00   800.00   0.00  25.00  0.00
              Total Enajenacion                                   2300.00  0.00  75.00  0.00
...
```

### D.2 Visual - Estructura típica
```
RESUMEN
┌────────┬──────────────┬─────┬─────────┬────────┬────────────┬───────┐
│ Moneda │ Ventas       │ FCI │ Opciones│ Rentas │ Dividendos │ Total │
├────────┼──────────────┼─────┼─────────┼────────┼────────────┼───────┤
│ ARS    │ 0.00         │ 0   │ 0       │ 0.00   │ (42.750,09)│ ...   │  ← Paréntesis!
│ USD    │ (14.670.053) │ 0   │ 0       │ 2073.32│ 0.14       │ ...   │
└────────┴──────────────┴─────┴─────────┴────────┴────────────┴───────┘

RESULTADO DE VENTAS EN DOLARES
┌──────────────────────────────────────┬────────┬──────────┬──────────────┐
│ Instrumento                          │ Moneda │ Cantidad │ Resultado    │
├──────────────────────────────────────┼────────┼──────────┼──────────────┤
│ CEDEAR APPLE INC. - Dolar MEP        │ USD    │ 100      │ (1.500,00)   │
│ CEDEAR NVIDIA - Dolar Cable          │ USD    │ 50       │ 2.300,50     │
└──────────────────────────────────────┴────────┴──────────┴──────────────┘
```
