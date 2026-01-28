"""
Extraction prompts for Gallo and Visual report types.
Contains prompt templates for each section type.
"""

# =============================================================================
# GALLO PROMPTS
# =============================================================================

GALLO_RESULTADO_TOTALES = """Extrae SOLO la tabla "RESULTADO TOTALES" del PDF de Gallo.
Es la tabla resumen al inicio con categorias y sus valores en pesos/USD.

MUY IMPORTANTE - NÚMEROS NEGATIVOS:
Si un número tiene el signo menos (-) AL FINAL del número, es un número NEGATIVO.
Ejemplo: "5,212,573.58-" en el PDF significa -5212573.58 en el JSON
Debes MOVER el signo menos al INICIO del número en tu respuesta.

FORMATO DE NÚMEROS:
- En el PDF: punto = miles, coma = decimales (formato europeo: "1.234,56")
- En el JSON: usar punto decimal americano (1234.56)
- Conversión: "4.000,00" → 4000.0

Devuelve JSON con esta estructura exacta:
{{"resultado_totales": [
    {{"categoria": "TIT.PRIVADOS EXENTOS (Enajenacion)", "valor_pesos": 1076527.46, "valor_usd": 0}},
    {{"categoria": "RENTA FIJA EN DOLARES (Enajenacion)", "valor_pesos": 0, "valor_usd": -5212573.58}}
]}}

Usa 0 para valores vacíos, NUNCA null.
NO incluyas la fila "TOTAL GENERAL" si existe.

TEXTO DEL PDF:
{text}"""


GALLO_TRANSACCIONES = """Extrae COMPLETAMENTE la sección "{section_name}" del PDF.
Esta sección puede abarcar MÚLTIPLES PÁGINAS - revisa TODAS hasta el final.

ESTRUCTURA:
- Cada especie tiene un encabezado con código y nombre (ej: "00007 ALUA ALUAR")
- Debajo del encabezado hay múltiples filas de transacciones
- Al final de cada especie puede haber filas "Total Enajenacion" y/o "Total Renta"

COLUMNAS (14 total):
tipo_fila, cod_especie, especie, fecha, operacion, numero, cantidad, precio, importe, costo, 
resultado_pesos, resultado_usd, gastos_pesos, gastos_usd

CAMPO tipo_fila:
- Vacío ("") para transacciones normales
- "Total Enajenacion" para filas de total de enajenación
- "Total Renta" para filas de total de renta

REGLAS CRÍTICAS:
1. CONTINUIDAD DE ESPECIE: El nombre de especie aparece solo en la primera fila. 
   REPITE el mismo valor para TODAS las filas siguientes hasta nuevo nombre.
2. CONTINUIDAD ENTRE PÁGINAS: Si una página comienza sin encabezado de especie, 
   las filas pertenecen a la ÚLTIMA especie de la página anterior.
3. NÚMEROS NEGATIVOS: Si un número tiene el signo - AL FINAL, es NEGATIVO.
   "5,212,573.58-" → -5212573.58
4. CELDAS VACÍAS = 0: Si una celda está vacía, usa 0 (cero). NUNCA null.
5. FORMATO NÚMEROS: Punto=miles, coma=decimales en PDF → punto decimal en JSON
6. EXTRAE TODO: Recorre TODAS las páginas de la sección. No te detengas.
7. NO INVENTES: Solo extrae datos que existen en el PDF.

Devuelve JSON:
{{"{section_key}": [
    {{"tipo_fila": "", "cod_especie": "00007", "especie": "ALUA ALUAR", "fecha": "01/03/2025", 
      "operacion": "VENTA", "numero": "123", "cantidad": 1000, "precio": 15.50, 
      "importe": 15500.00, "costo": 14000.00, "resultado_pesos": 1500.00, "resultado_usd": 0,
      "gastos_pesos": 50.00, "gastos_usd": 0}},
    {{"tipo_fila": "Total Enajenacion", "cod_especie": "00007", "especie": "ALUA ALUAR", 
      "fecha": "", "operacion": "", "numero": "", "cantidad": 0, "precio": 0,
      "importe": 0, "costo": 0, "resultado_pesos": 2300.00, "resultado_usd": 0,
      "gastos_pesos": 75.00, "gastos_usd": 0}}
]}}

TEXTO DEL PDF:
{text}"""


GALLO_TRANSACCIONES_CONTINUATION = """CONTEXTO DE CONTINUIDAD:
La última especie procesada fue:
  cod_especie: "{cod_especie}"
  especie: "{especie}"

Si la página comienza con transacciones sin encabezado de especie,
asigna estos valores a esas filas.

"""


GALLO_CAUCIONES = """Extrae COMPLETAMENTE la sección "CAUCIONES EN {currency}" del PDF.
Esta sección contiene operaciones de caución (colocaciones a plazo).

COLUMNAS (13 total):
tipo_fila, cod_especie, especie, fecha, vencimiento, operacion, numero, 
colocado, al_vencimiento, interes_pesos, interes_usd, gastos_pesos, gastos_usd

CAMPO tipo_fila:
- Vacío ("") para transacciones normales
- "Total" para filas de total

REGLAS CRÍTICAS:
1. CONTINUIDAD DE ESPECIE: Igual que transacciones
2. NÚMEROS NEGATIVOS: Signo - al final = negativo
3. CELDAS VACÍAS = 0: NUNCA null
4. FORMATO NÚMEROS: Punto=miles, coma=decimales → punto decimal

Devuelve JSON:
{{"cauciones_{currency_key}": [
    {{"tipo_fila": "", "cod_especie": "CAUCI", "especie": "CAUCION PESOS", 
      "fecha": "01/03/2025", "vencimiento": "08/03/2025", "operacion": "COLOCACION",
      "numero": "456", "colocado": 100000.00, "al_vencimiento": 100500.00,
      "interes_pesos": 500.00, "interes_usd": 0, "gastos_pesos": 10.00, "gastos_usd": 0}}
]}}

TEXTO DEL PDF:
{text}"""


GALLO_POSICION = """Extrae COMPLETAMENTE la tabla "POSICIÓN {position_type}" del PDF.
Esta tabla muestra la tenencia de títulos al inicio/final del período.

COLUMNAS (10 total):
tipo_especie, especie, detalle, custodia, cantidad, precio,
importe_pesos, pct_cartera_pesos, importe_dolares, pct_cartera_dolares

REGLAS:
1. tipo_especie puede ser: "ACCIONES", "BONOS", "CEDEARS", "FCI", etc.
2. Porcentajes: mantener como decimal (5.25% → 5.25)
3. CELDAS VACÍAS = 0 para números, "" para texto
4. FORMATO NÚMEROS: Punto=miles, coma=decimales → punto decimal

Devuelve JSON:
{{"posicion_{position_key}": [
    {{"tipo_especie": "ACCIONES", "especie": "ALUAR ALUMINIO", "detalle": "ALUA",
      "custodia": "CAJA DE VALORES", "cantidad": 1000, "precio": 850.50,
      "importe_pesos": 850500.00, "pct_cartera_pesos": 15.25,
      "importe_dolares": 850.50, "pct_cartera_dolares": 12.30}}
]}}

TEXTO DEL PDF:
{text}"""


# =============================================================================
# VISUAL PROMPTS
# =============================================================================

VISUAL_RESUMEN = """Extrae la tabla "RESUMEN" del PDF Visual.
Es la tabla al inicio con totales por moneda (ARS, USD) y categoría.

COLUMNAS: moneda, ventas, fci, opciones, rentas, dividendos, ef_cpd, pagares, futuros, cau_int, cau_cf, total

NÚMEROS NEGATIVOS: Valores entre paréntesis son NEGATIVOS.
"(42.750,09)" → -42750.09
"(14.670.053,15)" → -14670053.15

FORMATO: punto=miles, coma=decimales en PDF. Usar punto decimal en JSON.

Devuelve JSON:
{{"resumen": [
    {{"moneda": "ARS", "ventas": 0, "fci": 0, "opciones": 0, "rentas": 0, 
      "dividendos": -42750.09, "ef_cpd": 0, "pagares": 0, "futuros": 0, 
      "cau_int": 0, "cau_cf": 0, "total": -42750.09}},
    {{"moneda": "USD", "ventas": -14670053.15, "fci": 0, "opciones": 0, "rentas": 2073.32, 
      "dividendos": 0.14, "ef_cpd": 0, "pagares": 0, "futuros": 0,
      "cau_int": 0, "cau_cf": 0, "total": -14667979.69}}
]}}

TEXTO DEL PDF:
{text}"""


VISUAL_BOLETOS = """Extrae TODOS los boletos de la sección "BOLETOS" del PDF Visual.
Pueden ser múltiples páginas - extrae TODOS.

COLUMNAS (15 total):
tipo_instrumento, concertacion, liquidacion, nro_boleto, moneda, tipo_operacion,
cod_instrumento, instrumento, cantidad, precio, tipo_cambio, bruto, interes, gastos, neto

LIMPIEZA DE INSTRUMENTOS:
Quitar sufijo de moneda del nombre:
- "CEDEAR APPLE INC. - Pesos" → "CEDEAR APPLE INC."
- "BONO TX28 - Dolar MEP" → "BONO TX28"

NÚMEROS NEGATIVOS: Valores entre paréntesis son NEGATIVOS.

Devuelve JSON:
{{"boletos": [
    {{"tipo_instrumento": "CEDEAR", "concertacion": "15/03/2025", "liquidacion": "17/03/2025",
      "nro_boleto": "12345", "moneda": "ARS", "tipo_operacion": "COMPRA",
      "cod_instrumento": "AAPL", "instrumento": "CEDEAR APPLE INC.", "cantidad": 100,
      "precio": 15000.50, "tipo_cambio": 1050.00, "bruto": 1500050.00, 
      "interes": 0, "gastos": 1500.05, "neto": 1501550.05}}
]}}

TEXTO DEL PDF:
{text}"""


VISUAL_RESULTADO_VENTAS = """Extrae TODA la sección "RESULTADO DE VENTAS EN {currency}" del PDF Visual.
Pueden ser múltiples páginas - extrae TODAS las operaciones.

COLUMNAS (15 total):
tipo_instrumento, instrumento, cod_instrumento, concertacion, liquidacion,
moneda, tipo_operacion, cantidad, precio, bruto, interes, tipo_cambio, gastos, iva, resultado

LIMPIEZA DE INSTRUMENTOS:
Quitar sufijo de moneda: " - Pesos", " - Dolar MEP", " - Dolar Cable"

NÚMEROS NEGATIVOS: Valores entre paréntesis son NEGATIVOS.
"(1.500,00)" → -1500.00

FORMATO: punto=miles, coma=decimales → punto decimal en JSON

Devuelve JSON:
{{"resultado_ventas_{currency_key}": [
    {{"tipo_instrumento": "CEDEAR", "instrumento": "CEDEAR APPLE INC.", 
      "cod_instrumento": "AAPL", "concertacion": "15/03/2025", "liquidacion": "17/03/2025",
      "moneda": "{currency}", "tipo_operacion": "VENTA", "cantidad": 100,
      "precio": 15000.50, "bruto": 1500050.00, "interes": 0, "tipo_cambio": 1050.00,
      "gastos": 1500.05, "iva": 315.01, "resultado": -1500.00}}
]}}

TEXTO DEL PDF:
{text}"""


VISUAL_RENTAS_DIVIDENDOS = """Extrae TODA la sección "RENTAS Y DIVIDENDOS EN {currency}" del PDF Visual.
Pueden ser múltiples páginas - extrae TODOS los registros.

COLUMNAS (13 total):
instrumento, cod_instrumento, categoria, tipo_instrumento, concertacion, liquidacion,
nro_operacion, tipo_operacion, cantidad, moneda, tipo_cambio, gastos, importe

CAMPO categoria:
- "RENTAS" para pagos de intereses/rentas
- "DIVIDENDOS" para pagos de dividendos

LIMPIEZA DE INSTRUMENTOS:
Quitar sufijo de moneda: " - Pesos", " - Dolar MEP", " - Dolar Cable"

NÚMEROS NEGATIVOS: Valores entre paréntesis son NEGATIVOS.

Devuelve JSON:
{{"rentas_dividendos_{currency_key}": [
    {{"instrumento": "BONO AL30", "cod_instrumento": "AL30", "categoria": "RENTAS",
      "tipo_instrumento": "BONO", "concertacion": "01/04/2025", "liquidacion": "03/04/2025",
      "nro_operacion": "789", "tipo_operacion": "COBRO", "cantidad": 1000,
      "moneda": "{currency}", "tipo_cambio": 1050.00, "gastos": 10.50, "importe": 2073.32}}
]}}

TEXTO DEL PDF:
{text}"""


VISUAL_POSICION_TITULOS = """Extrae la tabla "POSICIÓN DE TÍTULOS" del PDF Visual.
Muestra la tenencia actual de instrumentos.

COLUMNAS (6 total):
instrumento, codigo, ticker, cantidad, importe, moneda

LIMPIEZA DE INSTRUMENTOS:
Quitar sufijo de moneda: " - Pesos", " - Dolar MEP", " - Dolar Cable"

FORMATO: punto=miles, coma=decimales → punto decimal

Devuelve JSON:
{{"posicion_titulos": [
    {{"instrumento": "CEDEAR APPLE INC.", "codigo": "AAPL", "ticker": "AAPL",
      "cantidad": 500, "importe": 7500000.00, "moneda": "ARS"}}
]}}

TEXTO DEL PDF:
{text}"""


# =============================================================================
# PROMPT REGISTRY
# =============================================================================

GALLO_PROMPTS = {
    "resultado_totales": GALLO_RESULTADO_TOTALES,
    "transacciones": GALLO_TRANSACCIONES,
    "transacciones_continuation": GALLO_TRANSACCIONES_CONTINUATION,
    "cauciones": GALLO_CAUCIONES,
    "posicion": GALLO_POSICION,
}

VISUAL_PROMPTS = {
    "resumen": VISUAL_RESUMEN,
    "boletos": VISUAL_BOLETOS,
    "resultado_ventas": VISUAL_RESULTADO_VENTAS,
    "rentas_dividendos": VISUAL_RENTAS_DIVIDENDOS,
    "posicion_titulos": VISUAL_POSICION_TITULOS,
}
