# Documentación Técnica: Merge Gallo + Visual

## Resumen Impositivo Anual - Guía Completa del Proceso

**Versión:** 1.0  
**Fecha:** Febrero 2026  
**Plataforma:** https://big-pdf-to-excel-converter.streamlit.app/

---

## Índice

1. [Flujo General del Proceso](#1-flujo-general-del-proceso)
2. [Archivos de Entrada](#2-archivos-de-entrada)
3. [Hojas Auxiliares](#3-hojas-auxiliares)
4. [Hojas Generadas en el Merge](#4-hojas-generadas-en-el-merge)
5. [Detalle de Cálculos por Hoja](#5-detalle-de-cálculos-por-hoja)
6. [Tratamientos Especiales](#6-tratamientos-especiales)
7. [Mapeo de Operaciones Gallo → Visual](#7-mapeo-de-operaciones-gallo--visual)
8. [Fórmulas Excel Utilizadas](#8-fórmulas-excel-utilizadas)
9. [Edge Cases Conocidos](#9-edge-cases-conocidos)

---

## 1. Flujo General del Proceso

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   PDF Gallo     │     │   PDF Visual    │     │  Hojas Aux.     │
│  (OCR → Excel)  │     │  (OCR → Excel)  │     │  (4 archivos)   │
└────────┬────────┘     └────────┬────────┘     └────────┬────────┘
         │                       │                       │
         └───────────────────────┴───────────────────────┘
                                 │
                                 ▼
                    ┌───────────────────────┐
                    │   MERGE CONSOLIDADO   │
                    │   (14 hojas Excel)    │
                    └───────────────────────┘
                                 │
                                 ▼
                    ┌───────────────────────┐
                    │   EXPORTAR A PDF      │
                    │   (formato Visual)    │
                    └───────────────────────┘
```

### Pasos del proceso:

1. **Upload PDFs**: El usuario sube los PDFs de Gallo y Visual
2. **OCR Fine-tuneado**: Motor de extracción convierte cada PDF a Excel estructurado
3. **Post-procesamiento**: Limpieza de datos, normalización de formatos
4. **Merge Automático**: Unificación de ambos Excel en un consolidado
5. **Exportación PDF**: Generación de PDF con formato Visual

---

## 2. Archivos de Entrada

### 2.1 Excel Gallo (generado del PDF)

| Hoja | Descripción |
|------|-------------|
| Posicion Inicial | Tenencias al inicio del período (1/1) |
| Posicion Final | Tenencias al cierre del período (1/7) |
| Tit Privados Exentos | Acciones, CEDEARs en pesos |
| Renta Fija Dolares | Bonos, ONs en dólares |
| Tit Privados Exterior | Bonos cable (exterior) |
| Cauciones | Operaciones de caución |
| Titulos Publicos | Letras, Lecaps en pesos |
| Cedears | CEDEARs (renta variable) |
| Resultados | Totales por categoría |

**Estructura de transacciones Gallo:**
| Columna | Contenido |
|---------|-----------|
| A | tipo_fila |
| B | cod_especie |
| C | especie (TICKER + Nombre) |
| D | fecha |
| E | operacion |
| F | numero |
| G | cantidad |
| H | precio |
| I | importe |
| J | costo |
| K | resultado_pesos |
| L | resultado_usd |
| M | gastos_pesos |
| N | gastos_usd |

### 2.2 Excel Visual (generado del PDF)

| Hoja | Descripción |
|------|-------------|
| Boletos | Comprobantes de operaciones |
| Resultado Ventas ARS | Operaciones con moneda_emision = Pesos |
| Resultado Ventas USD | Operaciones con moneda_emision = Dólar |
| Rentas Dividendos ARS | Rentas/Dividendos en pesos |
| Rentas Dividendos USD | Rentas/Dividendos en dólares |
| Cauciones Tomadoras | Cauciones donde el comitente toma prestado |
| Cauciones Colocadoras | Cauciones donde el comitente coloca fondos |
| Resumen | Totales consolidados |
| Posicion Titulos | Tenencias actuales |

---

## 3. Hojas Auxiliares

### 3.1 EspeciesVisual.xlsx

**Propósito:** Catálogo maestro de todas las especies del mercado con sus atributos.

| Columna | Campo | Uso |
|---------|-------|-----|
| C | Código | Clave primaria para VLOOKUP |
| G | moneda_emision | Determina si es ARS o USD |
| H | ticker | Símbolo corto |
| Q | nombre_con_moneda | Nombre completo + moneda |
| R | tipo_especie | Acciones, Títulos Públicos, ONs, etc. |

**¿Por qué es crítico?**
- Determina en qué hoja de Resultado Ventas va cada operación (ARS vs USD)
- Proporciona el nombre estandarizado del instrumento
- Clasifica el tipo de instrumento para agrupaciones

### 3.2 EspeciesGallo.xlsx

**Propósito:** Mapeo de códigos Gallo a información adicional.

| Columna | Campo | Uso |
|---------|-------|-----|
| A | Código | Clave primaria |
| B | Nombre | Descripción |
| J | Ticker | Símbolo |
| N | moneda_emision | Moneda de emisión |

### 3.3 Cotizacion_Dolar_Historica.xlsx

**Propósito:** Histórico de cotizaciones del dólar por fecha.

| Columna | Campo |
|---------|-------|
| A | Fecha |
| B | Cotización (en pesos) |
| C | Tipo (Dolar MEP local, Dolar Cable, etc.) |

**Uso principal:**
- Calcular tipo de cambio para operaciones
- Convertir precios de USD a pesos y viceversa
- Valorizar posiciones en diferentes monedas


### 3.4 PreciosInicialesEspecies.xlsx

**Propósito:** Precios de costo de las especies al inicio del período fiscal (1/1).

| Columna | Campo | Uso |
|---------|-------|-----|
| A | Código especie | Identificador |
| C | ORDEN/Ticker | Clave de búsqueda |
| G | Precio | Precio de costo inicial |

**¿Por qué es crítico para el costo de venta?**

El **costo de venta** se calcula usando el precio promedio ponderado del stock. Para la **primera operación de venta** de cada especie, necesitamos conocer:

1. **Cantidad inicial** (de Posicion Inicial Gallo)
2. **Precio inicial** (de PreciosInicialesEspecies)

Esto permite calcular:
```
Costo por venta = Cantidad vendida × Precio Promedio Stock
```


---

## 4. Hojas Generadas en el Merge

El merge genera **14 hojas** en el Excel consolidado:

| # | Hoja | Fuente | Descripción |
|---|------|--------|-------------|
| 1 | Posicion Inicial Gallo | Gallo | Tenencias al 31/5 con precios |
| 2 | Posicion Final Gallo | Gallo | Tenencias al 31/12 |
| 3 | Boletos | Gallo + Visual | Todas las transacciones de compra/venta |
| 4 | Cauciones Tomadoras | Gallo + Visual | Operaciones de caución tomadora (TOM) |
| 5 | Cauciones Colocadoras | Gallo + Visual | Operaciones de caución colocadora (COL) |
| 6 | Rentas y Dividendos Gallo | Gallo | Rentas y dividendos originales |
| 7 | Resultado Ventas ARS | Boletos filtrado | Operaciones en pesos con running stock |
| 8 | Resultado Ventas USD | Boletos filtrado | Operaciones en dólares con running stock |
| 9 | Rentas Dividendos ARS | R&D Gallo filtrado | Rentas/Dividendos en pesos |
| 10 | Rentas Dividendos USD | R&D Gallo filtrado | Rentas/Dividendos en dólares |
| 11 | Resumen | Calculado | Totales por categoría |
| 12 | Posicion Titulos | **Visual** | Tenencias finales (desde Visual, no Gallo) |
| 13 | EspeciesVisual | Auxiliar | Catálogo de especies |
| 14 | EspeciesGallo | Auxiliar | Mapeo especies Gallo |
| 15 | Cotizacion Dolar Historica | Auxiliar | Histórico TC |
| 16 | PreciosInicialesEspecies | Auxiliar | Precios de costo |

---

## 5. Detalle de Cálculos por Hoja

### 5.1 Posición Inicial/Final Gallo

**Estructura (20 columnas):**

| Col | Campo | Cálculo |
|-----|-------|---------|
| A | tipo_especie | Original de Gallo |
| B | Ticker | Primera palabra de especie |
| C | especie | Resto del nombre |
| D | Codigo especie | VLOOKUP en PreciosInicialesEspecies |
| E | Codigo Especie Origen | "PreciosInicialesEspecies" o "Gallo" |
| I | cantidad | Original |
| J | precio Tenencia Inicial Pesos | importe_pesos / cantidad |
| K | precio Tenencia Inicial USD | importe_usd / cantidad |
| L | Precio de PreciosIniciales | VLOOKUP por ticker |
| P | Precio a Utilizar | =PreciosInicialesEspecies |

**Nota especial para Renta Fija Dólares:**
- El precio viene dividido por 100 en Gallo
- Se multiplica x100 al importar: `precio_pesos = (importe_pesos / cantidad) * 100`

### 5.2 Boletos

**Columnas (19):**

| Col | Campo | Cálculo |
|-----|-------|---------|
| A | Tipo de Instrumento | `=VLOOKUP(G,EspeciesVisual!C:R,16,FALSE)` |
| B | Concertación | Fecha de la operación |
| C | Liquidación | Fecha liquidación (puede estar vacía) |
| D | Nro. Boleto | Número de comprobante |
| E | Moneda | Determinada por la hoja de origen |
| F | Tipo Operación | COMPRA, VENTA, LICITACION, etc. |
| G | Cod.Instrum | Código numérico de especie |
| H | Instrumento Crudo | Nombre original |
| I | InstrumentoConMoneda | `=VLOOKUP(G,EspeciesVisual!C:Q,15,FALSE)` |
| J | Cantidad | Cantidad operada (negativo = venta) |
| K | Precio | Precio unitario |
| L | Tipo Cambio | `=IF(E="Pesos",1,VLOOKUP(B,Cotizacion!A:B,2,FALSE))` |
| M | Bruto | `=J*K` |
| N | Interés | Intereses devengados |
| O | Gastos | Comisiones + aranceles |
| P | Neto Calculado | `=IF(J>0,J*K+O,J*K-O)` |
| Q | Origen | "gallo-[hoja]" o "Visual" |
| R | moneda emision | `=VLOOKUP(G,EspeciesVisual!C:Q,5,FALSE)` |
| S | Auditoría | Detalle para verificación |

**Determinación de Moneda (columna E):**
1. Si la hoja dice "Pesos" → "Pesos"
2. Si la hoja dice "Exterior" → "Dolar Cable"
3. Si la hoja dice "Dolares" → "Dolar MEP"
4. Si operación tiene "USD" → "Dolar MEP"

### 5.3 Resultado Ventas ARS

**Columnas (26):**

| Col | Campo | Cálculo |
|-----|-------|---------|
| A | Origen | "gallo-[hoja]" o "Visual" |
| B | Tipo de Instrumento | Del cache EspeciesVisual |
| C | Instrumento | Nombre con moneda |
| D | Cod.Instrum | Código especie |
| E-H | Fechas y operación | Copiados de Boletos |
| I | Cantidad | Cantidad operada |
| J | Precio | Precio original |
| K | Bruto | `=I*J` |
| L | Interés | Intereses |
| M | Tipo de Cambio | 1 (siempre para ARS) |
| N | Gastos | Comisiones |
| O | IVA | `=IF(N>0,N*0.1736,N*-0.1736)` |
| P | Resultado | (vacío) |
| **Q** | **Cantidad Stock Inicial** | Ver explicación abajo |
| **R** | **Precio Stock Inicial** | Ver explicación abajo |
| **S** | **Costo por venta** | `=IF(I<0,I*R,0)` |
| T | Neto Calculado | `=K-N` |
| **U** | **Resultado Calculado** | `=ABS(T)-ABS(S)` |
| V | Cantidad Stock Final | `=I+Q` |
| W | Precio Stock Final | Promedio ponderado |

**Cálculo del Running Stock (columnas Q-W):**

El sistema mantiene un "running stock" por especie para calcular el costo de venta correcto:

```
Si es la primera fila de la especie:
  Q = VLOOKUP(código, Posicion Inicial!D:I, 6)  → Cantidad inicial
  R = VLOOKUP(código, Posicion Inicial!D:P, 13) → Precio inicial

Si NO es la primera fila (misma especie que anterior):
  Q = V de la fila anterior  → Stock final anterior
  R = W de la fila anterior  → Precio promedio anterior
```

**Cálculo del Precio Stock Final (promedio ponderado):**
```excel
=IF(V=0, 0,
  IF(I>0,  // Es compra
    (I*J + Q*R) / (I+Q),  // Promedio ponderado
    R  // Es venta: mantiene precio anterior
  )
)
```

### 5.4 Resultado Ventas USD

**Diferencias con ARS:**

| Col | Campo | Cálculo |
|-----|-------|---------|
| K | Precio Standarizado | `precio * 100` si Visual, `precio` si Gallo |
| L | Precio Standarizado en USD | `=K*O` |
| M | Bruto en USD | `=I*L` |
| O | Tipo de Cambio | `=1` si "dolar" en moneda, sino `=1/P` |
| P | Valor USD Dia | `=VLOOKUP(fecha,Cotizacion!A:B,2,FALSE)` |
| Q | Gastos | Original |
| R | IVA | `=IF(Q>0,Q*0.1736,Q*-0.1736)` basado en Gastos |
| U | Precio Stock USD | `= Precio Posición / Valor USD Día` |

**¿Por qué Precio Standarizado x100?**
Visual reporta los precios de bonos como valor nominal/100 (ej: 0.68 = 68), mientras que Gallo los reporta directamente. El merge estandariza multiplicando x100 los de Visual.

### 5.5 Rentas Dividendos ARS/USD

**Columnas (14):**

| Col | Campo | Cálculo |
|-----|-------|---------|
| A | Instrumento | Nombre del instrumento |
| B | Cod.Instrum | Código especie |
| C | Categoría | "Rentas" o "Dividendos" |
| D | tipo_instrumento | Acciones, Títulos Públicos, ONs |
| E | Concertación | Fecha |
| F | Liquidación | Fecha liquidación |
| G | Nro. NDC | Número de operación |
| H | Tipo Operación | RENTA, DIVIDENDO, AMORTIZACION |
| I | Cantidad | Cantidad |
| J | Moneda | Pesos o tipo de dólar |
| K | Tipo de Cambio | 1 si Pesos, cotización si dólar |
| L | Gastos | Costo + Gastos originales |
| M | Importe | Resultado - Gastos - Costo |
| N | Origen | Hoja de procedencia |

**Categorización:**
```python
if tipo_operacion in ["RENTA", "AMORTIZACION", "AMORTIZACIÓN"]:
    categoria = "Rentas"
else:
    categoria = "Dividendos"
```

**Filtrado ARS vs USD:**
- Se usa `moneda_emision` del cache EspeciesVisual
- Si `moneda_emision == "Pesos"` → ARS
- Cualquier otra → USD

### 5.6 Cauciones (Tomadoras y Colocadoras)

**Separación por tipo de operación:**

| Origen | Condición | Destino PDF |
|--------|-----------|-------------|
| Gallo | Operación contiene "COL" | Cauciones Colocadoras |
| Gallo | Operación contiene "TOM" | Cauciones Tomadoras |
| Visual | Sección "Cauciones tomadoras" | Cauciones Tomadoras |
| Visual | Sección "Cauciones colocadoras" | Cauciones Colocadoras |

**Columnas:**

| Col | Campo | Cálculo |
|-----|-------|---------||
| A | Concertación | Fecha de la operación |
| B | Plazo | Días entre concertación y liquidación |
| C | Liquidación | Fecha de vencimiento |
| D | Operación | TOM CAU TER o COL CAU TER |
| E | Boleto | Número de comprobante |
| F | Contado | Monto colocado/tomado |
| G | Futuro | Monto al vencimiento |
| H | Tipo de Cambio | 1 si pesos, cotización si dólares |
| I | Tasa (%) | Tasa de interés |
| J | Interés Bruto | Intereses generados |
| K | Interés Devengado | Intereses devengados al cierre |
| L | Aranceles | Comisiones |
| M | Derechos | Derechos de mercado |
| N | Costo Financiero | -(Interés + Aranceles + Derechos) |

### 5.7 Resumen

Fórmulas de totales:

| Campo | Fórmula ARS | Fórmula USD |
|-------|-------------|-------------|
| Ventas | `=SUM('Resultado Ventas ARS'!U:U)` | `=SUM('Resultado Ventas USD'!X:X)` |
| Rentas | `=SUMIF('Rentas Dividendos ARS'!C:C,"Rentas",'Rentas Dividendos ARS'!M:M)` | Ídem USD |
| Dividendos | `=SUMIF('Rentas Dividendos ARS'!C:C,"Dividendos",'Rentas Dividendos ARS'!M:M)` | Ídem USD |
| Cau (int) | Suma de intereses de cauciones | Ídem USD |
| Cau (CF) | Suma de costo financiero de cauciones | Ídem USD |
| Total | `=SUM(B:K)` | `=SUM(B:K)` |

---

## 6. Tratamientos Especiales

### 6.1 Errores de OCR en Tickers (0 ↔ O)

El OCR frecuentemente confunde el número 0 con la letra O. El sistema genera variaciones:

```python
TLC10 → [TLC10, TLC1O]
TLOC0 → [TLOC0, TL0C0, TLOCO, TL0CO]
```

Esto se aplica al buscar en:
- PreciosInicialesEspecies
- EspeciesVisual
- EspeciesGallo

### 6.2 Cauciones Separadas

Las cauciones no van en Boletos. Se identifican por:
- Operación contiene "COL CAU" o "TOM CAU"
- Especie = "VARIAS"

Se envían a la hoja "Cauciones Colocadoras" o "Cauciones Tomadoras" con estructura especial.

### 6.3 Operaciones "Transferencia Externa"

Las transferencias externas (depósitos/retiros de títulos) se incluyen en Boletos con:
- Cantidad positiva = ingreso
- Cantidad negativa = egreso
- Sin precio ni gastos asociados

### 6.4 Precios de Bonos (x100)

Los bonos en Visual vienen con precio/100 :
- Visual: 0.68 
- Gallo: 68.00

El merge normaliza multiplicando x100 los de Visual en la columna "Precio Standarizado".

### 6.5 Costo Financiero en Cauciones

```
Costo Financiero = -(Interés + Aranceles + Derechos)
```

El costo financiero es **negativo** porque representa un gasto para el comitente.

---

## 7. Mapeo de Operaciones Gallo → Visual

### 7.1 Operaciones de Compra/Venta

| Gallo | Visual | Destino |
|-------|--------|---------|
| COMPRA | Compra Contado | Boletos |
| VENTA | Venta Contado | Boletos |
| CPRA | Compra | Boletos |
| CANJE | Canje | Boletos |
| LICITACION | Licitaciones MAE | Boletos |
| COMPRA USD | Compra (MEP) | Boletos |
| VENTA USD | Venta (MEP) | Boletos |
| CPRA CABLE | Compra (Cable) | Boletos |
| VENTA CABLE | Venta (Cable) | Boletos |

### 7.2 Operaciones de Rentas

| Gallo | Visual | Categoría |
|-------|--------|-----------|
| RENTA | Renta | Rentas |
| DIVIDENDO | Dividendo en efectivo | Dividendos |
| DIVIDENDOS | Dividendo en efectivo | Dividendos |
| AMORTIZACION | Amortización | Rentas |
| AMORTIZACIÓN | Amortización | Rentas |

### 7.3 Mapeo de Moneda por Hoja Origen

| Hoja Gallo | Moneda Resultante |
|------------|-------------------|
| Tit Privados Exentos | Pesos |
| Titulos Publicos | Pesos |
| Renta Fija Pesos | Pesos |
| Renta Fija Dolares | Dolar MEP |
| Tit Privados Exterior | Dolar Cable |
| Cauciones | Según tipo |

---

## 8. Fórmulas Excel Utilizadas

### 8.1 VLOOKUPs Principales

```excel
// Tipo de Instrumento
=IFERROR(VLOOKUP(G2,EspeciesVisual!C:R,16,FALSE),"")

// Instrumento con Moneda
=IFERROR(VLOOKUP(G2,EspeciesVisual!C:Q,15,FALSE),"")

// Moneda Emisión
=IFERROR(VLOOKUP(G2,EspeciesVisual!C:Q,5,FALSE),"")

// Tipo de Cambio
=IF(E2="Pesos",1,IFERROR(VLOOKUP(B2,'Cotizacion Dolar Historica'!A:B,2,FALSE),0))

// Precio Inicial de Posición
=IFERROR(VLOOKUP(D2,'Posicion Inicial Gallo'!D:P,13,FALSE),0)
```

### 8.2 Cálculos de Running Stock

```excel
// Cantidad Stock Inicial (primera fila)
=IF(LEFT(A2,5)="Gallo",
    IFERROR(VLOOKUP(D2,'Posicion Inicial Gallo'!D:I,6,FALSE),0),
    IFERROR(VLOOKUP(D2,'Posicion Final Gallo'!D:I,6,FALSE),0))

// Cantidad Stock Inicial (filas siguientes)
=IF(D3=D2,  // Si misma especie
    V2,     // Usar stock final anterior
    VLOOKUP(D3,'Posicion Inicial Gallo'!D:I,6,FALSE))

// Costo por Venta
=IF(I2<0,I2*R2,0)  // Si es venta, Cantidad × Precio Stock

// Resultado Calculado
=IF(V2<>0,ABS(T2)-ABS(S2),0)

// Precio Stock Final (promedio ponderado)
=IF(V2=0,0,
  IF(I2>0,
    (I2*J2+Q2*R2)/(I2+Q2),
    R2))
```

### 8.3 IVA sobre Gastos

```excel
// IVA = 17.36% de los gastos
=IF(N2>0,N2*0.1736,N2*-0.1736)
```

---

## 9. Edge Cases Conocidos

### 9.1 Reportar si encuentran:

| Situación | Qué reportar |
|-----------|--------------|
| Operaciones faltantes | Si una operación del PDF original no aparece en el Excel/PDF final |
| Números incorrectos | Cualquier valor que no coincida entre PDF original y resultado |
| Lógica errónea en cálculos | Si un resultado calculado no tiene sentido (ej: costo mayor que venta) |
| Moneda invertida | Si ven valores que deberían estar en dólares pero aparecen en pesos o viceversa |
| Precios claramente mal | Si el precio de un instrumento es absurdo (ej: acción a $0.01 o bono a $10000) |
| Código especie no encontrado | El código y el nombre del instrumento |
| Precio inicial = 0 | Ticker y código de especie |
| Cotización faltante | Fecha y tipo de dólar |
| Categoría incorrecta | Operación que debería ser Renta pero sale Dividendo o viceversa |
| Cantidad stock negativo | Si el stock final queda negativo, hay operaciones faltantes |
| Cauciones mal clasificadas | Si una caución tomadora aparece como colocadora o viceversa |

### 9.2 Casos especiales ya manejados:

- ✅ AMORTIZACIÓN con y sin tilde
- ✅ Errores OCR 0/O en tickers
- ✅ Precios x100 en bonos de Visual
- ✅ Cauciones separadas de Boletos
- ✅ Operaciones USD vs Cable
- ✅ Tipo de cambio = 1 para operaciones en dólares
- ✅ Running stock con promedio ponderado

### 9.3 Validaciones recomendadas:

1. **Verificar que el Total del Resumen coincida** con la suma manual de las operaciones
2. **Comparar cantidad inicial + operaciones = cantidad final** por especie
3. **Revisar operaciones con Resultado muy distinto a lo esperado**
4. **Buscar operaciones con campos vacíos** que deberían tener valor

---

## 📞 Soporte

Para reportar edge cases o problemas:
1. Captura de pantalla del PDF original
2. Fila exacta del Excel donde se ve el error
3. Valor esperado vs valor obtenido
4. Código de especie involucrado

---

*Documento generado automáticamente - Febrero 2026*
