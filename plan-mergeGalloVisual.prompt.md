# Plan: Automatizar Merge Gallo-Visual → Resumen Impositivo

**TL;DR:** Crear un módulo Python que tome los dos Excel generados (Gallo y Visual) y produzca un Excel consolidado con el resumen impositivo anual. El proceso traduce la estructura de Gallo al esquema de Visual, unifica movimientos en hojas de Boletos y Rentas/Dividendos, calcula resultados por operación, y genera la hoja Resumen final.

## Archivos de Input
- `12128_LANDRO_VERONICA_INES_Gallo_Generado_OK.xlsx` - Excel generado del PDF Gallo
- `12128_LANDRO_VERONICA_INES_Visual_Generado_OK.xlsx` - Excel generado del PDF Visual
- Hojas auxiliares (a exportar como archivos separados):
  - EspeciesVisual.xlsx (16258 filas)
  - EspeciesGallo.xlsx (2272 filas)
  - CotizacionDolarHistorica.xlsx (3490 filas)
  - PreciosInicialesEspecies.xlsx (726 filas)

## Estructura Final del Excel (14 hojas)

### Hojas Principales
| Hoja | Filas | Columnas |
|------|-------|----------|
| Posicion Inicial Gallo | 22 | 11 |
| Posicion Final Gallo | 22 | 20 |
| Boletos | 34 | 19 |
| Rentas y Dividendos Gallo | 27 | 20 |
| Resultado Ventas ARS | 147 | 24 |
| Resultado Ventas USD | 48 | 27 |
| Rentas Dividendos ARS | 38 | 24 |
| Rentas Dividendos USD | 30 | 14 |
| Resumen | 3 | 12 |
| Posicion Titulos | 20 | 6 |

### Hojas Auxiliares (separar como archivos)
- EspeciesVisual
- EspeciesGallo
- Cotizacion Dolar Historica
- PreciosInicialesEspecies

---

## ESQUEMA DE COLUMNAS POR HOJA

### Posicion Inicial Gallo (11 cols)
1. tipo_especie
2. Ticker (NUEVO: primera palabra de especie original)
3. especie (resto después del ticker)
4. detalle
5. custodia
6. cantidad
7. precio (calcular si falta: importe_pesos/cantidad)
8. importe_pesos
9. porc_cartera_pesos
10. importe_dolares
11. porc_cartera_dolares

### Posicion Final Gallo (20 cols)
1. tipo_especie
2. Ticker
3. especie
4. Codigo especie (match con otras hojas de Gallo, fallback EspeciesGallo)
5. Codigo Especie Origen ("Gallo" o "EspeciesGallo")
6. comentario especies (casos especiales, no va en resultado final)
7. detalle
8. custodia
9. cantidad
10. precio Tenencia Final Pesos (=importe_pesos/cantidad)
11. precio Tenencia Final USD (=importe_dolares/cantidad)
12. Precio Tenencia Inicial (=VLOOKUP ticker en PreciosInicialesEspecies, fijos: pesos=1, dolares=1167.806, dolar cable=1148.93)
13. precio costo (Σ(cantidad×precio compras) / cantidad_posicion_final, si hay canjes usar precio inicial)
14. Origen precio costo
15. comentarios precio costo
16. Precio a Utilizar (=precio_tenencia_inicial o precio_costo)
17. importe_pesos
18. porc_cartera_pesos
19. importe_dolares
20. porc_cartera_dolares

### Boletos (19 cols)
1. Tipo de Instrumento (=VLOOKUP cod en EspeciesVisual col TipoEspecie)
2. Concertación (fecha de Gallo, solo 2025)
3. Liquidación (vacío para Gallo)
4. Nro. Boleto (numero de Gallo)
5. Moneda (Pesos si resultado_pesos≠0, Dolar MEP/Cable según hoja origen)
6. Tipo Operación (solo compra/venta/canje, NO rentas/dividendos)
7. Cod.Instrum (cod_especie de Gallo)
8. Instrumento Crudo (especie de Gallo)
9. InstrumentoConMoneda (=VLOOKUP(G2,EspeciesVisual!C:Q,15,FALSE))
10. Cantidad
11. Precio
12. Tipo Cambio (=IF(E2="Pesos",1,INDEX/MATCH Cotizacion Dolar))
13. Bruto (=J2*K2)
14. Interés (siempre 0)
15. Gastos (gastos_pesos o gastos_usd según moneda)
16. Neto Calculado (=IF(J2>0,J2*K2+O2,J2*K2-O2))
17. Origen ("gallo-NombreHoja" o "visual")
18. moneda emision (=VLOOKUP(G2,EspeciesVisual!C:Q,5,FALSE))
19. (columna auditoría)

### Rentas y Dividendos Gallo (20 cols)
Misma estructura que Boletos pero:
- Filtrar por operación = renta/dividendos/amortización
- Col 16: Costo (de Gallo)
- Col 17: Neto Calculado (=M2-O2, o para amortización: =-M2-P2)
- Amortización: precio 100→1

### Resultado Ventas ARS (24 cols)
Filtro: InstrumentoConMoneda termina en "Pesos"

1. Origen (=Boletos!Q2)
2. Tipo de Instrumento (=IF(Boletos!R2="Pesos",Boletos!A2,""))
3. Instrumento (=IF(Boletos!R2="Pesos",Boletos!I2,""))
4. Cod.Instrum
5. Concertación
6. Liquidación
7. Moneda
8. Tipo Operación
9. Cantidad
10. Precio
11. Bruto
12. Interés
13. Tipo de Cambio
14. Gastos
15. IVA (=IF(N2>0,N2*0.1736,N2*-0.1736))
16. Resultado (vacío)
17. Cantidad Stock Inicial (=VLOOKUP(D2,'Posicion Final Gallo'!D:I,6,FALSE))
18. Precio Stock Inicial (=VLOOKUP(D2,'Posicion Final Gallo'!D:P,12,FALSE))
19. Costo por venta(gallo) (=IF(I<0,I*R,0) solo ventas)
20. Neto Calculado(visual) (=IF(S<>0,K+N,0) solo ventas)
21. Resultado Calculado(final) (=IF(S<>0,ABS(T)-ABS(S),0))
22. Cantidad de Stock Final (=I+Q, running database)
23. Precio Stock Final (=IF(V=0,0,IF(I>0,(I*J+Q*R)/(I+Q),R)) promedio ponderado)
24. chequeado (comentarios auditoría)

**LÓGICA RUNNING STOCK:**
- Fila 2: Stock Inicial = VLOOKUP desde Posicion Final Gallo
- Fila 3+: Si misma especie que fila anterior, Stock Inicial = Stock Final de fila anterior
- Fórmula fila 3: Cantidad Stock Inicial = =Q2+I2, Precio Stock Inicial = =(VLOOKUP*Q2+I2*J2)/Q3

### Resultado Ventas USD (27 cols)
Filtro: InstrumentoConMoneda incluye "Dolar"

1-10: Igual que ARS (con SEARCH("Dolar",...))
11. Precio Standarizado (=IF(A="visual",IF(U*P/J>80,J*100),J)) - ajuste x100 si nominal
12. Precio Standarizado en USD (=IF(G="Pesos",K/P,K))
13. Bruto en USD (=I*L)
14. Interés (0)
15. Tipo de Cambio (=IF(G="Pesos",Boletos!L/P,1)) - referencia USD=1
16. Valor USD Dia (manual o lookup)
17. Gastos
18. IVA
19. Resultado
20. Cantidad Stock Inicial
21. Precio Stock USD (=VLOOKUP(D,PreciosInicialesEspecies!A:G,7,FALSE)/P)
22. Costo por venta (=IF(I<0,I*U,0))
23. Neto Calculado (=IF(V<>0,M-Q,0))
24. Resultado Calculado (=IF(V<>0,ABS(W)-ABS(V),0))
25. Cantidad de Stock Final
26. Precio Stock Final
27. Comentarios (auditoría)

### Rentas Dividendos ARS (14 cols)
Filtro: moneda emision = "Pesos"

1. Instrumento
2. Cod.Instrum
3. Categoría (Rentas/Dividendos/AMORTIZACION)
4. tipo_instrumento
5. Concertación
6. Liquidación
7. Nro. NDC
8. Tipo Operación
9. Cantidad
10. Moneda
11. Tipo de Cambio
12. Gastos
13. Importe (=Neto Calculado de origen)
14. Origen

### Rentas Dividendos USD (14 cols)
Filtro: moneda emision incluye "Dolar"
Misma estructura que ARS

### Resumen (12 cols)
1. Moneda (ARS/USD)
2. Ventas (=SUM('Resultado Ventas ARS'!U:U))
3. FCI (0)
4. Opciones (0)
5. Rentas (=SUMIF('Rentas Dividendos ARS'!C:C,"Rentas",M:M)+SUMIF(...,"AMORTIZACION",...))
6. Dividendos (=SUMIF('Rentas Dividendos ARS'!C:C,"Dividendos",M:M))
7-11. (0)
12. Total

---

## CASOS ESPECIALES (de comentarios del Excel)

### Match de Código Especie
- Match NO es exacto: "KO CEDEAR COCA-COLA" en posición final puede ser "CEDEAR COCA-COLA" en transacciones
- Si nombre incluye fecha, es vencimiento y DEBE estar en match
- Si hay duplicados (ej: Ternium 8432 y 41946), probar cuál aparece en Visual
- CEDEAR Mercado Libre: 8445 vs 47887 → ir por match más cercano
- AMAZON: si no está cedear en EspeciesGallo, buscar por similitud en Visual
- TESLA: tipo_especie "Tit.Privados exterior" indica que NO es cedear

### Precio Estandarizado (USD)
- De Gallo viene nominal, de Visual varía
- Si precio_stock * precio_dia / cantidad > 80, entonces precio * 100
- Implementar lógica para detectar automáticamente

### Running Stock
- Compra y venta mismo día = operación instantánea, no calcular sobre stock anterior
- Compras sin venta: solo suman al stock, no generan resultado
- Para especies que vienen de Gallo: usar Posicion Inicial para cantidad/precio inicial

### Duplicados
- Si especie está repetida en Posicion Final (ej: ON Telecom), consolidar info

---

## Steps de Implementación

### Step 1: Exportar hojas auxiliares
Crear archivos en `pdf_converter/datalab/aux_data/`:
- EspeciesVisual.xlsx
- EspeciesGallo.xlsx
- CotizacionDolarHistorica.xlsx
- PreciosInicialesEspecies.xlsx

### Step 2: Crear clase GalloVisualMerger
```python
class GalloVisualMerger:
    def __init__(self, gallo_path, visual_path, aux_data_dir):
        self.gallo_wb = load_workbook(gallo_path)
        self.visual_wb = load_workbook(visual_path)
        self.aux_data = self._load_aux_data(aux_data_dir)
    
    def merge(self) -> Workbook:
        wb = Workbook()
        self._create_posicion_inicial(wb)
        self._create_posicion_final(wb)
        self._create_boletos(wb)
        self._create_rentas_dividendos_gallo(wb)
        self._create_resultado_ventas_ars(wb)
        self._create_resultado_ventas_usd(wb)
        self._create_rentas_dividendos_ars(wb)
        self._create_rentas_dividendos_usd(wb)
        self._create_resumen(wb)
        self._add_aux_sheets(wb)
        return wb
```

### Step 3-11: Implementar cada método según esquemas arriba

### Step 12: Agregar columna auditoría
Última columna de cada hoja con explicación:
```
"Bruto=J2*K2 | Gastos=col M Gallo | Neto=IF(J2>0,J2*K2+O2,J2*K2-O2)"
```

### Step 13: Fórmulas embebidas
Escribir fórmulas Excel literales, no valores calculados

### Step 14: Testing
Comparar output con `Visual_Estructurado_MergeGallo_Resultados_enproceso.xlsx`

### Step 15: Integrar en Streamlit
