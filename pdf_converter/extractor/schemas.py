"""
Schema definitions for Gallo and Visual report types.
Contains column lists and mappings for each section.
"""

# =============================================================================
# GALLO SCHEMAS
# =============================================================================

GALLO_RESULTADO_TOTALES_SCHEMA = [
    "categoria",
    "valor_pesos",
    "valor_usd"
]

GALLO_TRANSACCIONES_SCHEMA = [
    "tipo_fila",
    "cod_especie",
    "especie",
    "fecha",
    "operacion",
    "numero",
    "cantidad",
    "precio",
    "importe",
    "costo",
    "resultado_pesos",
    "resultado_usd",
    "gastos_pesos",
    "gastos_usd"
]

GALLO_CAUCIONES_SCHEMA = [
    "tipo_fila",
    "cod_especie",
    "especie",
    "fecha",
    "vencimiento",
    "operacion",
    "numero",
    "colocado",
    "al_vencimiento",
    "interes_pesos",
    "interes_usd",
    "gastos_pesos",
    "gastos_usd"
]

GALLO_POSICION_SCHEMA = [
    "tipo_especie",
    "especie",
    "detalle",
    "custodia",
    "cantidad",
    "precio",
    "importe_pesos",
    "pct_cartera_pesos",
    "importe_dolares",
    "pct_cartera_dolares"
]

# =============================================================================
# VISUAL SCHEMAS
# =============================================================================

VISUAL_BOLETOS_SCHEMA = [
    "tipo_instrumento",
    "concertacion",
    "liquidacion",
    "nro_boleto",
    "moneda",
    "tipo_operacion",
    "cod_instrumento",
    "instrumento",
    "cantidad",
    "precio",
    "tipo_cambio",
    "bruto",
    "interes",
    "gastos",
    "neto"
]

VISUAL_RESULTADO_VENTAS_SCHEMA = [
    "tipo_instrumento",
    "instrumento",
    "cod_instrumento",
    "concertacion",
    "liquidacion",
    "moneda",
    "tipo_operacion",
    "cantidad",
    "precio",
    "bruto",
    "interes",
    "tipo_cambio",
    "gastos",
    "iva",
    "resultado"
]

VISUAL_RENTAS_DIVIDENDOS_SCHEMA = [
    "instrumento",
    "cod_instrumento",
    "categoria",
    "tipo_instrumento",
    "concertacion",
    "liquidacion",
    "nro_operacion",
    "tipo_operacion",
    "cantidad",
    "moneda",
    "tipo_cambio",
    "gastos",
    "importe"
]

VISUAL_RESUMEN_SCHEMA = [
    "moneda",
    "ventas",
    "fci",
    "opciones",
    "rentas",
    "dividendos",
    "ef_cpd",
    "pagares",
    "futuros",
    "cau_int",
    "cau_cf",
    "total"
]

VISUAL_POSICION_TITULOS_SCHEMA = [
    "instrumento",
    "codigo",
    "ticker",
    "cantidad",
    "importe",
    "moneda"
]

# =============================================================================
# SCHEMA REGISTRY
# =============================================================================

GALLO_SCHEMAS = {
    "resultado_totales": GALLO_RESULTADO_TOTALES_SCHEMA,
    "tit_privados_exentos": GALLO_TRANSACCIONES_SCHEMA,
    "tit_privados_exterior": GALLO_TRANSACCIONES_SCHEMA,
    "renta_fija_pesos": GALLO_TRANSACCIONES_SCHEMA,
    "renta_fija_dolares": GALLO_TRANSACCIONES_SCHEMA,
    "fci": GALLO_TRANSACCIONES_SCHEMA,
    "opciones": GALLO_TRANSACCIONES_SCHEMA,
    "futuros": GALLO_TRANSACCIONES_SCHEMA,
    "cauciones_pesos": GALLO_CAUCIONES_SCHEMA,
    "cauciones_dolares": GALLO_CAUCIONES_SCHEMA,
    "posicion_inicial": GALLO_POSICION_SCHEMA,
    "posicion_final": GALLO_POSICION_SCHEMA,
}

VISUAL_SCHEMAS = {
    "boletos": VISUAL_BOLETOS_SCHEMA,
    "resultado_ventas_ars": VISUAL_RESULTADO_VENTAS_SCHEMA,
    "resultado_ventas_usd": VISUAL_RESULTADO_VENTAS_SCHEMA,
    "rentas_dividendos_ars": VISUAL_RENTAS_DIVIDENDOS_SCHEMA,
    "rentas_dividendos_usd": VISUAL_RENTAS_DIVIDENDOS_SCHEMA,
    "resumen": VISUAL_RESUMEN_SCHEMA,
    "posicion_titulos": VISUAL_POSICION_TITULOS_SCHEMA,
}

# =============================================================================
# SHEET NAME MAPPINGS
# =============================================================================

GALLO_SECTION_TO_SHEET = {
    "resultado_totales": "Resultado Totales",
    "tit_privados_exentos": "Tit.Privados Exentos",
    "tit_privados_exterior": "Tit.Privados Exterior",
    "renta_fija_pesos": "Renta Fija Pesos",
    "renta_fija_dolares": "Renta Fija Dolares",
    "fci": "FCI",
    "opciones": "Opciones",
    "futuros": "Futuros",
    "cauciones_pesos": "Cauciones Pesos",
    "cauciones_dolares": "Cauciones Dolares",
    "posicion_inicial": "Posicion Inicial",
    "posicion_final": "Posicion Final",
}

VISUAL_SECTION_TO_SHEET = {
    "boletos": "Boletos",
    "resultado_ventas_ars": "Resultado Ventas ARS",
    "resultado_ventas_usd": "Resultado Ventas USD",
    "rentas_dividendos_ars": "Rentas Dividendos ARS",
    "rentas_dividendos_usd": "Rentas Dividendos USD",
    "resumen": "Resumen",
    "posicion_titulos": "Posicion Titulos",
}

# =============================================================================
# CATEGORY TO SECTION MAPPING (for dynamic detection in Gallo)
# =============================================================================

CATEGORIA_TO_SECTION = {
    # TIT.PRIVADOS EXENTOS variations
    "tit.privados exentos": "tit_privados_exentos",
    "tit privados exentos": "tit_privados_exentos",
    "titulos privados exentos": "tit_privados_exentos",
    "títulos privados exentos": "tit_privados_exentos",
    
    # TIT.PRIVADOS DEL EXTERIOR variations
    "tit.privados del exterior": "tit_privados_exterior",
    "tit privados del exterior": "tit_privados_exterior",
    "titulos privados del exterior": "tit_privados_exterior",
    
    # RENTA FIJA variations
    "renta fija en pesos": "renta_fija_pesos",
    "renta fija pesos": "renta_fija_pesos",
    "renta fija en dolares": "renta_fija_dolares",
    "renta fija dolares": "renta_fija_dolares",
    "renta fija en dólares": "renta_fija_dolares",
    
    # CAUCIONES variations
    "cauciones en pesos": "cauciones_pesos",
    "cauciones pesos": "cauciones_pesos",
    "cauciones en dolares": "cauciones_dolares",
    "cauciones dolares": "cauciones_dolares",
    "cauciones en dólares": "cauciones_dolares",
    
    # FCI variations
    "fci": "fci",
    "fondos comunes de inversion": "fci",
    "fondos comunes de inversión": "fci",
    
    # OPCIONES
    "opciones": "opciones",
    
    # FUTUROS
    "futuros": "futuros",
}

# =============================================================================
# NUMERIC FIELDS (for post-processing)
# =============================================================================

GALLO_NUMERIC_FIELDS = {
    "resultado_totales": ["valor_pesos", "valor_usd"],
    "tit_privados_exentos": ["cantidad", "precio", "importe", "costo", "resultado_pesos", "resultado_usd", "gastos_pesos", "gastos_usd"],
    "tit_privados_exterior": ["cantidad", "precio", "importe", "costo", "resultado_pesos", "resultado_usd", "gastos_pesos", "gastos_usd"],
    "renta_fija_pesos": ["cantidad", "precio", "importe", "costo", "resultado_pesos", "resultado_usd", "gastos_pesos", "gastos_usd"],
    "renta_fija_dolares": ["cantidad", "precio", "importe", "costo", "resultado_pesos", "resultado_usd", "gastos_pesos", "gastos_usd"],
    "fci": ["cantidad", "precio", "importe", "costo", "resultado_pesos", "resultado_usd", "gastos_pesos", "gastos_usd"],
    "opciones": ["cantidad", "precio", "importe", "costo", "resultado_pesos", "resultado_usd", "gastos_pesos", "gastos_usd"],
    "futuros": ["cantidad", "precio", "importe", "costo", "resultado_pesos", "resultado_usd", "gastos_pesos", "gastos_usd"],
    "cauciones_pesos": ["colocado", "al_vencimiento", "interes_pesos", "interes_usd", "gastos_pesos", "gastos_usd"],
    "cauciones_dolares": ["colocado", "al_vencimiento", "interes_pesos", "interes_usd", "gastos_pesos", "gastos_usd"],
    "posicion_inicial": ["cantidad", "precio", "importe_pesos", "pct_cartera_pesos", "importe_dolares", "pct_cartera_dolares"],
    "posicion_final": ["cantidad", "precio", "importe_pesos", "pct_cartera_pesos", "importe_dolares", "pct_cartera_dolares"],
}

VISUAL_NUMERIC_FIELDS = {
    "boletos": ["cantidad", "precio", "tipo_cambio", "bruto", "interes", "gastos", "neto"],
    "resultado_ventas_ars": ["cantidad", "precio", "bruto", "interes", "tipo_cambio", "gastos", "iva", "resultado"],
    "resultado_ventas_usd": ["cantidad", "precio", "bruto", "interes", "tipo_cambio", "gastos", "iva", "resultado"],
    "rentas_dividendos_ars": ["cantidad", "tipo_cambio", "gastos", "importe"],
    "rentas_dividendos_usd": ["cantidad", "tipo_cambio", "gastos", "importe"],
    "resumen": ["ventas", "fci", "opciones", "rentas", "dividendos", "ef_cpd", "pagares", "futuros", "cau_int", "cau_cf", "total"],
    "posicion_titulos": ["cantidad", "importe"],
}

# =============================================================================
# DEDUPLICATION KEYS
# =============================================================================

GALLO_DEDUP_KEYS = {
    "tit_privados_exentos": ["cod_especie", "fecha", "operacion", "numero", "cantidad"],
    "tit_privados_exterior": ["cod_especie", "fecha", "operacion", "numero", "cantidad"],
    "renta_fija_pesos": ["cod_especie", "fecha", "operacion", "numero", "cantidad"],
    "renta_fija_dolares": ["cod_especie", "fecha", "operacion", "numero", "cantidad"],
    "fci": ["cod_especie", "fecha", "operacion", "numero", "cantidad"],
    "opciones": ["cod_especie", "fecha", "operacion", "numero", "cantidad"],
    "futuros": ["cod_especie", "fecha", "operacion", "numero", "cantidad"],
    "cauciones_pesos": ["cod_especie", "fecha", "vencimiento", "numero"],
    "cauciones_dolares": ["cod_especie", "fecha", "vencimiento", "numero"],
    "posicion_inicial": ["especie", "detalle", "cantidad"],
    "posicion_final": ["especie", "detalle", "cantidad"],
}

VISUAL_DEDUP_KEYS = {
    "boletos": ["nro_boleto", "cod_instrumento", "cantidad"],
    "resultado_ventas_ars": ["instrumento", "cod_instrumento", "concertacion", "tipo_operacion", "cantidad"],
    "resultado_ventas_usd": ["instrumento", "cod_instrumento", "concertacion", "tipo_operacion", "cantidad"],
    "rentas_dividendos_ars": ["instrumento", "cod_instrumento", "concertacion", "nro_operacion"],
    "rentas_dividendos_usd": ["instrumento", "cod_instrumento", "concertacion", "nro_operacion"],
    "posicion_titulos": ["instrumento", "codigo", "cantidad"],
}


def get_schema(report_type: str, section_key: str) -> list:
    """Get the column schema for a section."""
    if report_type == "gallo":
        return GALLO_SCHEMAS.get(section_key, [])
    else:
        return VISUAL_SCHEMAS.get(section_key, [])


def get_sheet_name(report_type: str, section_key: str) -> str:
    """Get the Excel sheet name for a section."""
    if report_type == "gallo":
        return GALLO_SECTION_TO_SHEET.get(section_key, section_key)
    else:
        return VISUAL_SECTION_TO_SHEET.get(section_key, section_key)


def get_numeric_fields(report_type: str, section_key: str) -> list:
    """Get the numeric fields for a section."""
    if report_type == "gallo":
        return GALLO_NUMERIC_FIELDS.get(section_key, [])
    else:
        return VISUAL_NUMERIC_FIELDS.get(section_key, [])


def get_dedup_keys(report_type: str, section_key: str) -> list:
    """Get the deduplication keys for a section."""
    if report_type == "gallo":
        return GALLO_DEDUP_KEYS.get(section_key, [])
    else:
        return VISUAL_DEDUP_KEYS.get(section_key, [])
