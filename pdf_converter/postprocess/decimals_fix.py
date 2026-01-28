"""
Decimal correction for x100 errors.
Fixes cases where LLM extracts values multiplied by 100.
"""

from typing import List, Dict, Any, Optional
from rich.console import Console

console = Console()


def detect_x100_error(extracted_value: float, expected_value: float, 
                      ratio_min: float = 95, ratio_max: float = 105) -> bool:
    """
    Detect if extracted value is ~100x the expected value.
    
    Args:
        extracted_value: Value extracted by LLM
        expected_value: Value calculated from detail data
        ratio_min: Minimum ratio to consider as x100 error
        ratio_max: Maximum ratio to consider as x100 error
    
    Returns:
        True if x100 error detected
    """
    if expected_value == 0:
        return False
    
    ratio = abs(extracted_value / expected_value)
    return ratio_min < ratio < ratio_max


def fix_resumen_decimals(
    resumen_rows: List[Dict],
    detail_data: Dict[str, List[Dict]],
    tolerance_ratio: float = 100.0
) -> List[Dict]:
    """
    Fix x100 decimal errors in Visual Resumen by comparing with detail data.
    
    The LLM sometimes extracts values multiplied by 100 (e.g., -4275009.00
    instead of -42750.09). This function detects and corrects such errors.
    
    Args:
        resumen_rows: List of resumen rows (typically 2: ARS and USD)
        detail_data: Dictionary with detail section data
        tolerance_ratio: Ratio range to detect x100 error (95-105)
    
    Returns:
        Corrected resumen rows
    """
    if not resumen_rows:
        return resumen_rows
    
    corrected = []
    
    for row in resumen_rows:
        new_row = row.copy()
        moneda = str(row.get("moneda", "")).upper()
        
        # Map currency to suffix
        if "ARS" in moneda or "PESOS" in moneda:
            suffix = "ars"
        elif "USD" in moneda or "DOLAR" in moneda:
            suffix = "usd"
        else:
            corrected.append(new_row)
            continue
        
        # Check each field that can be validated
        validations = {
            "ventas": _sum_resultado_ventas(detail_data, suffix),
            "rentas": _sum_rentas(detail_data, suffix),
            "dividendos": _sum_dividendos(detail_data, suffix),
        }
        
        for field, expected in validations.items():
            if expected == 0:
                continue
            
            extracted = float(row.get(field, 0))
            if extracted == 0:
                continue
            
            if detect_x100_error(extracted, expected):
                corrected_value = extracted / 100
                console.print(f"[yellow]  [FIX] {moneda} {field}: {extracted} → {corrected_value}[/yellow]")
                new_row[field] = corrected_value
        
        # Recalculate total
        total = sum(float(new_row.get(f, 0)) for f in [
            "ventas", "fci", "opciones", "rentas", "dividendos",
            "ef_cpd", "pagares", "futuros", "cau_int", "cau_cf"
        ])
        new_row["total"] = total
        
        corrected.append(new_row)
    
    return corrected


def _sum_resultado_ventas(detail_data: Dict, currency: str) -> float:
    """Sum resultado from resultado_ventas section."""
    key = f"resultado_ventas_{currency}"
    rows = detail_data.get(key, [])
    return sum(float(row.get("resultado", 0)) for row in rows)


def _sum_rentas(detail_data: Dict, currency: str) -> float:
    """Sum importe from rentas in rentas_dividendos section."""
    key = f"rentas_dividendos_{currency}"
    rows = detail_data.get(key, [])
    return sum(
        float(row.get("importe", 0)) 
        for row in rows 
        if str(row.get("categoria", "")).upper() in ["RENTAS", "RENTA"]
    )


def _sum_dividendos(detail_data: Dict, currency: str) -> float:
    """Sum importe from dividendos in rentas_dividendos section."""
    key = f"rentas_dividendos_{currency}"
    rows = detail_data.get(key, [])
    return sum(
        float(row.get("importe", 0)) 
        for row in rows 
        if str(row.get("categoria", "")).upper() in ["DIVIDENDOS", "DIVIDENDO"]
    )


def fix_gallo_totales(
    totales_rows: List[Dict],
    section_data: Dict[str, List[Dict]]
) -> List[Dict]:
    """
    Verify and potentially fix Gallo resultado_totales against section sums.
    
    This doesn't auto-correct but logs discrepancies for review.
    
    Args:
        totales_rows: Resultado totales rows
        section_data: All section data
    
    Returns:
        Totales rows (unchanged, but discrepancies logged)
    """
    import re
    
    for row in totales_rows:
        categoria = str(row.get("categoria", "")).upper()
        valor_pesos = float(row.get("valor_pesos", 0))
        valor_usd = float(row.get("valor_usd", 0))
        
        # Skip TOTAL GENERAL
        if "TOTAL GENERAL" in categoria:
            continue
        
        # Extract category base and type
        match = re.match(r'^(.+?)\s*\((.+?)\)\s*$', categoria)
        if not match:
            continue
        
        cat_base = match.group(1).strip()
        tipo = match.group(2).strip().lower()  # "enajenacion" or "renta"
        
        # Find corresponding section
        section_key = _map_categoria_to_section(cat_base)
        if not section_key or section_key not in section_data:
            continue
        
        section_rows = section_data[section_key]
        
        # Calculate expected value
        if "caucion" in section_key:
            # Cauciones: sum interes
            calc_pesos = sum(float(r.get("interes_pesos", 0)) for r in section_rows 
                           if "total" not in str(r.get("tipo_fila", "")).lower())
            calc_usd = sum(float(r.get("interes_usd", 0)) for r in section_rows 
                         if "total" not in str(r.get("tipo_fila", "")).lower())
        else:
            # Other sections: sum from Total rows matching tipo
            calc_pesos = sum(float(r.get("resultado_pesos", 0)) for r in section_rows 
                           if tipo in str(r.get("tipo_fila", "")).lower())
            calc_usd = sum(float(r.get("resultado_usd", 0)) for r in section_rows 
                         if tipo in str(r.get("tipo_fila", "")).lower())
        
        # Log discrepancies
        if abs(calc_pesos - valor_pesos) > 0.01:
            console.print(f"[yellow]  ⚠️ Discrepancy in {categoria} pesos: expected {calc_pesos}, got {valor_pesos}[/yellow]")
        if abs(calc_usd - valor_usd) > 0.01:
            console.print(f"[yellow]  ⚠️ Discrepancy in {categoria} usd: expected {calc_usd}, got {valor_usd}[/yellow]")
    
    return totales_rows


def _map_categoria_to_section(cat_base: str) -> Optional[str]:
    """Map categoria name to section key."""
    cat_lower = cat_base.lower()
    
    mappings = {
        "tit.privados exentos": "tit_privados_exentos",
        "tit privados exentos": "tit_privados_exentos",
        "tit.privados del exterior": "tit_privados_exterior",
        "renta fija en pesos": "renta_fija_pesos",
        "renta fija en dolares": "renta_fija_dolares",
        "renta fija en dólares": "renta_fija_dolares",
        "cauciones en pesos": "cauciones_pesos",
        "cauciones en dolares": "cauciones_dolares",
        "cauciones en dólares": "cauciones_dolares",
        "fci": "fci",
        "opciones": "opciones",
        "futuros": "futuros",
    }
    
    for pattern, section in mappings.items():
        if pattern in cat_lower:
            return section
    
    return None
