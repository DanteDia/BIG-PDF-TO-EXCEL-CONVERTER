"""
Módulo de Verificación Estandarizada para Archivos Procesados
=============================================================

Este módulo implementa las relaciones matemáticas verificadas entre las hojas
de los archivos de Gallo y Visual, permitiendo validar automáticamente
que los datos extraídos son correctos y consistentes.

RELACIONES MATEMÁTICAS:
=======================

VISUAL:
-------
| Celda Resumen | Valor         | Fórmula                                             |
|---------------|---------------|-----------------------------------------------------|
| B2 (ARS ventas)     | 0.00          | =SUM('Resultado Ventas ARS'!O:O)                    |
| B3 (USD ventas)     | -14,670,053.15| =SUM('Resultado Ventas USD'!O:O)                    |
| E2 (ARS rentas)     | 0.00          | =SUMIF(Rentas Div ARS, categoria="RENTAS", importe) |
| E3 (USD rentas)     | 2,073.32      | =SUMIF(Rentas Div USD, categoria="Rentas", importe) |
| F2 (ARS dividendos) | -42,750.09    | =SUMIF(Rentas Div ARS, categoria="DIVIDENDOS", importe)|
| F3 (USD dividendos) | 0.14          | =SUMIF(Rentas Div USD, categoria="Dividendos", importe)|
| L2/L3 (total)       |               | =SUM(B:K) por fila                                  |

GALLO:
------
| Celda Resultado Totales | Fórmula                                                     |
|-------------------------|-------------------------------------------------------------|
| TIT.PRIVADOS EXENTOS (Enajenacion) pesos | =K18+K26+K36+K43+K53+K57+K60 de 'Tit.Privados Exentos'|
| TIT.PRIVADOS EXENTOS (Renta) pesos       | =K9+K19+K27+K37+K44+K54 de 'Tit.Privados Exentos'     |
| TIT.PRIVADOS DEL EXTERIOR (Renta) usd    | =L8 de 'Tit.Privados Exterior'                        |
| RENTA FIJA EN PESOS (Enajenacion) pesos  | =K4+K7 de 'Renta Fija Pesos'                          |
| RENTA FIJA EN DOLARES (Enajenacion) usd  | =L16+L20+L26+L48+L60+L67 de 'Renta Fija Dolares'      |
| RENTA FIJA EN DOLARES (Renta) usd        | =L17+L27+L49+L61+L68 de 'Renta Fija Dolares'          |
| CAUCIONES EN PESOS (Enajenacion) pesos   | =J3 de 'Cauciones Pesos'                              |

Uso:
----
    from validation_module import validate_visual, validate_gallo, run_full_validation
    
    # Validar un archivo Visual
    result = validate_visual('archivo_visual.xlsx')
    
    # Validar un archivo Gallo  
    result = validate_gallo('archivo_gallo.xlsx')
    
    # Validación completa con reporte
    run_full_validation('archivo.xlsx', tipo='visual')  # o 'gallo'
"""

import pandas as pd
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional
import os


@dataclass
class ValidationResult:
    """Resultado de una validación individual"""
    field: str
    calculated: float
    expected: float
    match: bool
    tolerance: float = 0.01
    
    @property
    def difference(self) -> float:
        return abs(self.calculated - self.expected)


@dataclass  
class ValidationReport:
    """Reporte completo de validación"""
    file_path: str
    file_type: str  # 'visual' o 'gallo'
    results: List[ValidationResult]
    
    @property
    def all_passed(self) -> bool:
        return all(r.match for r in self.results)
    
    @property
    def passed_count(self) -> int:
        return sum(1 for r in self.results if r.match)
    
    @property
    def failed_count(self) -> int:
        return sum(1 for r in self.results if not r.match)
    
    def print_report(self):
        """Imprime el reporte de validación"""
        print("=" * 80)
        print(f"REPORTE DE VALIDACIÓN - {self.file_type.upper()}")
        print(f"Archivo: {self.file_path}")
        print("=" * 80)
        print()
        
        print(f"{'Campo':<45} {'Calculado':>18} {'Esperado':>18} {'Match':>6}")
        print("-" * 90)
        
        for r in self.results:
            status = '✓' if r.match else '✗'
            print(f"{r.field:<45} {r.calculated:>18,.2f} {r.expected:>18,.2f} {status:>6}")
        
        print()
        print(f"RESULTADO: {self.passed_count}/{len(self.results)} validaciones pasaron")
        if self.all_passed:
            print("✓ TODAS LAS VERIFICACIONES PASARON")
        else:
            print("✗ HAY DIFERENCIAS - REVISAR")
        print()


def validate_visual(file_path: str, tolerance: float = 0.01) -> ValidationReport:
    """
    Valida un archivo Visual verificando las relaciones matemáticas
    entre la hoja Resumen y las hojas de detalle.
    
    Args:
        file_path: Ruta al archivo Excel de Visual
        tolerance: Tolerancia para comparación de números flotantes
        
    Returns:
        ValidationReport con los resultados
    """
    results = []
    xls = pd.ExcelFile(file_path)
    
    # Cargar hoja Resumen
    resumen = pd.read_excel(xls, sheet_name='Resumen')
    
    # 1. Ventas ARS = SUM(Resultado Ventas ARS.resultado)
    if 'Resultado Ventas ARS' in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name='Resultado Ventas ARS')
        calc = df['resultado'].sum() if len(df) > 0 and 'resultado' in df.columns else 0
    else:
        calc = 0
    expected = resumen.iloc[0]['ventas']
    results.append(ValidationResult('ARS ventas (B2)', calc, expected, abs(calc - expected) < tolerance))
    
    # 2. Ventas USD = SUM(Resultado Ventas USD.resultado)
    if 'Resultado Ventas USD' in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name='Resultado Ventas USD')
        calc = df['resultado'].sum() if len(df) > 0 and 'resultado' in df.columns else 0
    else:
        calc = 0
    expected = resumen.iloc[1]['ventas']
    results.append(ValidationResult('USD ventas (B3)', calc, expected, abs(calc - expected) < tolerance))
    
    # 3. Rentas ARS = SUMIF(Rentas Dividendos ARS, categoria="RENTAS", importe)
    if 'Rentas Dividendos ARS' in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name='Rentas Dividendos ARS')
        if len(df) > 0 and 'categoria' in df.columns and 'importe' in df.columns:
            calc = df[df['categoria'].str.upper() == 'RENTAS']['importe'].sum()
        else:
            calc = 0
    else:
        calc = 0
    expected = resumen.iloc[0]['rentas']
    results.append(ValidationResult('ARS rentas (E2)', calc, expected, abs(calc - expected) < tolerance))
    
    # 4. Rentas USD = SUMIF(Rentas Dividendos USD, categoria="Rentas", importe)
    if 'Rentas Dividendos USD' in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name='Rentas Dividendos USD')
        if len(df) > 0 and 'categoria' in df.columns and 'importe' in df.columns:
            calc = df[df['categoria'].str.lower() == 'rentas']['importe'].sum()
        else:
            calc = 0
    else:
        calc = 0
    expected = resumen.iloc[1]['rentas']
    results.append(ValidationResult('USD rentas (E3)', calc, expected, abs(calc - expected) < tolerance))
    
    # 5. Dividendos ARS = SUMIF(Rentas Dividendos ARS, categoria="DIVIDENDOS", importe)
    if 'Rentas Dividendos ARS' in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name='Rentas Dividendos ARS')
        if len(df) > 0 and 'categoria' in df.columns and 'importe' in df.columns:
            calc = df[df['categoria'].str.upper() == 'DIVIDENDOS']['importe'].sum()
        else:
            calc = 0
    else:
        calc = 0
    expected = resumen.iloc[0]['dividendos']
    results.append(ValidationResult('ARS dividendos (F2)', calc, expected, abs(calc - expected) < tolerance))
    
    # 6. Dividendos USD = SUMIF(Rentas Dividendos USD, categoria="Dividendos", importe)
    if 'Rentas Dividendos USD' in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name='Rentas Dividendos USD')
        if len(df) > 0 and 'categoria' in df.columns and 'importe' in df.columns:
            calc = df[df['categoria'].str.lower() == 'dividendos']['importe'].sum()
        else:
            calc = 0
    else:
        calc = 0
    expected = resumen.iloc[1]['dividendos']
    results.append(ValidationResult('USD dividendos (F3)', calc, expected, abs(calc - expected) < tolerance))
    
    # 7. Total ARS = suma de columnas B-K fila 2
    total_calc = sum([resumen.iloc[0][col] for col in ['ventas', 'fci', 'opciones', 'rentas', 'dividendos', 
                                                         'ef_cpd', 'pagares', 'futuros', 'cau_int', 'cau_cf']])
    expected = resumen.iloc[0]['total']
    results.append(ValidationResult('ARS total (L2)', total_calc, expected, abs(total_calc - expected) < tolerance))
    
    # 8. Total USD = suma de columnas B-K fila 3
    total_calc = sum([resumen.iloc[1][col] for col in ['ventas', 'fci', 'opciones', 'rentas', 'dividendos',
                                                         'ef_cpd', 'pagares', 'futuros', 'cau_int', 'cau_cf']])
    expected = resumen.iloc[1]['total']
    results.append(ValidationResult('USD total (L3)', total_calc, expected, abs(total_calc - expected) < tolerance))
    
    return ValidationReport(file_path, 'visual', results)


def validate_gallo(file_path: str, tolerance: float = 0.01) -> ValidationReport:
    """
    Valida un archivo Gallo verificando las relaciones matemáticas
    entre la hoja Resultado Totales y las hojas de detalle.

    VERSIÓN DINÁMICA: Lee las categorías de Resultado Totales y valida
    automáticamente cada una contra su hoja de detalle correspondiente.

    Args:
        file_path: Ruta al archivo Excel de Gallo
        tolerance: Tolerancia para comparación de números flotantes

    Returns:
        ValidationReport con los resultados
    """
    results = []
    xls = pd.ExcelFile(file_path)

    # Cargar hoja Resultado Totales
    totales = pd.read_excel(xls, sheet_name='Resultado Totales')
    
    print(f"[DEBUG] Categorías en Resultado Totales: {totales['categoria'].tolist()}")

    # Mapeo de categorías a hojas de detalle
    categoria_to_sheet = {
        'tit.privados exentos': 'Tit.Privados Exentos',
        'tit privados exentos': 'Tit.Privados Exentos',
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
    }

    def get_sheet_for_categoria(cat_base):
        """Encuentra la hoja correspondiente a una categoría"""
        cat_lower = cat_base.lower()
        for key, sheet in categoria_to_sheet.items():
            if key in cat_lower or cat_lower in key:
                return sheet
        return None

    def clean_numeric_column(df, col_name):
        """Limpia una columna numérica quitando comas de miles"""
        def to_float(x):
            if pd.isna(x):
                return 0.0
            if isinstance(x, (int, float)):
                return float(x)
            s = str(x).replace(',', '').strip()
            # Manejar números negativos con signo al final (ej: "538.62-")
            if s.endswith('-'):
                s = '-' + s[:-1]
            try:
                return float(s)
            except:
                return 0.0
        df[col_name] = df[col_name].apply(to_float)
        return df

    def sum_transactions(sheet_name: str, tipo_pattern: str, col_name: str) -> float:
        """Suma valores de TRANSACCIONES individuales (excluyendo filas de Total).
        tipo_pattern: 'enajenacion' para compra/venta, 'renta' para rentas/dividendos/amortizaciones
        """
        if sheet_name not in xls.sheet_names:
            print(f"  [DEBUG] Hoja '{sheet_name}' no encontrada")
            return 0
        df = pd.read_excel(xls, sheet_name=sheet_name)
        if col_name not in df.columns:
            print(f"  [DEBUG] Columna '{col_name}' no encontrada en '{sheet_name}'")
            return 0
        
        # Limpiar números con comas
        df = clean_numeric_column(df, col_name)

        # Excluir filas de Total
        if 'tipo_fila' in df.columns:
            tipo_lower = df['tipo_fila'].astype(str).str.lower().fillna('')
            is_total = tipo_lower.str.contains('total', na=False)
        else:
            is_total = pd.Series([False] * len(df))
        
        # Filtrar por tipo de operación
        if 'operacion' in df.columns:
            oper_lower = df['operacion'].astype(str).str.lower().fillna('')
            
            if tipo_pattern.lower() == 'enajenacion':
                # Enajenación: compra, venta, amortización, ret ajuste (no renta/dividendo)
                is_enajenacion = (
                    oper_lower.str.contains('compra|venta|amortizacion|cpra|cable|ret', na=False, regex=True) &
                    ~oper_lower.str.contains('renta|dividendo', na=False, regex=True)
                )
                mask = ~is_total & is_enajenacion
            else:  # renta
                # Renta: solo operaciones de renta/dividendo
                is_renta = oper_lower.str.contains('renta|dividendo', na=False, regex=True)
                mask = ~is_total & is_renta
        else:
            mask = ~is_total
        
        return df[mask][col_name].sum()

    def sum_cauciones(sheet_name: str, col_name: str) -> float:
        """Suma intereses de cauciones (excluyendo filas de total)"""
        if sheet_name not in xls.sheet_names:
            return 0
        df = pd.read_excel(xls, sheet_name=sheet_name)
        if col_name not in df.columns:
            return 0
        
        # Limpiar números
        df = clean_numeric_column(df, col_name)
        
        if 'tipo_fila' in df.columns:
            tipo_lower = df['tipo_fila'].astype(str).str.lower()
            mask = ~tipo_lower.str.contains('total', na=False)
            return df[mask][col_name].sum()
        else:
            return df[col_name].sum()

    import re
    
    def clean_number(val):
        """Limpia un valor numérico con posibles comas de miles"""
        if pd.isna(val):
            return 0
        if isinstance(val, (int, float)):
            return float(val)
        # Quitar comas de miles
        val_str = str(val).replace(',', '')
        try:
            return float(val_str)
        except:
            return 0
    
    # Procesar cada categoría en Resultado Totales (excepto TOTAL GENERAL)
    for _, row in totales.iterrows():
        categoria = str(row['categoria']).strip()
        valor_pesos = clean_number(row['valor_pesos'])
        valor_usd = clean_number(row['valor_usd'])
        
        # Ignorar TOTAL GENERAL
        if 'total general' in categoria.lower():
            continue
        
        # Extraer el tipo (Enajenacion/Renta) y la categoría base
        match = re.match(r'(.+)\s*\((.+)\)', categoria)
        if match:
            cat_base = match.group(1).strip()
            tipo = match.group(2).strip()
        else:
            cat_base = categoria
            tipo = 'enajenacion'  # default
        
        # Encontrar la hoja correspondiente
        sheet_name = get_sheet_for_categoria(cat_base)
        if sheet_name is None:
            print(f"  [WARN] No se encontró hoja para categoría: {cat_base}")
            continue
        
        # Determinar qué columna y valor usar
        if 'caucion' in cat_base.lower():
            # Para cauciones, usar interes_pesos o interes_usd
            if valor_pesos != 0:
                col_name = 'interes_pesos'
                expected = valor_pesos
                field_suffix = 'pesos'
            else:
                col_name = 'interes_usd'
                expected = valor_usd
                field_suffix = 'usd'
            calc = sum_cauciones(sheet_name, col_name)
        else:
            # Para otras secciones, usar resultado_pesos o resultado_usd
            if valor_pesos != 0:
                col_name = 'resultado_pesos'
                expected = valor_pesos
                field_suffix = 'pesos'
            else:
                col_name = 'resultado_usd'
                expected = valor_usd
                field_suffix = 'usd'
            # Sumar transacciones individuales (no filas de Total)
            calc = sum_transactions(sheet_name, tipo, col_name)
        
        field_name = f"{categoria} {field_suffix}"
        match_result = abs(calc - expected) < tolerance
        
        print(f"  [VAL] {field_name}: calc={calc}, expected={expected}, match={match_result}")
        
        results.append(ValidationResult(field_name, calc, expected, match_result))

    return ValidationReport(file_path, 'gallo', results)


def detect_file_type(file_path: str) -> Optional[str]:
    """
    Detecta automáticamente si un archivo es de tipo Visual o Gallo
    basándose en los nombres de las hojas.
    
    Returns:
        'visual', 'gallo', o None si no se puede determinar
    """
    try:
        xls = pd.ExcelFile(file_path)
        sheets = set(xls.sheet_names)
        
        # Hojas características de Visual
        visual_sheets = {'Boletos', 'Resultado Ventas ARS', 'Resultado Ventas USD', 'Resumen'}
        
        # Hojas características de Gallo
        gallo_sheets = {'Resultado Totales', 'Tit.Privados Exentos', 'Cauciones Pesos'}
        
        visual_match = len(visual_sheets & sheets)
        gallo_match = len(gallo_sheets & sheets)
        
        if visual_match > gallo_match:
            return 'visual'
        elif gallo_match > visual_match:
            return 'gallo'
        else:
            return None
    except Exception:
        return None


def run_full_validation(file_path: str, file_type: Optional[str] = None) -> ValidationReport:
    """
    Ejecuta validación completa de un archivo.
    
    Args:
        file_path: Ruta al archivo Excel
        file_type: 'visual' o 'gallo'. Si es None, se detecta automáticamente.
        
    Returns:
        ValidationReport con los resultados
    """
    if file_type is None:
        file_type = detect_file_type(file_path)
        if file_type is None:
            raise ValueError(f"No se pudo determinar el tipo de archivo: {file_path}")
    
    if file_type.lower() == 'visual':
        report = validate_visual(file_path)
    elif file_type.lower() == 'gallo':
        report = validate_gallo(file_path)
    else:
        raise ValueError(f"Tipo de archivo no soportado: {file_type}")
    
    report.print_report()
    return report


if __name__ == "__main__":
    print("=" * 80)
    print("SISTEMA DE VERIFICACIÓN ESTANDARIZADA")
    print("=" * 80)
    print()
    
    # Validar archivo Visual
    if os.path.exists('test_visual_v2.xlsx'):
        print("\n>>> Validando archivo VISUAL...")
        report_visual = run_full_validation('test_visual_v2.xlsx', 'visual')
    
    # Validar archivo Gallo
    if os.path.exists('test_gallo_v2.xlsx'):
        print("\n>>> Validando archivo GALLO...")
        report_gallo = run_full_validation('test_gallo_v2.xlsx', 'gallo')
    
    print("\n" + "=" * 80)
    print("RESUMEN FINAL")
    print("=" * 80)
    
    if 'report_visual' in dir():
        print(f"Visual: {'✓ PASS' if report_visual.all_passed else '✗ FAIL'} ({report_visual.passed_count}/{len(report_visual.results)})")
    
    if 'report_gallo' in dir():
        print(f"Gallo:  {'✓ PASS' if report_gallo.all_passed else '✗ FAIL'} ({report_gallo.passed_count}/{len(report_gallo.results)})")
