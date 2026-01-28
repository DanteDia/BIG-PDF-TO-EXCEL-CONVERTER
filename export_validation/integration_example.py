"""
Ejemplo de integración del módulo de validación con cualquier conversor.

Este archivo muestra cómo integrar la validación automática con tu propio
conversor de PDF a Excel.
"""

from validation_module import (
    validate_visual, 
    validate_gallo, 
    run_full_validation,
    detect_file_type,
    ValidationReport
)


def example_integration_with_your_converter():
    """
    Ejemplo de cómo integrar con tu conversor.
    Reemplaza 'TuConversor' con tu implementación.
    """
    
    # 1. Tu conversor genera el Excel
    # from tu_conversor import TuConversor
    # conversor = TuConversor(api_key="tu_api_key")
    # conversor.convert('entrada.pdf', 'salida.xlsx', tipo='gallo')
    
    # 2. Validar el resultado
    excel_path = 'salida.xlsx'
    
    # Opción A: Detectar tipo automáticamente
    report = run_full_validation(excel_path)
    
    # Opción B: Especificar tipo manualmente
    # report = validate_gallo(excel_path)  # o validate_visual(excel_path)
    
    # 3. Verificar resultados
    print(f"\n{'='*60}")
    print(f"VALIDACIÓN: {excel_path}")
    print(f"{'='*60}")
    
    if report.all_passed:
        print("✅ TODAS LAS VALIDACIONES PASARON")
        return True
    else:
        print(f"⚠️ FALLARON {report.failed_count}/{len(report.results)} VALIDACIONES")
        print("\nDetalles de errores:")
        for r in report.results:
            if not r.match:
                diff = r.calculated - r.expected
                print(f"  ❌ {r.field}")
                print(f"     Calculado: {r.calculated:,.2f}")
                print(f"     Esperado:  {r.expected:,.2f}")
                print(f"     Diferencia: {diff:+,.2f}")
        return False


def batch_validation_example(excel_files: list):
    """
    Ejemplo de validación en lote para múltiples archivos.
    """
    results = {}
    
    for file_path in excel_files:
        try:
            report = run_full_validation(file_path)
            results[file_path] = {
                'passed': report.all_passed,
                'score': f"{report.passed_count}/{len(report.results)}",
                'failed_fields': [r.field for r in report.results if not r.match]
            }
        except Exception as e:
            results[file_path] = {
                'passed': False,
                'score': 'ERROR',
                'error': str(e)
            }
    
    # Resumen
    print("\n" + "="*70)
    print("RESUMEN DE VALIDACIÓN EN LOTE")
    print("="*70)
    
    total_passed = sum(1 for r in results.values() if r.get('passed'))
    total_files = len(results)
    
    print(f"Archivos procesados: {total_files}")
    print(f"Pasaron todas las validaciones: {total_passed}")
    print(f"Con errores: {total_files - total_passed}")
    
    print("\nDetalle:")
    for file_path, result in results.items():
        status = "✅" if result.get('passed') else "❌"
        print(f"  {status} {file_path}: {result['score']}")
        if result.get('failed_fields'):
            for field in result['failed_fields']:
                print(f"      - {field}")
        if result.get('error'):
            print(f"      ERROR: {result['error']}")
    
    return results


def custom_validation_callback(report: ValidationReport):
    """
    Ejemplo de callback personalizado para manejar resultados de validación.
    Útil para logging, alertas, o flujos de trabajo automatizados.
    """
    if report.all_passed:
        # Archivo válido - proceder con siguiente paso
        print(f"[OK] {report.file_path} validado correctamente")
        # Aquí podrías: mover a carpeta de éxito, notificar, etc.
    else:
        # Archivo con errores - requiere revisión
        print(f"[ERROR] {report.file_path} tiene {report.failed_count} errores")
        # Aquí podrías: mover a carpeta de revisión, enviar alerta, etc.
        
        # Generar reporte de errores
        errors = []
        for r in report.results:
            if not r.match:
                errors.append({
                    'campo': r.field,
                    'calculado': r.calculated,
                    'esperado': r.expected,
                    'diferencia': r.calculated - r.expected
                })
        return errors
    
    return None


if __name__ == "__main__":
    # Ejemplo de uso standalone
    import sys
    
    if len(sys.argv) > 1:
        # Validar archivo pasado como argumento
        for file_path in sys.argv[1:]:
            try:
                report = run_full_validation(file_path)
                report.print_report()
            except Exception as e:
                print(f"Error validando {file_path}: {e}")
    else:
        print("Uso: python integration_example.py archivo1.xlsx [archivo2.xlsx ...]")
        print("\nO importa el módulo:")
        print("  from validation_module import validate_gallo, validate_visual")
