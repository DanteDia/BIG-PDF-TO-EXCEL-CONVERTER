"""
Test de exportación a PDF usando Datalab API para valores de fórmulas.
"""
import sys
import os
sys.path.insert(0, '.')

# Simular que NO tenemos Excel COM deshabilitando Windows check
original_platform = sys.platform

from pdf_converter.datalab.excel_to_pdf import ExcelToPdfExporter

# Test
def test_pdf_export_with_datalab():
    api_key = "Z1JLPnKJIAcNosYvpyF-GZrjVizmsf5MsBlNGL7Szuk"
    
    cliente_info = {
        'numero': '12345',
        'nombre': 'CLIENTE TEST'
    }
    
    print("Creando exportador con Datalab API...")
    
    # Forzar uso de Datalab deshabilitando COM temporalmente
    import pdf_converter.datalab.excel_to_pdf as pdf_module
    original_func = pdf_module._read_excel_with_com
    pdf_module._read_excel_with_com = lambda x: None  # Forzar fallback a Datalab
    
    try:
        exporter = ExcelToPdfExporter(
            "TEST_Merge_v8.xlsx", 
            cliente_info,
            datalab_api_key=api_key
        )
        
        exporter.periodo_inicio = "Enero 1"
        exporter.periodo_fin = "Diciembre 31"
        exporter.anio = 2025
        
        print("Generando PDF...")
        pdf_bytes = exporter.export_to_pdf()
        
        # Guardar PDF
        output_path = "TEST_PDF_Datalab.pdf"
        with open(output_path, 'wb') as f:
            f.write(pdf_bytes)
        
        print(f"\n✅ PDF generado exitosamente: {output_path}")
        print(f"   Tamaño: {len(pdf_bytes):,} bytes")
        
    finally:
        # Restaurar función original
        pdf_module._read_excel_with_com = original_func


if __name__ == "__main__":
    test_pdf_export_with_datalab()
