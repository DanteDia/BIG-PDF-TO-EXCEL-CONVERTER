"""Test de extracción de valores del markdown existente."""
import sys
sys.path.insert(0, 'pdf_converter/datalab')

from datalab_excel_reader import DatalabExcelReader

# Leer el markdown que ya tenemos
with open('excel_to_markdown.md', 'r', encoding='utf-8') as f:
    markdown = f.read()

print(f"Markdown length: {len(markdown)}")

# Crear reader (sin API key, solo para parsing)
reader = DatalabExcelReader(api_key="dummy")

# Extraer valores
values = reader.extract_resumen_values(markdown)

print("\n=== VALORES DEL RESUMEN ===")
print(f"Ventas ARS: {values['ventas_ars']:,.2f}")
print(f"Ventas USD: {values['ventas_usd']:,.2f}")
print(f"Total ARS: {values['total_ars']:,.2f}")
print(f"Total USD: {values['total_usd']:,.2f}")

# Verificar contra V8
expected_ventas_ars = -54228.68
expected_ventas_usd = -313519.02

diff_ars = values['ventas_ars'] - expected_ventas_ars
diff_usd = values['ventas_usd'] - expected_ventas_usd

print("\n=== COMPARACIÓN CON V8 ===")
print(f"Ventas ARS esperado: {expected_ventas_ars:,.2f}")
print(f"Ventas ARS obtenido: {values['ventas_ars']:,.2f}")
print(f"Diferencia ARS: {diff_ars:,.2f}")

print(f"\nVentas USD esperado: {expected_ventas_usd:,.2f}")
print(f"Ventas USD obtenido: {values['ventas_usd']:,.2f}")
print(f"Diferencia USD: {diff_usd:,.2f}")

if abs(diff_ars) < 0.01 and abs(diff_usd) < 0.01:
    print("\n✅ ¡VALORES CORRECTOS! Coinciden con V8")
else:
    print("\n❌ DIFERENCIA DETECTADA")
