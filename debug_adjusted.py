#!/usr/bin/env python
"""Debug adjusted price lookup."""
import sys
sys.path.insert(0, '.')
from pathlib import Path
from pdf_converter.datalab.merge_gallo_visual import GalloVisualMerger

merger = GalloVisualMerger(
    '12128_LANDRO_VERONICA_INES_Gallo_Generado_OK.xlsx',
    '12128_LANDRO_VERONICA_INES_Visual_Generado_OK.xlsx',
    str(Path('pdf_converter/datalab/aux_data')),
    precio_tenencias_path='precio_estructurado.xlsx'
)

# Check what's in the caches
print("=== Raw price cache by codigo ===")
for k, v in merger._precio_tenencias_by_codigo.items():
    print(f"  {k}: {v}")

print("\n=== Raw price cache by ticker ===")
for k, v in merger._precio_tenencias_by_ticker.items():
    print(f"  {k}: {v}")

print("\n=== Adjusted price cache by codigo ===")
for k, v in merger._precio_tenencias_adjusted_by_codigo.items():
    print(f"  {k}: {v}")

print("\n=== Adjusted price cache by ticker ===")
for k, v in merger._precio_tenencias_adjusted_by_ticker.items():
    print(f"  {k}: {v}")

# Test lookups for NVDA-US (codigo 7044)
print("\n=== Lookups ===")
print(f"get_precio_tenencia_inicial('7044', 'NVDA-US'): {merger._get_precio_tenencia_inicial('7044', 'NVDA-US')}")
print(f"get_precio_tenencia_inicial_adjusted('7044', 'NVDA-US'): {merger._get_precio_tenencia_inicial_adjusted('7044', 'NVDA-US')}")
print(f"get_precio_tenencia_inicial('49315', 'TSLA-US'): {merger._get_precio_tenencia_inicial('49315', 'TSLA-US')}")
print(f"get_precio_tenencia_inicial_adjusted('49315', 'TSLA-US'): {merger._get_precio_tenencia_inicial_adjusted('49315', 'TSLA-US')}")

# Check is_accion_exterior
print(f"\nis_accion_exterior('7044'): {merger._is_accion_exterior('7044')}")
print(f"is_accion_exterior('49315'): {merger._is_accion_exterior('49315')}")

# Check ratio lookup
print(f"\nget_ratio_for_especie('NVIDIA', 'CORPORATION'): {merger._get_ratio_for_especie('NVIDIA', 'CORPORATION')}")
print(f"get_ratio_for_especie('TESLA', 'MOTORS INC.'): {merger._get_ratio_for_especie('TESLA', 'MOTORS INC.')}")
