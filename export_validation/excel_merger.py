"""
Merger de Excels Gallo + Visual
Combina los dos reportes estructurados en un Excel unificado
"""
import pandas as pd
from datetime import datetime

class ExcelMerger:
    def __init__(self, gallo_excel_path, visual_excel_path):
        self.gallo_path = gallo_excel_path
        self.visual_path = visual_excel_path
        
    def merge(self, output_path):
        """Combina los dos excels en uno unificado"""
        print("📊 Iniciando merge de reportes...")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 1. Boletos
            print("  🔄 Mergeando Boletos...")
            self._merge_sheet(writer, "Boletos", "Boletos")
            
            # 2. Resultado Ventas ARS
            print("  🔄 Mergeando Resultado Ventas ARS...")
            self._merge_sheet(writer, "Resultado Ventas ARS", "Resultado Ventas ARS")
            
            # 3. Resultado Ventas USD
            print("  🔄 Mergeando Resultado Ventas USD...")
            self._merge_sheet(writer, "Resultado Ventas USD", "Resultado Ventas USD")
            
            # 4. Rentas Dividendos ARS
            print("  🔄 Mergeando Rentas Dividendos ARS...")
            self._merge_sheet(writer, "Rentas Dividendos ARS", "Rentas Dividendos ARS")
            
            # 5. Rentas Dividendos USD
            print("  🔄 Mergeando Rentas Dividendos USD...")
            self._merge_sheet(writer, "Rentas Dividendos USD", "Rentas Dividendos USD")
            
            # 6. Resumen (sumar totales)
            print("  🔄 Calculando Resumen Consolidado...")
            self._merge_resumen(writer)
            
            # 7. Posición Títulos (sumar cantidades)
            print("  🔄 Consolidando Posición de Títulos...")
            self._merge_posicion(writer)
        
        print(f"✅ Merge completado: {output_path}")
        return output_path
    
    def _merge_sheet(self, writer, sheet_name, output_name):
        """Combina una hoja de ambos excels agregando columna Origen"""
        try:
            # Leer de Gallo
            df_gallo = pd.read_excel(self.gallo_path, sheet_name=sheet_name)
            df_gallo['Origen'] = 'Gallo'
        except Exception as e:
            print(f"    ⚠️ No se encontró {sheet_name} en Gallo: {e}")
            df_gallo = pd.DataFrame()
        
        try:
            # Leer de Visual
            df_visual = pd.read_excel(self.visual_path, sheet_name=sheet_name)
            df_visual['Origen'] = 'Visual'
        except Exception as e:
            print(f"    ⚠️ No se encontró {sheet_name} en Visual: {e}")
            df_visual = pd.DataFrame()
        
        # Combinar
        if not df_gallo.empty and not df_visual.empty:
            # Asegurar que tengan las mismas columnas
            all_cols = list(set(df_gallo.columns) | set(df_visual.columns))
            for col in all_cols:
                if col not in df_gallo.columns:
                    df_gallo[col] = None
                if col not in df_visual.columns:
                    df_visual[col] = None
            
            # Ordenar columnas igual
            df_gallo = df_gallo[all_cols]
            df_visual = df_visual[all_cols]
            
            df_merged = pd.concat([df_gallo, df_visual], ignore_index=True)
        elif not df_gallo.empty:
            df_merged = df_gallo
        elif not df_visual.empty:
            df_merged = df_visual
        else:
            df_merged = pd.DataFrame()
        
        # Ordenar por fecha si existe columna Concertación
        if 'Concertación' in df_merged.columns and not df_merged.empty:
            try:
                df_merged['_fecha_sort'] = pd.to_datetime(df_merged['Concertación'], format='%d/%m/%Y', errors='coerce')
                df_merged = df_merged.sort_values('_fecha_sort')
                df_merged = df_merged.drop(columns=['_fecha_sort'])
            except:
                pass
        
        # Reordenar para que Origen esté al principio
        if 'Origen' in df_merged.columns:
            cols = ['Origen'] + [c for c in df_merged.columns if c != 'Origen']
            df_merged = df_merged[cols]
        
        df_merged.to_excel(writer, sheet_name=output_name, index=False)
        print(f"    ✅ {len(df_merged)} filas en {output_name}")
    
    def _merge_resumen(self, writer):
        """Suma los totales del resumen de ambos reportes"""
        try:
            df_gallo = pd.read_excel(self.gallo_path, sheet_name="Resumen")
        except:
            df_gallo = pd.DataFrame()
        
        try:
            df_visual = pd.read_excel(self.visual_path, sheet_name="Resumen")
        except:
            df_visual = pd.DataFrame()
        
        if df_gallo.empty and df_visual.empty:
            pd.DataFrame().to_excel(writer, sheet_name="Resumen", index=False)
            return
        
        # Combinar sumando por moneda
        df_merged = pd.concat([df_gallo, df_visual], ignore_index=True)
        
        if 'Moneda' in df_merged.columns:
            # Agrupar por moneda y sumar columnas numéricas
            numeric_cols = df_merged.select_dtypes(include=['number']).columns.tolist()
            df_resumen = df_merged.groupby('Moneda', as_index=False)[numeric_cols].sum()
            
            # Mantener el orden: primero ARS, luego USD
            monedas_orden = ['ARS', 'USD']
            df_resumen['_orden'] = df_resumen['Moneda'].apply(lambda x: monedas_orden.index(x) if x in monedas_orden else 999)
            df_resumen = df_resumen.sort_values('_orden').drop(columns=['_orden'])
        else:
            df_resumen = df_merged
        
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
        print(f"    ✅ Resumen consolidado generado")
    
    def _merge_posicion(self, writer):
        """Combina posiciones de títulos sumando cantidades del mismo instrumento"""
        try:
            df_gallo = pd.read_excel(self.gallo_path, sheet_name="Posicion Titulos")
            df_gallo['Origen'] = 'Gallo'
        except:
            df_gallo = pd.DataFrame()
        
        try:
            df_visual = pd.read_excel(self.visual_path, sheet_name="Posicion Titulos")
            df_visual['Origen'] = 'Visual'
        except:
            df_visual = pd.DataFrame()
        
        if df_gallo.empty and df_visual.empty:
            pd.DataFrame().to_excel(writer, sheet_name="Posicion Titulos", index=False)
            return
        
        df_merged = pd.concat([df_gallo, df_visual], ignore_index=True)
        
        # Agrupar por Instrumento/Código y sumar cantidades
        if 'Código' in df_merged.columns and 'Cantidad' in df_merged.columns:
            # Conservar primera aparición de cada columna no numérica
            group_cols = ['Instrumento', 'Código']
            agg_dict = {
                'Cantidad': 'sum',
                'Importe': 'sum',
                'Ticker': 'first',
                'Moneda': 'first',
                'Origen': lambda x: ', '.join(sorted(set(x)))  # Combinar orígenes
            }
            
            # Solo agrupar si las columnas existen
            existing_agg = {k: v for k, v in agg_dict.items() if k in df_merged.columns}
            df_posicion = df_merged.groupby(group_cols, as_index=False).agg(existing_agg)
        else:
            df_posicion = df_merged
        
        df_posicion.to_excel(writer, sheet_name="Posicion Titulos", index=False)
        print(f"    ✅ {len(df_posicion)} instrumentos en posición consolidada")


def merge_excels(gallo_path, visual_path, output_path):
    """Función helper para merge directo"""
    merger = ExcelMerger(gallo_path, visual_path)
    return merger.merge(output_path)


if __name__ == "__main__":
    # Test
    gallo = "gallo_estructurado.xlsx"
    visual = "VeroLandro2025_final.xlsx"
    output = "Reporte_Unificado_FINAL.xlsx"
    
    merge_excels(gallo, visual, output)
