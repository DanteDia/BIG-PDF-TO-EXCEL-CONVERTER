"""
Script para separar la columna 'Especie' en 'Instrumento' y 'Cod.Instrum'
"""
import pandas as pd

def split_especie(excel_path, output_path):
    """Procesa un Excel separando la columna Especie"""
    
    sheets_to_process = ['Resultado Ventas ARS', 'Resultado Ventas USD', 
                          'Rentas Dividendos ARS', 'Rentas Dividendos USD']
    
    # Leer todas las hojas
    xl = pd.ExcelFile(excel_path)
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            
            # Si es una hoja que necesita procesamiento y tiene columna Especie
            if sheet_name in sheets_to_process and 'Especie' in df.columns:
                print(f"📊 Procesando {sheet_name}...")
                
                # Separar Especie en Instrumento y Cod.Instrum
                df['Instrumento'] = df['Especie'].str.split(' / ').str[0]
                df['Cod.Instrum'] = df['Especie'].str.split(' / ').str[1]
                
                # Eliminar columna Especie
                df = df.drop(columns=['Especie'])
                
                # Reordenar columnas para que Instrumento y Cod.Instrum estén al inicio
                cols = list(df.columns)
                # Mover Tipo de Instrumento, Instrumento, Cod.Instrum al principio
                if 'Tipo de Instrumento' in cols:
                    cols.remove('Tipo de Instrumento')
                    cols.remove('Instrumento')
                    cols.remove('Cod.Instrum')
                    cols = ['Tipo de Instrumento', 'Instrumento', 'Cod.Instrum'] + cols
                else:
                    cols.remove('Instrumento')
                    cols.remove('Cod.Instrum')
                    cols = ['Instrumento', 'Cod.Instrum'] + cols
                
                df = df[cols]
                print(f"  ✅ {len(df)} filas procesadas")
            
            # Guardar hoja (procesada o no)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"\n✅ Archivo guardado: {output_path}")

if __name__ == "__main__":
    input_file = "VeroLandro2025 (1)_claude_v2.xlsx"
    output_file = "VeroLandro2025_final.xlsx"
    
    split_especie(input_file, output_file)
