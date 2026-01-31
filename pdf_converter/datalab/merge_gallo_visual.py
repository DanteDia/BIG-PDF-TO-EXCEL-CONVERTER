"""
Módulo para unificar archivos Excel de Gallo y Visual en un resumen impositivo consolidado.
Traduce la estructura de Gallo al esquema de Visual y genera hojas de resultados.
"""

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import re


class GalloVisualMerger:
    """
    Clase principal para unificar Excel de Gallo y Visual.
    
    Genera un Excel consolidado con:
    - Posicion Inicial Gallo
    - Posicion Final Gallo  
    - Boletos (merge de transacciones)
    - Rentas y Dividendos Gallo
    - Resultado Ventas ARS
    - Resultado Ventas USD
    - Rentas Dividendos ARS
    - Rentas Dividendos USD
    - Resumen
    - Hojas auxiliares (EspeciesVisual, EspeciesGallo, etc.)
    """
    
    # Operaciones de compra/venta para Boletos
    OPERACIONES_BOLETOS = ['compra', 'venta', 'cpra', 'canje', 'licitacion', 'licitaciones', 
                           'compra usd', 'venta usd', 'cpra cable', 'venta cable']
    
    # Operaciones de rentas/dividendos
    OPERACIONES_RENTAS = ['renta', 'dividendo', 'dividendos', 'amortizacion', 'amortizaciones']
    
    # Hojas de transacciones en Gallo
    HOJAS_GALLO_TRANSACCIONES = ['Tit Privados Exentos', 'Renta Fija Dolares', 'Tit Privados Exterior',
                                  'Cauciones', 'Titulos Publicos', 'Cedears']
    
    def __init__(self, gallo_path: str, visual_path: str, aux_data_dir: str = None):
        """
        Inicializa el merger con las rutas a los archivos.
        
        Args:
            gallo_path: Ruta al Excel generado de Gallo
            visual_path: Ruta al Excel generado de Visual
            aux_data_dir: Directorio con hojas auxiliares (default: pdf_converter/datalab/aux_data)
        """
        self.gallo_path = Path(gallo_path)
        self.visual_path = Path(visual_path)
        
        if aux_data_dir is None:
            aux_data_dir = Path(__file__).parent / 'aux_data'
        self.aux_data_dir = Path(aux_data_dir)
        
        # Cargar workbooks
        self.gallo_wb = load_workbook(gallo_path)
        self.visual_wb = load_workbook(visual_path)
        
        # Cargar hojas auxiliares
        self.especies_visual = self._load_aux('EspeciesVisual.xlsx')
        self.especies_gallo = self._load_aux('EspeciesGallo.xlsx')
        self.cotizacion_dolar = self._load_aux('Cotizacion_Dolar_Historica.xlsx')
        self.precios_iniciales = self._load_aux('PreciosInicialesEspecies.xlsx')
        
        # Cache de mapeos
        self._especies_visual_cache = {}
        self._especies_gallo_cache = {}
        self._cotizacion_cache = {}
        self._precios_iniciales_cache = {}
        
        # Construir caches
        self._build_caches()
    
    def _load_aux(self, filename: str) -> Workbook:
        """Carga un archivo auxiliar."""
        path = self.aux_data_dir / filename
        if not path.exists():
            raise FileNotFoundError(f"Archivo auxiliar no encontrado: {path}")
        return load_workbook(path)
    
    def _build_caches(self):
        """Construye caches para búsquedas rápidas."""
        # Cache EspeciesVisual: codigo -> {nombre, moneda_emision, tipo_especie, ...}
        ws = self.especies_visual.active
        for row in range(2, ws.max_row + 1):
            codigo = ws.cell(row, 3).value  # Columna C = codigo
            if codigo:
                codigo_clean = self._clean_codigo(codigo)
                self._especies_visual_cache[codigo_clean] = {
                    'codigo': codigo,
                    'moneda_emision': ws.cell(row, 7).value,  # Col G
                    'ticker': ws.cell(row, 8).value,  # Col H
                    'nombre_con_moneda': ws.cell(row, 17).value,  # Col Q
                    'tipo_especie': ws.cell(row, 18).value,  # Col R
                }
        
        # Cache EspeciesGallo: codigo -> {nombre, ticker, moneda_emision}
        ws = self.especies_gallo.active
        for row in range(2, ws.max_row + 1):
            codigo = ws.cell(row, 1).value  # Columna A
            if codigo:
                codigo_clean = self._clean_codigo(codigo)
                self._especies_gallo_cache[codigo_clean] = {
                    'codigo': codigo,
                    'nombre': ws.cell(row, 2).value,  # Col B
                    'ticker': ws.cell(row, 10).value,  # Col J
                    'moneda_emision': ws.cell(row, 14).value,  # Col N
                }
        
        # Cache Cotización Dólar: (fecha, tipo_dolar) -> cotizacion
        ws = self.cotizacion_dolar.active
        for row in range(2, ws.max_row + 1):
            fecha = ws.cell(row, 1).value
            cotizacion = ws.cell(row, 2).value
            tipo_dolar = ws.cell(row, 3).value
            if fecha and cotizacion:
                # Normalizar fecha
                if isinstance(fecha, datetime):
                    fecha_key = fecha.date()
                else:
                    fecha_key = fecha
                self._cotizacion_cache[(fecha_key, tipo_dolar)] = cotizacion
        
        # Cache Precios Iniciales: ticker -> {codigo, precio}
        # Col A = codigo, Col B = nombre, Col C = ticker/ORDEN, Col G = precio
        ws = self.precios_iniciales.active
        for row in range(2, ws.max_row + 1):
            codigo = ws.cell(row, 1).value  # Col A = codigo especie
            ticker = ws.cell(row, 3).value  # Col C = ORDEN/ticker
            precio = ws.cell(row, 7).value  # Col G = precio
            if ticker:
                ticker_key = str(ticker).upper().strip()
                self._precios_iniciales_cache[ticker_key] = {
                    'codigo': int(codigo) if codigo else None,
                    'precio': precio if precio else 0
                }
    
    def _clean_codigo(self, codigo) -> str:
        """Limpia código de especie: quita puntos, ceros a izquierda, etc."""
        if codigo is None:
            return ""
        codigo_str = str(codigo).strip()
        # Quitar puntos
        codigo_str = codigo_str.replace('.', '').replace(',', '')
        # Quitar ceros a la izquierda
        try:
            return str(int(float(codigo_str)))
        except:
            return codigo_str
    
    def _split_especie(self, especie: str) -> Tuple[str, str]:
        """Divide especie en Ticker y resto del nombre."""
        if not especie:
            return "", ""
        parts = str(especie).strip().split(' ', 1)
        ticker = parts[0] if parts else ""
        resto = parts[1] if len(parts) > 1 else ""
        return ticker, resto
    
    def _get_moneda(self, resultado_pesos, resultado_usd, gastos_pesos, gastos_usd, hoja_origen: str, operacion: str = "") -> str:
        """Determina la moneda basándose en los valores, la hoja de origen y la operación."""
        operacion_lower = str(operacion).lower() if operacion else ""
        
        # Si la operación menciona USD/EXT/CABLE
        if 'usd' in operacion_lower or 'ext' in operacion_lower or 'cable' in operacion_lower:
            if 'exterior' in hoja_origen.lower() or 'cable' in operacion_lower:
                return "Dolar Cable"
            else:
                return "Dolar MEP"
        
        # Si hay valores en pesos
        if resultado_pesos and float(resultado_pesos) != 0:
            return "Pesos"
        if gastos_pesos and float(gastos_pesos) != 0:
            return "Pesos"
        
        # Si hay valores en USD, determinar tipo
        if resultado_usd or gastos_usd:
            if 'exterior' in hoja_origen.lower():
                return "Dolar Cable"
            else:
                return "Dolar MEP"
        
        # Por defecto usar la hoja de origen
        if 'dolar' in hoja_origen.lower():
            return "Dolar MEP"
        
        return "Pesos"  # Default
    
    def _parse_fecha(self, fecha_value) -> Tuple[datetime, int]:
        """Parsea fecha de varios formatos. Retorna (datetime, año)."""
        if fecha_value is None:
            return None, 0
        
        if isinstance(fecha_value, datetime):
            return fecha_value, fecha_value.year
        
        if isinstance(fecha_value, str):
            fecha_str = fecha_value.strip()
            # Formato dd/mm/yy o dd/mm/yyyy
            try:
                parts = fecha_str.split('/')
                if len(parts) == 3:
                    day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
                    # Ajustar año de 2 dígitos
                    if year < 100:
                        year = 2000 + year if year < 50 else 1900 + year
                    return datetime(year, month, day), year
            except:
                pass
        
        return None, 0
    
    def _is_year_2025(self, fecha_value) -> bool:
        """Verifica si una fecha corresponde a 2025."""
        _, year = self._parse_fecha(fecha_value)
        return year == 2025
    
    def _buscar_codigo_especie(self, especie: str, tipo_especie: str = None) -> Tuple[str, str]:
        """
        Busca el código de especie en las hojas de transacciones de Gallo.
        Retorna (codigo, origen) donde origen es 'Gallo' o 'EspeciesGallo'.
        """
        especie_upper = str(especie).upper().strip()
        
        # Primero buscar en transacciones de Gallo
        for sheet_name in self.gallo_wb.sheetnames:
            if sheet_name in ['Posicion Inicial', 'Posicion Final', 'Resultados']:
                continue
            try:
                ws = self.gallo_wb[sheet_name]
                for row in range(2, min(ws.max_row + 1, 500)):  # Limitar búsqueda
                    cod = ws.cell(row, 2).value  # Col B = cod_especie
                    esp = ws.cell(row, 3).value  # Col C = especie
                    if esp and self._match_especie(especie_upper, str(esp).upper()):
                        if cod:
                            return self._clean_codigo(cod), "Gallo"
            except:
                continue
        
        # Fallback: buscar en EspeciesGallo por nombre similar
        for codigo, data in self._especies_gallo_cache.items():
            nombre = data.get('nombre', '')
            if nombre and self._match_especie(especie_upper, str(nombre).upper()):
                return codigo, "EspeciesGallo"
        
        return "", "NoEncontrado"
    
    def _match_especie(self, especie1: str, especie2: str) -> bool:
        """Verifica si dos especies hacen match (fuzzy)."""
        # Limpieza básica
        e1 = especie1.replace('-', ' ').replace('/', ' ').strip()
        e2 = especie2.replace('-', ' ').replace('/', ' ').strip()
        
        # Match exacto
        if e1 == e2:
            return True
        
        # Match sin ticker (primera palabra)
        words1 = e1.split()
        words2 = e2.split()
        
        # Si uno contiene al otro
        if e1 in e2 or e2 in e1:
            return True
        
        # Match por palabras clave (excluyendo ticker)
        if len(words1) > 1 and len(words2) > 1:
            rest1 = ' '.join(words1[1:])
            rest2 = ' '.join(words2[1:])
            if rest1 == rest2 or rest1 in rest2 or rest2 in rest1:
                return True
        
        return False
    
    def _get_cotizacion(self, fecha, tipo_moneda: str) -> float:
        """Obtiene cotización del dólar para una fecha y tipo."""
        if tipo_moneda == "Pesos":
            return 1.0
        
        # Normalizar tipo
        tipo_key = "Dolar MEP" if "MEP" in tipo_moneda.upper() else "Dolar Cable"
        
        # Normalizar fecha
        if isinstance(fecha, datetime):
            fecha_key = fecha.date()
        else:
            fecha_key = fecha
        
        return self._cotizacion_cache.get((fecha_key, tipo_key), 1.0)
    
    def _generate_ticker_variations(self, ticker: str) -> List[str]:
        """
        Genera variaciones de ticker cambiando 0↔O para manejar errores de OCR.
        Ej: TLC10 -> [TLC10, TLC1O], TL0C0 -> [TL0C0, TLOCO, TL0CO, TLOC0]
        """
        ticker_upper = str(ticker).upper().strip()
        variations = [ticker_upper]
        
        # Encontrar posiciones de 0 y O
        positions_0 = [i for i, c in enumerate(ticker_upper) if c == '0']
        positions_O = [i for i, c in enumerate(ticker_upper) if c == 'O']
        
        # Si hay 0, generar versión con O
        for pos in positions_0:
            new_ticker = ticker_upper[:pos] + 'O' + ticker_upper[pos+1:]
            if new_ticker not in variations:
                variations.append(new_ticker)
        
        # Si hay O, generar versión con 0
        for pos in positions_O:
            new_ticker = ticker_upper[:pos] + '0' + ticker_upper[pos+1:]
            if new_ticker not in variations:
                variations.append(new_ticker)
        
        return variations
    
    def _get_precio_inicial(self, ticker: str) -> float:
        """Obtiene precio inicial de una especie por ticker."""
        ticker_upper = str(ticker).upper().strip()
        
        # Valores fijos para monedas
        if ticker_upper in ['PESOS', '$']:
            return 1.0
        if ticker_upper in ['DOLARES', 'USD', 'U$S', 'DOLAR']:
            return 1167.806
        if 'CABLE' in ticker_upper:
            return 1148.93
        
        # Primero probar ticker exacto
        data = self._precios_iniciales_cache.get(ticker_upper, {})
        if isinstance(data, dict) and data.get('precio'):
            return data.get('precio', 0)
        
        # Si no encuentra, probar variaciones de ticker (OCR 0↔O)
        for ticker_var in self._generate_ticker_variations(ticker_upper):
            data = self._precios_iniciales_cache.get(ticker_var, {})
            if isinstance(data, dict) and data.get('precio'):
                return data.get('precio', 0)
        
        return 0
    
    def _get_codigo_from_ticker(self, ticker: str) -> Optional[int]:
        """Obtiene código de especie desde el ticker usando PreciosInicialesEspecies."""
        ticker_upper = str(ticker).upper().strip()
        
        # Las monedas no tienen código
        if ticker_upper in ['PESOS', '$', 'DOLARES', 'USD', 'U$S', 'DOLAR', 'DOLAR CABLE']:
            return None
        
        # Primero probar ticker exacto
        data = self._precios_iniciales_cache.get(ticker_upper, {})
        if isinstance(data, dict) and data.get('codigo'):
            return data.get('codigo')
        
        # Si no encuentra, probar variaciones de ticker (OCR 0↔O)
        for ticker_var in self._generate_ticker_variations(ticker_upper):
            data = self._precios_iniciales_cache.get(ticker_var, {})
            if isinstance(data, dict) and data.get('codigo'):
                return data.get('codigo')
        
        return None
    
    def _is_moneda(self, ticker: str) -> bool:
        """Verifica si el ticker corresponde a una moneda (PESOS, DOLARES, DOLAR CABLE)."""
        ticker_upper = str(ticker).upper().strip()
        return ticker_upper in ['PESOS', '$', 'DOLARES', 'USD', 'U$S', 'DOLAR', 'DOLAR CABLE', 'CABLE']
    
    def _vlookup_especies_visual(self, codigo, columna: int):
        """Simula VLOOKUP en EspeciesVisual."""
        codigo_clean = self._clean_codigo(codigo)
        data = self._especies_visual_cache.get(codigo_clean, {})
        
        if columna == 5:  # Moneda emisión
            return data.get('moneda_emision', '')
        elif columna == 15:  # Nombre con moneda
            return data.get('nombre_con_moneda', '')
        elif columna == 16:  # Tipo especie
            return data.get('tipo_especie', '')
        
        return ''
    
    def merge(self) -> Workbook:
        """
        Ejecuta el merge completo y retorna el workbook consolidado.
        """
        wb = Workbook()
        # Eliminar hoja default
        wb.remove(wb.active)
        
        # Crear hojas en orden
        self._create_posicion_inicial(wb)
        self._create_posicion_final(wb)
        self._create_boletos(wb)
        self._create_cauciones(wb)  # Nueva hoja para cauciones
        self._create_rentas_dividendos_gallo(wb)
        self._create_resultado_ventas_ars(wb)
        self._create_resultado_ventas_usd(wb)
        self._create_rentas_dividendos_ars(wb)
        self._create_rentas_dividendos_usd(wb)
        self._create_resumen(wb)
        self._create_posicion_titulos(wb)
        
        # Agregar hojas auxiliares
        self._add_aux_sheets(wb)
        
        return wb
    
    def _create_posicion_inicial(self, wb: Workbook):
        """Crea hoja Posicion Inicial Gallo con las mismas columnas que Posicion Final."""
        ws = wb.create_sheet("Posicion Inicial Gallo")
        
        # Headers (20 columnas) - misma estructura que Posicion Final pero con nombres "Inicial"
        headers = ['tipo_especie', 'Ticker', 'especie', 'Codigo especie',
                   'Codigo Especie Origen', 'comentario especies', 'detalle', 'custodia', 'cantidad',
                   'precio Tenencia Inicial Pesos', 'precio Tenencia Inicial USD', 'Precio de PreciosIniciales',
                   'precio costo(en proceso)', 'Origen precio costo', 'comentarios precio costo',
                   'Precio a Utilizar', 'importe_pesos', 'porc_cartera_pesos', 'importe_dolares', 
                   'porc_cartera_dolares']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Copiar datos de Gallo
        try:
            gallo_ws = self.gallo_wb['Posicion Inicial']
        except KeyError:
            return
        
        row_out = 2
        for row in range(2, gallo_ws.max_row + 1):
            tipo_especie = gallo_ws.cell(row, 1).value
            especie_full = gallo_ws.cell(row, 2).value
            
            if not especie_full:
                continue
            
            ticker, especie = self._split_especie(especie_full)
            
            # Para monedas (PESOS, DOLARES, DOLAR CABLE), especie = ticker (no vacía)
            is_moneda = self._is_moneda(ticker)
            if is_moneda:
                especie = ticker  # Col C muestra PESOS, DOLARES, DOLAR CABLE
            
            # Buscar código de especie usando ticker en PreciosInicialesEspecies
            codigo = None
            codigo_origen = ""
            if not is_moneda:
                codigo = self._get_codigo_from_ticker(ticker)
                if codigo:
                    codigo_origen = "PreciosInicialesEspecies"
                else:
                    # Fallback: buscar en transacciones de Gallo
                    codigo_str, codigo_origen = self._buscar_codigo_especie(especie_full, tipo_especie)
                    if codigo_str:
                        try:
                            codigo = int(codigo_str)
                        except:
                            codigo = codigo_str
            
            # Datos originales
            detalle = gallo_ws.cell(row, 3).value
            custodia = gallo_ws.cell(row, 4).value
            cantidad = gallo_ws.cell(row, 5).value
            precio_orig = gallo_ws.cell(row, 6).value
            importe_pesos = gallo_ws.cell(row, 7).value
            porc_pesos = gallo_ws.cell(row, 8).value
            importe_usd = gallo_ws.cell(row, 9).value
            porc_usd = gallo_ws.cell(row, 10).value
            
            # Determinar si es renta fija dólares (precio dividido x100) o TIT.PRIVADOS EXTERIOR (precio en USD)
            tipo_lower = str(tipo_especie).lower() if tipo_especie else ""
            es_renta_fija_usd = 'renta fija' in tipo_lower and ('dolar' in tipo_lower or 'usd' in tipo_lower)
            es_tit_privados_ext = 'privados' in tipo_lower and 'exterior' in tipo_lower
            
            # Calcular precios tenencia inicial
            precio_pesos = 0
            precio_usd = 0
            if cantidad and float(cantidad) != 0:
                if importe_pesos:
                    try:
                        precio_pesos = float(importe_pesos) / float(cantidad)
                        # Para renta fija dólares, el precio viene dividido x100
                        if es_renta_fija_usd:
                            precio_pesos = precio_pesos * 100
                    except:
                        pass
                if importe_usd:
                    try:
                        precio_usd = float(importe_usd) / float(cantidad)
                    except:
                        pass
            elif precio_orig:
                precio_pesos = precio_orig
            
            # Precio de PreciosInicialesEspecies (via ticker)
            precio_inicial = self._get_precio_inicial(ticker)
            
            # Para TIT.PRIVADOS EXTERIOR, precio viene en USD - convertir a ARS
            # Usamos cotización del dólar cable al 31/12/2024 (inicio del año)
            if es_tit_privados_ext and precio_inicial > 0:
                # Cotización dólar cable al inicio del período (usamos el valor fijo)
                cotizacion_usd = 1148.93  # Dólar Cable 31/12/2024
                precio_inicial = precio_inicial * cotizacion_usd
            
            # Precio a utilizar = precio de PreciosInicialesEspecies
            precio_a_utilizar = precio_inicial
            
            # Escribir fila
            ws.cell(row_out, 1, tipo_especie)
            ws.cell(row_out, 2, ticker)
            ws.cell(row_out, 3, especie)
            ws.cell(row_out, 4, codigo)  # Forzar número
            ws.cell(row_out, 5, codigo_origen)
            ws.cell(row_out, 6, "")  # comentario especies
            ws.cell(row_out, 7, detalle)
            ws.cell(row_out, 8, custodia)
            ws.cell(row_out, 9, cantidad)
            ws.cell(row_out, 10, precio_pesos)
            ws.cell(row_out, 11, precio_usd)
            ws.cell(row_out, 12, precio_inicial)
            ws.cell(row_out, 13, "")  # precio costo (en proceso)
            ws.cell(row_out, 14, "")  # origen precio costo
            ws.cell(row_out, 15, "")  # comentarios precio costo
            ws.cell(row_out, 16, precio_a_utilizar)
            ws.cell(row_out, 17, importe_pesos)
            ws.cell(row_out, 18, porc_pesos)
            ws.cell(row_out, 19, importe_usd)
            ws.cell(row_out, 20, porc_usd)
            
            row_out += 1
    
    def _create_posicion_final(self, wb: Workbook):
        """Crea hoja Posicion Final Gallo con columnas adicionales."""
        ws = wb.create_sheet("Posicion Final Gallo")
        
        # Headers (20 columnas)
        headers = ['tipo_especie', 'Ticker', 'especie', 'Codigo especie',
                   'Codigo Especie Origen', 'comentario especies', 'detalle', 'custodia', 'cantidad',
                   'precio Tenencia Final Pesos', 'precio Tenencia Final USD', 'Precio Tenencia Inicial',
                   'precio costo(en proceso)', 'Origen precio costo', 'comentarios precio costo',
                   'Precio a Utilizar', 'importe_pesos', 'porc_cartera_pesos', 'importe_dolares', 
                   'porc_cartera_dolares']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Copiar datos de Gallo
        try:
            gallo_ws = self.gallo_wb['Posicion Final']
        except KeyError:
            return
        
        row_out = 2
        for row in range(2, gallo_ws.max_row + 1):
            tipo_especie = gallo_ws.cell(row, 1).value
            especie_full = gallo_ws.cell(row, 2).value
            
            if not especie_full:
                continue
            
            ticker, especie = self._split_especie(especie_full)
            
            # Para monedas (PESOS, DOLARES, DOLAR CABLE), especie = ticker
            is_moneda = self._is_moneda(ticker)
            if is_moneda:
                especie = ticker  # Col C muestra PESOS, DOLARES, DOLAR CABLE
            
            # Buscar código de especie usando ticker en PreciosInicialesEspecies
            codigo = None
            codigo_origen = ""
            if not is_moneda:
                codigo = self._get_codigo_from_ticker(ticker)
                if codigo:
                    codigo_origen = "PreciosInicialesEspecies"
                else:
                    # Fallback: buscar en transacciones de Gallo
                    codigo_str, codigo_origen = self._buscar_codigo_especie(especie_full, tipo_especie)
                    if codigo_str:
                        try:
                            codigo = int(codigo_str)
                        except:
                            codigo = codigo_str
            
            # Datos originales
            detalle = gallo_ws.cell(row, 3).value
            custodia = gallo_ws.cell(row, 4).value
            cantidad = gallo_ws.cell(row, 5).value
            importe_pesos = gallo_ws.cell(row, 7).value
            porc_pesos = gallo_ws.cell(row, 8).value
            importe_usd = gallo_ws.cell(row, 9).value
            porc_usd = gallo_ws.cell(row, 10).value
            
            # Determinar si es renta fija dólares (precio dividido x100) o TIT.PRIVADOS EXTERIOR (precio en USD)
            tipo_lower = str(tipo_especie).lower() if tipo_especie else ""
            es_tit_privados_ext = 'privados' in tipo_lower and 'exterior' in tipo_lower
            
            # Calcular precios tenencia final
            precio_pesos = 0
            precio_usd = 0
            if cantidad and float(cantidad) != 0:
                if importe_pesos:
                    try:
                        precio_pesos = float(importe_pesos) / float(cantidad)
                    except:
                        pass
                if importe_usd:
                    try:
                        precio_usd = float(importe_usd) / float(cantidad)
                    except:
                        pass
            
            # Precio tenencia inicial (de PreciosInicialesEspecies)
            precio_inicial = self._get_precio_inicial(ticker)
            
            # Para TIT.PRIVADOS EXTERIOR, precio viene en USD - convertir a ARS
            if es_tit_privados_ext and precio_inicial > 0:
                cotizacion_usd = 1148.93  # Dólar Cable 31/12/2024
                precio_inicial = precio_inicial * cotizacion_usd
            
            # Precio a utilizar = precio tenencia inicial
            precio_a_utilizar = precio_inicial
            
            # Escribir fila
            ws.cell(row_out, 1, tipo_especie)
            ws.cell(row_out, 2, ticker)
            ws.cell(row_out, 3, especie)
            ws.cell(row_out, 4, codigo)  # Forzar número (int)
            ws.cell(row_out, 5, codigo_origen)
            ws.cell(row_out, 6, "")  # comentario especies
            ws.cell(row_out, 7, detalle)
            ws.cell(row_out, 8, custodia)
            ws.cell(row_out, 9, cantidad)
            ws.cell(row_out, 10, precio_pesos)
            ws.cell(row_out, 11, precio_usd)
            ws.cell(row_out, 12, precio_inicial)
            ws.cell(row_out, 13, "")  # precio costo (en proceso)
            ws.cell(row_out, 14, "")  # origen precio costo
            ws.cell(row_out, 15, "")  # comentarios precio costo
            ws.cell(row_out, 16, precio_a_utilizar)
            ws.cell(row_out, 17, importe_pesos)
            ws.cell(row_out, 18, porc_pesos)
            ws.cell(row_out, 19, importe_usd)
            ws.cell(row_out, 20, porc_usd)
            
            row_out += 1
    
    def _create_boletos(self, wb: Workbook):
        """Crea hoja Boletos con transacciones de Gallo y Visual, ordenadas por Cod.Instrum y fecha."""
        ws = wb.create_sheet("Boletos")
        
        # Headers (19 columnas)
        headers = ['Tipo de Instrumento', 'Concertación', 'Liquidación', 'Nro. Boleto',
                   'Moneda', 'Tipo Operación', 'Cod.Instrum', 'Instrumento Crudo',
                   'InstrumentoConMoneda', 'Cantidad', 'Precio', 'Tipo Cambio',
                   'Bruto', 'Interés', 'Gastos', 'Neto Calculado', 'Origen', 
                   'moneda emision', 'Auditoría']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Recolectar todas las transacciones para ordenar
        all_transactions = []
        
        # Procesar hojas de Gallo
        for sheet_name in self.gallo_wb.sheetnames:
            # Saltear hojas de posición y resultados
            if any(skip in sheet_name for skip in ['Posicion', 'Resultado', 'Posición']):
                continue
            
            # SKIP cauciones - van en hoja separada
            if 'caucion' in sheet_name.lower():
                continue
            
            try:
                gallo_ws = self.gallo_wb[sheet_name]
            except:
                continue
            
            for row in range(2, gallo_ws.max_row + 1):
                operacion = gallo_ws.cell(row, 5).value  # Col E
                fecha = gallo_ws.cell(row, 4).value
                numero = gallo_ws.cell(row, 6).value
                
                if not operacion:
                    continue
                
                operacion_lower = str(operacion).lower().strip()
                
                # SKIP operaciones de cauciones (COL CAU TER con instrumento VARIAS)
                especie = gallo_ws.cell(row, 3).value
                if especie and 'varias' in str(especie).lower():
                    continue
                if 'col cau' in operacion_lower:
                    continue
                
                # Solo operaciones de compra/venta para Boletos
                operaciones_validas = ['compra', 'venta', 'cpra', 'canje', 'licitacion']
                if not any(op in operacion_lower for op in operaciones_validas):
                    continue
                
                # Filtrar solo 2025
                if not self._is_year_2025(fecha):
                    continue
                
                # Extraer datos
                cod_especie = gallo_ws.cell(row, 2).value
                cantidad = gallo_ws.cell(row, 7).value
                precio = gallo_ws.cell(row, 8).value
                resultado_pesos = gallo_ws.cell(row, 11).value
                resultado_usd = gallo_ws.cell(row, 12).value
                gastos_pesos = gallo_ws.cell(row, 13).value
                gastos_usd = gallo_ws.cell(row, 14).value
                
                # Determinar moneda PRIMERO basándose en el nombre de la hoja
                # Si la hoja dice "Pesos", es Pesos (ignorar "USD" en operación)
                # Si la hoja dice "Dolares", es Dolar MEP
                # Si la hoja dice "Exterior", es Dolar Cable
                sheet_lower = sheet_name.lower()
                if 'pesos' in sheet_lower:
                    moneda = "Pesos"
                elif 'exterior' in sheet_lower:
                    moneda = "Dolar Cable"
                elif 'dolar' in sheet_lower:
                    moneda = "Dolar MEP"
                else:
                    # Fallback: usar la lógica original
                    moneda = self._get_moneda(resultado_pesos, resultado_usd, gastos_pesos, gastos_usd, sheet_name, operacion)
                
                # Gastos según moneda
                gastos = gastos_pesos if moneda == "Pesos" else gastos_usd
                if gastos is None:
                    gastos = 0
                
                # Código limpio y forzar a número
                cod_clean = self._clean_codigo(cod_especie)
                try:
                    cod_num = int(cod_clean) if cod_clean else None
                except:
                    cod_num = cod_clean
                
                # Convertir fecha a datetime para Excel
                fecha_dt, _ = self._parse_fecha(fecha)
                
                # Auditoría
                auditoria = f"Origen: Gallo-{sheet_name} | Fecha: {fecha} | Cod: {cod_especie} | Op: {operacion}"
                
                # Guardar transacción (sin fórmulas, se generan al escribir)
                all_transactions.append({
                    'cod_instrum': cod_num,  # Forzado a número
                    'fecha': fecha_dt if fecha_dt else fecha,
                    'fecha_raw': fecha,
                    'liquidacion': "",
                    'numero': numero,
                    'moneda': moneda,
                    'operacion': operacion,
                    'especie': especie,
                    'cantidad': cantidad,
                    'precio': precio,
                    'interes': 0,
                    'gastos': gastos,
                    'origen': f"gallo-{sheet_name}",
                    'auditoria': auditoria,
                    'tipo_instrumento_val': None,  # Se usará fórmula
                })
        
        # Agregar transacciones de Visual
        try:
            visual_boletos = self.visual_wb['Boletos']
            for row in range(2, visual_boletos.max_row + 1):
                tipo_instrumento = visual_boletos.cell(row, 1).value
                concertacion = visual_boletos.cell(row, 2).value
                liquidacion = visual_boletos.cell(row, 3).value
                numero = visual_boletos.cell(row, 4).value
                moneda = visual_boletos.cell(row, 5).value
                operacion = visual_boletos.cell(row, 6).value
                cod_instrum = visual_boletos.cell(row, 7).value
                instrumento = visual_boletos.cell(row, 8).value
                cantidad = visual_boletos.cell(row, 9).value
                precio = visual_boletos.cell(row, 10).value
                interes = visual_boletos.cell(row, 13).value
                gastos = visual_boletos.cell(row, 14).value
                
                if not operacion:
                    continue
                
                # Parsear fecha
                fecha_dt, year = self._parse_fecha(concertacion)
                if year != 2025:
                    continue
                
                # Código limpio y forzar a número
                cod_clean = self._clean_codigo(cod_instrum)
                try:
                    cod_num = int(cod_clean) if cod_clean else None
                except:
                    cod_num = cod_clean
                
                auditoria = f"Origen: Visual | Fecha: {concertacion} | Cod: {cod_instrum} | Op: {operacion}"
                
                all_transactions.append({
                    'cod_instrum': cod_num,  # Forzado a número
                    'fecha': fecha_dt if fecha_dt else concertacion,
                    'fecha_raw': concertacion,
                    'liquidacion': liquidacion,
                    'numero': numero,
                    'moneda': moneda,
                    'operacion': operacion,
                    'especie': instrumento,
                    'cantidad': cantidad,
                    'precio': precio,
                    'interes': interes if interes else 0,
                    'gastos': gastos if gastos else 0,
                    'origen': "Visual",
                    'auditoria': auditoria,
                    'tipo_instrumento_val': tipo_instrumento,
                })
        except KeyError:
            pass  # Visual no tiene hoja Boletos
        
        # Ordenar por fecha de concertación
        def sort_key(t):
            fecha = t.get('fecha')
            if isinstance(fecha, datetime):
                return fecha
            else:
                return datetime.min
        
        all_transactions.sort(key=sort_key)
        
        # Escribir transacciones ordenadas
        for row_out, trans in enumerate(all_transactions, start=2):
            # Fórmulas con row_out correcto
            # Tipo de Instrumento: usa VLOOKUP si no viene de Visual
            tipo_instrumento = f'=IFERROR(VLOOKUP(G{row_out},EspeciesVisual!C:R,16,FALSE),"")' if not trans['tipo_instrumento_val'] else trans['tipo_instrumento_val']
            
            # InstrumentoConMoneda
            instrumento_con_moneda = f'=IFERROR(VLOOKUP(G{row_out},EspeciesVisual!C:Q,15,FALSE),"")'
            
            # Tipo Cambio: fórmula simplificada compatible con Excel 2013 español
            # Usa VLOOKUP simple por fecha (asumiendo que Cotización tiene fecha en col A, valor en col B)
            tipo_cambio = f'=IF(E{row_out}="Pesos",1,IFERROR(VLOOKUP(B{row_out},\'Cotizacion Dolar Historica\'!A:B,2,FALSE),0))'
            
            bruto = f'=J{row_out}*K{row_out}'
            neto = f'=IF(J{row_out}>0,J{row_out}*K{row_out}+O{row_out},J{row_out}*K{row_out}-O{row_out})'
            moneda_emision = f'=IFERROR(VLOOKUP(G{row_out},EspeciesVisual!C:Q,5,FALSE),"")'
            
            ws.cell(row_out, 1, tipo_instrumento)
            ws.cell(row_out, 2, trans['fecha'])
            ws.cell(row_out, 3, trans['liquidacion'])
            ws.cell(row_out, 4, trans['numero'])
            ws.cell(row_out, 5, trans['moneda'])
            ws.cell(row_out, 6, trans['operacion'])
            ws.cell(row_out, 7, trans['cod_instrum'])  # Ya es número
            ws.cell(row_out, 8, trans['especie'])
            ws.cell(row_out, 9, instrumento_con_moneda)
            ws.cell(row_out, 10, trans['cantidad'])
            ws.cell(row_out, 11, trans['precio'])
            ws.cell(row_out, 12, tipo_cambio)
            ws.cell(row_out, 13, bruto)
            ws.cell(row_out, 14, trans['interes'])
            ws.cell(row_out, 15, trans['gastos'])
            ws.cell(row_out, 16, neto)
            ws.cell(row_out, 17, trans['origen'])
            ws.cell(row_out, 18, moneda_emision)
            ws.cell(row_out, 19, trans['auditoria'])
    
    def _create_cauciones(self, wb: Workbook):
        """
        Crea hoja Cauciones con operaciones de caución separadas de Boletos.
        Columnas según estructura Visual con match a Gallo:
        - Concertación (Gallo: fecha)
        - Plazo (calculado: diferencia entre vencimiento y fecha)
        - Liquidación (Gallo: vencimiento)
        - Operación (Gallo: operacion)
        - Boleto (Gallo: numero)
        - Contado (Gallo: colocado)
        - Futuro (Gallo: al_vencimiento)
        - Tipo de Cambio (1 si pesos, cotización dólar si dólares)
        - Tasa (%) (no hay en Gallo)
        - Interés Bruto (no hay en Gallo)
        - Interés Devengado (Gallo: interes_pesos o interes_usd)
        - Aranceles (Gallo: gastos_pesos o gastos_usd)
        - Derechos (no hay en Gallo)
        - Costo Financiero (calculado: -(intereses + gastos))
        """
        ws = wb.create_sheet("Cauciones")
        
        # Headers según estructura Visual
        headers = ['Concertación', 'Plazo', 'Liquidación', 'Operación', 'Boleto',
                   'Contado', 'Futuro', 'Tipo de Cambio', 'Tasa (%)', 
                   'Interés Bruto', 'Interés Devengado', 'Aranceles', 'Derechos',
                   'Costo Financiero', 'Moneda', 'Origen', 'Auditoría']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Recolectar todas las cauciones para ordenar
        all_cauciones = []
        
        # Procesar hojas de Cauciones de Gallo
        for sheet_name in self.gallo_wb.sheetnames:
            if 'caucion' not in sheet_name.lower():
                continue
            
            # Determinar moneda del nombre de la hoja
            if 'pesos' in sheet_name.lower():
                moneda = "Pesos"
                tipo_cambio = 1
            elif 'dolar' in sheet_name.lower():
                moneda = "Dolar MEP"
                tipo_cambio = 1167.806  # Cotización dólar al 31/12/2024
            else:
                moneda = "Pesos"
                tipo_cambio = 1
            
            try:
                gallo_ws = self.gallo_wb[sheet_name]
            except:
                continue
            
            # Estructura Cauciones Gallo según OneShotSpec:
            # ["tipo_fila", "cod_especie", "especie", "fecha", "vencimiento", "operacion", 
            #  "numero", "colocado", "al_vencimiento", "interes_pesos", "interes_usd", 
            #  "gastos_pesos", "gastos_usd"]
            for row in range(2, gallo_ws.max_row + 1):
                tipo_fila = gallo_ws.cell(row, 1).value
                fecha = gallo_ws.cell(row, 4).value
                vencimiento = gallo_ws.cell(row, 5).value
                operacion = gallo_ws.cell(row, 6).value
                numero = gallo_ws.cell(row, 7).value
                colocado = gallo_ws.cell(row, 8).value
                al_vencimiento = gallo_ws.cell(row, 9).value
                interes_pesos = gallo_ws.cell(row, 10).value
                interes_usd = gallo_ws.cell(row, 11).value
                gastos_pesos = gallo_ws.cell(row, 12).value
                gastos_usd = gallo_ws.cell(row, 13).value
                
                # Saltear filas de total
                if tipo_fila and 'total' in str(tipo_fila).lower():
                    continue
                
                if not operacion:
                    continue
                
                # Filtrar solo 2025
                if not self._is_year_2025(fecha):
                    continue
                
                # Parsear fechas
                fecha_dt, _ = self._parse_fecha(fecha)
                venc_dt, _ = self._parse_fecha(vencimiento)
                
                # Calcular plazo (diferencia en días)
                plazo = 0
                if fecha_dt and venc_dt:
                    plazo = (venc_dt - fecha_dt).days
                
                # Interés según moneda
                interes = interes_pesos if moneda == "Pesos" else interes_usd
                interes = float(interes) if interes else 0
                
                # Gastos según moneda
                gastos = gastos_pesos if moneda == "Pesos" else gastos_usd
                gastos = float(gastos) if gastos else 0
                
                # Costo financiero = -(intereses + gastos)
                costo_financiero = -(interes + gastos)
                
                auditoria = f"Origen: Gallo-{sheet_name}"
                
                all_cauciones.append({
                    'fecha': fecha_dt if fecha_dt else fecha,
                    'plazo': plazo,
                    'liquidacion': venc_dt if venc_dt else vencimiento,
                    'operacion': operacion,
                    'boleto': numero,
                    'contado': colocado,
                    'futuro': al_vencimiento,
                    'tipo_cambio': tipo_cambio,
                    'tasa': None,  # No disponible en Gallo
                    'interes_bruto': None,  # No disponible en Gallo
                    'interes_devengado': interes,
                    'aranceles': gastos,
                    'derechos': None,  # No disponible en Gallo
                    'costo_financiero': costo_financiero,
                    'moneda': moneda,
                    'origen': f"Gallo-{sheet_name}",
                    'auditoria': auditoria,
                })
        
        # Ordenar por fecha de concertación
        def sort_key(t):
            fecha = t.get('fecha')
            if isinstance(fecha, datetime):
                return fecha
            else:
                return datetime.min
        
        all_cauciones.sort(key=sort_key)
        
        # Escribir cauciones ordenadas
        for row_out, cau in enumerate(all_cauciones, start=2):
            ws.cell(row_out, 1, cau['fecha'])
            ws.cell(row_out, 2, cau['plazo'])
            ws.cell(row_out, 3, cau['liquidacion'])
            ws.cell(row_out, 4, cau['operacion'])
            ws.cell(row_out, 5, cau['boleto'])
            ws.cell(row_out, 6, cau['contado'])
            ws.cell(row_out, 7, cau['futuro'])
            ws.cell(row_out, 8, cau['tipo_cambio'])
            ws.cell(row_out, 9, cau['tasa'])
            ws.cell(row_out, 10, cau['interes_bruto'])
            ws.cell(row_out, 11, cau['interes_devengado'])
            ws.cell(row_out, 12, cau['aranceles'])
            ws.cell(row_out, 13, cau['derechos'])
            ws.cell(row_out, 14, cau['costo_financiero'])
            ws.cell(row_out, 15, cau['moneda'])
            ws.cell(row_out, 16, cau['origen'])
            ws.cell(row_out, 17, cau['auditoria'])
    
    def _create_rentas_dividendos_gallo(self, wb: Workbook):
        """Crea hoja Rentas y Dividendos Gallo, ordenada por fecha de concertación."""
        ws = wb.create_sheet("Rentas y Dividendos Gallo")
        
        # Headers (20 columnas)
        headers = ['Tipo de Instrumento', 'Concertación', 'Liquidación', 'Nro. Boleto',
                   'Moneda', 'Tipo Operación', 'Cod.Instrum', 'Instrumento Crudo',
                   'InstrumentoConMoneda', 'Cantidad', 'Precio', 'Tipo Cambio',
                   'Bruto', 'Interés', 'Gastos', 'Costo', 'Neto Calculado', 
                   'Origen', 'moneda emision', 'Auditoría']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Recolectar todas las transacciones para ordenar
        all_rentas = []
        
        # Procesar hojas de Gallo
        for sheet_name in self.gallo_wb.sheetnames:
            # Saltear hojas de posición y resultados
            if any(skip in sheet_name for skip in ['Posicion', 'Resultado', 'Posición']):
                continue
            
            try:
                gallo_ws = self.gallo_wb[sheet_name]
            except:
                continue
            
            for row in range(2, gallo_ws.max_row + 1):
                operacion = gallo_ws.cell(row, 5).value  # Col E
                if not operacion:
                    continue
                
                operacion_lower = str(operacion).lower().strip()
                
                # Solo operaciones de rentas/dividendos/amortización
                if not any(op in operacion_lower for op in self.OPERACIONES_RENTAS):
                    continue
                
                # Extraer datos
                cod_especie = gallo_ws.cell(row, 2).value
                especie = gallo_ws.cell(row, 3).value
                fecha = gallo_ws.cell(row, 4).value
                numero = gallo_ws.cell(row, 6).value
                cantidad = gallo_ws.cell(row, 7).value
                precio = gallo_ws.cell(row, 8).value
                importe = gallo_ws.cell(row, 9).value
                costo = gallo_ws.cell(row, 10).value
                resultado_pesos = gallo_ws.cell(row, 11).value
                resultado_usd = gallo_ws.cell(row, 12).value
                gastos_pesos = gallo_ws.cell(row, 13).value
                gastos_usd = gallo_ws.cell(row, 14).value
                
                # Filtrar solo 2025
                if not self._is_year_2025(fecha):
                    continue
                
                # Determinar moneda basándose en nombre de hoja
                sheet_lower = sheet_name.lower()
                if 'pesos' in sheet_lower:
                    moneda = "Pesos"
                elif 'exterior' in sheet_lower:
                    moneda = "Dolar Cable"
                elif 'dolar' in sheet_lower:
                    moneda = "Dolar MEP"
                else:
                    moneda = self._get_moneda(resultado_pesos, resultado_usd, gastos_pesos, gastos_usd, sheet_name, operacion)
                
                # Gastos según moneda
                gastos = gastos_pesos if moneda == "Pesos" else gastos_usd
                if gastos is None:
                    gastos = 0
                
                # Ajustar precio para amortizaciones (100 -> 1)
                if 'amortizacion' in operacion_lower:
                    if precio and float(precio) == 100:
                        precio = 1
                
                # Código limpio y forzar a número
                cod_clean = self._clean_codigo(cod_especie)
                try:
                    cod_num = int(cod_clean) if cod_clean else None
                except:
                    cod_num = cod_clean
                
                # Bruto
                bruto = importe if importe else (cantidad * precio if cantidad and precio else 0)
                
                # Parsear fecha
                fecha_dt, _ = self._parse_fecha(fecha)
                
                auditoria = f"Origen: Gallo-{sheet_name} | Operación: {operacion}"
                
                all_rentas.append({
                    'tipo_instrumento_val': None,  # Usará fórmula
                    'fecha': fecha_dt if fecha_dt else fecha,
                    'liquidacion': "",
                    'numero': numero,
                    'moneda': moneda,
                    'operacion': operacion.upper(),
                    'cod_num': cod_num,
                    'especie': especie,
                    'cantidad': cantidad if cantidad else 0,
                    'precio': precio if precio else 0,
                    'bruto': bruto,
                    'interes': 0,
                    'gastos': gastos,
                    'costo': costo if costo else 0,
                    'origen': f"Gallo-{sheet_name}",
                    'is_amortizacion': 'amortizacion' in operacion_lower,
                    'auditoria': auditoria,
                })
        
        # Agregar Rentas/Dividendos de Visual
        visual_sheets = [('Rentas Dividendos ARS', 'Pesos'), ('Rentas Dividendos USD', 'Dolar')]
        for visual_sheet_name, moneda_default in visual_sheets:
            try:
                visual_ws = self.visual_wb[visual_sheet_name]
            except KeyError:
                continue
            
            for row in range(2, visual_ws.max_row + 1):
                instrumento = visual_ws.cell(row, 1).value
                cod_instrum = visual_ws.cell(row, 2).value
                categoria = visual_ws.cell(row, 3).value
                tipo_instrum = visual_ws.cell(row, 4).value
                concertacion = visual_ws.cell(row, 5).value
                liquidacion = visual_ws.cell(row, 6).value
                nro_ndc = visual_ws.cell(row, 7).value
                tipo_operacion = visual_ws.cell(row, 8).value
                cantidad = visual_ws.cell(row, 9).value
                moneda = visual_ws.cell(row, 10).value
                tipo_cambio_val = visual_ws.cell(row, 11).value
                gastos = visual_ws.cell(row, 12).value
                importe = visual_ws.cell(row, 13).value
                
                if not tipo_operacion:
                    continue
                
                # Parsear fecha y filtrar 2025
                fecha_dt, year = self._parse_fecha(concertacion)
                if year != 2025:
                    continue
                
                # Código limpio y convertir a número
                cod_clean = self._clean_codigo(cod_instrum)
                try:
                    cod_num = int(cod_clean) if cod_clean else None
                except (ValueError, TypeError):
                    cod_num = cod_clean
                
                # Determinar moneda correcta
                if moneda:
                    if 'peso' in str(moneda).lower():
                        moneda_final = 'Pesos'
                    elif 'mep' in str(moneda).lower():
                        moneda_final = 'Dolar MEP'
                    elif 'cable' in str(moneda).lower():
                        moneda_final = 'Dolar Cable'
                    else:
                        moneda_final = moneda_default
                else:
                    moneda_final = moneda_default
                
                bruto = importe if importe else 0
                auditoria = f"Origen: Visual-{visual_sheet_name} | Cat: {categoria} | Op: {tipo_operacion}"
                
                all_rentas.append({
                    'tipo_instrumento_val': tipo_instrum,
                    'fecha': fecha_dt if fecha_dt else concertacion,
                    'liquidacion': liquidacion,
                    'numero': nro_ndc,
                    'moneda': moneda_final,
                    'operacion': str(tipo_operacion).upper() if tipo_operacion else "",
                    'cod_num': cod_num,
                    'especie': instrumento,
                    'cantidad': cantidad if cantidad else 0,
                    'precio': 1,  # Precio = 1 para rentas/dividendos de Visual
                    'bruto': bruto,
                    'interes': 0,
                    'gastos': gastos if gastos else 0,
                    'costo': 0,
                    'origen': f"Visual-{visual_sheet_name}",
                    'is_amortizacion': False,
                    'auditoria': auditoria,
                })
        
        # Ordenar por fecha de concertación
        def sort_key(t):
            fecha = t.get('fecha')
            if isinstance(fecha, datetime):
                return fecha
            else:
                return datetime.min
        
        all_rentas.sort(key=sort_key)
        
        # Escribir transacciones ordenadas con fórmulas correctas
        for row_out, renta in enumerate(all_rentas, start=2):
            # Fórmulas con row_out correcto
            tipo_instrumento = renta['tipo_instrumento_val'] if renta['tipo_instrumento_val'] else f'=IFERROR(VLOOKUP(G{row_out},EspeciesVisual!C:R,16,FALSE),"")'
            instrumento_con_moneda = f'=IFERROR(VLOOKUP(G{row_out},EspeciesVisual!C:Q,15,FALSE),"")'
            tipo_cambio = f'=IF(E{row_out}="Pesos",1,IFERROR(VLOOKUP(B{row_out},\'Cotizacion Dolar Historica\'!A:B,2,FALSE),0))'
            moneda_emision = f'=IFERROR(VLOOKUP(G{row_out},EspeciesVisual!C:Q,5,FALSE),"")'
            
            # Neto calculado
            if renta['is_amortizacion']:
                neto = f'=-M{row_out}-P{row_out}'
            else:
                neto = f'=M{row_out}-O{row_out}'
            
            ws.cell(row_out, 1, tipo_instrumento)
            ws.cell(row_out, 2, renta['fecha'])
            ws.cell(row_out, 3, renta['liquidacion'])
            ws.cell(row_out, 4, renta['numero'])
            ws.cell(row_out, 5, renta['moneda'])
            ws.cell(row_out, 6, renta['operacion'])
            ws.cell(row_out, 7, renta['cod_num'])  # Forzado a número
            ws.cell(row_out, 8, renta['especie'])
            ws.cell(row_out, 9, instrumento_con_moneda)
            ws.cell(row_out, 10, renta['cantidad'])
            ws.cell(row_out, 11, renta['precio'])
            ws.cell(row_out, 12, tipo_cambio)
            ws.cell(row_out, 13, renta['bruto'])
            ws.cell(row_out, 14, renta['interes'])
            ws.cell(row_out, 15, renta['gastos'])
            ws.cell(row_out, 16, renta['costo'])
            ws.cell(row_out, 17, neto)
            ws.cell(row_out, 18, renta['origen'])
            ws.cell(row_out, 19, moneda_emision)
            ws.cell(row_out, 20, renta['auditoria'])
    
    def _create_resultado_ventas_ars(self, wb: Workbook):
        """Crea hoja Resultado Ventas ARS con transacciones de Boletos filtradas por Pesos."""
        ws = wb.create_sheet("Resultado Ventas ARS")
        
        # Headers (25 columnas - agregamos columna explicativa)
        headers = ['Origen', 'Tipo de Instrumento', 'Instrumento', 'Cod.Instrum',
                   'Concertación', 'Liquidación', 'Moneda', 'Tipo Operación',
                   'Cantidad', 'Precio', 'Bruto', 'Interés', 'Tipo de Cambio',
                   'Gastos', 'IVA', 'Resultado', 'Cantidad Stock Inicial',
                   'Precio Stock Inicial', 'Costo por venta(gallo)', 'Neto Calculado(visual)',
                   'Resultado Calculado(final)', 'Cantidad de Stock Final', 
                   'Precio Stock Final', 'Explicación Cálculo', 'chequeado']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Recolectar transacciones de Boletos con moneda = Pesos
        boletos_ws = wb['Boletos']
        transactions = []
        
        for boletos_row in range(2, boletos_ws.max_row + 1):
            moneda_emision = boletos_ws.cell(boletos_row, 18).value  # Col R = moneda_emision
            
            # Filtrar solo Pesos
            if moneda_emision != "Pesos":
                continue
            
            # Extraer valores
            origen = boletos_ws.cell(boletos_row, 17).value  # Col Q
            tipo_instrumento = boletos_ws.cell(boletos_row, 1).value
            instrumento = boletos_ws.cell(boletos_row, 9).value  # Col I
            cod_instrum = boletos_ws.cell(boletos_row, 7).value  # Col G
            concertacion = boletos_ws.cell(boletos_row, 2).value  # Col B
            liquidacion = boletos_ws.cell(boletos_row, 3).value  # Col C
            moneda = boletos_ws.cell(boletos_row, 5).value  # Col E
            tipo_operacion = boletos_ws.cell(boletos_row, 6).value  # Col F
            cantidad = boletos_ws.cell(boletos_row, 10).value  # Col J
            precio = boletos_ws.cell(boletos_row, 11).value  # Col K
            bruto_formula = boletos_ws.cell(boletos_row, 13).value  # Col M (puede ser fórmula)
            interes = boletos_ws.cell(boletos_row, 14).value  # Col N
            tipo_cambio_formula = boletos_ws.cell(boletos_row, 12).value  # Col L
            gastos = boletos_ws.cell(boletos_row, 15).value  # Col O
            
            transactions.append({
                'origen': origen,
                'tipo_instrumento': tipo_instrumento,
                'instrumento': instrumento,
                'cod_instrum': cod_instrum,
                'concertacion': concertacion,
                'liquidacion': liquidacion,
                'moneda': moneda,
                'tipo_operacion': tipo_operacion,
                'cantidad': cantidad,
                'precio': precio,
                'bruto_formula': bruto_formula,
                'interes': interes,
                'tipo_cambio_formula': tipo_cambio_formula,
                'gastos': gastos,
            })
        
        # Escribir transacciones
        for row_out, trans in enumerate(transactions, start=2):
            ws.cell(row_out, 1, trans['origen'])
            ws.cell(row_out, 2, trans['tipo_instrumento'])
            ws.cell(row_out, 3, trans['instrumento'])
            ws.cell(row_out, 4, trans['cod_instrum'])
            ws.cell(row_out, 5, trans['concertacion'])  # Fecha como datetime
            ws.cell(row_out, 6, trans['liquidacion'])
            ws.cell(row_out, 7, trans['moneda'])
            ws.cell(row_out, 8, trans['tipo_operacion'])
            ws.cell(row_out, 9, trans['cantidad'])
            ws.cell(row_out, 10, trans['precio'])
            ws.cell(row_out, 11, trans['bruto_formula'])
            ws.cell(row_out, 12, trans['interes'])
            ws.cell(row_out, 13, trans['tipo_cambio_formula'])
            ws.cell(row_out, 14, trans['gastos'])
            
            # IVA
            ws.cell(row_out, 15, f'=IF(N{row_out}>0,N{row_out}*0.1736,N{row_out}*-0.1736)')
            
            # Resultado (vacío)
            ws.cell(row_out, 16, "")
            
            # Running Stock Logic:
            if row_out == 2:
                # Primera fila: VLOOKUP a Posicion Inicial si Gallo, Posicion Final si Visual
                ws.cell(row_out, 17, f'=IF(LEFT(A{row_out},5)="Gallo",IFERROR(VLOOKUP(D{row_out},\'Posicion Inicial Gallo\'!D:I,6,FALSE),0),IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:I,6,FALSE),0))')
                ws.cell(row_out, 18, f'=IF(LEFT(A{row_out},5)="Gallo",IFERROR(VLOOKUP(D{row_out},\'Posicion Inicial Gallo\'!D:P,13,FALSE),0),IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:P,13,FALSE),0))')
            else:
                prev = row_out - 1
                ws.cell(row_out, 17, f'=IF(D{row_out}=D{prev},V{prev},IF(LEFT(A{row_out},5)="Gallo",IFERROR(VLOOKUP(D{row_out},\'Posicion Inicial Gallo\'!D:I,6,FALSE),0),IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:I,6,FALSE),0)))')
                ws.cell(row_out, 18, f'=IF(D{row_out}=D{prev},W{prev},IF(LEFT(A{row_out},5)="Gallo",IFERROR(VLOOKUP(D{row_out},\'Posicion Inicial Gallo\'!D:P,13,FALSE),0),IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:P,13,FALSE),0)))')
            
            # Costo por venta
            ws.cell(row_out, 19, f'=IFERROR(IF(I{row_out}<0,I{row_out}*R{row_out},0),"")')
            
            # Neto Calculado (para todas las ops: K+N)
            ws.cell(row_out, 20, f'=K{row_out}+N{row_out}')
            
            # Resultado Calculado
            ws.cell(row_out, 21, f'=IF(S{row_out}<>0,ABS(T{row_out})-ABS(S{row_out}),0)')
            
            # Cantidad Stock Final (running)
            ws.cell(row_out, 22, f'=I{row_out}+Q{row_out}')
            
            # Precio Stock Final (promedio ponderado)
            ws.cell(row_out, 23, f'=IF(V{row_out}=0,0,IF(I{row_out}>0,(I{row_out}*J{row_out}+Q{row_out}*R{row_out})/(I{row_out}+Q{row_out}),R{row_out}))')
            
            # Explicación Cálculo (columna nueva)
            ws.cell(row_out, 24, f'Q=Stock previo o VLOOKUP(D{row_out}→PosIni si Gallo/PosFin si Visual) | R=Precio previo o VLOOKUP col P | S=I{row_out}*R{row_out} si venta | T=K{row_out}+N{row_out} | U=|T|-|S| | V=I{row_out}+Q{row_out} | W=Promedio ponderado')
            
            # Chequeado/Auditoría
            if row_out == 2:
                auditoria = f"Primera op: Stock y Precio desde {'Posicion Inicial' if trans['origen'] and 'Gallo' in trans['origen'] else 'Posicion Final'}"
            else:
                auditoria = f"Running stock: Si mismo código usa stock/precio anterior (V{row_out-1},W{row_out-1})"
            ws.cell(row_out, 25, auditoria)
    
    def _create_resultado_ventas_usd(self, wb: Workbook):
        """Crea hoja Resultado Ventas USD con transacciones de Boletos filtradas por Dolar."""
        ws = wb.create_sheet("Resultado Ventas USD")
        
        # Headers (27 columnas)
        headers = ['Origen', 'Tipo de Instrumento', 'Instrumento', 'Cod.Instrum',
                   'Concertación', 'Liquidación', 'Moneda', 'Tipo Operación',
                   'Cantidad', 'Precio', 'Precio Standarizado', 'Precio Standarizado en USD',
                   'Bruto en USD', 'Interés', 'Tipo de Cambio', 'Valor USD Dia',
                   'Gastos', 'IVA', 'Resultado', 'Cantidad Stock Inicial',
                   'Precio Stock USD', 'Costo por venta(gallo)', 'Neto Calculado(visual)',
                   'Resultado Calculado(final)', 'Cantidad de Stock Final',
                   'Precio Stock Final', 'Comentarios']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Recolectar transacciones de Boletos con moneda contiene "Dolar"
        boletos_ws = wb['Boletos']
        transactions = []
        
        for boletos_row in range(2, boletos_ws.max_row + 1):
            moneda_emision = boletos_ws.cell(boletos_row, 18).value  # Col R = moneda_emision
            
            # Filtrar solo Dolar (MEP, Cable)
            if not moneda_emision or 'dolar' not in str(moneda_emision).lower():
                continue
            
            # Extraer valores
            origen = boletos_ws.cell(boletos_row, 17).value  # Col Q
            tipo_instrumento = boletos_ws.cell(boletos_row, 1).value
            instrumento = boletos_ws.cell(boletos_row, 9).value  # Col I
            cod_instrum = boletos_ws.cell(boletos_row, 7).value  # Col G
            concertacion = boletos_ws.cell(boletos_row, 2).value  # Col B
            liquidacion = boletos_ws.cell(boletos_row, 3).value  # Col C
            moneda = boletos_ws.cell(boletos_row, 5).value  # Col E
            tipo_operacion = boletos_ws.cell(boletos_row, 6).value  # Col F
            cantidad = boletos_ws.cell(boletos_row, 10).value  # Col J
            precio = boletos_ws.cell(boletos_row, 11).value  # Col K
            interes = boletos_ws.cell(boletos_row, 14).value  # Col N
            tipo_cambio_val = boletos_ws.cell(boletos_row, 12).value  # Col L (puede ser fórmula)
            gastos = boletos_ws.cell(boletos_row, 15).value  # Col O
            
            transactions.append({
                'origen': origen,
                'tipo_instrumento': tipo_instrumento,
                'instrumento': instrumento,
                'cod_instrum': cod_instrum,
                'concertacion': concertacion,
                'liquidacion': liquidacion,
                'moneda': moneda,
                'tipo_operacion': tipo_operacion,
                'cantidad': cantidad,
                'precio': precio,
                'interes': interes,
                'tipo_cambio_val': tipo_cambio_val,
                'gastos': gastos,
            })
        
        # Escribir transacciones
        for row_out, trans in enumerate(transactions, start=2):
            ws.cell(row_out, 1, trans['origen'])
            ws.cell(row_out, 2, trans['tipo_instrumento'])
            ws.cell(row_out, 3, trans['instrumento'])
            ws.cell(row_out, 4, trans['cod_instrum'])
            ws.cell(row_out, 5, trans['concertacion'])  # Fecha como datetime
            ws.cell(row_out, 6, trans['liquidacion'])
            ws.cell(row_out, 7, trans['moneda'])
            ws.cell(row_out, 8, trans['tipo_operacion'])
            ws.cell(row_out, 9, trans['cantidad'])
            ws.cell(row_out, 10, trans['precio'])
            
            # Precio estandarizado: Si viene de Visual, multiplicar x100
            # Si viene de Gallo, dejar como está
            is_visual = trans['origen'] and 'visual' in str(trans['origen']).lower()
            if is_visual:
                ws.cell(row_out, 11, f'=J{row_out}*100')
            else:
                ws.cell(row_out, 11, f'=J{row_out}')
            
            # Precio en USD
            ws.cell(row_out, 12, f'=IF(G{row_out}="Pesos",K{row_out}/O{row_out},K{row_out})')
            
            # Bruto en USD
            ws.cell(row_out, 13, f'=I{row_out}*L{row_out}')
            
            ws.cell(row_out, 14, trans['interes'] if trans['interes'] else 0)
            
            # Tipo cambio
            ws.cell(row_out, 15, trans['tipo_cambio_val'])
            
            # Valor USD Dia: VLOOKUP con fecha en Cotizacion Dolar Historica
            ws.cell(row_out, 16, f'=IFERROR(VLOOKUP(E{row_out},\'Cotizacion Dolar Historica\'!A:B,2,FALSE),0)')
            
            ws.cell(row_out, 17, trans['gastos'] if trans['gastos'] else 0)
            ws.cell(row_out, 18, 0)  # IVA
            ws.cell(row_out, 19, "")  # Resultado
            
            # Running Stock Logic para USD:
            if row_out == 2:
                # Primera fila: VLOOKUP a Posicion Inicial si Gallo, Posicion Final si Visual
                ws.cell(row_out, 20, f'=IF(LEFT(A{row_out},5)="Gallo",IFERROR(VLOOKUP(D{row_out},\'Posicion Inicial Gallo\'!D:I,6,FALSE),0),IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:I,6,FALSE),0))')
                # Precio Stock USD: Si P{row_out}=0, evitar división
                ws.cell(row_out, 21, f'=IF(P{row_out}=0,0,IF(LEFT(A{row_out},5)="Gallo",IFERROR(VLOOKUP(D{row_out},\'Posicion Inicial Gallo\'!D:P,13,FALSE),0),IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:P,13,FALSE),0))/P{row_out})')
            else:
                prev = row_out - 1
                ws.cell(row_out, 20, f'=IF(D{row_out}=D{prev},Y{prev},IF(LEFT(A{row_out},5)="Gallo",IFERROR(VLOOKUP(D{row_out},\'Posicion Inicial Gallo\'!D:I,6,FALSE),0),IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:I,6,FALSE),0)))')
                ws.cell(row_out, 21, f'=IF(D{row_out}=D{prev},Z{prev},IF(P{row_out}=0,0,IF(LEFT(A{row_out},5)="Gallo",IFERROR(VLOOKUP(D{row_out},\'Posicion Inicial Gallo\'!D:P,13,FALSE),0),IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:P,13,FALSE),0))/P{row_out}))')
            
            # Costo por venta
            ws.cell(row_out, 22, f'=IFERROR(IF(I{row_out}<0,I{row_out}*U{row_out},0),"")')
            
            # Neto Calculado
            ws.cell(row_out, 23, f'=M{row_out}-Q{row_out}')
            
            # Resultado Calculado
            ws.cell(row_out, 24, f'=IFERROR(IF(V{row_out}<>0,ABS(W{row_out})-ABS(V{row_out}),0),0)')
            
            # Cantidad Stock Final
            ws.cell(row_out, 25, f'=I{row_out}+T{row_out}')
            
            # Precio Stock Final: Evitar DIV/0!
            ws.cell(row_out, 26, f'=IF(Y{row_out}=0,0,IF(I{row_out}>0,IF((I{row_out}+T{row_out})=0,0,(I{row_out}*L{row_out}+T{row_out}*U{row_out})/(I{row_out}+T{row_out})),U{row_out}))')
            
            # Comentarios/Auditoría USD
            if row_out == 2:
                auditoria = f"Primera op USD: Stock desde {'Posicion Inicial' if trans['origen'] and 'Gallo' in trans['origen'] else 'Posicion Final'}. K=Precio×100 si Visual. P=VLOOKUP(E→Cotiz Dolar Hist)"
            else:
                auditoria = f"Running: Si mismo código usa Y{row_out-1},Z{row_out-1}. U,Z con IF(P=0,0,...) para evitar DIV/0!"
            ws.cell(row_out, 27, auditoria)
    
    def _create_rentas_dividendos_ars(self, wb: Workbook):
        """Crea hoja Rentas Dividendos ARS."""
        ws = wb.create_sheet("Rentas Dividendos ARS")
        
        headers = ['Instrumento', 'Cod.Instrum', 'Categoría', 'tipo_instrumento',
                   'Concertación', 'Liquidación', 'Nro. NDC', 'Tipo Operación',
                   'Cantidad', 'Moneda', 'Tipo de Cambio', 'Gastos', 'Importe', 'Origen']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Filtrar de Rentas y Dividendos Gallo por moneda = Pesos
        rentas_ws = wb['Rentas y Dividendos Gallo']
        
        row_out = 2
        for rentas_row in range(2, rentas_ws.max_row + 1):
            # Referencia con filtro por moneda emisión = Pesos
            ws.cell(row_out, 1, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!I{rentas_row},"")')
            ws.cell(row_out, 2, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!G{rentas_row},"")')
            
            # Categoría (Rentas/Dividendos/AMORTIZACION basado en tipo operación)
            ws.cell(row_out, 3, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",IF(OR(\'Rentas y Dividendos Gallo\'!F{rentas_row}="RENTA",\'Rentas y Dividendos Gallo\'!F{rentas_row}="AMORTIZACION"),"Rentas","Dividendos"),"")')
            
            ws.cell(row_out, 4, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!A{rentas_row},"")')
            ws.cell(row_out, 5, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!B{rentas_row},"")')
            ws.cell(row_out, 6, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!C{rentas_row},"")')
            ws.cell(row_out, 7, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!D{rentas_row},"")')
            ws.cell(row_out, 8, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!F{rentas_row},"")')
            ws.cell(row_out, 9, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!J{rentas_row},"")')
            ws.cell(row_out, 10, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!E{rentas_row},"")')
            ws.cell(row_out, 11, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!L{rentas_row},"")')
            ws.cell(row_out, 12, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!O{rentas_row},"")')
            ws.cell(row_out, 13, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!Q{rentas_row},"")')  # Neto calculado
            ws.cell(row_out, 14, f'=IF(\'Rentas y Dividendos Gallo\'!S{rentas_row}="Pesos",\'Rentas y Dividendos Gallo\'!R{rentas_row},"")')
            
            row_out += 1
    
    def _create_rentas_dividendos_usd(self, wb: Workbook):
        """Crea hoja Rentas Dividendos USD."""
        ws = wb.create_sheet("Rentas Dividendos USD")
        
        headers = ['Instrumento', 'Cod.Instrum', 'Categoría', 'tipo_instrumento',
                   'Concertación', 'Liquidación', 'Nro. NDC', 'Tipo Operación',
                   'Cantidad', 'Moneda', 'Tipo de Cambio', 'Gastos', 'Importe', 'Origen']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Filtrar de Rentas y Dividendos Gallo por moneda contiene Dolar
        rentas_ws = wb['Rentas y Dividendos Gallo']
        
        row_out = 2
        for rentas_row in range(2, rentas_ws.max_row + 1):
            # Usar SEARCH para detectar "Dolar" en moneda emisión
            ws.cell(row_out, 1, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!I{rentas_row},""),"")')
            ws.cell(row_out, 2, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!G{rentas_row},""),"")')
            ws.cell(row_out, 3, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),IF(OR(\'Rentas y Dividendos Gallo\'!F{rentas_row}="RENTA",\'Rentas y Dividendos Gallo\'!F{rentas_row}="AMORTIZACION"),"Rentas","Dividendos"),""),"")')
            ws.cell(row_out, 4, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!A{rentas_row},""),"")')
            ws.cell(row_out, 5, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!B{rentas_row},""),"")')
            ws.cell(row_out, 6, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!C{rentas_row},""),"")')
            ws.cell(row_out, 7, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!D{rentas_row},""),"")')
            ws.cell(row_out, 8, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!F{rentas_row},""),"")')
            ws.cell(row_out, 9, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!J{rentas_row},""),"")')
            ws.cell(row_out, 10, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!E{rentas_row},""),"")')
            ws.cell(row_out, 11, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!L{rentas_row},""),"")')
            ws.cell(row_out, 12, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!O{rentas_row},""),"")')
            ws.cell(row_out, 13, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!Q{rentas_row},""),"")')
            ws.cell(row_out, 14, f'=IFERROR(IF(SEARCH("Dolar",\'Rentas y Dividendos Gallo\'!S{rentas_row}),\'Rentas y Dividendos Gallo\'!R{rentas_row},""),"")')
            
            row_out += 1
    
    def _create_resumen(self, wb: Workbook):
        """Crea hoja Resumen con totales."""
        ws = wb.create_sheet("Resumen")
        
        headers = ['Moneda', 'Ventas', 'FCI', 'Opciones', 'Rentas', 'Dividendos',
                   'Ef. CPD', 'Pagarés', 'Futuros', 'Cau (int)', 'Cau (CF)', 'Total']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Fila ARS
        ws.cell(2, 1, "ARS")
        ws.cell(2, 2, "=SUM('Resultado Ventas ARS'!U:U)")  # Ventas
        ws.cell(2, 3, 0)  # FCI
        ws.cell(2, 4, 0)  # Opciones
        ws.cell(2, 5, "=SUMIF('Rentas Dividendos ARS'!C:C,\"Rentas\",'Rentas Dividendos ARS'!M:M)+SUMIF('Rentas Dividendos ARS'!C:C,\"AMORTIZACION\",'Rentas Dividendos ARS'!M:M)")
        ws.cell(2, 6, "=SUMIF('Rentas Dividendos ARS'!C:C,\"Dividendos\",'Rentas Dividendos ARS'!M:M)")
        ws.cell(2, 7, 0)  # Ef. CPD
        ws.cell(2, 8, 0)  # Pagarés
        ws.cell(2, 9, 0)  # Futuros
        ws.cell(2, 10, 0)  # Cau (int)
        ws.cell(2, 11, 0)  # Cau (CF)
        ws.cell(2, 12, "=SUM(B2:K2)")  # Total
        
        # Fila USD
        ws.cell(3, 1, "USD")
        ws.cell(3, 2, "=SUM('Resultado Ventas USD'!X:X)")  # Ventas
        ws.cell(3, 3, 0)
        ws.cell(3, 4, 0)
        ws.cell(3, 5, "=SUMIF('Rentas Dividendos USD'!C:C,\"Rentas\",'Rentas Dividendos USD'!M:M)+SUMIF('Rentas Dividendos USD'!C:C,\"AMORTIZACION\",'Rentas Dividendos USD'!M:M)")
        ws.cell(3, 6, "=SUMIF('Rentas Dividendos USD'!C:C,\"Dividendos\",'Rentas Dividendos USD'!M:M)")
        ws.cell(3, 7, 0)
        ws.cell(3, 8, 0)
        ws.cell(3, 9, 0)
        ws.cell(3, 10, 0)
        ws.cell(3, 11, 0)
        ws.cell(3, 12, "=SUM(B3:K3)")
    
    def _create_posicion_titulos(self, wb: Workbook):
        """Crea hoja Posicion Titulos con resumen simplificado de posiciones finales."""
        ws = wb.create_sheet("Posicion Titulos")
        
        headers = ['Instrumento', 'Código', 'Ticker', 'Cantidad', 'Importe', 'Moneda']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Obtener datos de Posicion Final Gallo
        try:
            pos_final = wb['Posicion Final Gallo']
        except KeyError:
            return
        
        row_out = 2
        for row in range(2, pos_final.max_row + 1):
            especie = pos_final.cell(row, 3).value  # especie (col 3)
            codigo = pos_final.cell(row, 4).value   # Codigo especie (col 4)
            ticker = pos_final.cell(row, 2).value   # Ticker (col 2)
            cantidad = pos_final.cell(row, 9).value  # cantidad (col 9)
            importe = pos_final.cell(row, 17).value  # importe_pesos (col 17)
            
            if not especie:
                continue
            
            # Determinar moneda basado en tipo_especie o nombre
            tipo_especie = pos_final.cell(row, 1).value or ''
            moneda = 'Pesos'  # Default
            if 'dolar' in str(especie).lower() or 'usd' in str(especie).lower():
                moneda = 'Dolar'
            
            ws.cell(row_out, 1, especie)
            ws.cell(row_out, 2, codigo)
            ws.cell(row_out, 3, ticker)
            ws.cell(row_out, 4, cantidad)
            ws.cell(row_out, 5, importe)
            ws.cell(row_out, 6, moneda)
            
            row_out += 1
    
    def _add_aux_sheets(self, wb: Workbook):
        """Agrega hojas auxiliares al workbook."""
        aux_files = {
            'EspeciesVisual': self.especies_visual,
            'EspeciesGallo': self.especies_gallo,
            'Cotizacion Dolar Historica': self.cotizacion_dolar,
            'PreciosInicialesEspecies': self.precios_iniciales
        }
        
        for sheet_name, aux_wb in aux_files.items():
            ws_src = aux_wb.active
            ws_dst = wb.create_sheet(sheet_name)
            
            for row in ws_src.iter_rows():
                for cell in row:
                    ws_dst.cell(row=cell.row, column=cell.column, value=cell.value)


def merge_gallo_visual(gallo_path: str, visual_path: str, output_path: str = None) -> str:
    """
    Función principal para ejecutar el merge.
    
    Args:
        gallo_path: Ruta al Excel de Gallo
        visual_path: Ruta al Excel de Visual
        output_path: Ruta de salida (opcional, genera nombre automático)
    
    Returns:
        Ruta del archivo generado
    """
    merger = GalloVisualMerger(gallo_path, visual_path)
    wb = merger.merge()
    
    if output_path is None:
        # Generar nombre basado en el archivo de entrada
        gallo_name = Path(gallo_path).stem.replace('_Gallo_Generado_OK', '')
        output_path = f"{gallo_name}_Merge_Consolidado.xlsx"
    
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    # Test con archivos de ejemplo
    import sys
    
    if len(sys.argv) >= 3:
        gallo = sys.argv[1]
        visual = sys.argv[2]
        output = sys.argv[3] if len(sys.argv) > 3 else None
        
        result = merge_gallo_visual(gallo, visual, output)
        print(f"Merge completado: {result}")
    else:
        print("Uso: python merge_gallo_visual.py <gallo.xlsx> <visual.xlsx> [output.xlsx]")
