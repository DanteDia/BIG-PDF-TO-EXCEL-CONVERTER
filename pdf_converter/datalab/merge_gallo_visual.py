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
        
        # Cache Precios Iniciales: ticker -> precio
        ws = self.precios_iniciales.active
        for row in range(2, ws.max_row + 1):
            ticker = ws.cell(row, 1).value
            precio = ws.cell(row, 7).value  # Col G = precio
            if ticker:
                self._precios_iniciales_cache[str(ticker).upper().strip()] = precio
    
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
    
    def _get_precio_inicial(self, ticker: str) -> float:
        """Obtiene precio inicial de una especie por ticker."""
        ticker_upper = str(ticker).upper().strip()
        
        # Valores fijos
        if ticker_upper in ['PESOS', '$']:
            return 1.0
        if ticker_upper in ['DOLARES', 'USD', 'U$S', 'DOLAR']:
            return 1167.806
        if 'CABLE' in ticker_upper:
            return 1148.93
        
        return self._precios_iniciales_cache.get(ticker_upper, 0)
    
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
        """Crea hoja Posicion Inicial Gallo."""
        ws = wb.create_sheet("Posicion Inicial Gallo")
        
        # Headers
        headers = ['tipo_especie', 'Ticker', 'especie', 'detalle', 'custodia', 
                   'cantidad', 'precio', 'importe_pesos', 'porc_cartera_pesos',
                   'importe_dolares', 'porc_cartera_dolares']
        
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
            
            # Datos originales
            detalle = gallo_ws.cell(row, 3).value
            custodia = gallo_ws.cell(row, 4).value
            cantidad = gallo_ws.cell(row, 5).value
            precio = gallo_ws.cell(row, 6).value
            importe_pesos = gallo_ws.cell(row, 7).value
            porc_pesos = gallo_ws.cell(row, 8).value
            importe_usd = gallo_ws.cell(row, 9).value
            porc_usd = gallo_ws.cell(row, 10).value
            
            # Calcular precio si falta
            if not precio and cantidad and importe_pesos:
                try:
                    precio = float(importe_pesos) / float(cantidad)
                except:
                    precio = 0
            
            # Escribir fila
            ws.cell(row_out, 1, tipo_especie)
            ws.cell(row_out, 2, ticker)
            ws.cell(row_out, 3, especie)
            ws.cell(row_out, 4, detalle)
            ws.cell(row_out, 5, custodia)
            ws.cell(row_out, 6, cantidad)
            ws.cell(row_out, 7, precio)
            ws.cell(row_out, 8, importe_pesos)
            ws.cell(row_out, 9, porc_pesos)
            ws.cell(row_out, 10, importe_usd)
            ws.cell(row_out, 11, porc_usd)
            
            row_out += 1
    
    def _create_posicion_final(self, wb: Workbook):
        """Crea hoja Posicion Final Gallo con columnas adicionales."""
        ws = wb.create_sheet("Posicion Final Gallo")
        
        # Headers (20 columnas)
        headers = ['tipo_especie', 'Ticker', 'especie', 'Codigo especie(gallo match con otras hojas)',
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
            
            # Buscar código de especie
            codigo, codigo_origen = self._buscar_codigo_especie(especie_full, tipo_especie)
            
            # Datos originales
            detalle = gallo_ws.cell(row, 3).value
            custodia = gallo_ws.cell(row, 4).value
            cantidad = gallo_ws.cell(row, 5).value
            importe_pesos = gallo_ws.cell(row, 7).value
            porc_pesos = gallo_ws.cell(row, 8).value
            importe_usd = gallo_ws.cell(row, 9).value
            porc_usd = gallo_ws.cell(row, 10).value
            
            # Calcular precios
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
            
            # Precio tenencia inicial
            precio_inicial = self._get_precio_inicial(ticker)
            
            # Por ahora, precio a utilizar = precio inicial
            precio_a_utilizar = precio_inicial
            
            # Escribir fila
            ws.cell(row_out, 1, tipo_especie)
            ws.cell(row_out, 2, ticker)
            ws.cell(row_out, 3, especie)
            ws.cell(row_out, 4, codigo)
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
            
            try:
                gallo_ws = self.gallo_wb[sheet_name]
            except:
                continue
            
            # Detectar si es hoja de cauciones (estructura diferente)
            is_caucion = 'caucion' in sheet_name.lower()
            
            for row in range(2, gallo_ws.max_row + 1):
                # En cauciones, la operación está en col 6; en otras hojas, col 5
                if is_caucion:
                    operacion = gallo_ws.cell(row, 6).value  # Col F para cauciones
                    fecha = gallo_ws.cell(row, 4).value
                    numero = gallo_ws.cell(row, 7).value
                else:
                    operacion = gallo_ws.cell(row, 5).value  # Col E
                    fecha = gallo_ws.cell(row, 4).value
                    numero = gallo_ws.cell(row, 6).value
                
                if not operacion:
                    continue
                
                operacion_lower = str(operacion).lower().strip()
                
                # Solo operaciones de compra/venta para Boletos
                operaciones_validas = ['compra', 'venta', 'cpra', 'canje', 'licitacion', 'col cau']
                if not any(op in operacion_lower for op in operaciones_validas):
                    continue
                
                # Filtrar solo 2025
                if not self._is_year_2025(fecha):
                    continue
                
                # Extraer datos
                cod_especie = gallo_ws.cell(row, 2).value
                especie = gallo_ws.cell(row, 3).value
                cantidad = gallo_ws.cell(row, 7).value if not is_caucion else gallo_ws.cell(row, 8).value
                precio = gallo_ws.cell(row, 8).value if not is_caucion else 1
                resultado_pesos = gallo_ws.cell(row, 11).value if not is_caucion else None
                resultado_usd = gallo_ws.cell(row, 12).value if not is_caucion else gallo_ws.cell(row, 11).value
                gastos_pesos = gallo_ws.cell(row, 13).value if not is_caucion else gallo_ws.cell(row, 12).value
                gastos_usd = gallo_ws.cell(row, 14).value if not is_caucion else gallo_ws.cell(row, 13).value
                
                # Determinar moneda
                moneda = self._get_moneda(resultado_pesos, resultado_usd, gastos_pesos, gastos_usd, sheet_name, operacion)
                
                # Gastos según moneda
                gastos = gastos_pesos if moneda == "Pesos" else gastos_usd
                if gastos is None:
                    gastos = 0
                
                # Código limpio
                cod_clean = self._clean_codigo(cod_especie)
                
                # Convertir fecha a datetime para Excel
                fecha_dt, _ = self._parse_fecha(fecha)
                
                # Auditoría
                auditoria = f"Origen: Gallo-{sheet_name} | Fecha: {fecha} | Cod: {cod_especie} | Op: {operacion}"
                
                # Guardar transacción (sin fórmulas, se generan al escribir)
                all_transactions.append({
                    'cod_instrum': cod_clean,
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
                
                # Código limpio
                cod_clean = self._clean_codigo(cod_instrum)
                
                auditoria = f"Origen: Visual | Fecha: {concertacion} | Cod: {cod_instrum} | Op: {operacion}"
                
                all_transactions.append({
                    'cod_instrum': cod_clean,
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
        
        # Ordenar por cod_instrum y fecha
        def sort_key(t):
            cod = t.get('cod_instrum') or 0
            try:
                cod_num = int(cod) if str(cod).isdigit() else 999999
            except:
                cod_num = 999999
            fecha = t.get('fecha')
            if isinstance(fecha, datetime):
                return (cod_num, fecha)
            else:
                return (cod_num, datetime.min)
        
        all_transactions.sort(key=sort_key)
        
        # Escribir transacciones ordenadas
        for row_out, trans in enumerate(all_transactions, start=2):
            # Fórmulas con row_out correcto
            tipo_instrumento = f'=VLOOKUP(G{row_out},EspeciesVisual!C:R,16,FALSE)' if not trans['tipo_instrumento_val'] else trans['tipo_instrumento_val']
            instrumento_con_moneda = f'=VLOOKUP(G{row_out},EspeciesVisual!C:Q,15,FALSE)'
            tipo_cambio = f'=IF(E{row_out}="Pesos",1,IFERROR(INDEX(\'Cotizacion Dolar Historica\'!$B:$B,MATCH(1,(\'Cotizacion Dolar Historica\'!$A:$A=B{row_out})*(\'Cotizacion Dolar Historica\'!$C:$C=E{row_out}),0)),""))'
            bruto = f'=J{row_out}*K{row_out}'
            neto = f'=IF(J{row_out}>0,J{row_out}*K{row_out}+O{row_out},J{row_out}*K{row_out}-O{row_out})'
            moneda_emision = f'=VLOOKUP(G{row_out},EspeciesVisual!C:Q,5,FALSE)'
            
            ws.cell(row_out, 1, tipo_instrumento)
            ws.cell(row_out, 2, trans['fecha'])
            ws.cell(row_out, 3, trans['liquidacion'])
            ws.cell(row_out, 4, trans['numero'])
            ws.cell(row_out, 5, trans['moneda'])
            ws.cell(row_out, 6, trans['operacion'])
            ws.cell(row_out, 7, trans['cod_instrum'])
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
    
    def _create_rentas_dividendos_gallo(self, wb: Workbook):
        """Crea hoja Rentas y Dividendos Gallo."""
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
        
        row_out = 2
        
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
                
                # Determinar moneda
                moneda = self._get_moneda(resultado_pesos, resultado_usd, gastos_pesos, gastos_usd, sheet_name, operacion)
                
                # Gastos según moneda
                gastos = gastos_pesos if moneda == "Pesos" else gastos_usd
                if gastos is None:
                    gastos = 0
                
                # Ajustar precio para amortizaciones (100 -> 1)
                if 'amortizacion' in operacion_lower:
                    if precio and float(precio) == 100:
                        precio = 1
                
                # Código limpio
                cod_clean = self._clean_codigo(cod_especie)
                
                # Bruto
                bruto = importe if importe else (cantidad * precio if cantidad and precio else 0)
                
                # Neto calculado
                if 'amortizacion' in operacion_lower:
                    # Para amortización: -bruto - costo
                    neto = f'=-M{row_out}-P{row_out}'
                else:
                    # Normal: bruto - gastos
                    neto = f'=M{row_out}-O{row_out}'
                
                # Fórmulas
                instrumento_con_moneda = f'=VLOOKUP(G{row_out},EspeciesVisual!C:Q,15,FALSE)'
                tipo_instrumento = f'=VLOOKUP(G{row_out},EspeciesVisual!C:R,16,FALSE)'
                tipo_cambio = f'=IF(E{row_out}="Pesos",1,IFERROR(INDEX(\'Cotizacion Dolar Historica\'!$B:$B,MATCH(1,(\'Cotizacion Dolar Historica\'!$A:$A=B{row_out})*(\'Cotizacion Dolar Historica\'!$C:$C=E{row_out}),0)),""))'
                moneda_emision = f'=VLOOKUP(G{row_out},EspeciesVisual!C:Q,5,FALSE)'
                
                auditoria = f"Origen: Gallo-{sheet_name} | Operación: {operacion}"
                
                # Parsear fecha
                fecha_dt, _ = self._parse_fecha(fecha)
                
                # Escribir fila
                ws.cell(row_out, 1, tipo_instrumento)
                ws.cell(row_out, 2, fecha_dt if fecha_dt else fecha)
                ws.cell(row_out, 3, "")
                ws.cell(row_out, 4, numero)
                ws.cell(row_out, 5, moneda)
                ws.cell(row_out, 6, operacion.upper())
                ws.cell(row_out, 7, cod_clean)
                ws.cell(row_out, 8, especie)
                ws.cell(row_out, 9, instrumento_con_moneda)
                ws.cell(row_out, 10, cantidad if cantidad else 0)
                ws.cell(row_out, 11, precio if precio else 0)
                ws.cell(row_out, 12, tipo_cambio)
                ws.cell(row_out, 13, bruto)
                ws.cell(row_out, 14, 0)  # Interés
                ws.cell(row_out, 15, gastos)
                ws.cell(row_out, 16, costo if costo else 0)
                ws.cell(row_out, 17, neto)
                ws.cell(row_out, 18, f"Gallo-{sheet_name}")
                ws.cell(row_out, 19, moneda_emision)
                ws.cell(row_out, 20, auditoria)
                
                row_out += 1
        
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
                categoria = visual_ws.cell(row, 3).value  # Rentas/Dividendos
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
                
                # Código limpio
                cod_clean = self._clean_codigo(cod_instrum)
                
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
                
                # Bruto es el importe
                bruto = importe if importe else 0
                
                # Neto calculado
                neto = f'=M{row_out}-O{row_out}'
                
                # Fórmulas
                instrumento_con_moneda = f'=VLOOKUP(G{row_out},EspeciesVisual!C:Q,15,FALSE)'
                tipo_instrumento_formula = f'=VLOOKUP(G{row_out},EspeciesVisual!C:R,16,FALSE)'
                tipo_cambio = f'=IF(E{row_out}="Pesos",1,IFERROR(INDEX(\'Cotizacion Dolar Historica\'!$B:$B,MATCH(1,(\'Cotizacion Dolar Historica\'!$A:$A=B{row_out})*(\'Cotizacion Dolar Historica\'!$C:$C=E{row_out}),0)),""))'
                moneda_emision = f'=VLOOKUP(G{row_out},EspeciesVisual!C:Q,5,FALSE)'
                
                auditoria = f"Origen: Visual-{visual_sheet_name} | Cat: {categoria} | Op: {tipo_operacion}"
                
                # Escribir fila
                ws.cell(row_out, 1, tipo_instrum if tipo_instrum else tipo_instrumento_formula)
                ws.cell(row_out, 2, fecha_dt if fecha_dt else concertacion)
                ws.cell(row_out, 3, liquidacion)
                ws.cell(row_out, 4, nro_ndc)
                ws.cell(row_out, 5, moneda_final)
                ws.cell(row_out, 6, str(tipo_operacion).upper() if tipo_operacion else "")
                ws.cell(row_out, 7, cod_clean)
                ws.cell(row_out, 8, instrumento)
                ws.cell(row_out, 9, instrumento_con_moneda)
                ws.cell(row_out, 10, cantidad if cantidad else 0)
                ws.cell(row_out, 11, 1)  # Precio = 1 para rentas/dividendos de Visual
                ws.cell(row_out, 12, tipo_cambio)
                ws.cell(row_out, 13, bruto)
                ws.cell(row_out, 14, 0)  # Interés
                ws.cell(row_out, 15, gastos if gastos else 0)
                ws.cell(row_out, 16, 0)  # Costo
                ws.cell(row_out, 17, neto)
                ws.cell(row_out, 18, f"Visual-{visual_sheet_name}")
                ws.cell(row_out, 19, moneda_emision)
                ws.cell(row_out, 20, auditoria)
                
                row_out += 1
    
    def _create_resultado_ventas_ars(self, wb: Workbook):
        """Crea hoja Resultado Ventas ARS."""
        ws = wb.create_sheet("Resultado Ventas ARS")
        
        # Headers (24 columnas)
        headers = ['Origen', 'Tipo de Instrumento', 'Instrumento', 'Cod.Instrum',
                   'Concertación', 'Liquidación', 'Moneda', 'Tipo Operación',
                   'Cantidad', 'Precio', 'Bruto', 'Interés', 'Tipo de Cambio',
                   'Gastos', 'IVA', 'Resultado', 'Cantidad Stock Inicial',
                   'Precio Stock Inicial', 'Costo por venta(gallo)', 'Neto Calculado(visual)',
                   'Resultado Calculado(final)', 'Cantidad de Stock Final', 
                   'Precio Stock Final', 'chequeado']
        
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
            ws.cell(1, col).font = Font(bold=True)
        
        # Las fórmulas referencian a Boletos filtrando por moneda = Pesos
        # Por ahora, creamos la estructura con fórmulas
        boletos_ws = wb['Boletos']
        
        row_out = 2
        for boletos_row in range(2, boletos_ws.max_row + 1):
            # Fórmulas que filtran por moneda = Pesos (columna R de Boletos)
            ws.cell(row_out, 1, f'=Boletos!Q{boletos_row}')  # Origen
            ws.cell(row_out, 2, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!A{boletos_row},"")')
            ws.cell(row_out, 3, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!I{boletos_row},"")')
            ws.cell(row_out, 4, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!G{boletos_row},"")')
            ws.cell(row_out, 5, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!B{boletos_row},"")')
            ws.cell(row_out, 6, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!C{boletos_row},"")')
            ws.cell(row_out, 7, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!E{boletos_row},"")')
            ws.cell(row_out, 8, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!F{boletos_row},"")')
            ws.cell(row_out, 9, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!J{boletos_row},"")')
            ws.cell(row_out, 10, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!K{boletos_row},"")')
            ws.cell(row_out, 11, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!M{boletos_row},"")')
            ws.cell(row_out, 12, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!N{boletos_row},"")')
            ws.cell(row_out, 13, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!L{boletos_row},"")')
            ws.cell(row_out, 14, f'=IF(Boletos!R{boletos_row}="Pesos",Boletos!O{boletos_row},"")')
            ws.cell(row_out, 15, f'=IF(N{row_out}>0,N{row_out}*0.1736,N{row_out}*-0.1736)')  # IVA
            ws.cell(row_out, 16, "")  # Resultado (vacío)
            
            # Running Stock Logic:
            # - Primera fila (row_out=2): VLOOKUP a Posicion Final Gallo
            # - Filas siguientes: Si mismo Cod.Instrum que fila anterior, usar Stock Final anterior
            
            if row_out == 2:
                # Primera fila: siempre VLOOKUP
                ws.cell(row_out, 17, f'=IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:I,6,FALSE),0)')
                ws.cell(row_out, 18, f'=IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:P,12,FALSE),0)')
            else:
                # Filas siguientes: condicional - si mismo código, usar stock final de fila anterior
                prev = row_out - 1
                ws.cell(row_out, 17, f'=IF(D{row_out}=D{prev},V{prev},IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:I,6,FALSE),0))')
                ws.cell(row_out, 18, f'=IF(D{row_out}=D{prev},W{prev},IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:P,12,FALSE),0))')
            
            # Costo por venta (solo si es venta: cantidad negativa)
            ws.cell(row_out, 19, f'=IFERROR(IF(I{row_out}<0,I{row_out}*R{row_out},0),"")')
            
            # Neto Calculado (solo ventas)
            ws.cell(row_out, 20, f'=IF(S{row_out}<>0,K{row_out}+N{row_out},0)')
            
            # Resultado Calculado
            ws.cell(row_out, 21, f'=IF(S{row_out}<>0,ABS(T{row_out})-ABS(S{row_out}),0)')
            
            # Cantidad Stock Final (running)
            ws.cell(row_out, 22, f'=I{row_out}+Q{row_out}')
            
            # Precio Stock Final (promedio ponderado)
            ws.cell(row_out, 23, f'=IF(V{row_out}=0,0,IF(I{row_out}>0,(I{row_out}*J{row_out}+Q{row_out}*R{row_out})/(I{row_out}+Q{row_out}),R{row_out}))')
            
            # Chequeado/Auditoría - Explicación de cómo se obtienen los resultados
            if row_out == 2:
                auditoria = (
                    "AUDITORÍA ARS: "
                    "Q=VLOOKUP(CodInstrum→PosicionFinal.cantidad) | "
                    "R=VLOOKUP(CodInstrum→PosicionFinal.PrecioInicial) | "
                    "S=Cantidad*PrecioStock (si venta) | "
                    "T=Bruto+Gastos (si venta) | "
                    "U=|T|-|S| (resultado) | "
                    "V=Cantidad+StockInicial (running) | "
                    "W=Promedio ponderado si compra, sino mantiene precio"
                )
            else:
                prev = row_out - 1
                auditoria = f"Si D{row_out}=D{prev}: Q=V{prev}, R=W{prev}; sino VLOOKUP. Running stock por especie."
            ws.cell(row_out, 24, auditoria)
            
            row_out += 1
    
    def _create_resultado_ventas_usd(self, wb: Workbook):
        """Crea hoja Resultado Ventas USD."""
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
        
        # Similar a ARS pero con SEARCH("Dolar",...) 
        boletos_ws = wb['Boletos']
        
        row_out = 2
        for boletos_row in range(2, boletos_ws.max_row + 1):
            # Fórmulas que filtran por moneda contiene "Dolar"
            ws.cell(row_out, 1, f'=Boletos!Q{boletos_row}')
            ws.cell(row_out, 2, f'=IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!A{boletos_row},"")')
            ws.cell(row_out, 3, f'=IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!I{boletos_row},"")')
            ws.cell(row_out, 4, f'=IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!G{boletos_row},"")')
            ws.cell(row_out, 5, f'=IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!B{boletos_row},"")')
            ws.cell(row_out, 6, f'=IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!C{boletos_row},"")')
            ws.cell(row_out, 7, f'=IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!E{boletos_row},"")')
            ws.cell(row_out, 8, f'=IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!F{boletos_row},"")')
            ws.cell(row_out, 9, f'=IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!J{boletos_row},"")')
            ws.cell(row_out, 10, f'=IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!K{boletos_row},"")')
            
            # Precio estandarizado (ajuste x100 si es nominal)
            ws.cell(row_out, 11, f'=IF(A{row_out}="visual",IF(U{row_out}*P{row_out}/J{row_out}>80,J{row_out}*100),J{row_out})')
            
            # Precio en USD
            ws.cell(row_out, 12, f'=IF(G{row_out}="Pesos",K{row_out}/P{row_out},K{row_out})')
            
            # Bruto en USD
            ws.cell(row_out, 13, f'=I{row_out}*L{row_out}')
            
            ws.cell(row_out, 14, 0)  # Interés
            
            # Tipo cambio (referencia USD=1)
            ws.cell(row_out, 15, f'=IF(G{row_out}="Pesos",IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!L{boletos_row},"")/P{row_out},1)')
            
            # Valor USD Dia (referencia manual o lookup)
            ws.cell(row_out, 16, "")
            
            ws.cell(row_out, 17, f'=IF(SEARCH("Dolar",Boletos!R{boletos_row}),Boletos!O{boletos_row},"")')
            ws.cell(row_out, 18, 0)  # IVA
            ws.cell(row_out, 19, "")  # Resultado
            
            # Running Stock Logic para USD:
            # - Primera fila (row_out=2): VLOOKUP a Posicion Final Gallo
            # - Filas siguientes: Si mismo Cod.Instrum que fila anterior, usar Stock Final anterior
            
            if row_out == 2:
                # Primera fila: siempre VLOOKUP
                ws.cell(row_out, 20, f'=IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:I,6,FALSE),0)')
                ws.cell(row_out, 21, f'=IFERROR(VLOOKUP(D{row_out},PreciosInicialesEspecies!A:G,7,FALSE),0)/P{row_out}')
            else:
                # Filas siguientes: condicional
                prev = row_out - 1
                ws.cell(row_out, 20, f'=IF(D{row_out}=D{prev},Y{prev},IFERROR(VLOOKUP(D{row_out},\'Posicion Final Gallo\'!D:I,6,FALSE),0))')
                ws.cell(row_out, 21, f'=IF(D{row_out}=D{prev},Z{prev},IFERROR(VLOOKUP(D{row_out},PreciosInicialesEspecies!A:G,7,FALSE),0)/P{row_out})')
            
            # Costo por venta
            ws.cell(row_out, 22, f'=IFERROR(IF(I{row_out}<0,I{row_out}*U{row_out},0),"")')
            
            # Neto Calculado
            ws.cell(row_out, 23, f'=IF(V{row_out}<>0,M{row_out}-Q{row_out},0)')
            
            # Resultado Calculado
            ws.cell(row_out, 24, f'=IFERROR(IF(V{row_out}<>0,ABS(W{row_out})-ABS(V{row_out}),0),0)')
            
            # Cantidad Stock Final
            ws.cell(row_out, 25, f'=I{row_out}+T{row_out}')
            
            # Precio Stock Final
            ws.cell(row_out, 26, f'=IF(Y{row_out}=0,0,IF(I{row_out}>0,(I{row_out}*L{row_out}+T{row_out}*U{row_out})/(I{row_out}+T{row_out}),U{row_out}))')
            
            # Comentarios/Auditoría USD
            if row_out == 2:
                auditoria = (
                    "AUDITORÍA USD: "
                    "T=VLOOKUP(CodInstrum→PosicionFinal.cantidad) | "
                    "U=VLOOKUP(CodInstrum→PreciosIniciales)/TipoCambio | "
                    "V=Cantidad*PrecioStock (si venta) | "
                    "W=BrutoUSD-Gastos (si venta) | "
                    "X=|W|-|V| (resultado) | "
                    "Y=Cantidad+StockInicial (running) | "
                    "Z=Promedio ponderado si compra"
                )
            else:
                prev = row_out - 1
                auditoria = f"Si D{row_out}=D{prev}: T=Y{prev}, U=Z{prev}; sino VLOOKUP. Running stock USD por especie."
            ws.cell(row_out, 27, auditoria)
            
            row_out += 1
    
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
