"""
Módulo para exportar el Excel consolidado a PDF con formato Visual.

Genera un PDF con todas las secciones del reporte financiero:
- Boletos (por tipo de instrumento)
- Resultado Ventas ARS
- Resultado Ventas USD
- Rentas y Dividendos ARS
- Rentas y Dividendos USD
- Cauciones Tomadoras
- Cauciones Colocadoras
- Resumen
- Posición de Títulos
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, 
    PageBreak, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
import io


class ExcelToPdfExporter:
    """
    Exporta un Excel consolidado (merge Gallo+Visual) a PDF con formato Visual.
    """
    
    # Colores corporativos
    HEADER_BG = colors.Color(0.2, 0.3, 0.5)  # Azul oscuro
    HEADER_TEXT = colors.white
    ROW_ALT_BG = colors.Color(0.95, 0.95, 0.95)  # Gris claro
    SECTION_BG = colors.Color(0.85, 0.85, 0.9)  # Gris azulado
    SUBSECTION_BG = colors.Color(0.9, 0.9, 0.95)  # Gris más claro
    
    def __init__(self, excel_path: str, cliente_info: Dict[str, str] = None):
        """
        Inicializa el exportador.
        
        Args:
            excel_path: Ruta al Excel consolidado
            cliente_info: Diccionario con info del cliente (numero, nombre)
        """
        self.excel_path = Path(excel_path)
        self.wb = load_workbook(excel_path, data_only=True)
        
        # Info del cliente
        self.cliente_info = cliente_info or {
            'numero': 'XXXXX',
            'nombre': 'CLIENTE'
        }
        
        # Período del reporte (detectar de datos o usar default)
        self.periodo_inicio = "Enero 1"
        self.periodo_fin = "Diciembre 31"
        self.anio = datetime.now().year
        
        # Estilos
        self.styles = getSampleStyleSheet()
        self._setup_styles()
    
    def _setup_styles(self):
        """Configura estilos personalizados."""
        # Título de sección
        self.styles.add(ParagraphStyle(
            'SectionTitle',
            parent=self.styles['Heading2'],
            fontSize=12,
            spaceAfter=6,
            spaceBefore=12,
            textColor=colors.Color(0.2, 0.2, 0.4),
            fontName='Helvetica-Bold'
        ))
        
        # Subtítulo (tipo de instrumento)
        self.styles.add(ParagraphStyle(
            'SubsectionTitle',
            parent=self.styles['Normal'],
            fontSize=10,
            spaceBefore=6,
            spaceAfter=3,
            textColor=colors.Color(0.3, 0.3, 0.5),
            fontName='Helvetica-Bold'
        ))
        
        # Normal pequeño para tablas
        self.styles.add(ParagraphStyle(
            'TableCell',
            parent=self.styles['Normal'],
            fontSize=7,
            leading=9
        ))
        
        # Header de tabla
        self.styles.add(ParagraphStyle(
            'TableHeader',
            parent=self.styles['Normal'],
            fontSize=7,
            fontName='Helvetica-Bold',
            textColor=colors.white
        ))
    
    def _format_number(self, value: Any, decimals: int = 2) -> str:
        """Formatea un número al estilo argentino (punto miles, coma decimales)."""
        if value is None:
            return ""
        try:
            num = float(value)
            # Manejar negativos con paréntesis
            is_negative = num < 0
            num = abs(num)
            
            # Formatear con separadores
            if decimals == 0:
                formatted = f"{num:,.0f}"
            else:
                formatted = f"{num:,.{decimals}f}"
            
            # Convertir a formato argentino
            # Primero reemplazar comas por placeholder, luego puntos por comas, finalmente placeholder por puntos
            formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
            
            if is_negative:
                return f"({formatted})"
            return formatted
        except (ValueError, TypeError):
            return str(value) if value else ""
    
    def _format_date(self, value: Any) -> str:
        """Formatea una fecha como D/M/AAAA."""
        if value is None:
            return ""
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y")
        return str(value)
    
    def _get_header_footer(self, page_num: int, total_pages: int) -> str:
        """Genera el encabezado del reporte."""
        return f"REPORTE DE GANANCIAS / Período {self.periodo_inicio} - {self.periodo_fin}, {self.anio}   Página {page_num} de {total_pages}       {self.cliente_info['numero']} - {self.cliente_info['nombre']}"
    
    def _read_sheet_data(self, sheet_name: str) -> Tuple[List[str], List[List[Any]]]:
        """
        Lee datos de una hoja Excel.
        
        Returns:
            Tuple de (headers, rows)
        """
        if sheet_name not in self.wb.sheetnames:
            return [], []
        
        ws = self.wb[sheet_name]
        
        # Headers
        headers = []
        for col in range(1, ws.max_column + 1):
            val = ws.cell(1, col).value
            headers.append(str(val) if val else "")
        
        # Data rows
        rows = []
        for row_num in range(2, ws.max_row + 1):
            row_data = []
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row_num, col).value
                row_data.append(val)
            rows.append(row_data)
        
        return headers, rows
    
    def _create_table(self, headers: List[str], rows: List[List[Any]], 
                      col_widths: List[float] = None,
                      col_formatters: Dict[int, str] = None) -> Table:
        """
        Crea una tabla formateada.
        
        Args:
            headers: Lista de encabezados
            rows: Lista de filas de datos
            col_widths: Anchos de columnas en mm
            col_formatters: Diccionario {col_index: 'date'|'number'|'integer'|'text'}
        """
        if not headers:
            return None
        
        col_formatters = col_formatters or {}
        
        # Formatear datos
        formatted_rows = [headers]
        for row in rows:
            formatted_row = []
            for i, val in enumerate(row):
                fmt_type = col_formatters.get(i, 'text')
                if fmt_type == 'date':
                    formatted_row.append(self._format_date(val))
                elif fmt_type == 'number':
                    formatted_row.append(self._format_number(val, 2))
                elif fmt_type == 'integer':
                    formatted_row.append(self._format_number(val, 0))
                else:
                    formatted_row.append(str(val) if val is not None else "")
            formatted_rows.append(formatted_row)
        
        # Crear tabla
        if col_widths:
            table = Table(formatted_rows, colWidths=[w * mm for w in col_widths])
        else:
            table = Table(formatted_rows)
        
        # Estilo
        style = TableStyle([
            # Header
            ('BACKGROUND', (0, 0), (-1, 0), self.HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), self.HEADER_TEXT),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 7),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            
            # Body
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
            
            # Números alineados a la derecha
            *[(('ALIGN', (i, 1), (i, -1), 'RIGHT')) for i in range(len(headers)) 
              if col_formatters.get(i) in ('number', 'integer')],
            
            # Bordes
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
            
            # Padding
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ])
        
        # Filas alternadas
        for i in range(1, len(formatted_rows)):
            if i % 2 == 0:
                style.add('BACKGROUND', (0, i), (-1, i), self.ROW_ALT_BG)
        
        table.setStyle(style)
        return table
    
    def _build_boletos_section(self) -> List:
        """Construye la sección de Boletos."""
        elements = []
        elements.append(Paragraph("Boletos", self.styles['SectionTitle']))
        
        headers, rows = self._read_sheet_data('Boletos')
        if not rows:
            elements.append(Paragraph("Sin operaciones en el período", self.styles['Normal']))
            elements.append(Spacer(1, 10*mm))
            return elements
        
        # Columnas a mostrar (indices del Excel)
        # Headers: Tipo de Instrumento(0), Concertación(1), Liquidación(2), Nro. Boleto(3), 
        # Moneda(4), Tipo Operación(5), Cod.Instrum(6), Instrumento Crudo(7), 
        # InstrumentoConMoneda(8), Cantidad(9), Precio(10), Tipo Cambio(11), 
        # Bruto(12), Interés(13), Gastos(14), Neto Calculado(15)
        
        # Mapeo de columnas a mostrar
        col_map = [
            (1, 'Concertación', 'date'),
            (2, 'Liquidación', 'date'),
            (3, 'Nro. Boleto', 'integer'),
            (4, 'Moneda', 'text'),
            (5, 'Tipo Operación', 'text'),
            (6, 'Cod.Instrum', 'integer'),
            (7, 'Instrumento', 'text'),
            (9, 'Cantidad', 'number'),
            (10, 'Precio', 'number'),
            (11, 'Tipo Cambio', 'number'),
            (12, 'Bruto', 'number'),
            (13, 'Interés', 'number'),
            (14, 'Gastos', 'number'),
            (15, 'Neto', 'number'),
        ]
        
        # Agrupar por tipo de instrumento
        by_tipo = {}
        for row in rows:
            tipo = row[0] if row[0] else "Otros"
            if tipo not in by_tipo:
                by_tipo[tipo] = []
            by_tipo[tipo].append(row)
        
        # Ordenar tipos
        tipos_order = ['Acciones', 'Títulos Públicos', 'Obligaciones Negociables', 
                       'Letras del Tesoro', 'CEDEAR', 'FCI', 'Otros']
        sorted_tipos = sorted(by_tipo.keys(), 
                            key=lambda x: tipos_order.index(x) if x in tipos_order else 999)
        
        for tipo in sorted_tipos:
            tipo_rows = by_tipo[tipo]
            elements.append(Paragraph(tipo, self.styles['SubsectionTitle']))
            
            # Preparar datos para tabla
            table_headers = [c[1] for c in col_map]
            table_rows = []
            col_formatters = {i: c[2] for i, c in enumerate(col_map)}
            
            for row in tipo_rows:
                table_row = [row[c[0]] if c[0] < len(row) else None for c in col_map]
                table_rows.append(table_row)
            
            # Anchos de columnas (total ~270mm para landscape A4)
            col_widths = [18, 18, 15, 20, 22, 14, 45, 18, 16, 16, 20, 14, 16, 18]
            
            table = self._create_table(table_headers, table_rows, col_widths, col_formatters)
            if table:
                elements.append(table)
            elements.append(Spacer(1, 3*mm))
        
        elements.append(Spacer(1, 10*mm))
        return elements
    
    def _build_resultado_ventas_section(self, moneda: str) -> List:
        """Construye la sección de Resultado Ventas (ARS o USD)."""
        elements = []
        sheet_name = f"Resultado Ventas {moneda}"
        
        elements.append(Paragraph(f"Resultado Ventas", self.styles['SectionTitle']))
        elements.append(Paragraph(moneda, self.styles['SubsectionTitle']))
        
        headers, rows = self._read_sheet_data(sheet_name)
        if not rows:
            elements.append(Paragraph("Sin operaciones en el período", self.styles['Normal']))
            elements.append(Spacer(1, 10*mm))
            return elements
        
        # Columnas para ARS: Origen(0), Tipo de Instrumento(1), Instrumento(2), Cod.Instrum(3),
        # Concertación(4), Liquidación(5), Moneda(6), Tipo Operación(7), Cantidad(8), Precio(9),
        # Bruto(10), Interés(11), Tipo de Cambio(12), Gastos(13), IVA(14), Resultado(15)
        
        if moneda == 'ARS':
            col_map = [
                (4, 'Concertación', 'date'),
                (5, 'Liquidación', 'date'),
                (6, 'Moneda', 'text'),
                (7, 'Tipo Operación', 'text'),
                (8, 'Cantidad', 'number'),
                (9, 'Precio', 'number'),
                (10, 'Bruto', 'number'),
                (11, 'Interés', 'number'),
                (12, 'Tipo de Cambio', 'number'),
                (13, 'Gastos', 'number'),
                (14, 'IVA', 'number'),
                (15, 'Resultado', 'number'),
            ]
        else:  # USD
            # Columnas USD: similar pero con columnas adicionales
            col_map = [
                (4, 'Concertación', 'date'),
                (5, 'Liquidación', 'date'),
                (6, 'Moneda', 'text'),
                (7, 'Tipo Operación', 'text'),
                (8, 'Cantidad', 'number'),
                (9, 'Precio', 'number'),
                (12, 'Bruto USD', 'number'),
                (13, 'Interés', 'number'),
                (14, 'Tipo de Cambio', 'number'),
                (16, 'Gastos', 'number'),
                (17, 'IVA', 'number'),
                (18, 'Resultado', 'number'),
            ]
        
        # Agrupar por tipo de instrumento e instrumento
        by_tipo = {}
        for row in rows:
            tipo_idx = 1  # Tipo de Instrumento
            instr_idx = 2  # Instrumento
            
            tipo = row[tipo_idx] if row[tipo_idx] else "Otros"
            instr = row[instr_idx] if row[instr_idx] else "Sin nombre"
            
            if tipo not in by_tipo:
                by_tipo[tipo] = {}
            if instr not in by_tipo[tipo]:
                by_tipo[tipo][instr] = []
            by_tipo[tipo][instr].append(row)
        
        for tipo in sorted(by_tipo.keys()):
            elements.append(Paragraph(tipo, self.styles['SubsectionTitle']))
            
            for instr in sorted(by_tipo[tipo].keys()):
                instr_rows = by_tipo[tipo][instr]
                
                # Nombre del instrumento
                elements.append(Paragraph(f"  {instr}", self.styles['Normal']))
                
                # Preparar tabla
                table_headers = [c[1] for c in col_map]
                table_rows = []
                col_formatters = {i: c[2] for i, c in enumerate(col_map)}
                
                for row in instr_rows:
                    table_row = [row[c[0]] if c[0] < len(row) else None for c in col_map]
                    table_rows.append(table_row)
                
                col_widths = [18, 18, 20, 25, 18, 16, 22, 14, 18, 18, 16, 22]
                
                table = self._create_table(table_headers, table_rows, col_widths, col_formatters)
                if table:
                    elements.append(table)
                elements.append(Spacer(1, 2*mm))
        
        elements.append(Spacer(1, 10*mm))
        return elements
    
    def _build_rentas_dividendos_section(self, moneda: str) -> List:
        """Construye la sección de Rentas y Dividendos (ARS o USD)."""
        elements = []
        sheet_name = f"Rentas Dividendos {moneda}"
        
        elements.append(Paragraph("Rentas y Dividendos", self.styles['SectionTitle']))
        elements.append(Paragraph(moneda, self.styles['SubsectionTitle']))
        
        headers, rows = self._read_sheet_data(sheet_name)
        if not rows:
            elements.append(Paragraph("Sin operaciones en el período", self.styles['Normal']))
            elements.append(Spacer(1, 10*mm))
            return elements
        
        # Headers: Instrumento(0), Cod.Instrum(1), Categoría(2), tipo_instrumento(3),
        # Concertación(4), Liquidación(5), Nro. NDC(6), Tipo Operación(7),
        # Cantidad(8), Moneda(9), Tipo de Cambio(10), Gastos(11), Importe(12)
        
        col_map = [
            (4, 'Concertación', 'date'),
            (5, 'Liquidación', 'date'),
            (6, 'Nro. NDC', 'integer'),
            (7, 'Tipo Operación', 'text'),
            (8, 'Cantidad', 'number'),
            (9, 'Moneda', 'text'),
            (10, 'Tipo de Cambio', 'number'),
            (11, 'Gastos', 'number'),
            (12, 'Importe', 'number'),
        ]
        
        # Agrupar por categoría y tipo_instrumento
        by_cat = {}
        for row in rows:
            cat = row[2] if row[2] else "Otros"  # Categoría
            tipo_instr = row[3] if row[3] else "Sin tipo"  # tipo_instrumento
            instr = row[0] if row[0] else "Sin nombre"  # Instrumento
            
            if cat not in by_cat:
                by_cat[cat] = {}
            if tipo_instr not in by_cat[cat]:
                by_cat[cat][tipo_instr] = {}
            if instr not in by_cat[cat][tipo_instr]:
                by_cat[cat][tipo_instr][instr] = []
            by_cat[cat][tipo_instr][instr].append(row)
        
        # Ordenar: Rentas primero, luego Dividendos
        cat_order = ['Rentas', 'Dividendos', 'Otros']
        sorted_cats = sorted(by_cat.keys(), 
                           key=lambda x: cat_order.index(x) if x in cat_order else 999)
        
        for cat in sorted_cats:
            elements.append(Paragraph(cat, self.styles['SubsectionTitle']))
            
            for tipo_instr in sorted(by_cat[cat].keys()):
                elements.append(Paragraph(f"  {tipo_instr}", self.styles['Normal']))
                
                for instr in sorted(by_cat[cat][tipo_instr].keys()):
                    instr_rows = by_cat[cat][tipo_instr][instr]
                    
                    # Nombre del instrumento
                    elements.append(Paragraph(f"    {instr}", 
                                            ParagraphStyle('InstrName', 
                                                          parent=self.styles['Normal'],
                                                          fontSize=8,
                                                          textColor=colors.grey)))
                    
                    table_headers = [c[1] for c in col_map]
                    table_rows = []
                    col_formatters = {i: c[2] for i, c in enumerate(col_map)}
                    
                    for row in instr_rows:
                        table_row = [row[c[0]] if c[0] < len(row) else None for c in col_map]
                        table_rows.append(table_row)
                    
                    col_widths = [18, 18, 15, 28, 18, 22, 18, 18, 25]
                    
                    table = self._create_table(table_headers, table_rows, col_widths, col_formatters)
                    if table:
                        elements.append(table)
                    elements.append(Spacer(1, 2*mm))
        
        elements.append(Spacer(1, 10*mm))
        return elements
    
    def _build_cauciones_section(self, tipo: str = "tomadoras") -> List:
        """
        Construye la sección de Cauciones.
        
        Args:
            tipo: 'tomadoras' o 'colocadoras'
        """
        elements = []
        titulo = f"Cauciones {tipo}"
        elements.append(Paragraph(titulo, self.styles['SectionTitle']))
        
        headers, rows = self._read_sheet_data('Cauciones')
        if not rows:
            elements.append(Paragraph("Sin operaciones en el período", self.styles['Normal']))
            elements.append(Spacer(1, 10*mm))
            return elements
        
        # Filtrar por tipo si hay columna de tipo
        # Headers: Concertación(0), Plazo(1), Liquidación(2), Operación(3), Boleto(4),
        # Contado(5), Futuro(6), Tipo de Cambio(7), Tasa (%)(8), Interés Bruto(9),
        # Interés Devengado(10), Aranceles(11), Derechos(12), Costo Financiero(13), Moneda(14)
        
        # Filtrar por tipo de operación
        filtered_rows = []
        for row in rows:
            operacion = str(row[3]).upper() if row[3] else ""
            if tipo == "tomadoras" and "TOM" in operacion:
                filtered_rows.append(row)
            elif tipo == "colocadoras" and "COL" in operacion:
                filtered_rows.append(row)
        
        if not filtered_rows:
            elements.append(Paragraph("Sin operaciones en el período", self.styles['Normal']))
            elements.append(Spacer(1, 10*mm))
            return elements
        
        col_map = [
            (0, 'Concertación', 'date'),
            (1, 'Plazo', 'integer'),
            (2, 'Liquidación', 'date'),
            (3, 'Operación', 'text'),
            (4, '# Boleto', 'integer'),
            (5, 'Contado', 'number'),
            (6, 'Futuro', 'number'),
            (7, 'Tipo de cambio', 'number'),
            (8, 'Tasa (%)', 'number'),
            (9, 'Interés Bruto', 'number'),
            (10, 'Interés Devengad', 'number'),
            (11, 'Aranceles', 'number'),
            (12, 'Derechos', 'number'),
            (13, 'Costo financiero', 'number'),
        ]
        
        # Agrupar por moneda
        by_moneda = {}
        for row in filtered_rows:
            moneda = row[14] if len(row) > 14 and row[14] else "Pesos"
            if moneda not in by_moneda:
                by_moneda[moneda] = []
            by_moneda[moneda].append(row)
        
        for moneda in ['Pesos', 'Dólares', 'Dolar MEP', 'Dolar Cable']:
            if moneda not in by_moneda:
                continue
            
            elements.append(Paragraph(f"  {moneda}", self.styles['SubsectionTitle']))
            
            table_headers = [c[1] for c in col_map]
            table_rows = []
            col_formatters = {i: c[2] for i, c in enumerate(col_map)}
            
            for row in by_moneda[moneda]:
                table_row = [row[c[0]] if c[0] < len(row) else None for c in col_map]
                table_rows.append(table_row)
            
            col_widths = [18, 10, 18, 25, 15, 22, 22, 16, 14, 18, 18, 15, 14, 18]
            
            table = self._create_table(table_headers, table_rows, col_widths, col_formatters)
            if table:
                elements.append(table)
            
            # Total
            total_cf = sum(float(r[13] or 0) for r in by_moneda[moneda])
            elements.append(Paragraph(f"Totales: {self._format_number(total_cf)}", 
                                     self.styles['Normal']))
            elements.append(Spacer(1, 3*mm))
        
        elements.append(Spacer(1, 10*mm))
        return elements
    
    def _build_resumen_section(self) -> List:
        """Construye la sección de Resumen."""
        elements = []
        elements.append(Paragraph("Resumen", self.styles['SectionTitle']))
        
        headers, rows = self._read_sheet_data('Resumen')
        if not rows:
            elements.append(Paragraph("Sin datos de resumen", self.styles['Normal']))
            return elements
        
        # Headers: Moneda(0), Ventas(1), FCI(2), Opciones(3), Rentas(4), Dividendos(5),
        # Ef. CPD(6), Pagarés(7), Futuros(8), Cau (int)(9), Cau (CF)(10), Total(11)
        
        table_headers = ['Moneda', 'Resultados', '', '', '', '', '', '', '', '', '', 'Total']
        sub_headers = ['', 'Ventas', 'FCI', 'Opciones', 'Rentas', 'Dividendos', 
                      'Ef. CPD', 'Pagarés', 'Futuros', 'Cau (int)', 'Cau (CF)', '']
        
        table_data = [table_headers, sub_headers]
        
        for row in rows:
            formatted_row = [
                row[0] if row[0] else "",  # Moneda
                self._format_number(row[1]),  # Ventas
                self._format_number(row[2]),  # FCI
                self._format_number(row[3]),  # Opciones
                self._format_number(row[4]),  # Rentas
                self._format_number(row[5]),  # Dividendos
                self._format_number(row[6]),  # Ef. CPD
                self._format_number(row[7]),  # Pagarés
                self._format_number(row[8]),  # Futuros
                self._format_number(row[9]),  # Cau (int)
                self._format_number(row[10]),  # Cau (CF)
                self._format_number(row[11]),  # Total
            ]
            table_data.append(formatted_row)
        
        col_widths = [20, 28, 20, 20, 22, 24, 20, 20, 20, 22, 22, 32]
        
        table = Table(table_data, colWidths=[w * mm for w in col_widths])
        
        style = TableStyle([
            # Header
            ('BACKGROUND', (0, 0), (-1, 1), self.HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 1), self.HEADER_TEXT),
            ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 1), 8),
            ('ALIGN', (0, 0), (-1, 1), 'CENTER'),
            
            # Merge "Resultados" header
            ('SPAN', (1, 0), (10, 0)),
            
            # Body
            ('FONTNAME', (0, 2), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 2), (-1, -1), 8),
            ('ALIGN', (1, 2), (-1, -1), 'RIGHT'),
            ('ALIGN', (0, 2), (0, -1), 'LEFT'),
            
            # Bordes
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('LINEBELOW', (0, 1), (-1, 1), 1, colors.black),
            
            # Padding
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ])
        
        table.setStyle(style)
        elements.append(table)
        elements.append(Spacer(1, 15*mm))
        
        return elements
    
    def _build_posicion_titulos_section(self) -> List:
        """Construye la sección de Posición de Títulos."""
        elements = []
        elements.append(Paragraph("Posición de Títulos", self.styles['SectionTitle']))
        
        headers, rows = self._read_sheet_data('Posicion Titulos')
        if not rows:
            elements.append(Paragraph("Sin posiciones", self.styles['Normal']))
            return elements
        
        # Agregar subtítulo "Es Disponible Sí"
        elements.append(Paragraph("Es Disponible Sí", self.styles['Normal']))
        elements.append(Spacer(1, 3*mm))
        
        # Headers: Instrumento(0), Código(1), Ticker(2), Cantidad(3), Importe(4), Moneda(5)
        
        col_map = [
            (0, 'Instrumento', 'text'),
            (1, 'Código', 'integer'),
            (2, 'Ticker', 'text'),
            (3, 'Cantidad', 'number'),
            (4, 'Importe', 'number'),
            (5, 'Moneda', 'text'),
        ]
        
        table_headers = [c[1] for c in col_map]
        table_rows = []
        col_formatters = {i: c[2] for i, c in enumerate(col_map)}
        
        for row in rows:
            table_row = [row[c[0]] if c[0] < len(row) else None for c in col_map]
            table_rows.append(table_row)
        
        col_widths = [90, 20, 20, 30, 35, 25]
        
        table = self._create_table(table_headers, table_rows, col_widths, col_formatters)
        if table:
            elements.append(table)
        
        return elements
    
    def export_to_pdf(self, output_path: str = None) -> bytes:
        """
        Exporta el Excel a PDF.
        
        Args:
            output_path: Ruta opcional para guardar el archivo
            
        Returns:
            bytes del PDF generado
        """
        # Buffer para el PDF
        buffer = io.BytesIO()
        
        # Crear documento
        doc = SimpleDocTemplate(
            buffer,
            pagesize=landscape(A4),
            rightMargin=10*mm,
            leftMargin=10*mm,
            topMargin=20*mm,
            bottomMargin=15*mm
        )
        
        # Construir contenido
        elements = []
        
        # Header principal (se repetirá en cada página via onPage)
        
        # Boletos
        elements.extend(self._build_boletos_section())
        elements.append(PageBreak())
        
        # Resultado Ventas ARS
        elements.extend(self._build_resultado_ventas_section('ARS'))
        elements.append(PageBreak())
        
        # Resultado Ventas USD
        elements.extend(self._build_resultado_ventas_section('USD'))
        elements.append(PageBreak())
        
        # Rentas y Dividendos ARS
        elements.extend(self._build_rentas_dividendos_section('ARS'))
        elements.append(PageBreak())
        
        # Rentas y Dividendos USD
        elements.extend(self._build_rentas_dividendos_section('USD'))
        elements.append(PageBreak())
        
        # Cauciones Tomadoras
        elements.extend(self._build_cauciones_section('tomadoras'))
        
        # Cauciones Colocadoras
        elements.extend(self._build_cauciones_section('colocadoras'))
        elements.append(PageBreak())
        
        # Resumen
        elements.extend(self._build_resumen_section())
        
        # Posición de Títulos
        elements.extend(self._build_posicion_titulos_section())
        
        # Generar PDF
        def add_header_footer(canvas, doc):
            canvas.saveState()
            # Header
            canvas.setFont('Helvetica', 8)
            header_text = f"REPORTE DE GANANCIAS / Período {self.periodo_inicio} - {self.periodo_fin}, {self.anio}       {self.cliente_info['numero']} - {self.cliente_info['nombre']}"
            canvas.drawString(10*mm, landscape(A4)[1] - 12*mm, header_text)
            
            # Page number
            page_num = f"Página {doc.page}"
            canvas.drawRightString(landscape(A4)[0] - 10*mm, landscape(A4)[1] - 12*mm, page_num)
            
            canvas.restoreState()
        
        doc.build(elements, onFirstPage=add_header_footer, onLaterPages=add_header_footer)
        
        # Obtener bytes
        pdf_bytes = buffer.getvalue()
        buffer.close()
        
        # Guardar si se especificó ruta
        if output_path:
            with open(output_path, 'wb') as f:
                f.write(pdf_bytes)
        
        return pdf_bytes


def export_excel_to_pdf(excel_path: str, output_path: str = None,
                        cliente_numero: str = None, cliente_nombre: str = None,
                        periodo_inicio: str = None, periodo_fin: str = None,
                        anio: int = None) -> bytes:
    """
    Función de conveniencia para exportar un Excel consolidado a PDF.
    
    Args:
        excel_path: Ruta al archivo Excel consolidado
        output_path: Ruta opcional para guardar el PDF
        cliente_numero: Número de comitente
        cliente_nombre: Nombre del cliente
        periodo_inicio: Fecha inicio del período (ej: "Junio 1")
        periodo_fin: Fecha fin del período (ej: "Diciembre 12")
        anio: Año del reporte
        
    Returns:
        bytes del PDF generado
    """
    cliente_info = {}
    if cliente_numero:
        cliente_info['numero'] = cliente_numero
    if cliente_nombre:
        cliente_info['nombre'] = cliente_nombre
    
    exporter = ExcelToPdfExporter(excel_path, cliente_info or None)
    
    if periodo_inicio:
        exporter.periodo_inicio = periodo_inicio
    if periodo_fin:
        exporter.periodo_fin = periodo_fin
    if anio:
        exporter.anio = anio
    
    return exporter.export_to_pdf(output_path)


if __name__ == "__main__":
    # Test
    import sys
    
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    else:
        excel_path = "TEST_Merge_v8.xlsx"
    
    output_path = excel_path.replace('.xlsx', '.pdf')
    
    pdf_bytes = export_excel_to_pdf(
        excel_path,
        output_path,
        cliente_numero="12345",
        cliente_nombre="TEST CLIENT",
        periodo_inicio="Junio 1",
        periodo_fin="Diciembre 12",
        anio=2025
    )
    
    print(f"PDF generado: {output_path} ({len(pdf_bytes)} bytes)")
