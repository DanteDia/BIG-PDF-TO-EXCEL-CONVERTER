"""
Módulo para leer valores calculados de Excel usando Datalab API.

Esto permite obtener valores de fórmulas evaluadas en entornos donde
no hay Excel instalado (ej: Streamlit Cloud, Linux).

Parsea el markdown completo de Datalab para extraer todas las secciones:
- Boletos (Compras/Ventas)
- Ventas ARS
- Ventas USD
- Rentas y Dividendos ARS
- Rentas y Dividendos USD
- Cauciones
- Resumen
"""

import httpx
import time
import re
from typing import Dict, Optional, Any, Tuple, List
from pathlib import Path


class DatalabExcelReader:
    """
    Lee valores de Excel usando Datalab API para evaluar fórmulas.
    Parsea el markdown para extraer todas las secciones del reporte.
    """
    
    API_URL = "https://www.datalab.to/api/v1/marker"
    
    def __init__(self, api_key: str = None):
        """
        Args:
            api_key: API key de Datalab (opcional si solo se va a parsear markdown)
        """
        self.api_key = api_key or ""
        self._markdown = None
        self._parsed_data = None
    
    def convert_to_markdown(self, excel_path: str, max_wait_seconds: int = 300) -> Optional[str]:
        """
        Convierte Excel a Markdown usando Datalab API.
        
        Args:
            excel_path: Ruta al archivo Excel
            max_wait_seconds: Tiempo máximo de espera en segundos
            
        Returns:
            Markdown content o None si falla
        """
        path = Path(excel_path)
        if not path.exists():
            print(f"[ERROR] Archivo no existe: {excel_path}")
            return None
        
        print(f"[INFO] Subiendo archivo a Datalab API...")
        
        # Subir archivo
        with open(path, "rb") as f:
            files = {"file": (path.name, f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            headers = {"X-Api-Key": self.api_key}
            
            response = httpx.post(
                self.API_URL,
                files=files,
                headers=headers,
                timeout=60,
                verify=False  # Workaround para SSL issues en algunos entornos
            )
        
        if response.status_code != 200:
            print(f"[ERROR] Datalab API error: {response.status_code}")
            return None
        
        data = response.json()
        if not data.get("success"):
            print(f"[ERROR] Datalab API failed: {data}")
            return None
        
        check_url = data.get("request_check_url")
        if not check_url:
            print("[ERROR] No check URL in response")
            return None
        
        # Polling para esperar resultado
        headers = {"X-Api-Key": self.api_key}
        start_time = time.time()
        poll_interval = 3
        attempt = 0
        
        print(f"[INFO] Esperando procesamiento de Datalab...")
        
        while True:
            elapsed = time.time() - start_time
            if elapsed > max_wait_seconds:
                print(f"[ERROR] Timeout esperando resultado de Datalab")
                return None
            
            time.sleep(poll_interval)
            attempt += 1
            
            poll_response = httpx.get(check_url, headers=headers, timeout=30, verify=False)
            poll_data = poll_response.json()
            status = poll_data.get("status")
            
            if attempt % 10 == 0:
                print(f"[INFO] Intento {attempt}, status: {status}")
            
            if status == "complete":
                self._markdown = poll_data.get("markdown", "")
                print(f"[INFO] Markdown recibido ({len(self._markdown):,} caracteres)")
                return self._markdown
            elif status == "error":
                print(f"[ERROR] Datalab processing error: {poll_data}")
                return None
            # else: sigue en "processing", continuar polling
    
    def parse_all_sections(self, markdown: str = None) -> Dict[str, Any]:
        """
        Parsea todas las secciones del markdown.
        
        Returns:
            Dict con todas las secciones parseadas:
            - boletos: List[Dict]
            - ventas_ars: List[Dict]
            - ventas_usd: List[Dict]
            - rentas_dividendos_ars: List[Dict]
            - rentas_dividendos_usd: List[Dict]
            - cauciones: List[Dict]
            - resumen: Dict
        """
        if markdown:
            self._markdown = markdown
        
        if not self._markdown:
            return {}
        
        # Resetear contador de secciones de rentas
        self._rentas_section_count = 0
        
        lines = self._markdown.split('\n')
        
        result = {
            'boletos': [],
            'ventas_ars': [],
            'ventas_usd': [],
            'rentas_dividendos_ars': [],
            'rentas_dividendos_usd': [],
            'cauciones': [],
            'resumen': {'ventas_ars': 0.0, 'ventas_usd': 0.0, 'total_ars': 0.0, 'total_usd': 0.0}
        }
        
        # Identificar secciones por sus headers
        current_section = None
        current_headers = []
        
        for i, line in enumerate(lines):
            line = line.strip()
            if not line or line.startswith('|---'):
                continue
            
            # Detectar headers de tabla (contienen <b>)
            if '<b>' in line and '</b>' in line:
                headers = self._parse_headers(line)
                if headers:
                    current_headers = headers
                    current_section = self._identify_section(headers)
                    continue
            
            # Si tenemos una sección activa, parsear datos
            if current_section and current_headers and line.startswith('|'):
                row_data = self._parse_data_row(line, current_headers)
                
                if current_section == 'boletos':
                    result['boletos'].append(row_data)
                elif current_section == 'ventas_ars':
                    result['ventas_ars'].append(row_data)
                elif current_section == 'ventas_usd':
                    result['ventas_usd'].append(row_data)
                elif current_section == 'rentas_ars':
                    result['rentas_dividendos_ars'].append(row_data)
                elif current_section == 'rentas_usd':
                    result['rentas_dividendos_usd'].append(row_data)
                elif current_section == 'cauciones':
                    result['cauciones'].append(row_data)
                elif current_section == 'resumen':
                    # Parsear resumen especialmente
                    self._parse_resumen_row(line, result['resumen'])
        
        self._parsed_data = result
        return result
    
    def _parse_headers(self, line: str) -> List[str]:
        """Extrae los nombres de las columnas de una línea de headers."""
        headers = []
        # Extraer contenido entre <b> y </b>
        pattern = r'<b>([^<]+)</b>'
        matches = re.findall(pattern, line)
        for match in matches:
            # Limpiar el nombre
            header = match.strip()
            headers.append(header)
        return headers
    
    def _identify_section(self, headers: List[str]) -> Optional[str]:
        """Identifica qué sección es basándose en los headers."""
        headers_lower = [h.lower() for h in headers]
        headers_str = ' '.join(headers_lower)
        
        # Resumen - tiene Moneda, Ventas, FCI, Opciones, etc.
        if 'moneda' in headers_str and 'ventas' in headers_str and 'fci' in headers_str:
            return 'resumen'
        
        # Ventas ARS - tiene "Resultado Calculado(final)" pero NO "Bruto en USD"
        if 'resultado calculado(final)' in headers_str and 'bruto en usd' not in headers_str:
            return 'ventas_ars'
        
        # Ventas USD - tiene "Bruto en USD" y "Resultado Calculado(final)"
        if 'bruto en usd' in headers_str and 'resultado calculado(final)' in headers_str:
            return 'ventas_usd'
        
        # Rentas y Dividendos - tiene "Categoría" y "Importe"
        # Hay dos tablas: una para ARS (primera) y una para USD (segunda)
        if 'categoría' in headers_str and 'importe' in headers_str:
            # La primera tabla es ARS (origen contiene "Gallo" o "Visual-Rentas Dividendos ARS")
            # La segunda tabla es USD (origen contiene "Visual-Rentas Dividendos USD")
            # Usamos un contador interno para distinguir
            if not hasattr(self, '_rentas_section_count'):
                self._rentas_section_count = 0
            self._rentas_section_count += 1
            if self._rentas_section_count == 1:
                return 'rentas_ars'
            else:
                return 'rentas_usd'
        
        # Cauciones - tiene "Tasa (%)" y "Costo Financiero"
        if 'tasa (%)' in headers_str or 'costo financiero' in headers_str:
            return 'cauciones'
        
        # Boletos - tiene Concertación, Liquidación, Nro. Boleto y "Neto Calculado"
        if 'concertación' in headers_str and 'nro. boleto' in headers_str:
            if 'neto calculado' in headers_str and 'resultado calculado' not in headers_str:
                return 'boletos'
        
        return None
    
    def _parse_data_row(self, line: str, headers: List[str]) -> Dict[str, Any]:
        """Parsea una fila de datos y la mapea a los headers."""
        parts = line.split('|')
        values = []
        
        for part in parts:
            part = part.strip()
            values.append(part)
        
        # Remover elementos vacíos del inicio y fin
        while values and values[0] == '':
            values.pop(0)
        while values and values[-1] == '':
            values.pop()
        
        # Crear diccionario mapeando headers a valores
        row = {}
        for i, header in enumerate(headers):
            if i < len(values):
                value = values[i]
                # Intentar convertir a número si es posible
                row[header] = self._convert_value(value)
            else:
                row[header] = None
        
        return row
    
    def _convert_value(self, value: str) -> Any:
        """Convierte un valor string a número si es posible."""
        if not value or value == '':
            return None
        
        # Limpiar espacios
        value = value.strip()
        
        # Intentar como float
        try:
            # Verificar si parece un número
            if re.match(r'^-?\d+\.?\d*$', value):
                return float(value)
        except:
            pass
        
        return value
    
    def _parse_resumen_row(self, line: str, resumen: Dict):
        """Parsea una fila del resumen (ARS o USD)."""
        line_stripped = line.strip()
        
        # Buscar fila ARS
        if line_stripped.startswith('| ARS') and '|' in line_stripped[5:]:
            pipe_count = line_stripped.count('|')
            if pipe_count >= 10:
                values = self._parse_numeric_values(line_stripped)
                if len(values) >= 2:
                    resumen['ventas_ars'] = values[0]
                    resumen['total_ars'] = values[-1]
        
        # Buscar fila USD
        elif line_stripped.startswith('| USD') and '|' in line_stripped[5:]:
            pipe_count = line_stripped.count('|')
            if pipe_count >= 10:
                values = self._parse_numeric_values(line_stripped)
                if len(values) >= 2:
                    resumen['ventas_usd'] = values[0]
                    resumen['total_usd'] = values[-1]
    
    def _parse_numeric_values(self, line: str) -> List[float]:
        """Extrae solo valores numéricos de una línea."""
        values = []
        parts = line.split('|')
        
        for part in parts:
            part = part.strip()
            if not part:
                continue
            
            if part in ['ARS', 'USD']:
                continue
            
            try:
                if re.match(r'^-?\d+\.?\d*$', part):
                    values.append(float(part))
                elif part == '0':
                    values.append(0.0)
            except:
                pass
        
        return values
    
    def get_boletos(self) -> List[Dict]:
        """Retorna los boletos parseados."""
        if not self._parsed_data:
            self.parse_all_sections()
        return self._parsed_data.get('boletos', [])
    
    def get_ventas_ars(self) -> List[Dict]:
        """Retorna las ventas ARS parseadas."""
        if not self._parsed_data:
            self.parse_all_sections()
        return self._parsed_data.get('ventas_ars', [])
    
    def get_ventas_usd(self) -> List[Dict]:
        """Retorna las ventas USD parseadas."""
        if not self._parsed_data:
            self.parse_all_sections()
        return self._parsed_data.get('ventas_usd', [])
    
    def get_rentas_dividendos_ars(self) -> List[Dict]:
        """Retorna las rentas y dividendos ARS parseadas."""
        if not self._parsed_data:
            self.parse_all_sections()
        return self._parsed_data.get('rentas_dividendos_ars', [])
    
    def get_rentas_dividendos_usd(self) -> List[Dict]:
        """Retorna las rentas y dividendos USD parseadas."""
        if not self._parsed_data:
            self.parse_all_sections()
        return self._parsed_data.get('rentas_dividendos_usd', [])
    
    def get_cauciones(self) -> List[Dict]:
        """Retorna las cauciones parseadas."""
        if not self._parsed_data:
            self.parse_all_sections()
        return self._parsed_data.get('cauciones', [])
    
    def get_resumen(self) -> Dict:
        """Retorna el resumen parseado."""
        if not self._parsed_data:
            self.parse_all_sections()
        return self._parsed_data.get('resumen', {})


def read_excel_with_datalab(excel_path: str, api_key: str) -> Tuple[Optional['DatalabExcelReader'], Optional[str]]:
    """
    Función de conveniencia para leer Excel con Datalab.
    
    Args:
        excel_path: Ruta al Excel
        api_key: API key de Datalab
        
    Returns:
        Tuple de (DatalabExcelReader con datos parseados, markdown raw)
    """
    reader = DatalabExcelReader(api_key)
    markdown = reader.convert_to_markdown(excel_path)
    
    if not markdown:
        return None, None
    
    # Parsear todas las secciones
    reader.parse_all_sections(markdown)
    
    return reader, markdown


if __name__ == "__main__":
    # Test con markdown existente
    import sys
    
    # Si hay un archivo markdown existente, usarlo
    if len(sys.argv) >= 2 and sys.argv[1].endswith('.md'):
        md_path = sys.argv[1]
        print(f"Leyendo markdown desde {md_path}...")
        with open(md_path, 'r', encoding='utf-8') as f:
            markdown = f.read()
        
        reader = DatalabExcelReader("dummy")
        data = reader.parse_all_sections(markdown)
        
        print("\n=== SECCIONES ENCONTRADAS ===")
        print(f"Boletos: {len(data['boletos'])} filas")
        print(f"Ventas ARS: {len(data['ventas_ars'])} filas")
        print(f"Ventas USD: {len(data['ventas_usd'])} filas")
        print(f"Rentas/Div ARS: {len(data['rentas_dividendos_ars'])} filas")
        print(f"Rentas/Div USD: {len(data['rentas_dividendos_usd'])} filas")
        print(f"Cauciones: {len(data['cauciones'])} filas")
        
        print("\n=== RESUMEN ===")
        print(f"Ventas ARS: {data['resumen']['ventas_ars']:,.2f}")
        print(f"Ventas USD: {data['resumen']['ventas_usd']:,.2f}")
        print(f"Total ARS: {data['resumen']['total_ars']:,.2f}")
        print(f"Total USD: {data['resumen']['total_usd']:,.2f}")
        
        # Mostrar ejemplo de cada sección
        if data['boletos']:
            print("\n=== EJEMPLO BOLETO ===")
            for k, v in list(data['boletos'][0].items())[:10]:
                print(f"  {k}: {v}")
        
        if data['ventas_ars']:
            print("\n=== EJEMPLO VENTAS ARS ===")
            for k, v in list(data['ventas_ars'][0].items())[:15]:
                print(f"  {k}: {v}")
        
        if data['ventas_usd']:
            print("\n=== EJEMPLO VENTAS USD ===")
            for k, v in list(data['ventas_usd'][0].items())[:15]:
                print(f"  {k}: {v}")
    
    elif len(sys.argv) >= 3:
        excel_path = sys.argv[1]
        api_key = sys.argv[2]
        
        reader, markdown = read_excel_with_datalab(excel_path, api_key)
        
        if reader:
            resumen = reader.get_resumen()
            print("\n=== VALORES DEL RESUMEN ===")
            print(f"Ventas ARS: {resumen['ventas_ars']:,.2f}")
            print(f"Ventas USD: {resumen['ventas_usd']:,.2f}")
            print(f"Total ARS: {resumen['total_ars']:,.2f}")
            print(f"Total USD: {resumen['total_usd']:,.2f}")
    else:
        print("Uso:")
        print("  python datalab_excel_reader.py <markdown.md>")
        print("  python datalab_excel_reader.py <excel_path> <api_key>")
