"""
Extraction context for maintaining state across page chunks.
Critical for multi-page sections where entity names span across pages.
"""

from dataclasses import dataclass, field
from typing import Optional, Any
import copy


@dataclass
class ExtractionContext:
    """
    Maintains extraction state across page chunks.
    Ensures continuity of entity names (e.g., especie) when processing
    multi-page sections in chunks.
    """
    
    # Current entity being processed (e.g., "especie" in Gallo)
    current_entity: dict = field(default_factory=dict)
    
    # Last N rows for context overlap detection
    last_rows: list = field(default_factory=list)
    max_context_rows: int = 5
    
    # Accumulated results for this section
    section_data: list = field(default_factory=list)
    
    # Metadata
    current_section: str = ""
    pages_processed: list = field(default_factory=list)
    
    def update(self, rows: list):
        """
        Update context after processing a chunk.
        
        Args:
            rows: List of row dictionaries from the chunk
        """
        if not rows:
            return
        
        # Track last entity for continuity
        # Look for the last non-total row with an entity name
        for row in reversed(rows):
            entity_name = row.get("especie") or row.get("instrumento")
            tipo_fila = str(row.get("tipo_fila", "")).lower()
            
            if entity_name and "total" not in tipo_fila:
                self.current_entity = {
                    "cod_especie": row.get("cod_especie", ""),
                    "especie": entity_name,
                    "cod_instrumento": row.get("cod_instrumento", ""),
                    "instrumento": entity_name,
                }
                break
        
        # Keep last N rows for overlap detection
        self.last_rows = rows[-self.max_context_rows:]
        
        # Accumulate data
        self.section_data.extend(rows)
    
    def get_continuation_hint(self, report_type: str = "gallo") -> str:
        """
        Generate prompt hint for next chunk.
        Injects the last entity context into the prompt.
        
        Args:
            report_type: Either "gallo" or "visual"
        
        Returns:
            Continuation context string to prepend to prompt
        """
        if not self.current_entity:
            return ""
        
        if report_type == "gallo":
            cod_especie = self.current_entity.get("cod_especie", "")
            especie = self.current_entity.get("especie", "")
            
            if especie:
                return f"""CONTEXTO DE CONTINUIDAD:
La última especie procesada fue:
  cod_especie: "{cod_especie}"
  especie: "{especie}"

Si la página comienza con transacciones sin encabezado de especie,
asigna estos valores a esas filas.

"""
        else:  # visual
            instrumento = self.current_entity.get("instrumento", "")
            cod_instrumento = self.current_entity.get("cod_instrumento", "")
            
            if instrumento:
                return f"""CONTEXTO DE CONTINUIDAD:
El último instrumento procesado fue:
  cod_instrumento: "{cod_instrumento}"
  instrumento: "{instrumento}"

Si la página continúa con datos del mismo instrumento, mantén estos valores.

"""
        
        return ""
    
    def get_dedup_keys(self) -> list:
        """
        Get the last few row signatures for deduplication.
        Used to detect and remove duplicates from chunk overlap.
        """
        keys = []
        for row in self.last_rows:
            # Create a signature from key fields
            sig = "|".join(str(row.get(k, "")) for k in [
                "especie", "instrumento", "fecha", "concertacion",
                "operacion", "tipo_operacion", "numero", "cantidad"
            ])
            keys.append(sig)
        return keys
    
    def reset_section(self, section_name: str):
        """Reset context for a new section."""
        self.current_entity = {}
        self.last_rows = []
        self.section_data = []
        self.current_section = section_name
        self.pages_processed = []
    
    def add_processed_pages(self, pages: list):
        """Track which pages have been processed."""
        self.pages_processed.extend(pages)
    
    def copy(self) -> "ExtractionContext":
        """Create a deep copy of this context."""
        return copy.deepcopy(self)
    
    def get_stats(self) -> dict:
        """Get extraction statistics."""
        return {
            "section": self.current_section,
            "pages_processed": len(self.pages_processed),
            "rows_extracted": len(self.section_data),
            "current_entity": self.current_entity.get("especie") or self.current_entity.get("instrumento", ""),
        }


@dataclass
class SectionBoundary:
    """Represents the page boundaries of a section in the PDF."""
    name: str
    start_page: int  # 0-indexed
    end_page: int    # 0-indexed, inclusive
    section_key: str  # Key used in JSON output
    
    def __post_init__(self):
        if self.end_page < self.start_page:
            self.end_page = self.start_page
    
    @property
    def page_count(self) -> int:
        return self.end_page - self.start_page + 1
    
    def contains_page(self, page_num: int) -> bool:
        return self.start_page <= page_num <= self.end_page


class SectionDetector:
    """
    Detects section boundaries in PDFs by looking for section headers.
    """
    
    # Section header patterns for Gallo reports
    # These patterns identify DETAIL sections, not the summary in "RESULTADOS TOTALES"
    GALLO_HEADERS = {
        # Resultados Totales is on page 2, but we detect by looking for standalone "RESULTADOS TOTALES"
        "RESULTADOS TOTALES": "resultado_totales",
        # Detail sections (exclude if followed by values like "(Renta)" on same line - that's summary)
        "TIT.PRIVADOS EXENTOS": "tit_privados_exentos",
        "TIT PRIVADOS EXENTOS": "tit_privados_exentos",
        "TIT.PRIVADOS DEL EXTERIOR": "tit_privados_exterior",
        "RENTA FIJA EN PESOS": "renta_fija_pesos",
        "RENTA FIJA EN DOLARES": "renta_fija_dolares",
        "RENTA FIJA EN DÓLARES": "renta_fija_dolares",
        "CAUCIONES EN PESOS": "cauciones_pesos",
        "CAUCIONES EN DOLARES": "cauciones_dolares",
        "CAUCIONES EN DÓLARES": "cauciones_dolares",
        "FCI": "fci",
        "FONDOS COMUNES": "fci",
        "OPCIONES": "opciones",
        "FUTUROS": "futuros",
        # Position sections - various formats
        "POSICION AL 01/01": "posicion_inicial",
        "POSICION AL 31/12": "posicion_final",
        "POSICIÓN AL 01/01": "posicion_inicial",
        "POSICIÓN AL 31/12": "posicion_final",
        "POSICION INICIAL": "posicion_inicial",
        "POSICIÓN INICIAL": "posicion_inicial",
        "POSICION FINAL": "posicion_final",
        "POSICIÓN FINAL": "posicion_final",
    }
    
    # Section header patterns for Visual reports
    VISUAL_HEADERS = {
        "BOLETOS": "boletos",
        "RESULTADO DE VENTAS EN PESOS": "resultado_ventas_ars",
        "RESULTADO DE VENTAS EN ARS": "resultado_ventas_ars",
        "RESULTADO DE VENTAS EN DOLARES": "resultado_ventas_usd",
        "RESULTADO DE VENTAS EN DÓLARES": "resultado_ventas_usd",
        "RESULTADO DE VENTAS EN USD": "resultado_ventas_usd",
        "RENTAS Y DIVIDENDOS EN PESOS": "rentas_dividendos_ars",
        "RENTAS Y DIVIDENDOS EN ARS": "rentas_dividendos_ars",
        "RENTAS DIVIDENDOS ARS": "rentas_dividendos_ars",
        "RENTAS Y DIVIDENDOS EN DOLARES": "rentas_dividendos_usd",
        "RENTAS Y DIVIDENDOS EN DÓLARES": "rentas_dividendos_usd",
        "RENTAS DIVIDENDOS USD": "rentas_dividendos_usd",
        "RESUMEN": "resumen",
        "POSICION DE TITULOS": "posicion_titulos",
        "POSICIÓN DE TÍTULOS": "posicion_titulos",
        "POSICION TITULOS": "posicion_titulos",
    }
    
    def __init__(self, report_type: str):
        """
        Initialize detector for a specific report type.
        
        Args:
            report_type: Either "gallo" or "visual"
        """
        self.report_type = report_type
        self.headers = self.GALLO_HEADERS if report_type == "gallo" else self.VISUAL_HEADERS
    
    def detect_sections(self, pdf_reader) -> list[SectionBoundary]:
        """
        Detect all section boundaries in the PDF.
        
        Args:
            pdf_reader: PDFReader instance
        
        Returns:
            List of SectionBoundary objects
        """
        sections = []
        found_sections = {}  # section_key -> start_page
        
        total_pages = pdf_reader.get_page_count()
        
        for page_num in range(total_pages):
            text = pdf_reader.extract_page_text(page_num)
            text_upper = text.upper()
            lines = text.split('\n')
            
            for header, section_key in self.headers.items():
                if header in text_upper:
                    # For detail sections, check if header is standalone (not part of summary)
                    is_summary_line = False
                    for line in lines:
                        line_upper = line.upper()
                        if header in line_upper:
                            # Summary lines have numeric values on same line like "(Renta) 1,799.21"
                            # or are in a line that starts with the category
                            if "(RENTA)" in line_upper and any(c.isdigit() for c in line):
                                is_summary_line = True
                            elif "(ENAJENACION)" in line_upper and any(c.isdigit() for c in line):
                                is_summary_line = True
                            # Check if it's a standalone section header (no numbers or at start)
                            elif header in ["RESULTADOS TOTALES", "POSICION AL 01/01", "POSICION AL 31/12"]:
                                is_summary_line = False
                            else:
                                # If line is short and is just the header, it's a detail section
                                stripped = line.strip()
                                if len(stripped) < 50 and not any(c.isdigit() for c in stripped[-10:] if c != '/'):
                                    is_summary_line = False
                                    break
                    
                    if not is_summary_line and section_key not in found_sections:
                        found_sections[section_key] = {
                            "name": header,
                            "start_page": page_num,
                            "section_key": section_key
                        }
        
        # Convert to boundaries (each section ends when the next begins or at EOF)
        sorted_sections = sorted(found_sections.values(), key=lambda x: x["start_page"])
        
        for i, section_info in enumerate(sorted_sections):
            if i + 1 < len(sorted_sections):
                end_page = sorted_sections[i + 1]["start_page"] - 1
            else:
                end_page = total_pages - 1
            
            sections.append(SectionBoundary(
                name=section_info["name"],
                start_page=section_info["start_page"],
                end_page=end_page,
                section_key=section_info["section_key"]
            ))
        
        return sections
    
    def find_section(self, sections: list[SectionBoundary], section_key: str) -> Optional[SectionBoundary]:
        """Find a specific section by key."""
        for section in sections:
            if section.section_key == section_key:
                return section
        return None
