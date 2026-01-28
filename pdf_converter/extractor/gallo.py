"""
Gallo Report Extractor.
Handles dynamic section detection and extraction for Gallo (tax summary) reports.
"""

import re
from typing import Optional
from rich.console import Console

try:
    from ..pdf.reader import PDFReader
    from ..llm.client import LLMClient, ChunkedExtractor
    from ..llm.prompts import GALLO_PROMPTS
    from .context import ExtractionContext, SectionBoundary, SectionDetector
    from .schemas import (
        GALLO_SCHEMAS,
        GALLO_SECTION_TO_SHEET,
        CATEGORIA_TO_SECTION,
        get_schema,
        get_numeric_fields,
    )
except ImportError:
    from pdf.reader import PDFReader
    from llm.client import LLMClient, ChunkedExtractor
    from llm.prompts import GALLO_PROMPTS
    from extractor.context import ExtractionContext, SectionBoundary, SectionDetector
    from extractor.schemas import (
        GALLO_SCHEMAS,
        GALLO_SECTION_TO_SHEET,
        CATEGORIA_TO_SECTION,
        get_schema,
        get_numeric_fields,
    )

console = Console()


class GalloExtractor:
    """
    Extractor for Gallo (Resumen Impositivo) reports.
    
    Features:
    - Dynamic section detection from Resultado Totales
    - Chunked extraction for large PDFs (50+ pages)
    - Context continuity for multi-page sections
    """
    
    # Sections that are always present
    FIXED_SECTIONS = ["resultado_totales", "posicion_inicial", "posicion_final"]
    
    # Sections that are transaction-based (use transacciones schema)
    TRANSACTION_SECTIONS = [
        "tit_privados_exentos",
        "tit_privados_exterior", 
        "renta_fija_pesos",
        "renta_fija_dolares",
        "fci",
        "opciones",
        "futuros",
    ]
    
    # Sections that are caución-based
    CAUCION_SECTIONS = ["cauciones_pesos", "cauciones_dolares"]
    
    def __init__(self, pdf_path: str, llm_client: LLMClient, max_pages_per_chunk: int = 5):
        """
        Initialize the Gallo extractor.
        
        Args:
            pdf_path: Path to the PDF file
            llm_client: LLMClient instance for AI extraction
            max_pages_per_chunk: Maximum pages to process per LLM call
        """
        self.pdf_reader = PDFReader(pdf_path)
        self.llm = llm_client
        self.chunked_extractor = ChunkedExtractor(llm_client, max_pages_per_chunk)
        self.context = ExtractionContext()
        self.section_detector = SectionDetector("gallo")
        
        # Detected sections from Resultado Totales
        self.detected_sections = set()
        
        # Extraction results
        self.results = {}
    
    def extract_all(self) -> dict:
        """
        Extract all sections from the Gallo PDF.
        
        Returns:
            Dictionary with section_key -> list of rows
        """
        console.print(f"[bold blue]📄 Processing Gallo report: {self.pdf_reader.path.name}[/bold blue]")
        console.print(f"   Total pages: {self.pdf_reader.total_pages}")
        
        # Step 1: Detect sections by scanning the PDF
        sections = self.section_detector.detect_sections(self.pdf_reader)
        console.print(f"   Sections found: {len(sections)}")
        
        # Step 2: Extract Resultado Totales first to know which sections have data
        console.print("\n[cyan]Extracting Resultado Totales...[/cyan]")
        resultado_totales = self._extract_resultado_totales(sections)
        self.results["resultado_totales"] = resultado_totales
        
        # Step 3: Determine which sections to extract based on categories
        self._detect_active_sections(resultado_totales)
        console.print(f"   Active sections: {', '.join(self.detected_sections)}")
        
        # Step 4: Extract each detected section
        for section_key in self.detected_sections:
            if section_key == "resultado_totales":
                continue  # Already extracted
            
            section = self.section_detector.find_section(sections, section_key)
            if not section:
                console.print(f"[yellow]⚠️ Section {section_key} not found in PDF[/yellow]")
                continue
            
            console.print(f"\n[cyan]Extracting {section_key} (pages {section.start_page + 1}-{section.end_page + 1})...[/cyan]")
            
            if section_key in self.TRANSACTION_SECTIONS:
                rows = self._extract_transacciones(section)
            elif section_key in self.CAUCION_SECTIONS:
                rows = self._extract_cauciones(section)
            elif section_key in ["posicion_inicial", "posicion_final"]:
                rows = self._extract_posicion(section)
            else:
                console.print(f"[yellow]Unknown section type: {section_key}[/yellow]")
                rows = []
            
            self.results[section_key] = rows
            console.print(f"   → {len(rows)} rows extracted")
        
        # Step 5: Extract positions if not already detected
        for pos_section in ["posicion_inicial", "posicion_final"]:
            if pos_section not in self.results:
                section = self.section_detector.find_section(sections, pos_section)
                if section:
                    console.print(f"\n[cyan]Extracting {pos_section}...[/cyan]")
                    rows = self._extract_posicion(section)
                    self.results[pos_section] = rows
                    console.print(f"   → {len(rows)} rows extracted")
        
        return self.results
    
    def _extract_resultado_totales(self, sections: list[SectionBoundary]) -> list[dict]:
        """Extract the Resultado Totales summary table."""
        section = self.section_detector.find_section(sections, "resultado_totales")
        
        if not section:
            # Try first 2 pages if section not explicitly found
            start_page = 0
            end_page = min(1, self.pdf_reader.total_pages - 1)
        else:
            start_page = section.start_page
            end_page = min(section.start_page + 1, section.end_page)
        
        text = self.pdf_reader.extract_pages_text(start_page, end_page)
        prompt = GALLO_PROMPTS["resultado_totales"].format(text=text)
        
        result = self.llm.extract(prompt, expected_keys=["resultado_totales"])
        
        if result.success and "resultado_totales" in result.data:
            return result.data["resultado_totales"]
        
        console.print(f"[red]Failed to extract Resultado Totales: {result.error}[/red]")
        return []
    
    def _detect_active_sections(self, resultado_totales: list[dict]):
        """
        Detect which sections have data based on Resultado Totales categories.
        """
        self.detected_sections = {"resultado_totales"}
        
        for row in resultado_totales:
            categoria = row.get("categoria", "").upper()
            
            # Skip TOTAL GENERAL
            if "TOTAL GENERAL" in categoria:
                continue
            
            # Extract base category (remove suffix in parentheses)
            match = re.match(r'^(.+?)\s*\(', categoria)
            if match:
                cat_base = match.group(1).strip()
            else:
                cat_base = categoria.strip()
            
            # Map to section key
            cat_lower = cat_base.lower()
            for pattern, section_key in CATEGORIA_TO_SECTION.items():
                if pattern in cat_lower or cat_lower in pattern:
                    self.detected_sections.add(section_key)
                    break
    
    def _extract_transacciones(self, section: SectionBoundary) -> list[dict]:
        """
        Extract a transaction-based section with chunking.
        Handles context continuity for especie names across pages.
        """
        self.context.reset_section(section.section_key)
        
        all_rows = []
        current_page = section.start_page
        
        while current_page <= section.end_page:
            chunk_end = min(
                current_page + self.chunked_extractor.max_pages_per_chunk - 1,
                section.end_page
            )
            
            console.print(f"  [dim]Pages {current_page + 1}-{chunk_end + 1}...[/dim]")
            
            # Get text for this chunk
            text = self.pdf_reader.extract_pages_text(current_page, chunk_end)
            
            # Build prompt with continuation context
            continuation = self.context.get_continuation_hint("gallo")
            
            # Map section to display name
            section_display_name = GALLO_SECTION_TO_SHEET.get(
                section.section_key, 
                section.section_key.upper().replace("_", " ")
            )
            
            prompt_template = GALLO_PROMPTS["transacciones"]
            prompt = prompt_template.format(
                section_name=section_display_name,
                section_key=section.section_key,
                text=text
            )
            
            if continuation:
                prompt = continuation + prompt
            
            # Extract
            result = self.llm.extract(prompt, expected_keys=[section.section_key])
            
            if result.success and section.section_key in result.data:
                rows = result.data[section.section_key]
                
                # Update context for next chunk
                self.context.update(rows)
                all_rows.extend(rows)
            else:
                console.print(f"  [yellow]Warning: Chunk extraction failed[/yellow]")
            
            # Move to next chunk
            current_page = chunk_end + 1
            self.context.add_processed_pages(list(range(current_page - 1, chunk_end + 1)))
        
        return all_rows
    
    def _extract_cauciones(self, section: SectionBoundary) -> list[dict]:
        """
        Extract a caución section.
        """
        self.context.reset_section(section.section_key)
        
        # Determine currency from section key
        if "dolares" in section.section_key:
            currency = "DOLARES"
            currency_key = "dolares"
        else:
            currency = "PESOS"
            currency_key = "pesos"
        
        all_rows = []
        current_page = section.start_page
        
        while current_page <= section.end_page:
            chunk_end = min(
                current_page + self.chunked_extractor.max_pages_per_chunk - 1,
                section.end_page
            )
            
            console.print(f"  [dim]Pages {current_page + 1}-{chunk_end + 1}...[/dim]")
            
            text = self.pdf_reader.extract_pages_text(current_page, chunk_end)
            
            prompt_template = GALLO_PROMPTS["cauciones"]
            prompt = prompt_template.format(
                currency=currency,
                currency_key=currency_key,
                text=text
            )
            
            continuation = self.context.get_continuation_hint("gallo")
            if continuation:
                prompt = continuation + prompt
            
            result = self.llm.extract(prompt, expected_keys=[f"cauciones_{currency_key}"])
            
            if result.success:
                key = f"cauciones_{currency_key}"
                if key in result.data:
                    rows = result.data[key]
                    self.context.update(rows)
                    all_rows.extend(rows)
            
            current_page = chunk_end + 1
        
        return all_rows
    
    def _extract_posicion(self, section: SectionBoundary) -> list[dict]:
        """
        Extract a position section (inicial or final).
        """
        # Determine position type
        if "inicial" in section.section_key:
            position_type = "INICIAL"
            position_key = "inicial"
        else:
            position_type = "FINAL"
            position_key = "final"
        
        all_rows = []
        current_page = section.start_page
        
        while current_page <= section.end_page:
            chunk_end = min(
                current_page + self.chunked_extractor.max_pages_per_chunk - 1,
                section.end_page
            )
            
            console.print(f"  [dim]Pages {current_page + 1}-{chunk_end + 1}...[/dim]")
            
            text = self.pdf_reader.extract_pages_text(current_page, chunk_end)
            
            prompt_template = GALLO_PROMPTS["posicion"]
            prompt = prompt_template.format(
                position_type=position_type,
                position_key=position_key,
                text=text
            )
            
            result = self.llm.extract(prompt, expected_keys=[f"posicion_{position_key}"])
            
            if result.success:
                key = f"posicion_{position_key}"
                if key in result.data:
                    all_rows.extend(result.data[key])
            
            current_page = chunk_end + 1
        
        return all_rows
    
    def get_results(self) -> dict:
        """Get the extraction results."""
        return self.results
    
    def get_detected_sections(self) -> set:
        """Get the set of detected sections."""
        return self.detected_sections
    
    def close(self):
        """Close the PDF reader."""
        self.pdf_reader.close()
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
