"""
Visual Report Extractor.
Handles extraction for Visual (broker) reports with fixed sections.
"""

from typing import Optional
from rich.console import Console

try:
    from ..pdf.reader import PDFReader
    from ..llm.client import LLMClient, ChunkedExtractor
    from ..llm.prompts import VISUAL_PROMPTS
    from .context import ExtractionContext, SectionBoundary, SectionDetector
    from .schemas import (
        VISUAL_SCHEMAS,
        VISUAL_SECTION_TO_SHEET,
        get_schema,
    )
except ImportError:
    from pdf.reader import PDFReader
    from llm.client import LLMClient, ChunkedExtractor
    from llm.prompts import VISUAL_PROMPTS
    from extractor.context import ExtractionContext, SectionBoundary, SectionDetector
    from extractor.schemas import (
        VISUAL_SCHEMAS,
        VISUAL_SECTION_TO_SHEET,
        get_schema,
    )

console = Console()


class VisualExtractor:
    """
    Extractor for Visual (broker) reports.
    
    Features:
    - Fixed sections (Boletos, Resultado Ventas, Rentas Dividendos, Resumen, Posicion Titulos)
    - Chunked extraction for large PDFs
    - Handles parenthesis negatives
    """
    
    # Fixed sections in order they typically appear
    SECTIONS_ORDER = [
        "resumen",
        "boletos",
        "resultado_ventas_ars",
        "resultado_ventas_usd",
        "rentas_dividendos_ars",
        "rentas_dividendos_usd",
        "posicion_titulos",
    ]
    
    def __init__(self, pdf_path: str, llm_client: LLMClient, max_pages_per_chunk: int = 5):
        """
        Initialize the Visual extractor.
        
        Args:
            pdf_path: Path to the PDF file
            llm_client: LLMClient instance for AI extraction
            max_pages_per_chunk: Maximum pages to process per LLM call
        """
        self.pdf_reader = PDFReader(pdf_path)
        self.llm = llm_client
        self.chunked_extractor = ChunkedExtractor(llm_client, max_pages_per_chunk)
        self.context = ExtractionContext()
        self.section_detector = SectionDetector("visual")
        
        # Extraction results
        self.results = {}
    
    def extract_all(self) -> dict:
        """
        Extract all sections from the Visual PDF.
        
        Returns:
            Dictionary with section_key -> list of rows
        """
        console.print(f"[bold blue]📄 Processing Visual report: {self.pdf_reader.path.name}[/bold blue]")
        console.print(f"   Total pages: {self.pdf_reader.total_pages}")
        
        # Detect sections
        sections = self.section_detector.detect_sections(self.pdf_reader)
        console.print(f"   Sections found: {len(sections)}")
        
        for section_key in self.SECTIONS_ORDER:
            section = self.section_detector.find_section(sections, section_key)
            
            if not section:
                console.print(f"[dim]   Section {section_key} not found in PDF[/dim]")
                continue
            
            console.print(f"\n[cyan]Extracting {section_key} (pages {section.start_page + 1}-{section.end_page + 1})...[/cyan]")
            
            if section_key == "resumen":
                rows = self._extract_resumen(section)
            elif section_key == "boletos":
                rows = self._extract_boletos(section)
            elif section_key.startswith("resultado_ventas"):
                rows = self._extract_resultado_ventas(section)
            elif section_key.startswith("rentas_dividendos"):
                rows = self._extract_rentas_dividendos(section)
            elif section_key == "posicion_titulos":
                rows = self._extract_posicion_titulos(section)
            else:
                console.print(f"[yellow]Unknown section type: {section_key}[/yellow]")
                rows = []
            
            self.results[section_key] = rows
            console.print(f"   → {len(rows)} rows extracted")
        
        return self.results
    
    def _extract_resumen(self, section: SectionBoundary) -> list[dict]:
        """Extract the Resumen summary table."""
        # Resumen is typically on one page
        text = self.pdf_reader.extract_pages_text(section.start_page, section.start_page)
        
        prompt = VISUAL_PROMPTS["resumen"].format(text=text)
        result = self.llm.extract(prompt, expected_keys=["resumen"])
        
        if result.success and "resumen" in result.data:
            return result.data["resumen"]
        
        console.print(f"[red]Failed to extract Resumen: {result.error}[/red]")
        return []
    
    def _extract_boletos(self, section: SectionBoundary) -> list[dict]:
        """Extract boletos with chunking for multi-page sections."""
        self.context.reset_section("boletos")
        
        all_rows = []
        current_page = section.start_page
        
        while current_page <= section.end_page:
            chunk_end = min(
                current_page + self.chunked_extractor.max_pages_per_chunk - 1,
                section.end_page
            )
            
            console.print(f"  [dim]Pages {current_page + 1}-{chunk_end + 1}...[/dim]")
            
            text = self.pdf_reader.extract_pages_text(current_page, chunk_end)
            prompt = VISUAL_PROMPTS["boletos"].format(text=text)
            
            result = self.llm.extract(prompt, expected_keys=["boletos"])
            
            if result.success and "boletos" in result.data:
                rows = result.data["boletos"]
                self.context.update(rows)
                all_rows.extend(rows)
            
            current_page = chunk_end + 1
        
        return all_rows
    
    def _extract_resultado_ventas(self, section: SectionBoundary) -> list[dict]:
        """Extract resultado de ventas section."""
        # Determine currency from section key
        if "usd" in section.section_key:
            currency = "DOLARES"
            currency_key = "usd"
        else:
            currency = "PESOS"
            currency_key = "ars"
        
        self.context.reset_section(section.section_key)
        
        all_rows = []
        current_page = section.start_page
        
        while current_page <= section.end_page:
            chunk_end = min(
                current_page + self.chunked_extractor.max_pages_per_chunk - 1,
                section.end_page
            )
            
            console.print(f"  [dim]Pages {current_page + 1}-{chunk_end + 1}...[/dim]")
            
            text = self.pdf_reader.extract_pages_text(current_page, chunk_end)
            
            prompt_template = VISUAL_PROMPTS["resultado_ventas"]
            prompt = prompt_template.format(
                currency=currency,
                currency_key=currency_key,
                text=text
            )
            
            continuation = self.context.get_continuation_hint("visual")
            if continuation:
                prompt = continuation + prompt
            
            result = self.llm.extract(prompt, expected_keys=[f"resultado_ventas_{currency_key}"])
            
            if result.success:
                key = f"resultado_ventas_{currency_key}"
                if key in result.data:
                    rows = result.data[key]
                    self.context.update(rows)
                    all_rows.extend(rows)
            
            current_page = chunk_end + 1
        
        return all_rows
    
    def _extract_rentas_dividendos(self, section: SectionBoundary) -> list[dict]:
        """Extract rentas y dividendos section."""
        # Determine currency from section key
        if "usd" in section.section_key:
            currency = "DOLARES"
            currency_key = "usd"
        else:
            currency = "PESOS"
            currency_key = "ars"
        
        self.context.reset_section(section.section_key)
        
        all_rows = []
        current_page = section.start_page
        
        while current_page <= section.end_page:
            chunk_end = min(
                current_page + self.chunked_extractor.max_pages_per_chunk - 1,
                section.end_page
            )
            
            console.print(f"  [dim]Pages {current_page + 1}-{chunk_end + 1}...[/dim]")
            
            text = self.pdf_reader.extract_pages_text(current_page, chunk_end)
            
            prompt_template = VISUAL_PROMPTS["rentas_dividendos"]
            prompt = prompt_template.format(
                currency=currency,
                currency_key=currency_key,
                text=text
            )
            
            continuation = self.context.get_continuation_hint("visual")
            if continuation:
                prompt = continuation + prompt
            
            result = self.llm.extract(prompt, expected_keys=[f"rentas_dividendos_{currency_key}"])
            
            if result.success:
                key = f"rentas_dividendos_{currency_key}"
                if key in result.data:
                    rows = result.data[key]
                    self.context.update(rows)
                    all_rows.extend(rows)
            
            current_page = chunk_end + 1
        
        return all_rows
    
    def _extract_posicion_titulos(self, section: SectionBoundary) -> list[dict]:
        """Extract posicion de titulos section."""
        all_rows = []
        current_page = section.start_page
        
        while current_page <= section.end_page:
            chunk_end = min(
                current_page + self.chunked_extractor.max_pages_per_chunk - 1,
                section.end_page
            )
            
            console.print(f"  [dim]Pages {current_page + 1}-{chunk_end + 1}...[/dim]")
            
            text = self.pdf_reader.extract_pages_text(current_page, chunk_end)
            prompt = VISUAL_PROMPTS["posicion_titulos"].format(text=text)
            
            result = self.llm.extract(prompt, expected_keys=["posicion_titulos"])
            
            if result.success and "posicion_titulos" in result.data:
                all_rows.extend(result.data["posicion_titulos"])
            
            current_page = chunk_end + 1
        
        return all_rows
    
    def get_results(self) -> dict:
        """Get the extraction results."""
        return self.results
    
    def close(self):
        """Close the PDF reader."""
        self.pdf_reader.close()
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
