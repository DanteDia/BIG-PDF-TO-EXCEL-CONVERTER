"""
PDF Reader with native text extraction and OCR fallback.
Handles both text-based and scanned PDFs.
"""

import os
from pathlib import Path
from typing import Optional
import pdfplumber
import fitz  # PyMuPDF
from rich.console import Console

console = Console()


class PDFReader:
    """
    PDF Reader that extracts text from PDFs.
    Automatically detects if OCR is needed for scanned documents.
    """
    
    def __init__(self, pdf_path: str):
        self.path = Path(pdf_path)
        if not self.path.exists():
            raise FileNotFoundError(f"PDF not found: {pdf_path}")
        
        self.doc = fitz.open(str(self.path))
        self.total_pages = len(self.doc)
        self.is_ocr_needed = self._detect_ocr_need()
        
        if self.is_ocr_needed:
            console.print(f"[yellow]⚠️ OCR detected as needed for {self.path.name}[/yellow]")
    
    def _detect_ocr_need(self) -> bool:
        """
        Detect if PDF is image-based (needs OCR).
        Checks first 3 pages for text density.
        """
        pages_to_check = min(3, self.total_pages)
        
        for page_num in range(pages_to_check):
            page = self.doc[page_num]
            text = page.get_text()
            
            # If very little text but has images, needs OCR
            if len(text.strip()) < 100:
                images = page.get_images()
                if images:
                    return True
        
        return False
    
    def extract_page_text(self, page_num: int) -> str:
        """
        Extract text from a single page.
        Uses pdfplumber for better table extraction.
        """
        if page_num < 0 or page_num >= self.total_pages:
            raise ValueError(f"Page {page_num} out of range (0-{self.total_pages-1})")
        
        if self.is_ocr_needed:
            return self._ocr_page(page_num)
        else:
            return self._extract_native(page_num)
    
    def _extract_native(self, page_num: int) -> str:
        """Native text extraction with pdfplumber for better table handling."""
        try:
            with pdfplumber.open(str(self.path)) as pdf:
                page = pdf.pages[page_num]
                
                # Try table extraction first
                tables = page.extract_tables()
                if tables:
                    return self._tables_to_text(tables)
                
                # Fall back to raw text
                text = page.extract_text() or ""
                return text
        except Exception as e:
            console.print(f"[red]Error extracting page {page_num}: {e}[/red]")
            # Fallback to PyMuPDF
            return self.doc[page_num].get_text()
    
    def _ocr_page(self, page_num: int) -> str:
        """
        OCR extraction for scanned PDFs.
        Requires pdf2image and pytesseract installed.
        """
        try:
            from pdf2image import convert_from_path
            import pytesseract
            
            images = convert_from_path(
                str(self.path),
                first_page=page_num + 1,
                last_page=page_num + 1,
                dpi=300  # Higher DPI for better OCR
            )
            
            if images:
                # Use Spanish language model for better accuracy
                text = pytesseract.image_to_string(
                    images[0],
                    lang='spa',
                    config='--psm 6'  # Assume uniform block of text
                )
                return text
            
            return ""
        except ImportError as e:
            console.print(f"[red]OCR libraries not installed: {e}[/red]")
            console.print("[yellow]Install with: pip install pdf2image pytesseract[/yellow]")
            # Fallback to basic extraction
            return self.doc[page_num].get_text()
        except Exception as e:
            console.print(f"[red]OCR error on page {page_num}: {e}[/red]")
            return self.doc[page_num].get_text()
    
    def _tables_to_text(self, tables: list) -> str:
        """Convert extracted tables to structured text."""
        result = []
        for table in tables:
            for row in table:
                # Filter out None values
                cells = [str(cell).strip() if cell else "" for cell in row]
                # Skip empty rows
                if any(cells):
                    result.append(" | ".join(cells))
        return "\n".join(result)
    
    def extract_pages_text(self, start_page: int, end_page: int) -> str:
        """
        Extract text from a range of pages.
        Includes page markers for context.
        """
        text_parts = []
        
        for page_num in range(start_page, min(end_page + 1, self.total_pages)):
            page_text = self.extract_page_text(page_num)
            text_parts.append(f"--- PÁGINA {page_num + 1} ---\n{page_text}")
        
        return "\n\n".join(text_parts)
    
    def extract_all_text(self) -> str:
        """Extract text from all pages."""
        return self.extract_pages_text(0, self.total_pages - 1)
    
    def get_page_count(self) -> int:
        """Return total number of pages."""
        return self.total_pages
    
    def close(self):
        """Close the PDF document."""
        if self.doc:
            self.doc.close()
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
    
    def detect_report_type(self) -> str:
        """
        Detect if PDF is Gallo or Visual format.
        Returns 'gallo', 'visual', or 'unknown'.
        """
        # Check first few pages for distinctive markers
        first_pages_text = self.extract_pages_text(0, min(2, self.total_pages - 1)).upper()
        
        # Gallo markers
        gallo_markers = [
            "RESULTADO TOTALES",
            "RESUMEN IMPOSITIVO",
            "TIT.PRIVADOS EXENTOS",
            "RENTA FIJA EN PESOS",
            "RENTA FIJA EN DOLARES",
            "POSICION INICIAL",
            "POSICION FINAL"
        ]
        
        # Visual markers
        visual_markers = [
            "BOLETOS",
            "RESULTADO DE VENTAS",
            "RENTAS DIVIDENDOS",
            "POSICION TITULOS",
            "DOLAR MEP",
            "DOLAR CABLE"
        ]
        
        gallo_score = sum(1 for marker in gallo_markers if marker in first_pages_text)
        visual_score = sum(1 for marker in visual_markers if marker in first_pages_text)
        
        if gallo_score > visual_score:
            return "gallo"
        elif visual_score > gallo_score:
            return "visual"
        else:
            # Check for more specific patterns
            if "RESULTADO TOTALES" in first_pages_text:
                return "gallo"
            elif "BOLETOS" in first_pages_text:
                return "visual"
            return "unknown"


def get_pdf_info(pdf_path: str) -> dict:
    """
    Get basic information about a PDF file.
    """
    with PDFReader(pdf_path) as reader:
        return {
            "path": str(reader.path),
            "name": reader.path.name,
            "pages": reader.total_pages,
            "needs_ocr": reader.is_ocr_needed,
            "report_type": reader.detect_report_type()
        }
