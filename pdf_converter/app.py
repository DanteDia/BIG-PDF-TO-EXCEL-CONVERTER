"""
PDF to Excel Converter - Main Application
Converts Gallo and Visual financial reports from PDF to structured Excel.

Usage:
    python app.py <pdf_file> [--output <output_file>] [--type gallo|visual]

Example:
    python app.py Vero_2025_gallo.PDF
    python app.py VeroLandro2025.pdf --output output.xlsx --type visual
"""

import argparse
import sys
from pathlib import Path
from typing import Optional, Dict, List, Any
from rich.console import Console
from rich.panel import Panel

# Add parent directory to path for imports
app_dir = Path(__file__).parent
sys.path.insert(0, str(app_dir))
sys.path.insert(0, str(app_dir.parent))

try:
    from pdf.reader import PDFReader, get_pdf_info
    from llm.client import LLMClient
    from extractor.gallo import GalloExtractor
    from extractor.visual import VisualExtractor
    from extractor.schemas import (
        GALLO_SCHEMAS, VISUAL_SCHEMAS,
        GALLO_SECTION_TO_SHEET, VISUAL_SECTION_TO_SHEET,
        GALLO_NUMERIC_FIELDS, VISUAL_NUMERIC_FIELDS,
        GALLO_DEDUP_KEYS, VISUAL_DEDUP_KEYS,
    )
    from postprocess.cleanup import cleanup_section_data
    from postprocess.decimals_fix import fix_resumen_decimals, fix_gallo_totales
    from validation.gallo import validate_gallo, print_validation_report, validation_report_to_dict
    from validation.visual import validate_visual
    from validation.visual import print_validation_report as print_visual_validation
    from export.excel_writer import create_excel_from_data
except ImportError:
    from pdf_converter.pdf.reader import PDFReader, get_pdf_info
    from pdf_converter.llm.client import LLMClient
    from pdf_converter.extractor.gallo import GalloExtractor
    from pdf_converter.extractor.visual import VisualExtractor
    from pdf_converter.extractor.schemas import (
        GALLO_SCHEMAS, VISUAL_SCHEMAS,
        GALLO_SECTION_TO_SHEET, VISUAL_SECTION_TO_SHEET,
        GALLO_NUMERIC_FIELDS, VISUAL_NUMERIC_FIELDS,
        GALLO_DEDUP_KEYS, VISUAL_DEDUP_KEYS,
    )
    from pdf_converter.postprocess.cleanup import cleanup_section_data
    from pdf_converter.postprocess.decimals_fix import fix_resumen_decimals, fix_gallo_totales
    from pdf_converter.validation.gallo import validate_gallo, print_validation_report, validation_report_to_dict
    from pdf_converter.validation.visual import validate_visual
    from pdf_converter.validation.visual import print_validation_report as print_visual_validation
    from pdf_converter.export.excel_writer import create_excel_from_data

console = Console()


class PDFConverter:
    """
    Main converter class that orchestrates the PDF to Excel conversion.
    """
    
    def __init__(self, max_pages_per_chunk: int = 5):
        """
        Initialize the converter.
        
        Args:
            max_pages_per_chunk: Maximum pages to process per LLM call
        """
        self.max_pages_per_chunk = max_pages_per_chunk
        self.llm = LLMClient()
    
    def convert(
        self,
        pdf_path: str,
        output_path: Optional[str] = None,
        report_type: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Convert a PDF file to Excel.
        
        Args:
            pdf_path: Path to the PDF file
            output_path: Optional output path (auto-generated if not provided)
            report_type: Optional report type ('gallo' or 'visual', auto-detected if not provided)
        
        Returns:
            Dictionary with conversion results
        """
        pdf_path = Path(pdf_path)
        
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF not found: {pdf_path}")
        
        # Display header
        console.print(Panel.fit(
            f"[bold blue]📄 PDF to Excel Converter[/bold blue]\n"
            f"File: {pdf_path.name}",
            border_style="blue"
        ))
        
        # Get PDF info and detect type
        pdf_info = get_pdf_info(str(pdf_path))
        console.print(f"\n[dim]Pages: {pdf_info['pages']} | OCR needed: {pdf_info['needs_ocr']}[/dim]")
        
        # Determine report type
        if report_type:
            detected_type = report_type.lower()
        else:
            detected_type = pdf_info['report_type']
        
        if detected_type not in ['gallo', 'visual']:
            console.print(f"[red]❌ Could not detect report type. Please specify with --type[/red]")
            return {"success": False, "error": "Unknown report type"}
        
        console.print(f"[cyan]Report type: {detected_type.upper()}[/cyan]")
        
        # Generate output path if not provided
        if not output_path:
            output_path = pdf_path.parent / f"{pdf_path.stem}_Estructurado.xlsx"
        
        # Extract data
        console.print(f"\n[bold]📊 Extracting data...[/bold]")
        
        if detected_type == 'gallo':
            data = self._extract_gallo(str(pdf_path))
        else:
            data = self._extract_visual(str(pdf_path))
        
        # Post-process
        console.print(f"\n[bold]🔧 Post-processing...[/bold]")
        data = self._postprocess(data, detected_type)
        
        # Validate
        console.print(f"\n[bold]✅ Validating...[/bold]")
        validation = self._validate(data, detected_type)
        
        # Generate Excel
        console.print(f"\n[bold]📥 Generating Excel...[/bold]")
        
        if detected_type == 'gallo':
            schemas = GALLO_SCHEMAS
            sheet_names = GALLO_SECTION_TO_SHEET
            numeric_fields = GALLO_NUMERIC_FIELDS
        else:
            schemas = VISUAL_SCHEMAS
            sheet_names = VISUAL_SECTION_TO_SHEET
            numeric_fields = VISUAL_NUMERIC_FIELDS
        
        output_file = create_excel_from_data(
            data=data,
            output_path=str(output_path),
            report_type=detected_type,
            schemas=schemas,
            sheet_names=sheet_names,
            numeric_fields=numeric_fields,
            validation_results=validation.get("results") if validation else None
        )
        
        # Summary
        total_rows = sum(len(rows) for rows in data.values())
        console.print(Panel.fit(
            f"[bold green]✅ Conversion Complete![/bold green]\n\n"
            f"📄 Input: {pdf_path.name}\n"
            f"📥 Output: {Path(output_file).name}\n"
            f"📊 Sections: {len(data)}\n"
            f"📝 Total rows: {total_rows}\n"
            f"✅ Validation: {validation.get('passed', 0)}/{validation.get('passed', 0) + validation.get('failed', 0)} passed",
            border_style="green"
        ))
        
        return {
            "success": True,
            "input_file": str(pdf_path),
            "output_file": output_file,
            "report_type": detected_type,
            "sections": list(data.keys()),
            "total_rows": total_rows,
            "validation": validation
        }
    
    def _extract_gallo(self, pdf_path: str) -> Dict[str, List[Dict]]:
        """Extract data from a Gallo PDF."""
        with GalloExtractor(pdf_path, self.llm, self.max_pages_per_chunk) as extractor:
            return extractor.extract_all()
    
    def _extract_visual(self, pdf_path: str) -> Dict[str, List[Dict]]:
        """Extract data from a Visual PDF."""
        with VisualExtractor(pdf_path, self.llm, self.max_pages_per_chunk) as extractor:
            return extractor.extract_all()
    
    def _postprocess(self, data: Dict[str, List[Dict]], report_type: str) -> Dict[str, List[Dict]]:
        """Apply post-processing to extracted data."""
        
        if report_type == 'gallo':
            numeric_fields = GALLO_NUMERIC_FIELDS
            dedup_keys = GALLO_DEDUP_KEYS
        else:
            numeric_fields = VISUAL_NUMERIC_FIELDS
            dedup_keys = VISUAL_DEDUP_KEYS
        
        processed = {}
        
        for section_key, rows in data.items():
            if not rows:
                processed[section_key] = rows
                continue
            
            # Apply cleanup
            cleaned = cleanup_section_data(
                rows=rows,
                numeric_fields=numeric_fields.get(section_key, []),
                dedup_keys=dedup_keys.get(section_key, []),
                report_type=report_type,
                fill_entity=(report_type == 'gallo' and section_key not in ['resultado_totales', 'posicion_inicial', 'posicion_final'])
            )
            
            processed[section_key] = cleaned
            
            if len(rows) != len(cleaned):
                console.print(f"  [dim]{section_key}: {len(rows)} → {len(cleaned)} rows (dedup)[/dim]")
        
        # Apply decimal fixes
        if report_type == 'visual' and 'resumen' in processed:
            processed['resumen'] = fix_resumen_decimals(processed['resumen'], processed)
        elif report_type == 'gallo' and 'resultado_totales' in processed:
            fix_gallo_totales(processed['resultado_totales'], processed)
        
        return processed
    
    def _validate(self, data: Dict[str, List[Dict]], report_type: str) -> Dict[str, Any]:
        """Validate extracted data."""
        
        if report_type == 'gallo':
            report = validate_gallo(data)
            print_validation_report(report)
        else:
            report = validate_visual(data)
            print_visual_validation(report)
        
        return validation_report_to_dict(report)


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Convert financial PDF reports to Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python app.py Vero_2025_gallo.PDF
    python app.py VeroLandro2025.pdf --output output.xlsx
    python app.py report.pdf --type gallo
        """
    )
    
    parser.add_argument(
        "pdf_file",
        help="Path to the PDF file to convert"
    )
    
    parser.add_argument(
        "--output", "-o",
        help="Output Excel file path (default: <input>_Estructurado.xlsx)"
    )
    
    parser.add_argument(
        "--type", "-t",
        choices=["gallo", "visual"],
        help="Report type (auto-detected if not specified)"
    )
    
    parser.add_argument(
        "--chunk-size", "-c",
        type=int,
        default=5,
        help="Maximum pages per LLM call (default: 5)"
    )
    
    args = parser.parse_args()
    
    try:
        converter = PDFConverter(max_pages_per_chunk=args.chunk_size)
        result = converter.convert(
            pdf_path=args.pdf_file,
            output_path=args.output,
            report_type=args.type
        )
        
        if result["success"]:
            sys.exit(0)
        else:
            console.print(f"[red]Error: {result.get('error', 'Unknown error')}[/red]")
            sys.exit(1)
            
    except FileNotFoundError as e:
        console.print(f"[red]Error: {e}[/red]")
        sys.exit(1)
    except KeyboardInterrupt:
        console.print("\n[yellow]Cancelled by user[/yellow]")
        sys.exit(130)
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
