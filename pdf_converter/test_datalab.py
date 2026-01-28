#!/usr/bin/env python
"""
Test script for Datalab PDF to Markdown conversion.
Uses the Datalab API for high-accuracy OCR of financial PDFs.
"""

import os
import sys
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv
from rich.console import Console
from rich.markdown import Markdown

from datalab import DatalabClient

console = Console()


def test_datalab_conversion(pdf_path: str, mode: str = "balanced"):
    """
    Test Datalab conversion on a PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        mode: Processing mode (fast, balanced, accurate)
    """
    load_dotenv()
    
    api_key = os.environ.get("DATALAB_API_KEY", "").strip()
    if not api_key:
        console.print("[red]❌ DATALAB_API_KEY not found in environment.[/red]")
        console.print("[dim]Set it in .env file or as environment variable.[/dim]")
        console.print("[dim]Get your key at: https://www.datalab.to[/dim]")
        return
    
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        console.print(f"[red]❌ File not found: {pdf_path}[/red]")
        return
    
    console.print(f"\n[bold cyan]🔬 Testing Datalab Conversion[/bold cyan]")
    console.print(f"[dim]PDF: {pdf_path.name}[/dim]")
    console.print(f"[dim]Mode: {mode}[/dim]")
    console.print()
    
    with DatalabClient(api_key=api_key, mode=mode) as client:
        # Check API health
        console.print("[dim]Checking API status...[/dim]")
        if not client.check_health():
            console.print("[yellow]⚠️ API health check failed, proceeding anyway...[/yellow]")
        else:
            console.print("[green]✓ API is healthy[/green]")
        
        console.print()
        
        # Convert PDF
        result = client.convert_pdf(str(pdf_path), paginate=True)
        
        if not result.success:
            console.print(f"[red]❌ Conversion failed: {result.error}[/red]")
            return
        
        console.print()
        console.print(f"[green]✓ Conversion successful![/green]")
        console.print(f"[dim]Pages: {result.page_count}[/dim]")
        console.print(f"[dim]Runtime: {result.runtime:.1f}s[/dim]")
        
        if result.cost_breakdown:
            console.print(f"[dim]Cost: {result.cost_breakdown}[/dim]")
        
        # Save markdown output
        output_path = pdf_path.with_suffix(".datalab.md")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(result.markdown or "")
        
        console.print(f"\n[cyan]📄 Markdown saved to: {output_path}[/cyan]")
        
        # Show preview
        console.print("\n[bold]Preview (first 2000 chars):[/bold]")
        console.print("-" * 60)
        preview = (result.markdown or "")[:2000]
        console.print(preview)
        if len(result.markdown or "") > 2000:
            console.print("\n[dim]... (truncated)[/dim]")
        console.print("-" * 60)
        
        return result


def main():
    """Main entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Test Datalab PDF conversion")
    parser.add_argument("pdf", nargs="?", help="PDF file to convert")
    parser.add_argument("-m", "--mode", choices=["fast", "balanced", "accurate"],
                       default="balanced", help="Processing mode (default: balanced)")
    
    args = parser.parse_args()
    
    # Default test PDF
    if not args.pdf:
        base_dir = Path(__file__).parent.parent
        test_pdfs = [
            base_dir / "Vero_2025_gallo.PDF",
            base_dir / "Aguiar_2025_Gallo.PDF",
        ]
        for pdf in test_pdfs:
            if pdf.exists():
                args.pdf = str(pdf)
                break
    
    if not args.pdf:
        console.print("[red]❌ No PDF file specified and no test file found.[/red]")
        console.print("Usage: python test_datalab.py <pdf_file> [-m mode]")
        return 1
    
    test_datalab_conversion(args.pdf, args.mode)
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
