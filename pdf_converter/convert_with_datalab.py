#!/usr/bin/env python
"""
Complete PDF to Excel converter using Datalab API.
Handles both Gallo and Visual format financial reports.
"""

import os
import sys
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel

from datalab import DatalabClient
from datalab.md_to_excel import convert_markdown_to_excel

console = Console()


def convert_pdf_to_excel(
    pdf_path: str,
    output_path: Optional[str] = None,
    mode: str = "accurate",
    keep_markdown: bool = True
) -> str:
    """
    Convert a PDF financial report to structured Excel.
    
    Args:
        pdf_path: Path to the input PDF
        output_path: Optional path for output Excel
        mode: Datalab processing mode (fast, balanced, accurate)
        keep_markdown: Keep the intermediate markdown file
    
    Returns:
        Path to the generated Excel file
    """
    load_dotenv()
    
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")
    
    api_key = os.environ.get("DATALAB_API_KEY", "").strip()
    if not api_key:
        raise ValueError(
            "DATALAB_API_KEY not found. "
            "Set it in .env file or as environment variable. "
            "Get your key at: https://www.datalab.to"
        )
    
    # Default output path
    if not output_path:
        output_path = str(pdf_path.with_suffix('.xlsx'))
    
    console.print(Panel.fit(
        f"[bold cyan]PDF to Excel Converter[/bold cyan]\n"
        f"[dim]Using Datalab API ({mode} mode)[/dim]",
        border_style="cyan"
    ))
    
    console.print(f"\n[bold]Input:[/bold] {pdf_path.name}")
    console.print(f"[bold]Output:[/bold] {Path(output_path).name}")
    console.print()
    
    # Step 1: Convert PDF to Markdown using Datalab
    console.print("[cyan]Step 1:[/cyan] Converting PDF to Markdown...")
    
    with DatalabClient(api_key=api_key, mode=mode) as client:
        result = client.convert_pdf(str(pdf_path), paginate=True)
        
        if not result.success:
            raise RuntimeError(f"PDF conversion failed: {result.error}")
        
        console.print(f"  [green]✓[/green] Converted {result.page_count} pages in {result.runtime:.1f}s")
        
        if result.cost_breakdown:
            cost = result.cost_breakdown.get('final_cost_cents', 0) / 100
            console.print(f"  [dim]Cost: ${cost:.2f}[/dim]")
    
    # Save markdown
    md_path = pdf_path.with_suffix('.datalab.md')
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(result.markdown or "")
    
    console.print(f"  [dim]Markdown saved: {md_path.name}[/dim]")
    
    # Step 2: Parse Markdown and create Excel
    console.print("\n[cyan]Step 2:[/cyan] Creating Excel from Markdown...")
    
    excel_path = convert_markdown_to_excel(str(md_path), output_path)
    
    # Cleanup if not keeping markdown
    if not keep_markdown:
        md_path.unlink()
        console.print("  [dim]Cleaned up intermediate files[/dim]")
    
    console.print(Panel.fit(
        f"[bold green]✓ Conversion Complete[/bold green]\n"
        f"Output: [cyan]{excel_path}[/cyan]",
        border_style="green"
    ))
    
    return excel_path


def main():
    """Main entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Convert PDF financial reports to Excel using Datalab API"
    )
    parser.add_argument("pdf", help="Path to the PDF file")
    parser.add_argument("-o", "--output", help="Output Excel path")
    parser.add_argument(
        "-m", "--mode",
        choices=["fast", "balanced", "accurate"],
        default="accurate",
        help="Processing mode (default: accurate)"
    )
    parser.add_argument(
        "--no-keep-md",
        action="store_true",
        help="Delete intermediate markdown file"
    )
    
    args = parser.parse_args()
    
    try:
        convert_pdf_to_excel(
            args.pdf,
            args.output,
            mode=args.mode,
            keep_markdown=not args.no_keep_md
        )
        return 0
    except Exception as e:
        console.print(f"[red]❌ Error: {e}[/red]")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
