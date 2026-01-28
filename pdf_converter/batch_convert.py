"""
Batch Processing Script for PDF to Excel Conversion
Processes multiple PDFs in a folder or from a list.

Usage:
    python batch_convert.py <folder_or_file> [--output-dir <dir>] [--pattern "*.pdf"]

Example:
    python batch_convert.py ./pdfs/
    python batch_convert.py ./pdfs/ --output-dir ./output/
    python batch_convert.py file_list.txt
"""

import argparse
import sys
import json
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any
from rich.console import Console
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn

sys.path.insert(0, str(Path(__file__).parent))

from app import PDFConverter

console = Console()


def find_pdfs(source: str, pattern: str = "*.pdf") -> List[Path]:
    """
    Find PDF files from a folder or file list.
    
    Args:
        source: Folder path or text file with PDF paths
        pattern: Glob pattern for finding PDFs
    
    Returns:
        List of PDF file paths
    """
    source_path = Path(source)
    
    if source_path.is_dir():
        # Find all PDFs in directory
        pdfs = list(source_path.glob(pattern))
        # Also check subdirectories
        pdfs.extend(source_path.rglob(pattern))
        # Remove duplicates
        pdfs = list(set(pdfs))
    elif source_path.is_file() and source_path.suffix == '.txt':
        # Read paths from text file
        with open(source_path, 'r', encoding='utf-8') as f:
            pdfs = [Path(line.strip()) for line in f if line.strip() and Path(line.strip()).exists()]
    elif source_path.is_file() and source_path.suffix.lower() == '.pdf':
        pdfs = [source_path]
    else:
        pdfs = []
    
    return sorted(pdfs)


def batch_convert(
    pdf_files: List[Path],
    output_dir: Path,
    max_pages_per_chunk: int = 5
) -> List[Dict[str, Any]]:
    """
    Convert multiple PDF files to Excel.
    
    Args:
        pdf_files: List of PDF file paths
        output_dir: Output directory for Excel files
        max_pages_per_chunk: Maximum pages per LLM call
    
    Returns:
        List of conversion results
    """
    results = []
    converter = PDFConverter(max_pages_per_chunk=max_pages_per_chunk)
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        console=console
    ) as progress:
        task = progress.add_task("Converting PDFs...", total=len(pdf_files))
        
        for pdf_path in pdf_files:
            progress.update(task, description=f"Processing {pdf_path.name}...")
            
            try:
                # Generate output path
                output_path = output_dir / f"{pdf_path.stem}_Estructurado.xlsx"
                
                # Convert
                result = converter.convert(
                    pdf_path=str(pdf_path),
                    output_path=str(output_path)
                )
                
                results.append({
                    "file": pdf_path.name,
                    "status": "success" if result["success"] else "failed",
                    "output": result.get("output_file", ""),
                    "sections": len(result.get("sections", [])),
                    "rows": result.get("total_rows", 0),
                    "validation_passed": result.get("validation", {}).get("passed", 0),
                    "validation_failed": result.get("validation", {}).get("failed", 0),
                    "error": result.get("error", "")
                })
                
            except Exception as e:
                results.append({
                    "file": pdf_path.name,
                    "status": "error",
                    "error": str(e)
                })
            
            progress.advance(task)
    
    return results


def print_summary(results: List[Dict[str, Any]]):
    """Print a summary table of conversion results."""
    table = Table(title="Batch Conversion Results")
    
    table.add_column("File", style="cyan")
    table.add_column("Status", justify="center")
    table.add_column("Sections", justify="right")
    table.add_column("Rows", justify="right")
    table.add_column("Validation", justify="center")
    table.add_column("Error", style="red")
    
    success_count = 0
    error_count = 0
    
    for result in results:
        status = result.get("status", "unknown")
        
        if status == "success":
            status_icon = "✅"
            success_count += 1
        elif status == "failed":
            status_icon = "⚠️"
            error_count += 1
        else:
            status_icon = "❌"
            error_count += 1
        
        validation = f"{result.get('validation_passed', 0)}/{result.get('validation_passed', 0) + result.get('validation_failed', 0)}"
        
        table.add_row(
            result.get("file", ""),
            status_icon,
            str(result.get("sections", "")),
            str(result.get("rows", "")),
            validation if status == "success" else "",
            result.get("error", "")[:50] if result.get("error") else ""
        )
    
    console.print(table)
    console.print(f"\n[green]Success: {success_count}[/green] | [red]Errors: {error_count}[/red] | Total: {len(results)}")


def save_results_json(results: List[Dict[str, Any]], output_path: Path):
    """Save results to a JSON file."""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump({
            "timestamp": datetime.now().isoformat(),
            "total": len(results),
            "success": sum(1 for r in results if r.get("status") == "success"),
            "failed": sum(1 for r in results if r.get("status") != "success"),
            "results": results
        }, f, indent=2, ensure_ascii=False)
    
    console.print(f"[dim]Results saved to: {output_path}[/dim]")


def main():
    parser = argparse.ArgumentParser(
        description="Batch convert PDF reports to Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument(
        "source",
        help="Folder with PDFs, text file with paths, or single PDF file"
    )
    
    parser.add_argument(
        "--output-dir", "-o",
        default="./output",
        help="Output directory for Excel files (default: ./output)"
    )
    
    parser.add_argument(
        "--pattern", "-p",
        default="*.pdf",
        help="Glob pattern for finding PDFs (default: *.pdf)"
    )
    
    parser.add_argument(
        "--chunk-size", "-c",
        type=int,
        default=5,
        help="Maximum pages per LLM call (default: 5)"
    )
    
    parser.add_argument(
        "--save-results",
        action="store_true",
        help="Save results to JSON file"
    )
    
    args = parser.parse_args()
    
    # Find PDFs
    console.print(f"[cyan]Searching for PDFs in: {args.source}[/cyan]")
    pdf_files = find_pdfs(args.source, args.pattern)
    
    if not pdf_files:
        console.print("[red]No PDF files found.[/red]")
        sys.exit(1)
    
    console.print(f"[green]Found {len(pdf_files)} PDF files[/green]")
    
    # Convert
    output_dir = Path(args.output_dir)
    results = batch_convert(pdf_files, output_dir, args.chunk_size)
    
    # Summary
    print_summary(results)
    
    # Save results
    if args.save_results:
        results_path = output_dir / f"batch_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        save_results_json(results, results_path)
    
    # Exit code
    error_count = sum(1 for r in results if r.get("status") != "success")
    sys.exit(1 if error_count > 0 else 0)


if __name__ == "__main__":
    main()
