"""
Validation engine for Gallo reports.
Cross-checks Resultado Totales against section sums.
"""

import re
from typing import List, Dict, Any, Optional
from dataclasses import dataclass
from rich.console import Console
from rich.table import Table

console = Console()


@dataclass
class ValidationResult:
    """Result of a single validation check."""
    field: str
    expected: float
    calculated: float
    match: bool
    difference: float


@dataclass
class ValidationReport:
    """Complete validation report."""
    report_type: str
    passed: int
    failed: int
    results: List[ValidationResult]
    
    @property
    def success(self) -> bool:
        return self.failed == 0


TOLERANCE = 0.01  # Maximum acceptable difference


def validate_gallo(data: Dict[str, List[Dict]]) -> ValidationReport:
    """
    Validate Gallo report data.
    
    Cross-checks each category in Resultado Totales against the
    corresponding section's Total rows.
    
    Args:
        data: Dictionary with section_key -> list of rows
    
    Returns:
        ValidationReport with all validation results
    """
    results = []
    resultado_totales = data.get("resultado_totales", [])
    
    for row in resultado_totales:
        categoria = str(row.get("categoria", "")).upper()
        valor_pesos = float(row.get("valor_pesos", 0))
        valor_usd = float(row.get("valor_usd", 0))
        
        # Skip TOTAL GENERAL
        if "TOTAL GENERAL" in categoria:
            continue
        
        # Extract category base and type
        match = re.match(r'^(.+?)\s*\((.+?)\)\s*$', categoria)
        if not match:
            continue
        
        cat_base = match.group(1).strip()
        tipo = match.group(2).strip().lower()  # "enajenacion" or "renta"
        
        # Find corresponding section
        section_key = _map_gallo_categoria_to_section(cat_base)
        if not section_key or section_key not in data:
            console.print(f"[dim]  Skipping validation for {categoria} (section not found)[/dim]")
            continue
        
        section_rows = data[section_key]
        
        # Calculate expected values
        if "caucion" in section_key:
            # Cauciones: sum interes from non-total rows
            calc_pesos = sum(
                float(r.get("interes_pesos", 0)) 
                for r in section_rows 
                if "total" not in str(r.get("tipo_fila", "")).lower()
            )
            calc_usd = sum(
                float(r.get("interes_usd", 0)) 
                for r in section_rows 
                if "total" not in str(r.get("tipo_fila", "")).lower()
            )
        else:
            # Other sections: sum from Total rows matching tipo
            calc_pesos = sum(
                float(r.get("resultado_pesos", 0)) 
                for r in section_rows 
                if tipo in str(r.get("tipo_fila", "")).lower()
            )
            calc_usd = sum(
                float(r.get("resultado_usd", 0)) 
                for r in section_rows 
                if tipo in str(r.get("tipo_fila", "")).lower()
            )
        
        # Validate pesos
        diff_pesos = abs(calc_pesos - valor_pesos)
        results.append(ValidationResult(
            field=f"{categoria} pesos",
            expected=valor_pesos,
            calculated=calc_pesos,
            match=diff_pesos <= TOLERANCE,
            difference=diff_pesos
        ))
        
        # Validate USD
        diff_usd = abs(calc_usd - valor_usd)
        results.append(ValidationResult(
            field=f"{categoria} usd",
            expected=valor_usd,
            calculated=calc_usd,
            match=diff_usd <= TOLERANCE,
            difference=diff_usd
        ))
    
    passed = sum(1 for r in results if r.match)
    failed = sum(1 for r in results if not r.match)
    
    return ValidationReport(
        report_type="gallo",
        passed=passed,
        failed=failed,
        results=results
    )


def _map_gallo_categoria_to_section(cat_base: str) -> Optional[str]:
    """Map categoria name to section key."""
    cat_lower = cat_base.lower()
    
    mappings = {
        "tit.privados exentos": "tit_privados_exentos",
        "tit privados exentos": "tit_privados_exentos",
        "titulos privados exentos": "tit_privados_exentos",
        "tit.privados del exterior": "tit_privados_exterior",
        "tit privados del exterior": "tit_privados_exterior",
        "renta fija en pesos": "renta_fija_pesos",
        "renta fija en dolares": "renta_fija_dolares",
        "renta fija en dólares": "renta_fija_dolares",
        "cauciones en pesos": "cauciones_pesos",
        "cauciones en dolares": "cauciones_dolares",
        "cauciones en dólares": "cauciones_dolares",
        "fci": "fci",
        "fondos comunes": "fci",
        "opciones": "opciones",
        "futuros": "futuros",
    }
    
    for pattern, section in mappings.items():
        if pattern in cat_lower:
            return section
    
    return None


def print_validation_report(report: ValidationReport):
    """Print validation report as a formatted table."""
    table = Table(title=f"Validation Report ({report.report_type.upper()})")
    
    table.add_column("Campo", style="cyan")
    table.add_column("Calculado", justify="right")
    table.add_column("Esperado", justify="right")
    table.add_column("Match", justify="center")
    
    for result in report.results:
        match_icon = "✅" if result.match else "❌"
        style = "" if result.match else "red"
        
        table.add_row(
            result.field,
            f"{result.calculated:,.2f}",
            f"{result.expected:,.2f}",
            match_icon,
            style=style
        )
    
    console.print(table)
    
    if report.success:
        console.print(f"\n[green]✅ All validations passed ({report.passed}/{report.passed + report.failed})[/green]")
    else:
        console.print(f"\n[red]❌ Validation failed: {report.passed} passed, {report.failed} failed[/red]")


def validation_report_to_dict(report: ValidationReport) -> dict:
    """Convert validation report to dictionary for JSON output."""
    return {
        "report_type": report.report_type,
        "passed": report.passed,
        "failed": report.failed,
        "success": report.success,
        "results": [
            {
                "field": r.field,
                "expected": r.expected,
                "calculated": r.calculated,
                "match": r.match,
                "difference": r.difference
            }
            for r in report.results
        ]
    }
