"""
Validation engine for Visual reports.
Cross-checks Resumen totals against detail section sums.
"""

from typing import List, Dict, Any
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


def validate_visual(data: Dict[str, List[Dict]]) -> ValidationReport:
    """
    Validate Visual report data.
    
    Cross-checks Resumen totals against detail section sums:
    - ventas = SUM(resultado_ventas.resultado)
    - rentas = SUMIF(rentas_dividendos, categoria="RENTAS", importe)
    - dividendos = SUMIF(rentas_dividendos, categoria="DIVIDENDOS", importe)
    - total = SUM(all fields)
    
    Args:
        data: Dictionary with section_key -> list of rows
    
    Returns:
        ValidationReport with all validation results
    """
    results = []
    resumen = data.get("resumen", [])
    
    for row in resumen:
        moneda = str(row.get("moneda", "")).upper()
        
        # Determine currency suffix
        if "ARS" in moneda or "PESOS" in moneda:
            suffix = "ars"
            currency_name = "ARS"
        elif "USD" in moneda or "DOLAR" in moneda:
            suffix = "usd"
            currency_name = "USD"
        else:
            continue
        
        # Validate ventas
        ventas_expected = float(row.get("ventas", 0))
        ventas_key = f"resultado_ventas_{suffix}"
        ventas_rows = data.get(ventas_key, [])
        ventas_calculated = sum(float(r.get("resultado", 0)) for r in ventas_rows)
        
        diff_ventas = abs(ventas_calculated - ventas_expected)
        results.append(ValidationResult(
            field=f"ventas {currency_name}",
            expected=ventas_expected,
            calculated=ventas_calculated,
            match=diff_ventas <= TOLERANCE,
            difference=diff_ventas
        ))
        
        # Validate rentas
        rentas_expected = float(row.get("rentas", 0))
        rentas_key = f"rentas_dividendos_{suffix}"
        rentas_rows = data.get(rentas_key, [])
        rentas_calculated = sum(
            float(r.get("importe", 0)) 
            for r in rentas_rows 
            if str(r.get("categoria", "")).upper() in ["RENTAS", "RENTA"]
        )
        
        diff_rentas = abs(rentas_calculated - rentas_expected)
        results.append(ValidationResult(
            field=f"rentas {currency_name}",
            expected=rentas_expected,
            calculated=rentas_calculated,
            match=diff_rentas <= TOLERANCE,
            difference=diff_rentas
        ))
        
        # Validate dividendos
        dividendos_expected = float(row.get("dividendos", 0))
        dividendos_calculated = sum(
            float(r.get("importe", 0)) 
            for r in rentas_rows 
            if str(r.get("categoria", "")).upper() in ["DIVIDENDOS", "DIVIDENDO"]
        )
        
        diff_dividendos = abs(dividendos_calculated - dividendos_expected)
        results.append(ValidationResult(
            field=f"dividendos {currency_name}",
            expected=dividendos_expected,
            calculated=dividendos_calculated,
            match=diff_dividendos <= TOLERANCE,
            difference=diff_dividendos
        ))
        
        # Validate total
        total_expected = float(row.get("total", 0))
        total_calculated = sum(float(row.get(f, 0)) for f in [
            "ventas", "fci", "opciones", "rentas", "dividendos",
            "ef_cpd", "pagares", "futuros", "cau_int", "cau_cf"
        ])
        
        diff_total = abs(total_calculated - total_expected)
        results.append(ValidationResult(
            field=f"total {currency_name}",
            expected=total_expected,
            calculated=total_calculated,
            match=diff_total <= TOLERANCE,
            difference=diff_total
        ))
    
    passed = sum(1 for r in results if r.match)
    failed = sum(1 for r in results if not r.match)
    
    return ValidationReport(
        report_type="visual",
        passed=passed,
        failed=failed,
        results=results
    )


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
