from __future__ import annotations

import argparse
import json
from pathlib import Path

from dotenv import load_dotenv

from pdf_converter.convert_with_datalab import convert_pdf_to_excel
from pdf_converter.datalab.excel_to_pdf import ExcelToPdfExporter
from pdf_converter.datalab.economic_sanity import add_validation_sheet, validate_workbook
from pdf_converter.datalab.merge_gallo_visual import GalloVisualMerger


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate final merged Excel and PDF for a case.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--visual-pdf", type=Path)
    parser.add_argument("--gallo-pdf", type=Path)
    parser.add_argument("--precio-pdf", type=Path)
    parser.add_argument("--visual-xlsx", type=Path)
    parser.add_argument("--gallo-xlsx", type=Path)
    parser.add_argument("--precio-xlsx", type=Path)
    parser.add_argument("--case-prefix", required=True)
    parser.add_argument("--client-number", required=True)
    parser.add_argument("--client-name", required=True)
    parser.add_argument("--year", type=int, default=2025)
    parser.add_argument("--period-start", default="Enero 1")
    parser.add_argument("--period-end", default="Diciembre 31")
    parser.add_argument(
        "--prefer-precio-tenencias-usd-basis",
        action="store_true",
        help="Backwards-compatible no-op: Precio Tenencias USD fixed-income basis is enabled by default.",
    )
    parser.add_argument(
        "--auto-fallback-usd-basis-on-validation",
        action="store_true",
        help="Backwards-compatible no-op: USD basis fallback on validation is enabled by default.",
    )
    parser.add_argument(
        "--precio-tenencias-usd-basis-fallback-code",
        action="append",
        default=[],
        help="Code that should use Gallo position USD even when --prefer-precio-tenencias-usd-basis is enabled.",
    )
    args = parser.parse_args()

    root = args.root.resolve()
    load_dotenv(root / "pdf_converter" / ".env")

    def resolve_input(path: Path | None) -> Path | None:
        if path is None:
            return None
        return (root / path).resolve() if not path.is_absolute() else path.resolve()

    visual_pdf = resolve_input(args.visual_pdf)
    gallo_pdf = resolve_input(args.gallo_pdf)
    precio_pdf = resolve_input(args.precio_pdf)
    visual_xlsx = resolve_input(args.visual_xlsx)
    gallo_xlsx = resolve_input(args.gallo_xlsx)
    precio_xlsx = resolve_input(args.precio_xlsx)

    visual_excel = visual_xlsx or (root / f"{args.case_prefix}_Visual_from_PDF.xlsx")
    gallo_excel = gallo_xlsx or (root / f"{args.case_prefix}_Gallo_from_PDF.xlsx")
    precio_excel = precio_xlsx or ((root / f"{args.case_prefix}_PrecioTenencias_from_PDF.xlsx") if precio_pdf else None)
    merge_formulas = root / f"{args.case_prefix}_Resumen_Impositivo_FIXED_formulas.xlsx"
    merge_values = root / f"{args.case_prefix}_Resumen_Impositivo_FIXED_values.xlsx"
    pdf_output = root / f"{args.case_prefix}_Resumen_Impositivo_FIXED.pdf"

    for output_path in [visual_excel, gallo_excel, precio_excel, merge_formulas, merge_values, pdf_output]:
        if output_path is not None:
            output_path.parent.mkdir(parents=True, exist_ok=True)

    missing_inputs = []
    if visual_excel is None or not visual_excel.exists():
        if visual_pdf is None:
            missing_inputs.append("--visual-pdf or --visual-xlsx")
    if gallo_excel is None or not gallo_excel.exists():
        if gallo_pdf is None:
            missing_inputs.append("--gallo-pdf or --gallo-xlsx")
    if missing_inputs:
        parser.error("Missing required inputs: " + ", ".join(missing_inputs))

    print(f"Generating case: {args.case_prefix}")
    print(f"Visual source: {visual_excel.name if visual_excel.exists() else visual_pdf.name}")
    print(f"Gallo source: {gallo_excel.name if gallo_excel.exists() else gallo_pdf.name}")
    print(f"Precio source: {precio_excel.name if precio_excel and precio_excel.exists() else (precio_pdf.name if precio_pdf else 'omitted')}")

    if not visual_excel.exists():
        convert_pdf_to_excel(str(visual_pdf), str(visual_excel))
    if not gallo_excel.exists():
        convert_pdf_to_excel(str(gallo_pdf), str(gallo_excel))
    if precio_excel is not None and not precio_excel.exists():
        convert_pdf_to_excel(str(precio_pdf), str(precio_excel))

    aux_dir = root / "pdf_converter" / "datalab" / "aux_data"
    merger = GalloVisualMerger(
        str(gallo_excel),
        str(visual_excel),
        str(aux_dir),
        precio_tenencias_path=str(precio_excel) if precio_excel else None,
        prefer_precio_tenencias_usd_cost_basis=True,
        precio_tenencias_usd_basis_fallback_codes=list(args.precio_tenencias_usd_basis_fallback_code),
    )
    wb_formulas, wb_values = merger.merge(output_mode="both")
    validation_report = validate_workbook(wb_values)

    add_validation_sheet(wb_values, validation_report)
    wb_formulas.save(merge_formulas)
    wb_values.save(merge_values)

    validation_output = root / f"{args.case_prefix}_Resumen_Impositivo_VALIDATION.json"
    validation_output.write_text(
        json.dumps(validation_report.to_dict(), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    exporter = ExcelToPdfExporter(
        str(merge_values),
        {"numero": args.client_number, "nombre": args.client_name},
    )
    exporter.periodo_inicio = args.period_start
    exporter.periodo_fin = args.period_end
    exporter.anio = args.year
    exporter.export_to_pdf(str(pdf_output))

    print("DONE")
    print(visual_excel)
    print(gallo_excel)
    if precio_excel:
        print(precio_excel)
    print(merge_formulas)
    print(merge_values)
    print(validation_output)
    print(pdf_output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())