from datetime import date

from openpyxl import Workbook

from pdf_converter.datalab.merge_gallo_visual import GalloVisualMerger


def _minimal_merger() -> GalloVisualMerger:
    merger = object.__new__(GalloVisualMerger)
    merger._precio_tenencias_by_codigo = {}
    merger._precio_tenencias_by_ticker = {}
    merger._precio_tenencias_qty_by_codigo = {}
    merger._precio_tenencias_qty_by_ticker = {}
    merger._precio_tenencias_zero_cost_codes = set()
    merger._precio_tenencias_zero_cost_tickers = set()
    merger._precios_iniciales_by_codigo = {"8499": {"precio": 13900}}
    merger._especies_visual_cache = {"8499": {"tipo_especie": "Cedears"}}
    merger._gallo_position_dates = {"Posicion Final": date(2025, 5, 31)}
    return merger


def _workbook_with_final_globant_position() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Posicion Final Gallo"
    headers = [
        "tipo_especie", "Ticker", "especie", "Codigo especie", "Codigo Especie Origen",
        "comentario especies", "detalle", "custodia", "cantidad", "precio Tenencia Final Pesos",
        "precio Tenencia Final USD", "Precio Tenencia Inicial", "precio costo(en proceso)",
        "Origen precio costo", "comentarios precio costo", "Precio a Utilizar", "importe_pesos",
        "porc_cartera_pesos", "importe_dolares", "porc_cartera_dolares", "Tipo Instrumento",
        "Precio Nominal",
    ]
    ws.append(headers)
    ws.append([
        "TITULOS PRIVADOS LOCALES", "GLOB", "CEDEAR GLOBANT", 8499, "PreciosInicialesEspecies",
        "", "", "CAJA VALORES", 880, 6560, 5.497386363636363, 13900, "",
        "PreciosInicialesEspecies", "", 13900, 5772800, 4.86, 4837.7, 4.86, "Cedears", 13900,
    ])
    return wb


def test_dated_intermediate_gallo_position_seeds_later_ars_sale():
    merger = _minimal_merger()
    wb = _workbook_with_final_globant_position()

    cantidad, precio = merger._get_posicion_inicial(
        wb,
        "8499",
        is_gallo=False,
        for_usd=False,
        operation_date=date(2025, 6, 30),
    )

    assert cantidad == 880
    assert precio == 6560


def test_dated_intermediate_gallo_position_does_not_seed_prior_operation():
    merger = _minimal_merger()
    wb = _workbook_with_final_globant_position()

    cantidad, precio = merger._get_posicion_inicial(
        wb,
        "8499",
        is_gallo=False,
        for_usd=False,
        operation_date=date(2025, 5, 1),
    )

    assert cantidad == 0
    assert precio == 13900
