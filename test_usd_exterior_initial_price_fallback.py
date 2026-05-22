from pdf_converter.datalab.merge_gallo_visual import GalloVisualMerger


def _minimal_merger() -> GalloVisualMerger:
    merger = object.__new__(GalloVisualMerger)
    merger._precio_tenencias_by_codigo = {}
    merger._precio_tenencias_by_ticker = {}
    merger._precio_tenencias_qty_by_codigo = {}
    merger._precio_tenencias_qty_by_ticker = {}
    merger._precio_tenencias_zero_cost_codes = set()
    merger._precio_tenencias_zero_cost_tickers = set()
    merger._precios_iniciales_by_codigo = {
        "7877": {"precio": 250.44},
        "95806": {"precio": 103.5},
    }
    merger._especies_visual_cache = {
        "7877": {"moneda_emision": "Dolar Cable (exterior)", "tipo_especie": "Acciones"},
        "95806": {"moneda_emision": "Dolar MEP (local)", "tipo_especie": "Obligaciones Negociables"},
    }
    merger._gallo_position_dates = {}
    return merger


def test_usd_exterior_action_precios_iniciales_fallback_stays_in_usd():
    merger = _minimal_merger()

    cantidad, precio = merger._get_posicion_inicial(
        wb=type("WorkbookStub", (), {"sheetnames": []})(),
        cod_instrum="7877",
        is_gallo=True,
        for_usd=True,
    )

    assert cantidad == 0
    assert precio == 250.44


def test_usd_non_exterior_precios_iniciales_fallback_is_converted_to_usd():
    merger = _minimal_merger()

    cantidad, precio = merger._get_posicion_inicial(
        wb=type("WorkbookStub", (), {"sheetnames": []})(),
        cod_instrum="95806",
        is_gallo=True,
        for_usd=True,
    )

    assert cantidad == 0
    assert precio == 1.035 / GalloVisualMerger.COTIZACION_INICIO_PERIODO