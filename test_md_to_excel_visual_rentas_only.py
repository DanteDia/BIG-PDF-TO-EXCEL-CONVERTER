import textwrap

from pdf_converter.datalab.md_to_excel import MarkdownTableParser


def test_visual_rentas_only_report_is_detected_as_visual():
    markdown = textwrap.dedent(
        r"""
        REPORTE DE GANANCIAS / Periodo Junio 1 - Diciembre 31, 2025
        13969 - SZYCHOWSKI, MARIA INES

        ### Rentas y Dividendos

        | Concertación | Liquidación | Nro. NDC | Tipo Operación | Cantidad | Moneda | Tipo de Cambio | Gastos | Importe |
        |---|---|---|---|---|---|---|---|---|
        | <b>USD</b> | | | | | | | | |
        | <b>Rentas</b> | | | | | | | | |
        | <b>Títulos Públicos</b> | | | | | | | | |
        | <b>BONO REP. ARGENTINA USD STEP UP 2030 - Dolar MEP (local) / 5.921</b> | | | | | | | | |
        | 10/7/2025 | 10/7/2025 | 10709 | Renta | | Dolar MEP (local) | 1,0000000000 | 0,04 | 25,77 |
        | <b>BONOS REP. ARG. U$S STEP UP V.09/07/30 - Dolar Cable (exterior) / 81.086</b> | | | | | | | | |
        | 10/7/2025 | 10/7/2025 | 12498 | Renta | | Dolar Cable (exte | 1,0000000000 | | 3,46 |
        | <b>Total Rentas</b> | | | | | | | | <b>29,23</b> |
        """
    )

    parser = MarkdownTableParser(markdown)
    tables = parser.parse()

    assert parser.format_type == "visual"
    assert "Rentas Dividendos USD" in tables
    rentas_usd = tables["Rentas Dividendos USD"]
    assert len(rentas_usd.rows) == 2
    assert rentas_usd.rows[0][0] == "BONO REP. ARGENTINA USD STEP UP 2030 - Dolar MEP (local)"
    assert rentas_usd.rows[0][1] == "5921"
    assert rentas_usd.rows[0][2] == "Rentas"
    assert rentas_usd.rows[0][3] == "Títulos Públicos"
    assert rentas_usd.rows[0][7] == "Renta"
    assert rentas_usd.rows[0][12] == "25,77"
    assert rentas_usd.rows[1][0] == "BONOS REP. ARG. U$S STEP UP V.09/07/30 - Dolar Cable (exterior)"
    assert rentas_usd.rows[1][1] == "81086"
    assert rentas_usd.rows[1][12] == "3,46"


def test_visual_cauciones_preserves_repeated_rows_across_fragments():
    markdown = textwrap.dedent(
        r"""
        REPORTE DE GANANCIAS / Periodo Junio 1 - Diciembre 31, 2025
        13056 - AGUIAR, JUAN MARTIN

        ### Cauciones tomadoras

        | Concertación | Pla | Liquidación | Operación | # Boleto | Contado | Futuro | Tipo de cambio | Tasa (%) | Interés Bruto | Interés Devenga | Aranceles | Derechos | Costo financiero |
        |---|---|---|---|---|---|---|---|---|---|---|---|---|---|
        | <b>1 / Pesos</b> | | | | | | | | | | | | | |
        | 22/8/2025 | 3 | 25/8/2025 | Apertura Tomador Cauci | 103.226 | 100.000,00 | 100.402,74 | 1,00 | 49,00 | 402,74 | 402,74 | 19,97 | 3,64 | (379,13) |

        ### Cauciones tomadoras

        | Concertación | Pla | Liquidación | Operación | # Boleto | Contado | Futuro | Tipo de cambio | Tasa (%) | Interés Bruto | Interés Devenga | Aranceles | Derechos | Costo financiero |
        |---|---|---|---|---|---|---|---|---|---|---|---|---|---|
        | <b>1 / Pesos</b> | | | | | | | | | | | | | |
        | 22/8/2025 | 3 | 25/8/2025 | Apertura Tomador Cauci | 103.226 | 100.000,00 | 100.402,74 | 1,00 | 49,00 | 402,74 | 402,74 | 19,97 | 3,64 | (379,13) |
        """
    )

    parser = MarkdownTableParser(markdown)
    tables = parser.parse()

    assert "Cauciones Tomadoras" in tables
    assert len(tables["Cauciones Tomadoras"].rows) == 2
