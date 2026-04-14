import math
import sys
from openpyxl import Workbook

sys.path.insert(0, '.')

from pdf_converter.datalab.postprocess import postprocess_visual_workbook


BOLETOS_HEADERS = [
    'Tipo de Instrumento', 'Concertación', 'Liquidación', 'Nro. Boleto', 'Moneda',
    'Tipo Operación', 'Cod.Instru', 'Instrumento', 'Cantidad', 'Precio',
    'Tipo Cambio', 'Bruto', 'Interés', 'Gastos', 'Neto'
]

RESULTADO_HEADERS = [
    'Tipo de Instrumento', 'Instrumento', 'Cod.Instrum', 'Concertación', 'Liquidación',
    'Moneda', 'Tipo Operación', 'Cantidad', 'Precio', 'Bruto', 'Interés',
    'Tipo de Cambio', 'Gastos', 'IVA', 'Resultado'
]


def _find_boleto_row(ws, boleto):
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, 4).value == boleto:
            return row
    raise AssertionError(f'Boleto not found: {boleto}')


def _find_resultado_row(ws, code, concertacion, moneda, operacion):
    for row in range(2, ws.max_row + 1):
        if (
            ws.cell(row, 3).value == code
            and ws.cell(row, 4).value == concertacion
            and ws.cell(row, 6).value == moneda
            and ws.cell(row, 7).value == operacion
        ):
            return row
    raise AssertionError(f'Resultado row not found: {(code, concertacion, moneda, operacion)}')


def test_visual_quantity_rescue_from_resultado_anchors():
    wb = Workbook()
    ws_boletos = wb.active
    ws_boletos.title = 'Boletos'
    ws_boletos.append(BOLETOS_HEADERS)
    ws_boletos.append([
        'Letras del Tesoro nac', '10/7/2025', '10/7/2025', 80261, 'Pesos',
        'Venta Contado Continuo No', 9305, 'LETRAS DEL TESO', -2173913, 1.4568,
        1, -3166999.9, 0, 0, -3166999.9
    ])
    ws_boletos.append([
        'Letras del Tesoro nac', '10/7/2025', '10/7/2025', 80260, 'Dolar Ca',
        'Compra Contado Continuo', 9305, 'LETRAS DEL TESO', 2173913, 0.0012,
        1255, 2500000, 0, 0, 2500000
    ])
    ws_boletos.append([
        'Letras del Tesoro nac', '24/7/2025', '24/7/2025', 88186, 'Pesos',
        'Venta Contado Continuo No', 9305, 'LETRAS DEL TESO', '(2.217.391.3', 1.4544,
        1, '(3.224.973.9', 0, 0, '(3.224.973.9'
    ])
    ws_boletos.append([
        'Letras del Tesoro nac', '24/7/2025', '24/7/2025', 88185, 'Dolar Ca',
        'Compra Contado Continuo', 9305, 'LETRAS DEL TESO', '2.217.391.3', 0.0012,
        1258, 2550000, 0, 0, 2550000
    ])
    ws_boletos.append([
        'Letras del Tesoro nac', '5/9/2025', '5/9/2025', 111965, 'Pesos',
        'Venta Contado Continuo No', 9301, 'LETRAS DEL TESO', -352670.14, 1.574,
        1, -555102.8, 0, 0, -555102.8
    ])
    ws_boletos.append([
        'Letras del Tesoro nac', '29/8/2025', '29/8/2025', 107310, 'Pesos',
        'Compra Contado Continuo', 9301, 'LETRAS DEL TESO', 352670140, 1.5627,
        1, 551117627.78, 0, 0, 551117627.78
    ])
    ws_boletos.append([
        'Obligaciones Negociables', '12/9/2025', '12/9/2025', 116455, 'Pesos',
        'Venta Contado Continuo No', 58449, 'ON BCO GALICIA C', -1285714.2, 1.0197,
        1, -1311029.9, 0, 0, -1311029.9
    ])
    ws_boletos.append([
        'Obligaciones Negociables', '12/9/2025', '12/9/2025', 116454, 'Dolar Ca',
        'Compra Contado Continuo', 58449, 'ON BCO GALICIA C', 1285714.2, 0.0007,
        1432, 900000, 0, 0, 900000
    ])
    ws_boletos.append([
        'Letras del Tesoro nac', '17/6/2025', '17/6/2025', 68292, 'Pesos',
        'Venta Contado Continuo No', 9295, 'LT REP ARGENTIN', -1632072.1, 1.4439,
        1, -2356499.9, 0, 0, -2356499.9
    ])
    ws_boletos.append([
        'Letras del Tesoro nac', '17/6/2025', '17/6/2025', 68291, 'Dolar Ca',
        'Compra Contado Continuo', 9295, 'LT REP ARGENTIN', 1632072.1, 0.0012,
        1182, 1991128, 0, 0, 1991128
    ])
    ws_boletos.append([
        'Títulos Públicos', '29/7/2025', '29/7/2025', 90061, 'Dolar ME',
        'Venta Contado Continuo No', 5921, 'BONO REP. ARGE', -2080213, 0.6009,
        1295, -1249999.95, 0, 0, -1249999.95
    ])
    ws_boletos.append([
        'Títulos Públicos', '29/7/2025', '29/7/2025', 90060, 'Dolar Ca',
        'Compra Contado Continuo', 5921, 'BONO REP. ARGE', 2080213, 0.6,
        1295, 1248127.8, 0, 0, 1248127.8
    ])
    ws_boletos.append([
        'Títulos Públicos', '21/8/2025', '21/8/2025', 102074, 'Dolar ME',
        'Venta Contado Continuo No', 5921, 'BONO REP. ARGE', -4500000, 0.6006,
        1301, -2702700.09, 0, 0, -2702700.09
    ])
    ws_boletos.append([
        'Títulos Públicos', '21/8/2025', '21/8/2025', 102073, 'Dolar Ca',
        'Compra Contado Continuo', 5921, 'BONO REP. ARGE', 4500000, 0.6,
        1301, 2700000, 0, 0, 2700000
    ])
    ws_boletos.append([
        'Títulos Públicos', '3/10/2025', '3/10/2025', 128850, 'Dolar ME',
        'Venta Contado Continuo No', 5921, 'BONO REP. ARGE', -462963, 0.5508,
        1450, -255000.02, 0, 0, -255000.02
    ])
    ws_boletos.append([
        'Títulos Públicos', '3/10/2025', '3/10/2025', 128849, 'Dolar Ca',
        'Compra Contado Continuo', 5921, 'BONO REP. ARGE', 462963, 0.54,
        1424.5, 250000.02, 0, 0, 250000.02
    ])

    ws_result = wb.create_sheet('Resultado Ventas ARS')
    ws_result.append(RESULTADO_HEADERS)
    ws_result.append([
        'Letras del Tesoro nac', 'LETRAS DEL TESORO CAP $ V31/07/25', 9305, '10/7/2025',
        '10/7/2025', 'Pesos', 'Venta Contado Continuo', '(2.173.913.043,)', 1.4568,
        '3.166.999.999,3', 0, 1, 0, 0, '29.499.978,25'
    ])
    ws_result.append([
        'Letras del Tesoro nac', 'LETRAS DEL TESORO CAP $ V31/07/25', 9305, '10/7/2025',
        '10/7/2025', 'Dolar Cable', 'Compra Contado Continuo', '2.173.913.043,0', 0.0012,
        '(3.137.500.000,)', 0, '1.255,00000000', 0, 0, 0
    ])
    ws_result.append([
        'Letras del Tesoro nac', 'LETRAS DEL TESORO CAP $ V31/07/25', 9305, '24/7/2025',
        '24/7/2025', 'Pesos', 'Venta Contado Continuo', '(2.217.391.304,)', 1.4544,
        '3.224.973.912,5', 0, 1, 0, 0, '18.480.703,04'
    ])
    ws_result.append([
        'Letras del Tesoro nac', 'LETRAS DEL TESORO CAP $ V31/07/25', 9305, '24/7/2025',
        '24/7/2025', 'Dolar Cable', 'Compra Contado Continuo', '2.217.391.304,0', 0.0012,
        '(3.207.900.000,)', 0, '1.258,00000000', 0, 0, 0
    ])
    ws_result.append([
        'Letras del Tesoro nac', 'LETRAS DEL TESORO CAP $ V12/09/25', 9301, '29/8/2025',
        '29/8/2025', 'Pesos', 'Compra Contado Continuo', '352.670.140.000,0', 1.5627,
        '(551.117.627,78)', 0, 1, 0, 0, 0
    ])
    ws_result.append([
        'Letras del Tesoro nac', 'LETRAS DEL TESORO CAP $ V12/09/25', 9301, '5/9/2025',
        '5/9/2025', 'Pesos', 'Venta Contado Continuo', '(352.670.140,00)', 1.574,
        '555.102.800,36', 0, 1, 0, 0, '3.985.172,58'
    ])
    ws_result.append([
        'Obligaciones Negociables', 'ON BCO GALICIA CL.21 V10/02/26 $ CG', 58449, '12/9/2025',
        '12/9/2025', 'Pesos', 'Venta Contado Continuo', '(1.285.714.285,0)', 1.0197,
        '1.311.029.999,2', 0, 1, 0, 0, '22.229.987,13'
    ])
    ws_result.append([
        'Obligaciones Negociables', 'ON BCO GALICIA CL.21 V10/02/26 $ CG', 58449, '12/9/2025',
        '12/9/2025', 'Dolar Cable', 'Compra Contado Continuo', '1.285.714.285,0', 0.0007,
        '(1.288.800.000,)', 0, '1.432,00000000', 0, 0, 0
    ])

    ws_result_usd = wb.create_sheet('Resultado Ventas USD')
    ws_result_usd.append(RESULTADO_HEADERS)
    ws_result_usd.append([
        'Títulos Públicos', 'BONO REP. ARGENTINA USD STEP UP 2030', 5921, '29/7/2025',
        '29/7/2025', 'Dolar MEP (I)', 'Venta Contado Continuo', -2080213000, 0.6009,
        1249999.95, 0, 1, 0, 0, 1872.11
    ])
    ws_result_usd.append([
        'Títulos Públicos', 'BONO REP. ARGENTINA USD STEP UP 2030', 5921, '29/7/2025',
        '29/7/2025', 'Dolar Cable', 'Compra Contado Continuo', 2080213000, 0.6,
        -1248127.84, 0, 1, 0, 0, 0
    ])
    ws_result_usd.append([
        'Títulos Públicos', 'BONO REP. ARGENTINA USD STEP UP 2030', 5921, '21/8/2025',
        '21/8/2025', 'Dolar MEP (I)', 'Venta Contado Continuo', -4500000000, 0.6006,
        2702700.09, 0, 1, 0, 0, 2700.14
    ])
    ws_result_usd.append([
        'Títulos Públicos', 'BONO REP. ARGENTINA USD STEP UP 2030', 5921, '21/8/2025',
        '21/8/2025', 'Dolar Cable', 'Compra Contado Continuo', 4500000000, 0.6,
        -2699999.97, 0, 1, 0, 0, 0
    ])
    ws_result_usd.append([
        'Títulos Públicos', 'BONO REP. ARGENTINA USD STEP UP 2030', 5921, '3/10/2025',
        '3/10/2025', 'Dolar MEP (I)', 'Venta Contado Continuo', -462963000, 0.5508,
        255000.03, 0, 1, 0, 0, 5000.01
    ])
    ws_result_usd.append([
        'Títulos Públicos', 'BONO REP. ARGENTINA USD STEP UP 2030', 5921, '3/10/2025',
        '3/10/2025', 'Dolar Cable', 'Compra Contado Continuo', 462963000, 0.54,
        -250000.02, 0, 1, 0, 0, 0
    ])
    ws_result.append([
        'Letras del Tesoro nac', 'LT REP ARGENTINA CAP V30/06/25 $ CG', 9295, '17/6/2025',
        '17/6/2025', 'Pesos', 'Venta Contado Continuo', -1632072131, 1.4439,
        2356499987.7, 0, 1, 0, 0, 2986692
    ])
    ws_result.append([
        'Letras del Tesoro nac', 'LT REP ARGENTINA CAP V30/06/25 $ CG', 9295, '17/6/2025',
        '17/6/2025', 'Dolar Cable', 'Compra Contado Continuo', 1632072131, 0.0012,
        -2353513296, 0, 1182, 0, 0, 0
    ])

    postprocess_visual_workbook(wb)

    row_80261 = _find_boleto_row(ws_boletos, 80261)
    row_80260 = _find_boleto_row(ws_boletos, 80260)
    row_88186 = _find_boleto_row(ws_boletos, 88186)
    row_88185 = _find_boleto_row(ws_boletos, 88185)
    row_111965 = _find_boleto_row(ws_boletos, 111965)
    row_107310 = _find_boleto_row(ws_boletos, 107310)
    row_116455 = _find_boleto_row(ws_boletos, 116455)
    row_116454 = _find_boleto_row(ws_boletos, 116454)
    row_68292 = _find_boleto_row(ws_boletos, 68292)
    row_68291 = _find_boleto_row(ws_boletos, 68291)
    row_90060 = _find_boleto_row(ws_boletos, 90060)
    row_102073 = _find_boleto_row(ws_boletos, 102073)
    row_128850 = _find_boleto_row(ws_boletos, 128850)
    row_128849 = _find_boleto_row(ws_boletos, 128849)

    assert ws_boletos.cell(row_80261, 9).value == -2173913043
    assert ws_boletos.cell(row_80260, 9).value == 2173913043
    assert ws_boletos.cell(row_88186, 9).value == -2217391304
    assert ws_boletos.cell(row_88185, 9).value == 2217391304
    assert ws_boletos.cell(row_111965, 9).value == -352670140
    assert ws_boletos.cell(row_107310, 9).value == 352670140
    assert ws_boletos.cell(row_116455, 9).value == -1285714.2
    assert ws_boletos.cell(row_116454, 9).value == 1285714.2
    assert ws_boletos.cell(row_68292, 9).value == -1632072.1
    assert ws_boletos.cell(row_68291, 9).value == 1632072.1
    assert ws_boletos.cell(row_90060, 9).value == 2080213
    assert ws_boletos.cell(row_102073, 9).value == 4500000
    assert ws_boletos.cell(row_128850, 9).value == -462963
    assert ws_boletos.cell(row_128849, 9).value == 462963

    assert math.isclose(ws_boletos.cell(row_111965, 12).value, -555102800.36, rel_tol=0, abs_tol=0.01)
    assert math.isclose(ws_boletos.cell(row_111965, 15).value, -555102800.36, rel_tol=0, abs_tol=0.01)
    assert math.isclose(ws_boletos.cell(row_88186, 12).value, -3224973912.5376, rel_tol=0, abs_tol=0.01)
    assert math.isclose(ws_boletos.cell(row_88186, 15).value, -3224973912.5376, rel_tol=0, abs_tol=0.01)

    # The anchor sheets should keep integer quantities instead of inflating zero decimal tails.
    row_result_9305 = _find_resultado_row(ws_result, 9305, '10/7/2025', 'Pesos', 'Venta Contado Continuo')
    row_result_9301 = _find_resultado_row(ws_result, 9301, '5/9/2025', 'Pesos', 'Venta Contado Continuo')
    row_result_58449 = _find_resultado_row(ws_result, 58449, '12/9/2025', 'Pesos', 'Venta Contado Continuo')
    row_result_9295 = _find_resultado_row(ws_result, 9295, '17/6/2025', 'Pesos', 'Venta Contado Continuo')
    assert ws_result.cell(row_result_9305, 8).value == -2173913043
    assert ws_result.cell(row_result_9301, 8).value == -352670140
    assert ws_result.cell(row_result_58449, 8).value == -1285714285
    assert ws_result.cell(row_result_9295, 8).value == -1632072131


if __name__ == '__main__':
    test_visual_quantity_rescue_from_resultado_anchors()
    print('ok')
