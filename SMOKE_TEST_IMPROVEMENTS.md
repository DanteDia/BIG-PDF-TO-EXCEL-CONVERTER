# Smoke Test Improvements Tracker

Each section documents what improvements a smoke test baseline protects, so regressions can be traced back to the original fix.

---

## Sigal 10374 — `SMOKE_BASELINE/SIGAL_20260414_APPROVED/`

**Baseline**: 21 sheets, 382,647 cells  
**Script**: `smoke_test_sigal.py`

### Improvements Protected

| ID | Fix | Species / Boletos | What Changed | Before Fix | After Fix |
|----|-----|-------------------|--------------|------------|-----------|
| 1 | Column-shift detection for USD cada-100 instruments | B=82749 (esp 81090), B=84357 (esp 81090) | Precio Nominal was wrong because OCR dropped a cell, shifting TC→Precio and Bruto→TC | PN=1280 / PN=1261 (wrong — those were TC values) | PN=0.0011 (correct micro-price) |
| 2 | 1B USD guardrail | B=81990, B=79603 | Bruto USD capped at 1B to prevent impossible values from column-shift residuals | Bruto could reach trillions | Bruto=5686800000 / -5722800000 (correct) |

### Key Invariants
- B=79603: Precio Nominal=1.4217, Bruto=5686800000
- B=81990: Precio Nominal=1.4307, Bruto=-5722800000
- B=82749: Precio Nominal=0.0011, TC=1280, Bruto=2600000
- B=84357: Precio Nominal=0.0011, TC=1261, Bruto=2550000

---

## Canullo 10600 — `SMOKE_BASELINE/CANULLO_20260326_APPROVED/`

**Baseline**: Existing approved baseline  

### Improvements Protected

| ID | Fix | Area | What Changed |
|----|-----|------|--------------|
| 1 | ARS raw price preservation | Resultado Ventas ARS | TP/ON/Letras with low nominal price not divided by 100 again |
| 2 | Rentas/dividendos currency routing | Rentas Dividendos ARS/USD | Source sheet (ARS/USD) wins for rentas placement |

### Key Invariants
- Tenaris dividend stays in ARS despite USD-like text
- Rentas Dividendos ARS contains peso-sourced rentas only

---

## Prida 10488 — `SMOKE_BASELINE/PRIDA_????????_APPROVED/` *(pending)*

**Baseline**: TBD  
**Script**: TBD

### Improvements Protected

| ID | Fix | Species | What Changed | Before Fix | After Fix |
|----|-----|---------|--------------|------------|-----------|
| 1 | TRF TITULOS treated as COMPRA for PPP | 81086, 81090, 81092, 81274 | Title transfers now enter running stock as buys, fixing cost basis | TRF skipped → no cost basis → wrong resultado (loss instead of gain) | TRF enters PPP → correct cost → correct resultado (gain) |

### Key Invariants
- *(to be filled after implementation)*
