# PDF to Excel Converter for Financial Reports

Converts Gallo (Resumen Impositivo) and Visual (Broker) financial PDF reports to structured Excel files with AI-powered extraction.

## Features

- 🔍 **Automatic report type detection** (Gallo vs Visual)
- 📄 **Handles large PDFs (50+ pages)** with intelligent chunking
- 🔄 **Context continuity** across page chunks for multi-page sections
- ✅ **Mathematical validation** cross-checks totals against detail sums
- 📊 **Structured Excel output** with proper formatting and multiple sheets
- 🚀 **Batch processing** for multiple PDFs

## Installation

```bash
cd pdf_converter
pip install -r requirements.txt
```

### Required: Set up API Key

Create a `.env` file in the `pdf_converter` directory:

```bash
cp .env.example .env
# Edit .env and add your Anthropic API key
```

Or set as environment variable:

```bash
set ANTHROPIC_API_KEY=your_api_key_here
```

## Usage

### Single File Conversion

```bash
# Auto-detect report type
python app.py "C:\path\to\report.pdf"

# Specify output file
python app.py report.pdf --output result.xlsx

# Specify report type
python app.py report.pdf --type gallo
```

### Batch Processing

```bash
# Convert all PDFs in a folder
python batch_convert.py ./pdfs/ --output-dir ./output/

# Convert specific pattern
python batch_convert.py ./pdfs/ --pattern "*gallo*.pdf"

# Save results to JSON
python batch_convert.py ./pdfs/ --save-results
```

## Supported Report Types

### Gallo (Resumen Impositivo)

Tax summary reports with dynamic sections based on transaction categories.

**Sections:**
- Resultado Totales (summary)
- Tit.Privados Exentos
- Tit.Privados del Exterior
- Renta Fija (Pesos/Dolares)
- Cauciones (Pesos/Dolares)
- FCI, Opciones, Futuros
- Posición Inicial/Final

### Visual (Broker)

Broker trading reports with fixed sections.

**Sections:**
- Resumen
- Boletos
- Resultado de Ventas (ARS/USD)
- Rentas y Dividendos (ARS/USD)
- Posición de Títulos

## How It Works

```
┌─────────────┐     ┌──────────────┐     ┌─────────────┐     ┌──────────────┐
│  Upload PDF │ ──▶ │ Detect Type  │ ──▶ │ Extract by  │ ──▶ │ Post-process │
│             │     │ & Sections   │     │ Section     │     │ & Validate   │
└─────────────┘     └──────────────┘     └─────────────┘     └──────────────┘
                                                                     │
                                                                     ▼
                                                            ┌──────────────┐
                                                            │ Generate XLSX│
                                                            │ Multi-sheet  │
                                                            └──────────────┘
```

### Chunking Strategy (Avoids Token Truncation)

For large PDFs (50+ pages), the tool processes 5 pages at a time:

1. **Section Detection** - Identifies section boundaries in the PDF
2. **Chunked Extraction** - Processes each section in 5-page chunks
3. **Context Continuity** - Maintains entity names (especie) across chunks
4. **Deduplication** - Removes duplicates from chunk overlaps

## Configuration

Edit `config.yaml` to customize:

```yaml
extraction:
  max_pages_per_chunk: 5    # Pages per LLM call
  max_retries: 3            # Retries on failure
  temperature: 0.0          # Deterministic extraction

validation:
  tolerance: 0.01           # Max acceptable difference
```

## Output Format

The Excel file contains:
- **Multiple sheets** - One per section
- **Formatted headers** - Bold, colored headers
- **Number formatting** - Proper decimal alignment
- **Validation sheet** - Cross-check results

## Project Structure

```
pdf_converter/
├── app.py                  # Main entry point
├── batch_convert.py        # Batch processing
├── config.yaml             # Configuration
├── requirements.txt        # Dependencies
├── pdf/
│   └── reader.py           # PDF text extraction
├── llm/
│   ├── client.py           # LLM client with chunking
│   └── prompts.py          # Extraction prompts
├── extractor/
│   ├── context.py          # Extraction context
│   ├── schemas.py          # Column schemas
│   ├── gallo.py            # Gallo extractor
│   └── visual.py           # Visual extractor
├── postprocess/
│   ├── numbers.py          # Number parsing
│   ├── cleanup.py          # Deduplication
│   └── decimals_fix.py     # x100 error fix
├── validation/
│   ├── gallo.py            # Gallo validation
│   └── visual.py           # Visual validation
└── export/
    └── excel_writer.py     # Excel generation
```

## Troubleshooting

### "ANTHROPIC_API_KEY not set"

Set your API key in `.env` file or as environment variable.

### "OCR detected as needed"

The PDF appears to be scanned. Install OCR dependencies:

```bash
pip install pdf2image pytesseract
# Also install Tesseract OCR: https://github.com/tesseract-ocr/tesseract
```

### "Validation failed"

Check the Validation sheet in the Excel output for discrepancies.
This usually indicates missing data or extraction errors.

## License

MIT
