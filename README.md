# pptx-report

Convert PowerPoint (`.pptx`) to **HTML** + **Markdown** reports with quality evaluation. Compares 3 extraction engines side-by-side.

| Engine | Format | |
|--------|--------|-|
| **python-pptx** | Plain text | Always available |
| **MarkItDown** | Markdown | Optional (`pip install markitdown`) |
| **pptx2md** | Markdown | Optional (`pip install pptx2md`) |

## Quick Start

```bash
git clone https://github.com/Azizjazz/SDSI-Plateforme.git
cd SDSI-Plateforme
python -m venv .venv

# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install python-pptx
```

## Usage

```bash
python generate_report.py presentation.pptx
```

Output (in `outputs/presentation/`):

```
outputs/{name}/
├── report.html     # Interactive HTML: scores + side-by-side slide viewer
└── report.md       # Clean Markdown: extracted text + evaluation tables
```

### Options

| Flag | Description |
|------|-------------|
| `-o DIR` | Output directory (default: `outputs/`) |
| `--no-md` | Skip MarkItDown extraction |
| `--no-p2` | Skip pptx2md extraction |

### Examples

```bash
# Basic
python generate_report.py slides.pptx

# Custom output dir, skip optional tools
python generate_report.py slides.pptx -o my_report --no-md --no-p2
```

## HTML Report

Two tabs:
- **Scores** — 3 score cards with weighted metrics (Content 35%, Cleanliness 30%, Coverage 35%), formula explanations, progress bars
- **Slide-by-Slide** — Navigate with Prev/Next or arrow keys, 3 columns comparing tool outputs, speaker notes panel

## Markdown Report

Per-slide extraction (text, tables, charts, hyperlinks, notes) plus evaluation section:

```
## Evaluation

### Extraction Stats
| Total Slides | 90 | Tables | 45 | ...

### Tool Comparison Scores
| Score | python-pptx | MarkItDown | pptx2md |
| Content (35%) | 32.2% | 100% | 100% |

### Quality Score
| Coverage (50%) | 98.9% | Overall: 85.1% |
```

## Requirements

- Python 3.10+
- `python-pptx` (required)
- `markitdown` (optional, for comparison)
- `pptx2md` (optional, for comparison)

## Project Structure

```
├── generate_report.py   # Single script, zero config
├── requirements.txt
├── pyproject.toml
└── README.md
```

## License

MIT
