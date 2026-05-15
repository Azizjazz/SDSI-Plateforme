# pptx-report

Convert PowerPoint (`.pptx`) files to clean Markdown with automatic quality evaluation.

Extracts text, tables, charts, hyperlinks, and speaker notes from every slide and produces a well-formatted `.md` file with per-slide breakdown and overall quality scores.

## Quick Start

```bash
# 1. Clone & enter
git clone https://github.com/anomalyco/pptx-report.git
cd pptx-report

# 2. Create virtual environment
python -m venv .venv

# 3. Activate it
#    Windows:
.venv\Scripts\activate
#    macOS / Linux:
source .venv/bin/activate

# 4. Install dependency
pip install python-pptx
```

## Usage

```bash
python generate_report.py presentation.pptx
```

This creates `presentation.md` in the same directory, containing:

- Extracted text from every slide
- Formatted markdown tables
- Chart descriptions with data series
- Hyperlinks as `[text](url)`
- Speaker notes
- **Evaluation section** with quality metrics and overall score

### Options

| Flag | Description |
|------|-------------|
| `-o output.md` | Custom output path |

### Example

```bash
python generate_report.py slides.pptx -o report/slides.md
```

## Output Example

```markdown
# slides

> Generated 2026-05-15 14:30 · 12 slides

---

## Slide 1

Project Overview

### Tables

| Name | Status |
|------|--------|
| Alpha | Done |
| Beta | In Progress |

---

## Evaluation

| Metric | Value |
|--------|-------|
| Total Slides | 12 |
| Total Characters | 8450 |
| Coverage | 100% |
| **Overall** | **92.4%** |
```

## Quality Scores

The evaluation section scores each extraction:

| Score | Weight | What it measures |
|-------|--------|------------------|
| **Coverage** | 50% | % of non-empty slides |
| **Text Density** | 30% | Characters extracted vs expected minimum |
| **Richness** | 20% | Tables, charts, links, and notes per slide |

Each slide is evaluated individually — empty slides, missing tables, or sparse text
are reflected in the final score.

## Requirements

- Python 3.10+
- [python-pptx](https://python-pptx.readthedocs.io/) — the only dependency

## Project Structure

```
pptx-report/
├── generate_report.py   # Main script (single file, zero configuration)
├── requirements.txt     # Dependency list
└── README.md            # This file
```

## License

MIT
