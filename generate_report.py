import sys, re
from pathlib import Path
from datetime import datetime
from pptx import Presentation


def extract_pptx(pptx_path):
    prs = Presentation(pptx_path)
    slides = []

    for idx, slide in enumerate(prs.slides, start=1):
        texts, tables, charts, links = [], [], [], []
        notes = ""

        sorted_shapes = sorted(
            slide.shapes, key=lambda s: (int(s.top or 0), int(s.left or 0))
        )

        for shape in sorted_shapes:
            if hasattr(shape, "text") and shape.text.strip():
                texts.append(shape.text.strip())

            if shape.has_table:
                table = shape.table
                rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
                tables.append(
                    {"rows": len(table.rows), "cols": len(table.columns), "data": rows}
                )

            if hasattr(shape, "has_chart") and shape.has_chart:
                try:
                    chart = shape.chart
                    series = []
                    for s in chart.series:
                        series.append(
                            {
                                "name": str(s.name) if s.name is not None else "",
                                "values": [
                                    round(v, 4) if isinstance(v, float) else v
                                    for v in s.values
                                ],
                            }
                        )
                    charts.append({"type": str(chart.chart_type), "series": series})
                except:
                    pass

            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.hyperlink and run.hyperlink.address:
                            links.append(
                                {"text": run.text, "url": run.hyperlink.address}
                            )

        try:
            if slide.has_notes_slide:
                t = slide.notes_slide.notes_text_frame.text.strip()
                if t:
                    notes = t
        except:
            pass

        slides.append(
            {
                "num": idx,
                "texts": texts,
                "tables": tables,
                "charts": charts,
                "links": links,
                "notes": notes,
            }
        )

    return slides


def fmt_markdown(name, slides):
    total = len(slides)
    lines = []
    total_chars = 0
    total_tables = 0
    total_charts = 0
    total_links = 0
    total_notes = 0
    empty_slides = 0
    per_slide_chars = []

    lines.append(f"# {name}")
    lines.append(f"")
    lines.append(
        f"> Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} &middot; {total} slides"
    )
    lines.append(f"")
    lines.append(f"---")
    lines.append(f"")

    for s in slides:
        lines.append(f"## Slide {s['num']}")
        lines.append(f"")

        slide_chars = 0
        has = bool(s["texts"] or s["tables"] or s["charts"] or s["links"])

        if not has and not s["notes"]:
            lines.append(f"*empty*")
            lines.append(f"")
            empty_slides += 1
            per_slide_chars.append(0)
            lines.append(f"---")
            lines.append(f"")
            continue

        for t in s["texts"]:
            lines.append(f"{t}")
            lines.append(f"")
            slide_chars += len(t)

        for t in s["tables"]:
            total_tables += 1
            h = t["data"][0]
            pad = lambda r: r + [""] * (len(h) - len(r))
            lines.append(f"| {' | '.join(h)} |")
            lines.append(f"| {' | '.join(['---'] * len(h))} |")
            for r in t["data"][1:]:
                p = pad(r)
                lines.append(f"| {' | '.join(p)} |")
            lines.append(f"")

        for c in s["charts"]:
            total_charts += 1
            ct = re.sub(r"\((\d+)\)", "", c["type"]).strip()
            lines.append(f"**Chart:** {ct}")
            for sr in c["series"]:
                vals = ", ".join(str(v) for v in sr["values"][:12])
                label = sr["name"] or "Series"
                lines.append(f"- {label}: {vals}")
                if len(sr["values"]) > 12:
                    lines.append(f"  *(+{len(sr['values']) - 12} more)*")
            lines.append(f"")

        for lk in s["links"]:
            total_links += 1
            lines.append(f"- [{lk['text']}]({lk['url']})")
        if s["links"]:
            lines.append(f"")

        if s["notes"]:
            total_notes += 1
            lines.append(f"> **Notes:** {s['notes']}")
            lines.append(f"")

        per_slide_chars.append(slide_chars)
        total_chars += slide_chars
        lines.append(f"---")
        lines.append(f"")

    # ── Evaluation ──────────────────────────────────────────────
    coverage = round((total - empty_slides) / total * 100, 1) if total else 0
    avg = round(total_chars / total, 1) if total else 0
    max_c = max(per_slide_chars) if per_slide_chars else 0
    min_c = (
        min(c for c in per_slide_chars if c > 0)
        if any(c > 0 for c in per_slide_chars)
        else 0
    )
    std = (
        round((sum((c - avg) ** 2 for c in per_slide_chars) / total) ** 0.5, 1)
        if total
        else 0
    )

    lines.append(f"## Evaluation")
    lines.append(f"")
    lines.append(f"| Metric | Value |")
    lines.append(f"|--------|-------|")
    lines.append(f"| Total Slides | {total} |")
    lines.append(f"| Total Characters | {total_chars} |")
    lines.append(f"| Average per Slide | {avg} |")
    lines.append(f"| Max per Slide | {max_c} |")
    lines.append(f"| Min (non-empty) | {min_c} |")
    lines.append(f"| Std Deviation | {std} |")
    lines.append(f"| Total Tables | {total_tables} |")
    lines.append(f"| Total Charts | {total_charts} |")
    lines.append(f"| Total Hyperlinks | {total_links} |")
    lines.append(f"| Slides with Notes | {total_notes} |")
    lines.append(f"| Empty Slides | {empty_slides} |")
    lines.append(f"| Coverage | {coverage}% |")
    lines.append(f"")

    # Content score (0-100): weighted by density, coverage, richness
    density_score = (
        min(100, round(total_chars / max(1, total * 50) * 100, 1)) if total else 0
    )
    richness = total_tables + total_charts + total_links + total_notes
    richness_score = min(100, round(richness / max(1, total) * 50, 1)) if total else 0
    overall = round(0.5 * coverage + 0.3 * density_score + 0.2 * richness_score, 1)

    lines.append(f"### Quality Scores")
    lines.append(f"")
    lines.append(f"| Score | Value |")
    lines.append(f"|-------|-------|")
    lines.append(f"| Coverage | {coverage}% |")
    lines.append(f"| Text Density | {density_score}% |")
    lines.append(f"| Richness | {richness_score}% |")
    lines.append(f"| **Overall** | **{overall}%** |")
    lines.append(f"")

    lines.append(f"---")
    lines.append(
        f"*Generated by [pptx-report](https://github.com/anomalyco/pptx-report)*"
    )
    lines.append(f"")

    return "\n".join(lines), {
        "total": total,
        "chars": total_chars,
        "tables": total_tables,
        "charts": total_charts,
        "links": total_links,
        "notes": total_notes,
        "empty": empty_slides,
        "coverage": coverage,
        "overall": overall,
    }


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert PPTX to Markdown with quality evaluation"
    )
    parser.add_argument("input", help="Path to .pptx file")
    parser.add_argument(
        "-o", "--output", help="Output markdown path (default: <name>.md)"
    )
    args = parser.parse_args()

    path = Path(args.input)
    if not path.exists():
        print(f"Error: '{path}' not found", file=sys.stderr)
        sys.exit(1)

    print(f"Reading: {path.name}")
    slides = extract_pptx(str(path.resolve()))
    print(f"  {len(slides)} slides extracted")

    md, stats = fmt_markdown(path.stem, slides)
    out = Path(args.output) if args.output else path.with_suffix(".md")
    out.write_text(md, encoding="utf-8")

    print(f"  Output: {out.resolve()}")
    print(
        f"  Chars: {stats['chars']}  |  Tables: {stats['tables']}  |  Charts: {stats['charts']}"
    )
    print(
        f"  Links: {stats['links']}  |  Notes: {stats['notes']}  |  Empty: {stats['empty']}"
    )
    print(f"  Coverage: {stats['coverage']}%  |  Overall Score: {stats['overall']}%")


if __name__ == "__main__":
    main()
