"""
Build a PPT-like summary table from an Excel source sheet and save it as an image.

The script:
1) Detects the first sheet whose name starts with '8_' (or a user-specified sheet).
2) Reconstructs the intermediate table (Evol o dif cant) logic from "00 -Intervalos de Confianza".
3) Produces a PPT-style table (like the provided screenshot) as a PNG image.

Usage:
    python build_ppt_image.py input.xlsx --sheet 8_Pan_Bimbo --output output.png

Design goals:
- Modular functions so it can be embedded into a larger PPTX generator.
- Resilient to additional brands/blocks in the input sheet.
"""

from __future__ import annotations

import argparse
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import matplotlib.pyplot as plt
import pandas as pd

# ---------- Data structures ----------


@dataclass
class MonthlyBlock:
    brand: str
    rows: pd.DataFrame  # columns: 0=date, 1=Weighted PENET, 2=BUYERS


@dataclass
class AggBlock:
    brand: str
    block: pd.DataFrame  # columns: 4=label, 5=MAT Aug-24, 6=MAT Aug-25


@dataclass
class PptRow:
    marca: str
    mat_sep25: float
    lim_inf: float
    lim_sup: float
    compradores_prom: float
    pen12m: float
    nivel_pen: str
    error_muestral: float


# ---------- Parsing helpers ----------


def parse_monthly_blocks(df: pd.DataFrame) -> List[MonthlyBlock]:
    """Find brand-delimited monthly data blocks in columns 0-2."""
    brand_rows = df[(df[0].notna()) & df[1].isna() & df[2].isna()].index.tolist()
    blocks: List[MonthlyBlock] = []
    for i, start in enumerate(brand_rows):
        end = brand_rows[i + 1] if i + 1 < len(brand_rows) else len(df)
        data = df.loc[start + 1 : end - 1]
        data = data[pd.to_numeric(data[1], errors="coerce").notna()]
        blocks.append(MonthlyBlock(brand=str(df.loc[start, 0]).strip(), rows=data))
    return blocks


def parse_agg_blocks(df: pd.DataFrame) -> List[AggBlock]:
    """Find brand-delimited aggregate blocks in columns 4-6."""
    agg_brand_rows = df[
        (df[4].notna())
        & (~df[4].astype(str).str.contains("Weighted"))
        & (~df[4].astype(str).str.contains("table"))
        & df[5].isna()
        & df[6].isna()
    ].index.tolist()
    blocks: List[AggBlock] = []
    for i, start in enumerate(agg_brand_rows):
        end = agg_brand_rows[i + 1] if i + 1 < len(agg_brand_rows) else len(df)
        block = df.loc[start + 1 : end - 1]
        block = block[df[4].notna()]
        blocks.append(AggBlock(brand=str(df.loc[start, 4]).strip(), block=block))
    return blocks


def extract_metrics(block: pd.DataFrame) -> Dict[str, Tuple[float, float]]:
    """Map metric label -> (MAT Aug-24, MAT Aug-25)."""
    metrics: Dict[str, Tuple[float, float]] = {}
    for _, row in block.iterrows():
        label = str(row[4]).strip().replace("Weighted ", "").strip().upper()
        metrics[label] = (float(row[5]), float(row[6]))
    return metrics


# ---------- Computation ----------


def classify_penetration(p: float) -> str:
    """Classify penetration ratio p (0-1)."""
    if p <= 0:
        return "Fuera de rango"
    if 0.01 <= p <= 0.10:
        return "Marca Pequeña (1% a 10%)"
    if 0.10 < p <= 0.30:
        return "Marca Mediana (11% a 30%)"
    if 0.30 < p <= 0.50:
        return "Marca Grande (31% a 50%)"
    if 0.50 < p <= 0.99:
        return "Super Marca (51% a 99%)"
    return "Fuera de rango"


def compute_ppt_rows(
    monthly_blocks: Iterable[MonthlyBlock], agg_blocks: Iterable[AggBlock]
) -> List[PptRow]:
    """Zip monthly + aggregate blocks to produce PPT rows."""
    rows: List[PptRow] = []
    for mblock, ablock in zip(monthly_blocks, agg_blocks):
        monthly = mblock.rows
        last12_pen = monthly[1].astype(float).tail(12)
        last12_buy = monthly[2].astype(float).tail(12)
        pen_avg = last12_pen.mean()
        buyers_avg = last12_buy.mean()

        metrics = extract_metrics(ablock.block)
        vol1, vol2 = metrics["R_VOL1"]
        pen1, pen2 = metrics["PENET"]
        freq1, freq2 = metrics["FREQ"]
        _, hh = metrics.get("HHOLDS", (None, None))

        dif = vol2 - vol1
        b = (pen1 + pen2) / 200
        w = (freq1 + freq2) / 2
        m_val = b * w
        se_rel = math.sqrt(2 / (m_val * hh))
        avg_vol = (vol1 + vol2) / 2
        int_dif = 2 * se_rel * avg_vol
        lim_cant2_minus = vol1 + (dif - int_dif)
        lim_cant2_plus = vol1 + (dif + int_dif)

        evol = (vol2 / vol1 - 1) * 100
        lim_evol_minus = (lim_cant2_minus / vol1 - 1) * 100
        lim_evol_plus = (lim_cant2_plus / vol1 - 1) * 100

        p_ratio = pen_avg / 100
        err = 1.96 * math.sqrt((p_ratio * (1 - p_ratio)) / hh) / p_ratio
        nivel = classify_penetration(p_ratio)

        rows.append(
            PptRow(
                marca=mblock.brand,
                mat_sep25=evol,
                lim_inf=lim_evol_minus,
                lim_sup=lim_evol_plus,
                compradores_prom=buyers_avg,
                pen12m=pen_avg,
                nivel_pen=nivel,
                error_muestral=err,
            )
        )
    return rows


# ---------- Rendering ----------


def render_ppt_table(rows: List[PptRow], output_path: Path) -> None:
    """Render the PPT table to an image using matplotlib.

    Columns shown: Marca, %VAR (MAT/LIM INF/LIM SUP), Nivel de penetracion, Error muestral.
    (Compradores promedio y Penetracion 12m se calculan pero se ocultan en la imagen.)
    """
    # Colors (hex) roughly matching the screenshot
    COLOR_HEADER = "#1F4E78"
    COLOR_B = "#DCE6F1"
    COLOR_C = "#F2DCDB"
    COLOR_D = "#E2EFDA"
    COLOR_G = "#DCE6F1"  # level
    COLOR_H = "#E2EFDA"  # error

    # Single-row header
    header_bottom = [
        "Marca",
        "% VAR MAT Sep-25",
        "% VAR LIM INFERIOR",
        "% VAR LIM SUPERIOR",
        "Nivel de penetracion",
        "Error muestral",
    ]

    data = [
        [
            r.marca,
            f"{r.mat_sep25:.1f}",
            f"{r.lim_inf:.1f}",
            f"{r.lim_sup:.1f}",
            r.nivel_pen,
            f"{r.error_muestral:.1%}",
        ]
        for r in rows
    ]

    fig_height = 1.4 + 0.55 * len(rows)
    fig, ax = plt.subplots(figsize=(12, fig_height))
    ax.axis("off")

    table = ax.table(
        cellText=[header_bottom] + data,
        colLabels=None,
        cellLoc="center",
        loc="center",
    )

    # Font sizing / scaling
    table.auto_set_font_size(False)
    table.set_fontsize(13)
    table.scale(1.0, 1.15)

    # Column widths to mimic the screenshot proportions, omitting hidden columns
    widths = [0.24, 0.14, 0.14, 0.14, 0.24, 0.14]
    for i, w in enumerate(widths):
        table.auto_set_column_width(i)
        table._cells[(0, i)].set_width(w)
    # Edge styling for all cells
    for (r, c), cell in table._cells.items():
        cell.set_edgecolor("black")
        cell.set_linewidth(1.5 if r == 0 else 1.2)

    # Styling headers (single row)
    for col in range(len(header_bottom)):
        cell = table[0, col]
        cell.set_facecolor(COLOR_HEADER)
        cell.set_text_props(color="white", weight="bold")

    # Styling body rows
    start_body = 1
    for r_idx in range(start_body, len(data) + start_body):
        table[r_idx, 1].set_facecolor(COLOR_B)
        table[r_idx, 2].set_facecolor(COLOR_C)
        table[r_idx, 3].set_facecolor(COLOR_D)
        table[r_idx, 4].set_facecolor(COLOR_G)
        table[r_idx, 5].set_facecolor(COLOR_H)
        # Green for error
        table[r_idx, 5].set_text_props(color="#006100")

    plt.tight_layout()
    fig.savefig(output_path, dpi=200, bbox_inches="tight")
    plt.close(fig)


# ---------- Orchestration ----------


def find_sheet_name(xlsx: Path, user_sheet: Optional[str]) -> str:
    xls = pd.ExcelFile(xlsx)
    if user_sheet:
        if user_sheet not in xls.sheet_names:
            raise ValueError(f"Sheet '{user_sheet}' not found. Available: {xls.sheet_names}")
        return user_sheet
    for name in xls.sheet_names:
        if str(name).startswith("8_"):
            return name
    raise ValueError("No sheet starting with '8_' found. Specify --sheet.")


def build_image_from_excel(xlsx: Path, sheet: Optional[str], output: Path) -> None:
    sheet_name = find_sheet_name(xlsx, sheet)
    df = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)
    monthly_blocks = parse_monthly_blocks(df)
    agg_blocks = parse_agg_blocks(df)
    if not monthly_blocks or not agg_blocks:
        raise ValueError("No brand blocks detected; check sheet structure.")
    rows = compute_ppt_rows(monthly_blocks, agg_blocks)
    render_ppt_table(rows, output)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate PPT-like table image from an Excel sheet.")
    parser.add_argument("xlsx", type=Path, help="Input Excel file.")
    parser.add_argument("--sheet", help="Sheet name (default: first starting with '8_').")
    parser.add_argument("--output", type=Path, help="Output PNG path (default: <xlsx>_ppt.png).")
    args = parser.parse_args()

    output = args.output or args.xlsx.with_name(f"{args.xlsx.stem}_ppt.png")
    build_image_from_excel(args.xlsx, args.sheet, output)
    print(f"Saved image to {output}")


if __name__ == "__main__":
    main()
