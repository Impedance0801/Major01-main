import json
import math
from pathlib import Path
from textwrap import fill

import pandas as pd
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
DATA_FILE = ROOT / "methodology_comparison_data.json"
OUTPUT_DIR = ROOT / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


KPI_LABELS = {
    "throughput_output": "Throughput /\nOutput",
    "availability_downtime": "Availability /\nDowntime",
    "waste_reduction": "Waste\nReduction",
    "lead_time_flow": "Lead Time /\nFlow",
    "inventory_wip": "Inventory /\nWIP",
    "quality_defect": "Quality /\nDefect",
    "labor_human": "Labor /\nHuman",
    "cost_resource": "Cost /\nResource",
    "flexibility_responsiveness": "Flexibility /\nResponsiveness",
}

CAPABILITY_LABELS = {
    "real_data_grounding": "Real Data\nGrounding",
    "root_cause_specificity": "Root-Cause\nSpecificity",
    "uncertainty_handling": "Uncertainty\nHandling",
    "human_factor_modeling": "Human-Factor\nModeling",
    "real_time_support": "Real-Time\nSupport",
    "sme_accessibility": "SME\nAccessibility",
}

INDUSTRY_LABELS = {
    "beverage_packaging": "Beverage /\nPackaging",
    "food_beverage_smes": "Food and\nBeverage SMEs",
    "apparel_labor_intensive": "Apparel /\nLabor-Intensive",
    "high_mix_aerospace_jobshop": "High-Mix Aerospace\n/ Job Shop",
    "digital_twin_cells": "Digital-Twin\nCells",
    "continuous_process": "Continuous\nProcess",
}

SCORE_COLORS = {
    0: "#f1f5f9",
    1: "#cbd5e1",
    2: "#60a5fa",
    3: "#1d4ed8",
}

CAPABILITY_SCORE_COLORS = {
    0: "#fff7ed",
    1: "#fdba74",
    2: "#f97316",
    3: "#9a3412",
}

INDUSTRY_SCORE_COLORS = {
    0: "#f0fdfa",
    1: "#99f6e4",
    2: "#14b8a6",
    3: "#115e59",
}


def load_data():
    with DATA_FILE.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def build_score_frame(studies, section_name, label_map):
    rows = []
    for study in studies:
        row = {"study": study["short_name"]}
        for key in label_map:
            row[key] = study[section_name][key]
        rows.append(row)
    df = pd.DataFrame(rows).set_index("study")
    return df.rename(columns=label_map)


def save_score_table(df, filename):
    df.reset_index().to_csv(OUTPUT_DIR / filename, index=False)


def save_summary_table(studies):
    rows = []
    for study in studies:
        rows.append(
            {
                "study": study["short_name"],
                "year": study["year"],
                "method_family": study["method_family"],
                "industry_basis": study["industry_basis"],
                "primary_kpis": "; ".join(study["primary_kpis"]),
                "best_fit": study["best_fit"],
                "main_strength": study["main_strength"],
                "main_limitation": study["main_limitation"],
                "source_url": study["source_url"],
            }
        )
    pd.DataFrame(rows).to_csv(OUTPUT_DIR / "methodology_comparison_summary.csv", index=False)


def save_notes_file(payload):
    notes_path = OUTPUT_DIR / "methodology_comparison_notes.txt"
    with notes_path.open("w", encoding="utf-8") as handle:
        handle.write("Scoring scale used in the comparative methodology figures\n")
        handle.write("0 = not explicitly modeled or measured\n")
        handle.write("1 = indirect, qualitative, or weak emphasis\n")
        handle.write("2 = explicit secondary emphasis\n")
        handle.write("3 = core quantitative emphasis or strongest fit\n\n")
        handle.write(f"Comparison basis: {payload['notes']['comparison_basis']}\n\n")
        handle.write(f"Scope note: {payload['notes']['scored_studies_scope']}\n")


def get_font(size, bold=False):
    candidates = []
    if bold:
        candidates.extend(
            [
                r"C:\Windows\Fonts\arialbd.ttf",
                r"C:\Windows\Fonts\calibrib.ttf",
                r"C:\Windows\Fonts\segoeuib.ttf",
            ]
        )
    candidates.extend(
        [
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\calibri.ttf",
            r"C:\Windows\Fonts\segoeui.ttf",
        ]
    )
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size=size)
    return ImageFont.load_default()


def draw_multiline_text(draw, position, text, font, fill, anchor="la", spacing=4, align="left"):
    draw.multiline_text(position, text, font=font, fill=fill, anchor=anchor, spacing=spacing, align=align)


def wrapped_label(text, width):
    return fill(text.replace("\n", " "), width=width).replace("\n", "\n")


def make_heatmap(df, title, subtitle, output_name, color_map):
    row_label_width = 280
    legend_width = 170
    cell_width = 130
    cell_height = 56
    header_height = 150
    footer_height = 36
    margin = 30

    img_width = margin * 2 + row_label_width + df.shape[1] * cell_width + legend_width
    img_height = margin * 2 + header_height + df.shape[0] * cell_height + footer_height

    image = Image.new("RGB", (img_width, img_height), "white")
    draw = ImageDraw.Draw(image)

    title_font = get_font(28, bold=True)
    subtitle_font = get_font(16)
    axis_font = get_font(15, bold=True)
    row_font = get_font(16)
    cell_font = get_font(18, bold=True)
    legend_font = get_font(15)

    draw_multiline_text(draw, (margin, margin), title, title_font, fill="#0f172a")
    draw_multiline_text(draw, (margin, margin + 42), subtitle, subtitle_font, fill="#475569")

    top = margin + header_height
    left = margin + row_label_width

    for col_idx, col_name in enumerate(df.columns):
        text = col_name
        cell_left = left + col_idx * cell_width
        text_x = cell_left + cell_width / 2
        text_y = top - 70
        draw_multiline_text(draw, (text_x, text_y), text, axis_font, fill="#0f172a", anchor="mm", align="center")

    for row_idx, row_name in enumerate(df.index):
        cell_top = top + row_idx * cell_height
        text_y = cell_top + cell_height / 2
        draw_multiline_text(draw, (margin, text_y), row_name, row_font, fill="#0f172a", anchor="lm")

        for col_idx, value in enumerate(df.iloc[row_idx].tolist()):
            cell_left = left + col_idx * cell_width
            box = (
                cell_left,
                cell_top,
                cell_left + cell_width,
                cell_top + cell_height,
            )
            draw.rectangle(box, fill=color_map[int(value)], outline="#cbd5e1", width=1)
            text_color = "white" if int(value) >= 2 else "#0f172a"
            draw.text(
                (cell_left + cell_width / 2, cell_top + cell_height / 2),
                str(int(value)),
                font=cell_font,
                fill=text_color,
                anchor="mm",
            )

    legend_x = left + df.shape[1] * cell_width + 35
    legend_y = top + 8
    draw.text((legend_x, legend_y), "Score legend", font=axis_font, fill="#0f172a")
    legend_items = [
        (0, "0 = not explicit"),
        (1, "1 = weak / indirect"),
        (2, "2 = explicit"),
        (3, "3 = strongest fit"),
    ]
    for idx, (score, label) in enumerate(legend_items):
        y = legend_y + 34 + idx * 34
        draw.rectangle((legend_x, y, legend_x + 24, y + 24), fill=color_map[score], outline="#94a3b8")
        draw.text((legend_x + 34, y + 12), label, font=legend_font, fill="#334155", anchor="lm")

    image.save(OUTPUT_DIR / output_name)


def make_breadth_chart(kpi_df, capability_df):
    kpi_totals = kpi_df.sum(axis=1).sort_values(ascending=False)
    capability_totals = capability_df.loc[kpi_totals.index].sum(axis=1)

    width = 1300
    height = 760
    margin_left = 90
    margin_right = 70
    margin_top = 90
    margin_bottom = 210
    chart_width = width - margin_left - margin_right
    chart_height = height - margin_top - margin_bottom

    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)

    title_font = get_font(28, bold=True)
    subtitle_font = get_font(16)
    axis_font = get_font(16, bold=True)
    label_font = get_font(15)
    value_font = get_font(14, bold=True)

    draw.text((margin_left, 26), "Overall Comparative Breadth of the Scored Methodologies", font=title_font, fill="#0f172a")
    draw.text((margin_left, 64), "Blue bars sum KPI-emphasis scores; orange bars sum methodological-capability scores.", font=subtitle_font, fill="#475569")

    chart_left = margin_left
    chart_top = margin_top
    chart_bottom = chart_top + chart_height
    chart_right = chart_left + chart_width
    draw.line((chart_left, chart_top, chart_left, chart_bottom), fill="#64748b", width=2)
    draw.line((chart_left, chart_bottom, chart_right, chart_bottom), fill="#64748b", width=2)

    max_value = max(float(kpi_totals.max()), float(capability_totals.max()))
    max_axis = int(math.ceil(max_value + 1))
    tick_count = 5
    for idx in range(tick_count + 1):
        value = max_axis * idx / tick_count
        y = chart_bottom - (value / max_axis) * chart_height
        draw.line((chart_left - 6, y, chart_right, y), fill="#e2e8f0", width=1)
        draw.text((chart_left - 12, y), f"{value:.0f}", font=label_font, fill="#475569", anchor="rm")

    group_count = len(kpi_totals)
    group_width = chart_width / group_count
    bar_width = group_width * 0.28

    for idx, study in enumerate(kpi_totals.index):
        kpi_value = float(kpi_totals.iloc[idx])
        capability_value = float(capability_totals.iloc[idx])

        center_x = chart_left + idx * group_width + group_width / 2
        kpi_left = center_x - bar_width - 8
        cap_left = center_x + 8

        kpi_top = chart_bottom - (kpi_value / max_axis) * chart_height
        cap_top = chart_bottom - (capability_value / max_axis) * chart_height

        draw.rectangle((kpi_left, kpi_top, kpi_left + bar_width, chart_bottom), fill="#2563eb", outline="#1d4ed8")
        draw.rectangle((cap_left, cap_top, cap_left + bar_width, chart_bottom), fill="#f97316", outline="#c2410c")

        draw.text((kpi_left + bar_width / 2, kpi_top - 8), f"{kpi_value:.0f}", font=value_font, fill="#1e293b", anchor="ms")
        draw.text((cap_left + bar_width / 2, cap_top - 8), f"{capability_value:.0f}", font=value_font, fill="#1e293b", anchor="ms")

        label = wrapped_label(study, 16)
        draw_multiline_text(draw, (center_x, chart_bottom + 16), label, label_font, fill="#0f172a", anchor="ma", align="center")

    legend_x = chart_right - 250
    legend_y = 105
    draw.rectangle((legend_x, legend_y, legend_x + 22, legend_y + 22), fill="#2563eb")
    draw.text((legend_x + 32, legend_y + 11), "KPI coverage score", font=label_font, fill="#334155", anchor="lm")
    draw.rectangle((legend_x, legend_y + 34, legend_x + 22, legend_y + 56), fill="#f97316")
    draw.text((legend_x + 32, legend_y + 45), "Method capability score", font=label_font, fill="#334155", anchor="lm")

    image.save(OUTPUT_DIR / "methodology_comparison_breadth.png")


def main():
    payload = load_data()
    studies = payload["studies"]

    kpi_df = build_score_frame(studies, "kpis", KPI_LABELS)
    capability_df = build_score_frame(studies, "capabilities", CAPABILITY_LABELS)
    industry_df = build_score_frame(studies, "industries", INDUSTRY_LABELS)

    save_score_table(kpi_df, "methodology_comparison_kpis.csv")
    save_score_table(capability_df, "methodology_comparison_capabilities.csv")
    save_score_table(industry_df, "methodology_comparison_industry_fit.csv")
    save_summary_table(studies)
    save_notes_file(payload)

    make_heatmap(
        kpi_df,
        "KPI Emphasis Across Comparable Lean-Evaluation Methodologies",
        "Comparative scores are based on reported KPI scope in each study.",
        "methodology_kpi_heatmap.png",
        SCORE_COLORS,
    )
    make_heatmap(
        capability_df,
        "Methodological Capability Comparison",
        "Scores summarize comparative strengths in each reported workflow.",
        "methodology_capability_heatmap.png",
        CAPABILITY_SCORE_COLORS,
    )
    make_heatmap(
        industry_df,
        "Industry Fit of the Compared Methodologies",
        "Higher scores indicate a stronger methodological fit for that industry archetype.",
        "methodology_industry_fit_heatmap.png",
        INDUSTRY_SCORE_COLORS,
    )
    make_breadth_chart(kpi_df, capability_df)


if __name__ == "__main__":
    main()
