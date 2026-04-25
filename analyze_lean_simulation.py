import json
from pathlib import Path

import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont


DATA_FILE = Path("production_raw.xlsx")
OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

SCENARIO_ORDER = ["baseline", "s1_5s_visual", "s2_tpm", "s3_integrated"]
SCENARIO_LABELS = ["Baseline", "5S + Visual", "TPM", "Integrated Lean"]


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


def load_data():
    daily = pd.read_excel(DATA_FILE, sheet_name="daily_operation_summary")
    hourly = pd.read_excel(DATA_FILE, sheet_name="hourly_operation_breakdown")
    downtime = pd.read_excel(DATA_FILE, sheet_name="downtime_event_log")
    production = pd.read_excel(DATA_FILE, sheet_name="processed_hourly")

    for df in (daily, hourly, downtime, production):
        if "date" in df.columns:
            df["date"] = pd.to_datetime(df["date"])

    downtime["downtime_min"] = pd.to_timedelta(downtime["downtime_time"].astype(str)).dt.total_seconds() / 60.0
    return daily, hourly, downtime, production


def build_day_profiles(daily, downtime):
    daily_by_date = (
        daily.groupby("date", as_index=False)
        .agg(
            production_units=("production_units", "sum"),
            liters_produced=("liters_produced", "sum"),
            monitored_time_dec=("monitored_time_dec", "sum"),
            operation_time_dec=("operation_time_dec", "sum"),
            pause_time_dec=("pause_time_dec", "sum"),
        )
    )
    daily_by_date["efficiency"] = daily_by_date["operation_time_dec"] / daily_by_date["monitored_time_dec"].replace(0, np.nan)

    event_sums = downtime.groupby("date")["downtime_min"].sum().div(60.0).rename("event_downtime_h")
    profiles = daily_by_date.merge(event_sums, on="date", how="left")
    profiles["event_downtime_h"] = profiles["event_downtime_h"].fillna(0.0)
    profiles["event_scale"] = np.where(
        profiles["event_downtime_h"] > 0,
        np.minimum(1.0, profiles["pause_time_dec"] / profiles["event_downtime_h"]),
        1.0,
    )
    profiles["effective_event_h"] = profiles["event_downtime_h"] * profiles["event_scale"]
    profiles["residual_pause_h"] = (profiles["pause_time_dec"] - profiles["effective_event_h"]).clip(lower=0.0)
    profiles["prod_rate_lph"] = profiles["liters_produced"] / profiles["operation_time_dec"].replace(0, np.nan)
    profiles["prod_rate_lph"] = profiles["prod_rate_lph"].fillna(0.0)
    return profiles


def scenario_downtime_minutes(event_minutes, scenario):
    event_minutes = np.array(event_minutes, dtype=float)
    if scenario == "baseline":
        return event_minutes

    transformed = event_minutes.copy()
    micro = transformed <= 2.0
    minor = (transformed > 2.0) & (transformed <= 10.0)
    major = transformed > 10.0

    if scenario == "s1_5s_visual":
        keep = np.ones(len(transformed), dtype=bool)
        micro_idx = np.where(micro)[0]
        if len(micro_idx):
            remove_count = int(np.floor(0.15 * len(micro_idx)))
            keep[micro_idx[:remove_count]] = False
        transformed = transformed[keep]
    elif scenario == "s2_tpm":
        transformed[major] = transformed[major] * 0.75
    elif scenario == "s3_integrated":
        keep = np.ones(len(transformed), dtype=bool)
        micro_idx = np.where(micro)[0]
        if len(micro_idx):
            remove_count = int(np.floor(0.20 * len(micro_idx)))
            keep[micro_idx[:remove_count]] = False
        transformed = transformed[keep]
        transformed[(transformed > 2.0) & (transformed <= 10.0)] *= 0.85
        transformed[transformed > 10.0] *= 0.70
    else:
        raise ValueError(f"Unknown scenario: {scenario}")

    return transformed


def bootstrap_simulation(profiles, downtime, replications=5000, seed=42):
    rng = np.random.default_rng(seed)
    events_by_date = {
        date: group["downtime_min"].tolist()
        for date, group in downtime.groupby("date")
    }
    records = []
    scenarios = SCENARIO_ORDER

    profile_rows = profiles.to_dict("records")
    for rep in range(replications):
        sampled_idx = rng.integers(0, len(profile_rows), size=len(profile_rows))
        for scenario in scenarios:
            total_liters = 0.0
            total_op_h = 0.0
            total_pause_h = 0.0
            total_mon_h = 0.0

            for idx in sampled_idx:
                row = profile_rows[idx]
                event_minutes = np.array(events_by_date.get(row["date"], []), dtype=float) * row["event_scale"]
                adjusted_events_h = scenario_downtime_minutes(event_minutes, scenario).sum() / 60.0
                total_pause = row["residual_pause_h"] + adjusted_events_h
                total_pause = min(total_pause, row["monitored_time_dec"])
                op_time = max(row["monitored_time_dec"] - total_pause, 0.0)

                total_liters += op_time * row["prod_rate_lph"]
                total_op_h += op_time
                total_pause_h += total_pause
                total_mon_h += row["monitored_time_dec"]

            records.append(
                {
                    "replication": rep,
                    "scenario": scenario,
                    "liters": total_liters,
                    "operation_h": total_op_h,
                    "pause_h": total_pause_h,
                    "monitored_h": total_mon_h,
                    "efficiency": total_op_h / total_mon_h if total_mon_h else 0.0,
                }
            )

    return pd.DataFrame(records)


def summarize_results(sim_df):
    summary = (
        sim_df.groupby("scenario")
        .agg(
            liters_mean=("liters", "mean"),
            liters_p05=("liters", lambda s: np.percentile(s, 5)),
            liters_p95=("liters", lambda s: np.percentile(s, 95)),
            eff_mean=("efficiency", "mean"),
            eff_p05=("efficiency", lambda s: np.percentile(s, 5)),
            eff_p95=("efficiency", lambda s: np.percentile(s, 95)),
            pause_mean=("pause_h", "mean"),
            monitored_mean=("monitored_h", "mean"),
        )
        .reset_index()
    )
    baseline = summary.loc[summary["scenario"] == "baseline"].iloc[0]
    summary["liters_gain_pct"] = (summary["liters_mean"] - baseline["liters_mean"]) / baseline["liters_mean"] * 100.0
    summary["eff_gain_pp"] = (summary["eff_mean"] - baseline["eff_mean"]) * 100.0
    summary["pause_reduction_pct"] = (baseline["pause_mean"] - summary["pause_mean"]) / baseline["pause_mean"] * 100.0
    summary["waiting_waste_h"] = summary["pause_mean"]
    summary["waiting_waste_share_pct"] = summary["pause_mean"] / summary["monitored_mean"] * 100.0
    summary["waste_recovered_h"] = baseline["pause_mean"] - summary["pause_mean"]
    summary["waste_reduction_pct"] = summary["pause_reduction_pct"]
    return summary.set_index("scenario").loc[SCENARIO_ORDER].reset_index()


def create_plots(daily, downtime, summary):
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ModuleNotFoundError:
        plt = None

    if plt is not None:
        daily_plot = daily.copy()
        daily_plot["month"] = daily_plot["date"].dt.to_period("M").astype(str)
        monthly = (
            daily_plot.groupby("month")
            .agg(
                avg_efficiency=("efficiency", "mean"),
                liters=("liters_produced", "sum"),
            )
            .reset_index()
        )

        plt.figure(figsize=(9, 4.8))
        plt.plot(monthly["month"], monthly["avg_efficiency"] * 100, marker="o", linewidth=2, color="#1f77b4")
        plt.title("Monthly Average Operational Efficiency")
        plt.ylabel("Efficiency (%)")
        plt.xlabel("Month")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(OUTPUT_DIR / "monthly_efficiency.png", dpi=180)
        plt.close()

        plt.figure(figsize=(9, 4.8))
        plt.hist(downtime["downtime_min"], bins=50, color="#ff7f0e", edgecolor="black")
        plt.xlim(0, 60)
        plt.title("Downtime Event Duration Distribution")
        plt.xlabel("Downtime duration (minutes)")
        plt.ylabel("Frequency")
        plt.tight_layout()
        plt.savefig(OUTPUT_DIR / "downtime_distribution.png", dpi=180)
        plt.close()

        plot_df = summary.set_index("scenario").loc[SCENARIO_ORDER].reset_index()

        plt.figure(figsize=(8.8, 4.8))
        bars = plt.bar(SCENARIO_LABELS, plot_df["liters_mean"], color=["#7f7f7f", "#2ca02c", "#d62728", "#1f77b4"])
        plt.title("Simulated Production Output by Scenario")
        plt.ylabel("Total liters over 59 sampled days")
        plt.tight_layout()
        for bar, value in zip(bars, plot_df["liters_mean"]):
            plt.text(bar.get_x() + bar.get_width() / 2, value + 1000, f"{value:,.0f}", ha="center", va="bottom", fontsize=9)
        plt.savefig(OUTPUT_DIR / "scenario_output.png", dpi=180)
        plt.close()

        by_date = (
            daily.groupby("date", as_index=False)
            .agg(
                liters=("liters_produced", "sum"),
                monitored=("monitored_time_dec", "sum"),
                operation=("operation_time_dec", "sum"),
                pause=("pause_time_dec", "sum"),
            )
        )
        by_date["efficiency"] = by_date["operation"] / by_date["monitored"].replace(0, np.nan)

        plt.figure(figsize=(8.8, 4.8))
        plt.scatter(by_date["pause"], by_date["efficiency"] * 100, alpha=0.8, color="#9467bd")
        coeff = np.polyfit(by_date["pause"], by_date["efficiency"] * 100, 1)
        x_vals = np.linspace(by_date["pause"].min(), by_date["pause"].max(), 100)
        y_vals = coeff[0] * x_vals + coeff[1]
        plt.plot(x_vals, y_vals, color="black", linewidth=1.5)
        plt.title("Pause Time vs Operational Efficiency")
        plt.xlabel("Pause time per day (hours)")
        plt.ylabel("Efficiency (%)")
        plt.tight_layout()
        plt.savefig(OUTPUT_DIR / "pause_vs_efficiency.png", dpi=180)
        plt.close()

        downtime_plot = downtime.copy()
        downtime_plot["category"] = np.select(
            [
                downtime_plot["downtime_min"] <= 2.0,
                (downtime_plot["downtime_min"] > 2.0) & (downtime_plot["downtime_min"] <= 10.0),
                downtime_plot["downtime_min"] > 10.0,
            ],
            ["Micro (<=2 min)", "Minor (2-10 min)", "Major (>10 min)"],
            default="Other",
        )
        cat = (
            downtime_plot.groupby("category")["downtime_min"]
            .agg(event_count="count", lost_min="sum")
            .reindex(["Micro (<=2 min)", "Minor (2-10 min)", "Major (>10 min)"])
            .reset_index()
        )
        cat["lost_share_pct"] = cat["lost_min"] / cat["lost_min"].sum() * 100

        fig, ax1 = plt.subplots(figsize=(8.8, 4.8))
        bars = ax1.bar(cat["category"], cat["event_count"], color=["#2ca02c", "#ff7f0e", "#d62728"])
        ax1.set_ylabel("Event count")
        ax1.set_title("Downtime Pareto by Event Category")
        ax1.tick_params(axis="x", rotation=15)
        ax2 = ax1.twinx()
        ax2.plot(cat["category"], cat["lost_share_pct"], color="black", marker="o", linewidth=2)
        ax2.set_ylabel("Share of downtime minutes (%)")
        for bar, count in zip(bars, cat["event_count"]):
            ax1.text(bar.get_x() + bar.get_width() / 2, count + 10, f"{int(count)}", ha="center", va="bottom", fontsize=9)
        plt.tight_layout()
        plt.savefig(OUTPUT_DIR / "downtime_category_pareto.png", dpi=180)
        plt.close()

        rate_plot = (
            daily.assign(prod_rate_lph=daily["liters_produced"] / daily["operation_time_dec"].replace(0, np.nan))
            .groupby("product_type_l")["prod_rate_lph"]
            .agg(mean="mean", median="median")
            .reset_index()
        )
        labels = [f"{int(x)} L" for x in rate_plot["product_type_l"]]
        x = np.arange(len(labels))
        width = 0.35
        plt.figure(figsize=(7.5, 4.8))
        plt.bar(x - width / 2, rate_plot["mean"], width=width, label="Mean", color="#1f77b4")
        plt.bar(x + width / 2, rate_plot["median"], width=width, label="Median", color="#17becf")
        plt.xticks(x, labels)
        plt.ylabel("Liters per operating hour")
        plt.title("Production Rate by Product Size")
        plt.legend()
        plt.tight_layout()
        plt.savefig(OUTPUT_DIR / "product_rate_comparison.png", dpi=180)
        plt.close()

    create_waste_reduction_plot(summary)


def create_waste_reduction_plot(summary):
    plot_df = summary.set_index("scenario").loc[SCENARIO_ORDER].reset_index()

    width, height = 1220, 720
    margin_left, margin_right, margin_top, margin_bottom = 110, 120, 110, 140
    chart_left = margin_left
    chart_top = margin_top
    chart_right = width - margin_right
    chart_bottom = height - margin_bottom
    chart_width = chart_right - chart_left
    chart_height = chart_bottom - chart_top

    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)

    title_font = get_font(28, bold=True)
    subtitle_font = get_font(16)
    axis_font = get_font(16, bold=True)
    label_font = get_font(15)
    value_font = get_font(14, bold=True)

    draw.text((margin_left, 28), "Waste Reduction Achieved by Each Lean Scenario", font=title_font, fill="#0f172a")
    draw.text(
        (margin_left, 66),
        "Bars show recovered waiting-waste hours; the black line shows waste reduction relative to the simulated baseline.",
        font=subtitle_font,
        fill="#475569",
    )

    draw.line((chart_left, chart_top, chart_left, chart_bottom), fill="#64748b", width=2)
    draw.line((chart_left, chart_bottom, chart_right, chart_bottom), fill="#64748b", width=2)
    draw.line((chart_right, chart_top, chart_right, chart_bottom), fill="#64748b", width=2)

    max_hours = max(float(plot_df["waste_recovered_h"].max()), 1.0)
    max_pct = max(float(plot_df["waste_reduction_pct"].max()), 1.0)
    scale_max_hours = max_hours * 1.15
    scale_max_pct = max_pct * 1.15

    tick_count = 5
    for idx in range(tick_count + 1):
        value = scale_max_hours * idx / tick_count
        y = chart_bottom - (value / scale_max_hours) * chart_height
        draw.line((chart_left, y, chart_right, y), fill="#e2e8f0", width=1)
        draw.text((chart_left - 12, y), f"{value:.1f}", font=label_font, fill="#475569", anchor="rm")

    for idx in range(tick_count + 1):
        value = scale_max_pct * idx / tick_count
        y = chart_bottom - (value / scale_max_pct) * chart_height
        draw.text((chart_right + 12, y), f"{value:.0f}%", font=label_font, fill="#475569", anchor="lm")

    bar_colors = ["#cbd5e1", "#86efac", "#fb7185", "#38bdf8"]
    group_count = len(plot_df)
    group_width = chart_width / group_count
    bar_width = group_width * 0.42
    line_points = []

    for idx, (_, row) in enumerate(plot_df.iterrows()):
        center_x = chart_left + idx * group_width + group_width / 2
        bar_height = 0 if scale_max_hours == 0 else (float(row["waste_recovered_h"]) / scale_max_hours) * chart_height
        bar_top = chart_bottom - bar_height
        draw.rectangle(
            (center_x - bar_width / 2, bar_top, center_x + bar_width / 2, chart_bottom),
            fill=bar_colors[idx],
            outline="#475569",
            width=2,
        )
        draw.text((center_x, bar_top - 10), f"{row['waste_recovered_h']:.2f} h", font=value_font, fill="#0f172a", anchor="ms")
        draw.text((center_x, chart_bottom + 16), SCENARIO_LABELS[idx], font=label_font, fill="#0f172a", anchor="ma")

        pct_y = chart_bottom - (float(row["waste_reduction_pct"]) / scale_max_pct) * chart_height if scale_max_pct else chart_bottom
        line_points.append((center_x, pct_y))

    if len(line_points) > 1:
        draw.line(line_points, fill="black", width=4)
    for x, y in line_points:
        draw.ellipse((x - 6, y - 6, x + 6, y + 6), fill="black", outline="white", width=1)

    draw.text((chart_left - 70, chart_top - 30), "Recovered waste (h)", font=axis_font, fill="#0f172a")
    draw.text((chart_right + 18, chart_top - 30), "Reduction (%)", font=axis_font, fill="#0f172a")

    image.save(OUTPUT_DIR / "scenario_waste_reduction.png")


def build_report(daily, hourly, downtime, production, profiles, summary):
    baseline = summary.loc[summary["scenario"] == "baseline"].iloc[0]
    integrated = summary.loc[summary["scenario"] == "s3_integrated"].iloc[0]
    tpm = summary.loc[summary["scenario"] == "s2_tpm"].iloc[0]
    visual = summary.loc[summary["scenario"] == "s1_5s_visual"].iloc[0]

    total_downtime_min = downtime["downtime_min"].sum()
    micro = downtime.loc[downtime["downtime_min"] <= 2, "downtime_min"]
    minor = downtime.loc[(downtime["downtime_min"] > 2) & (downtime["downtime_min"] <= 10), "downtime_min"]
    major = downtime.loc[downtime["downtime_min"] > 10, "downtime_min"]

    report = {
        "dataset": {
            "days": int(len(daily)),
            "hourly_rows": int(len(hourly)),
            "downtime_events": int(len(downtime)),
            "total_liters": float(daily["liters_produced"].sum()),
            "total_units": int(daily["production_units"].sum()),
            "total_monitored_hours": float(daily["monitored_time_dec"].sum()),
            "total_operation_hours": float(daily["operation_time_dec"].sum()),
            "total_pause_hours": float(daily["pause_time_dec"].sum()),
            "overall_efficiency": float(daily["operation_time_dec"].sum() / daily["monitored_time_dec"].sum()),
            "mean_daily_efficiency": float(daily["efficiency"].mean()),
        },
        "downtime_profile": {
            "median_min": float(downtime["downtime_min"].median()),
            "p90_min": float(np.percentile(downtime["downtime_min"], 90)),
            "p95_min": float(np.percentile(downtime["downtime_min"], 95)),
            "max_min": float(downtime["downtime_min"].max()),
            "micro_count": int(len(micro)),
            "micro_share_time_pct": float(micro.sum() / total_downtime_min * 100.0),
            "minor_count": int(len(minor)),
            "minor_share_time_pct": float(minor.sum() / total_downtime_min * 100.0),
            "major_count": int(len(major)),
            "major_share_time_pct": float(major.sum() / total_downtime_min * 100.0),
        },
        "waste_baseline": {
            "waiting_waste_hours": float(daily["pause_time_dec"].sum()),
            "waiting_waste_share_pct": float(daily["pause_time_dec"].sum() / daily["monitored_time_dec"].sum() * 100.0),
            "value_adding_time_share_pct": float(daily["operation_time_dec"].sum() / daily["monitored_time_dec"].sum() * 100.0),
        },
        "scenarios": {
            "baseline_liters": float(baseline["liters_mean"]),
            "visual_liters_gain_pct": float(visual["liters_gain_pct"]),
            "tpm_liters_gain_pct": float(tpm["liters_gain_pct"]),
            "integrated_liters_gain_pct": float(integrated["liters_gain_pct"]),
            "baseline_efficiency_pct": float(baseline["eff_mean"] * 100.0),
            "integrated_efficiency_pct": float(integrated["eff_mean"] * 100.0),
            "integrated_eff_gain_pp": float(integrated["eff_gain_pp"]),
            "integrated_pause_reduction_pct": float(integrated["pause_reduction_pct"]),
            "integrated_waiting_waste_hours": float(integrated["waiting_waste_h"]),
            "integrated_waiting_waste_share_pct": float(integrated["waiting_waste_share_pct"]),
            "integrated_waste_recovered_h": float(integrated["waste_recovered_h"]),
        },
        "scenario_comparison": (
            summary.assign(
                efficiency_pct=summary["eff_mean"] * 100.0,
            )[
                [
                    "scenario",
                    "liters_mean",
                    "liters_gain_pct",
                    "efficiency_pct",
                    "pause_mean",
                    "waiting_waste_h",
                    "waiting_waste_share_pct",
                    "waste_recovered_h",
                    "waste_reduction_pct",
                ]
            ]
            .round(4)
            .to_dict(orient="records")
        ),
        "product_rates_lph": (
            daily.assign(prod_rate_lph=daily["liters_produced"] / daily["operation_time_dec"].replace(0, np.nan))
            .groupby("product_type_l")["prod_rate_lph"]
            .agg(["mean", "median"])
            .round(3)
            .to_dict(orient="index")
        ),
    }
    return report


def main():
    daily, hourly, downtime, production = load_data()
    profiles = build_day_profiles(daily, downtime)
    sim_df = bootstrap_simulation(profiles, downtime)
    summary = summarize_results(sim_df)
    create_plots(daily, downtime, summary)
    report = build_report(daily, hourly, downtime, production, profiles, summary)

    summary.to_csv(OUTPUT_DIR / "scenario_summary.csv", index=False)
    sim_df.to_csv(OUTPUT_DIR / "simulation_runs.csv", index=False)
    with open(OUTPUT_DIR / "analysis_summary.json", "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)

    print(json.dumps(report, indent=2))
    print("\nScenario summary:")
    print(summary.round(4).to_string(index=False))


if __name__ == "__main__":
    main()
