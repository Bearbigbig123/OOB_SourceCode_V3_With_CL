"""Generate 300 deterministic CL/Tighten test charts for CL_limit_class.py."""

from __future__ import annotations

import contextlib
import io
from pathlib import Path

import numpy as np
import pandas as pd

from CL_limit_class import CLTightenCalculator


OUTPUT_DIR = Path("tighten_300_test_data")
RAW_DIR = OUTPUT_DIR / "raw_charts"
GROUP_NAME = "Tighten300"
SEED = 20260806
TOTAL_CHARTS = 300

FAMILY_COUNTS = [
    ("HardRule_Constant", 24),
    ("HardRule_TwoCategory", 24),
    ("HardRule_ThreeCategory", 24),
    ("Normal", 36),
    ("Skew_Right", 30),
    ("Skew_Left", 30),
    ("Bimodal", 30),
    ("Attribute", 30),
    ("Near_Constant", 24),
    ("Outlier_Contaminated", 24),
    ("Uniform", 12),
    ("Heavy_Tail", 12),
]


def make_values(family, n, center, scale, resolution, rng):
    if family == "HardRule_Constant":
        values = np.full(n, center)
    elif family == "HardRule_TwoCategory":
        values = np.resize([center, center + resolution], n)
    elif family == "HardRule_ThreeCategory":
        values = np.resize(
            [center - resolution, center, center + resolution], n
        )
    elif family == "Normal":
        values = rng.normal(center, scale, n)
    elif family == "Skew_Right":
        values = center - 2 * scale + rng.gamma(2.0, scale, n)
    elif family == "Skew_Left":
        values = center + 2 * scale - rng.gamma(2.0, scale, n)
    elif family == "Bimodal":
        left_n = n // 2
        values = np.r_[
            rng.normal(center - 2.7 * scale, 0.35 * scale, left_n),
            rng.normal(center + 2.7 * scale, 0.35 * scale, n - left_n),
        ]
        rng.shuffle(values)
    elif family == "Attribute":
        categories = center + resolution * np.arange(-3, 3)
        values = rng.choice(
            categories, n, p=[0.08, 0.15, 0.27, 0.25, 0.17, 0.08]
        )
    elif family == "Near_Constant":
        values = np.full(n, center)
        values[-2:] = [center - 2 * resolution, center + 2 * resolution]
        rng.shuffle(values)
    elif family == "Outlier_Contaminated":
        values = rng.normal(center, 0.65 * scale, n)
        values[-2:] = [center - 12 * scale, center + 14 * scale]
        rng.shuffle(values)
    elif family == "Uniform":
        values = rng.uniform(center - 2.5 * scale, center + 2.5 * scale, n)
    elif family == "Heavy_Tail":
        values = center + scale * rng.standard_t(df=2.5, size=n)
    else:
        raise ValueError(f"Unknown family: {family}")

    decimals = max(0, -int(np.floor(np.log10(resolution))))
    return np.round(np.asarray(values, dtype=float), decimals)


def configured_limits(values, center, resolution, characteristic, limit_mode):
    observed_sigma = float(np.std(values, ddof=1)) if len(values) > 1 else 0.0
    effective_sigma = max(observed_sigma, resolution)
    half_range = max(center - float(np.min(values)), float(np.max(values)) - center)

    if limit_mode == "Wide_Expect_Tighten":
        active_distance = max(14 * effective_sigma, half_range + 6 * resolution)
    else:
        active_distance = max(2.2 * effective_sigma, half_range + resolution)

    inactive_distance = max(16 * effective_sigma, half_range + 8 * resolution)
    if characteristic == "Smaller":
        ucl = center + active_distance
        lcl = center - inactive_distance
    elif characteristic == "Bigger":
        ucl = center + inactive_distance
        lcl = center - active_distance
    else:
        ucl = center + active_distance
        lcl = center - active_distance

    decimals = max(0, -int(np.floor(np.log10(resolution))))
    return round(ucl, decimals), round(lcl, decimals)


def build_cases():
    rng = np.random.default_rng(SEED)
    today = pd.Timestamp.today().normalize()
    families = [family for family, count in FAMILY_COUNTS for _ in range(count)]
    assert len(families) == TOTAL_CHARTS

    chart_rows = []
    raw_frames = {}
    n_options = [30, 45, 60, 90, 120, 180, 240, 300]
    resolutions = [1.0, 0.5, 0.1, 0.05, 0.01]
    characteristics = ["Nominal", "Smaller", "Bigger"]

    for index, family in enumerate(families, start=1):
        n = n_options[(index - 1) % len(n_options)]
        if family in {"Attribute", "Near_Constant"}:
            n = max(n, 30)
        resolution = resolutions[(index - 1) % len(resolutions)]
        characteristic = characteristics[(index - 1) % len(characteristics)]
        limit_mode = (
            "Wide_Expect_Tighten" if index % 2 else "Tight_Expect_No"
        )
        center = round(50 + (index % 7) * 5, 2)
        scale = 0.8 + (index % 5) * 0.35
        values = make_values(
            family, n, center, scale, resolution, rng
        )
        ucl, lcl = configured_limits(
            values, center, resolution, characteristic, limit_mode
        )

        chart_name = f"T{index:03d}_{family}"
        create_time = (
            today - pd.DateOffset(months=6)
            if index % 3 == 0
            else today - pd.DateOffset(years=2)
        )
        margin = max(abs(ucl - center), abs(center - lcl), scale) + 10 * scale
        usl = center + margin
        lsl = center - margin
        detection_limit = center - 3.5 * scale if characteristic == "Smaller" else np.nan

        chart_rows.append({
            "GroupName": GROUP_NAME,
            "ChartName": chart_name,
            "ChartID": f"TIGHTEN_{index:03d}",
            "Material_no": f"TEST_MAT_{(index - 1) // 10 + 1:02d}",
            "Target": center,
            "UCL": ucl,
            "LCL": lcl,
            "USL": usl,
            "LSL": lsl,
            "Characteristics": characteristic,
            "DetectionLimit": detection_limit,
            "Resolution": resolution,
            "CHART_CREATE_TIME": create_time,
            "DesignedFamily": family,
            "DesignedLimitMode": limit_mode,
            "DesignedSampleCount": n,
        })

        start = today - pd.Timedelta(days=n)
        raw_frames[chart_name] = pd.DataFrame({
            "GroupName": GROUP_NAME,
            "ChartName": chart_name,
            "point_time": pd.date_range(start, periods=n, freq="D"),
            "point_val": values,
            "Batch_ID": [f"BATCH_{index:03d}_{i:04d}" for i in range(1, n + 1)],
            "Matching": [f"TOOL_{(i % 3) + 1}" for i in range(n)],
            "Customer": [["TSMC", "UMC", "ASE"][i % 3] for i in range(n)],
        })

    return pd.DataFrame(chart_rows), raw_frames


def verify(chart_df, raw_frames):
    results = []
    for _, chart in chart_df.iterrows():
        raw = raw_frames[chart["ChartName"]].copy()
        raw = raw.rename(columns={"point_time": "date", "point_val": "value"})
        raw["DetectionLimit"] = chart["DetectionLimit"]
        raw["Target"] = chart["Target"]
        raw["UCL"] = chart["UCL"]
        raw["LCL"] = chart["LCL"]
        raw["Resolution"] = chart["Resolution"]
        raw["oos_flag"] = False
        calculator = CLTightenCalculator()
        with contextlib.redirect_stdout(io.StringIO()):
            result = calculator.process_chart(
                raw,
                value_col="value",
                date_col="date",
                oos_col="oos_flag",
                characteristic=chart["Characteristics"],
                chart_create_time=chart["CHART_CREATE_TIME"],
            )
        results.append({
            "ChartName": chart["ChartName"],
            "DesignedFamily": chart["DesignedFamily"],
            "DesignedLimitMode": chart["DesignedLimitMode"],
            "Characteristics": chart["Characteristics"],
            "SampleCount": result["TotalDataCount"],
            "DataCountUsed": result["DataCountUsed"],
            "ActualPattern": result["Pattern"],
            "HardRule": result["HardRule"],
            "Original_UCL": chart["UCL"],
            "Original_LCL": chart["LCL"],
            "Static_UCL": result["Static UCL"],
            "Static_LCL": result["Static LCL"],
            "Suggest_UCL": result["Suggest UCL"],
            "Suggest_LCL": result["Suggest LCL"],
            "Ori_OOC": result["Ori_OOC_Count"],
            "Static_OOC": result["Static_OOC_Count"],
            "Final_OOC": result["Final_OOC_Count"],
            "Original_Tolerance": result["Original_Tolerance"],
            "New_Tolerance": result["New_Tolerance"],
            "Diff_Ratio_%": result["Diff_Ratio_%"],
            "Tighten_Threshold_%": result["Tighten_Threshold_%"],
            "TightenNeeded": result["TightenNeeded"],
        })
    return pd.DataFrame(results)


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    chart_df, raw_frames = build_cases()

    for chart_name, raw in raw_frames.items():
        raw.to_csv(
            RAW_DIR / f"{GROUP_NAME}_{chart_name}.csv",
            index=False,
            encoding="utf-8-sig",
        )

    result_df = verify(chart_df, raw_frames)
    workbook_path = OUTPUT_DIR / "All_Chart_Information.xlsx"
    with pd.ExcelWriter(workbook_path, engine="openpyxl") as writer:
        chart_df.to_excel(writer, sheet_name="Chart", index=False)
        result_df.to_excel(writer, sheet_name="Verification", index=False)

    result_df.to_csv(
        OUTPUT_DIR / "Tighten_300_Verification.csv",
        index=False,
        encoding="utf-8-sig",
    )
    summary = (
        result_df.groupby(["DesignedFamily", "TightenNeeded"], dropna=False)
        .size()
        .rename("ChartCount")
        .reset_index()
    )
    summary.to_csv(
        OUTPUT_DIR / "Tighten_300_Summary.csv",
        index=False,
        encoding="utf-8-sig",
    )

    print(f"Created {len(chart_df)} charts at {OUTPUT_DIR.resolve()}")
    print(result_df["TightenNeeded"].value_counts(dropna=False).to_string())
    print("Actual patterns:")
    print(result_df["ActualPattern"].value_counts(dropna=False).to_string())


if __name__ == "__main__":
    main()
