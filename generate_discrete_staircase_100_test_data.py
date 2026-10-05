"""Generate 100 deterministic discrete staircase charts for CL testing."""

from __future__ import annotations

import contextlib
import io
from pathlib import Path

import numpy as np
import pandas as pd

from CL_limit_class import CLTightenCalculator


OUTPUT_DIR = Path("discrete_staircase_100_test_data")
RAW_DIR = OUTPUT_DIR / "raw_charts"
GROUP_NAME = "DiscreteStair100"
TOTAL_CHARTS = 100
SEED = 20260806

PATTERNS = [
    "Step_Up",
    "Step_Down",
    "Up_Then_Down",
    "Down_Then_Up",
    "Cyclic_Up",
    "Cyclic_Down",
    "Late_Level_Shift_Up",
    "Late_Level_Shift_Down",
    "Uneven_Plateaus_Up",
    "Alternating_Plateaus",
]


def quantize(values, resolution):
    return np.round(np.asarray(values, dtype=float) / resolution) * resolution


def expand_levels(levels, n, uneven=False):
    levels = np.asarray(levels, dtype=float)
    if uneven:
        weights = np.arange(1, len(levels) + 1, dtype=float)
        weights = weights / weights.sum()
        counts = np.maximum(2, np.floor(weights * n).astype(int))
        counts[-1] += n - counts.sum()
        while counts[-1] < 2:
            donor = int(np.argmax(counts[:-1]))
            counts[donor] -= 1
            counts[-1] += 1
        values = np.repeat(levels, counts)
        return np.resize(values, n)
    return np.resize(np.repeat(levels, int(np.ceil(n / len(levels)))), n)


def make_staircase(pattern, n, center, resolution, level_count):
    # Integer-resolution offsets keep every designed stair level distinct.
    offsets = np.arange(level_count) - ((level_count - 1) // 2)
    levels = quantize(center + offsets * resolution, resolution)

    if pattern == "Step_Up":
        values = expand_levels(levels, n)
    elif pattern == "Step_Down":
        values = expand_levels(levels[::-1], n)
    elif pattern == "Up_Then_Down":
        route = np.r_[levels, levels[-2:0:-1]]
        values = expand_levels(route, n)
    elif pattern == "Down_Then_Up":
        route = np.r_[levels[::-1], levels[1:-1]]
        values = expand_levels(route, n)
    elif pattern == "Cyclic_Up":
        block = max(3, n // (level_count * 4))
        values = np.resize(np.repeat(levels, block), n)
    elif pattern == "Cyclic_Down":
        block = max(3, n // (level_count * 4))
        values = np.resize(np.repeat(levels[::-1], block), n)
    elif pattern == "Late_Level_Shift_Up":
        cut = int(n * 0.72)
        values = np.r_[np.full(cut, levels[1]), np.full(n - cut, levels[-2])]
    elif pattern == "Late_Level_Shift_Down":
        cut = int(n * 0.72)
        values = np.r_[np.full(cut, levels[-2]), np.full(n - cut, levels[1])]
    elif pattern == "Uneven_Plateaus_Up":
        values = expand_levels(levels, n, uneven=True)
    elif pattern == "Alternating_Plateaus":
        route = np.ravel(np.column_stack((levels, levels[::-1])))
        values = expand_levels(route, n)
    else:
        raise ValueError(f"Unknown staircase pattern: {pattern}")

    return quantize(values, resolution)


def configured_limits(values, center, resolution, characteristic, wide):
    distance = max(
        center - float(np.min(values)),
        float(np.max(values)) - center,
        resolution,
    )
    active = distance + (8 if wide else 1) * resolution
    inactive = distance + 10 * resolution
    if characteristic == "Smaller":
        ucl, lcl = center + active, center - inactive
    elif characteristic == "Bigger":
        ucl, lcl = center + inactive, center - active
    else:
        ucl, lcl = center + active, center - active
    decimals = max(0, -int(np.floor(np.log10(resolution))))
    return round(ucl, decimals), round(lcl, decimals)


def build_cases():
    today = pd.Timestamp.today().normalize()
    point_options = [60, 90, 120, 150, 180, 240, 300]
    resolutions = [1.0, 0.5, 0.1, 0.05, 0.01]
    characteristics = ["Nominal", "Smaller", "Bigger"]
    chart_rows = []
    raw_frames = {}

    for index in range(1, TOTAL_CHARTS + 1):
        pattern = PATTERNS[(index - 1) % len(PATTERNS)]
        n = point_options[(index - 1) % len(point_options)]
        resolution = resolutions[(index - 1) % len(resolutions)]
        characteristic = characteristics[(index - 1) % len(characteristics)]
        level_count = 4 + ((index - 1) % 7)
        center = quantize([40 + (index % 9) * 5], resolution)[0]
        values = make_staircase(
            pattern, n, center, resolution, level_count
        )
        wide = index % 2 == 1
        ucl, lcl = configured_limits(
            values, center, resolution, characteristic, wide
        )
        chart_name = f"DS{index:03d}_{pattern}"
        margin = max(abs(ucl - center), abs(center - lcl)) + 12 * resolution

        chart_rows.append({
            "GroupName": GROUP_NAME,
            "ChartName": chart_name,
            "ChartID": f"DISCRETE_STAIR_{index:03d}",
            "Material_no": f"STAIR_MAT_{(index - 1) // 10 + 1:02d}",
            "Target": center,
            "UCL": ucl,
            "LCL": lcl,
            "USL": center + margin,
            "LSL": center - margin,
            "Characteristics": characteristic,
            "DetectionLimit": np.nan,
            "Resolution": resolution,
            "CHART_CREATE_TIME": today - pd.DateOffset(years=2),
            "DesignedFamily": "Discrete_Staircase",
            "DesignedPattern": pattern,
            "DesignedLimitMode": "Wide_Expect_Tighten" if wide else "Tight_Expect_No",
            "DesignedLevelCount": level_count,
            "DesignedSampleCount": n,
        })

        start = today - pd.Timedelta(days=n)
        raw_frames[chart_name] = pd.DataFrame({
            "GroupName": GROUP_NAME,
            "ChartName": chart_name,
            "point_time": pd.date_range(start, periods=n, freq="D"),
            "point_val": values,
            "Batch_ID": [f"STAIR_BATCH_{index:03d}_{i:04d}" for i in range(1, n + 1)],
            "Matching": [f"TOOL_{(i % 4) + 1}" for i in range(n)],
            "Customer": [["TSMC", "UMC", "ASE", "PSMC"][i % 4] for i in range(n)],
        })

    return pd.DataFrame(chart_rows), raw_frames


def verify(chart_df, raw_frames):
    rows = []
    for _, chart in chart_df.iterrows():
        raw = raw_frames[chart["ChartName"]].rename(
            columns={"point_time": "date", "point_val": "value"}
        )
        raw["DetectionLimit"] = chart["DetectionLimit"]
        raw["Target"] = chart["Target"]
        raw["UCL"] = chart["UCL"]
        raw["LCL"] = chart["LCL"]
        raw["Resolution"] = chart["Resolution"]
        raw["oos_flag"] = False
        with contextlib.redirect_stdout(io.StringIO()):
            result = CLTightenCalculator().process_chart(
                raw,
                value_col="value",
                date_col="date",
                oos_col="oos_flag",
                characteristic=chart["Characteristics"],
                chart_create_time=chart["CHART_CREATE_TIME"],
            )
        rows.append({
            "ChartName": chart["ChartName"],
            "DesignedPattern": chart["DesignedPattern"],
            "DesignedLimitMode": chart["DesignedLimitMode"],
            "Characteristics": chart["Characteristics"],
            "Resolution": chart["Resolution"],
            "LevelCount": chart["DesignedLevelCount"],
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
            "TightenNeeded": result["TightenNeeded"],
        })
    return pd.DataFrame(rows)


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
    with pd.ExcelWriter(
        OUTPUT_DIR / "All_Chart_Information.xlsx", engine="openpyxl"
    ) as writer:
        chart_df.to_excel(writer, sheet_name="Chart", index=False)
        result_df.to_excel(writer, sheet_name="Verification", index=False)

    result_df.to_csv(
        OUTPUT_DIR / "Discrete_Staircase_100_Verification.csv",
        index=False,
        encoding="utf-8-sig",
    )
    summary = (
        result_df.groupby(
            ["DesignedPattern", "TightenNeeded"], dropna=False
        )
        .size()
        .rename("ChartCount")
        .reset_index()
    )
    summary.to_csv(
        OUTPUT_DIR / "Discrete_Staircase_100_Summary.csv",
        index=False,
        encoding="utf-8-sig",
    )

    print(f"Created {len(chart_df)} charts at {OUTPUT_DIR.resolve()}")
    print(result_df["TightenNeeded"].value_counts(dropna=False).to_string())
    print("Actual patterns:")
    print(result_df["ActualPattern"].value_counts(dropna=False).to_string())


if __name__ == "__main__":
    main()
