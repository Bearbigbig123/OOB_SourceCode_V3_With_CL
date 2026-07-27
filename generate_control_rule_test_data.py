"""產生桌面版 Control 倍數與 Hard Rule 的手動驗收資料。"""

from pathlib import Path

import pandas as pd


OUTPUT_DIR = Path("control_rule_test_data")
RAW_DIR = OUTPUT_DIR / "raw_charts"


CASES = [
    {
        "chart": "HR1_8Sigma_Young",
        "values": [10] * 5,
        "create_age": "young",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": 12.0,
        "expected_lcl": 8.0,
        "expected_rule": "Hard Rule 1: Constant/Near Constant",
        "notes": "N=5、未滿一年：8σ，Hard Rule 向外 ±2×resolution",
    },
    {
        "chart": "HR1_5Sigma_Old",
        "values": [10] * 5,
        "create_age": "old",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": 11.0,
        "expected_lcl": 9.0,
        "expected_rule": "Hard Rule 1: Constant/Near Constant",
        "notes": "N=5、超過一年：5σ，Hard Rule 向外 ±1×resolution",
    },
    {
        "chart": "HR2_8Sigma_Young",
        "values": [10, 11] * 4,
        "create_age": "young",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": 13.0,
        "expected_lcl": 8.0,
        "expected_rule": "Hard Rule 2: Two Categories",
        "notes": "N=8、未滿一年：8σ，max/min 再向外 ±2×resolution",
    },
    {
        "chart": "HR2_5Sigma_N16",
        "values": [10, 11] * 8,
        "create_age": "young",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": 12.0,
        "expected_lcl": 9.0,
        "expected_rule": "Hard Rule 2: Two Categories",
        "notes": "N=16：固定 5σ，max/min 再向外 ±1×resolution",
    },
    {
        "chart": "HR3_8Sigma_Young",
        "values": [10, 11, 12] * 4,
        "create_age": "young",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": 14.0,
        "expected_lcl": 8.0,
        "expected_rule": "Hard Rule 3: Three Categories Spaced by Resolution",
        "notes": "N=12、未滿一年：8σ，三階資料再向外 ±2×resolution",
    },
    {
        "chart": "HR3_5Sigma_N16",
        "values": [10, 11, 12, 10, 11, 12, 10, 11] * 2,
        "create_age": "old",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": 13.0,
        "expected_lcl": 9.0,
        "expected_rule": "Hard Rule 3: Three Categories Spaced by Resolution",
        "notes": "N=16：固定 5σ，三階資料再向外 ±1×resolution",
    },
    {
        "chart": "HR2_Smaller_N16",
        "values": [10, 11] * 8,
        "create_age": "old",
        "characteristic": "Smaller",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": 12.0,
        "expected_lcl": -100.0,
        "expected_rule": "Hard Rule 2: Two Categories",
        "notes": "Smaller 只調 UCL；LCL 保留原始值",
    },
    {
        "chart": "HR2_Bigger_N16",
        "values": [10, 11] * 8,
        "create_age": "old",
        "characteristic": "Bigger",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": 100.0,
        "expected_lcl": 9.0,
        "expected_rule": "Hard Rule 2: Two Categories",
        "notes": "Bigger 只調 LCL；UCL 保留原始值",
    },
    {
        "chart": "HR2_Clamp_N16",
        "values": [10, 11] * 8,
        "create_age": "old",
        "characteristic": "Nominal",
        "ucl": 11.5,
        "lcl": 10.5,
        "expected_ucl": 11.5,
        "expected_lcl": 10.5,
        "expected_rule": "Hard Rule 2: Two Categories",
        "notes": "Hard Rule 算出 12/9 後，被原始 UCL/LCL Clamp",
    },
    {
        "chart": "Insufficient_N3",
        "values": [9.8, 10.0, 10.2],
        "create_age": "young",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": None,
        "expected_lcl": None,
        "expected_rule": "Insufficient Data",
        "notes": "N=3：點數不足，不計算 Suggest UCL/LCL",
    },
    {
        "chart": "NonHardRule_N4_Young",
        "values": [9.4, 10.0, 10.7, 11.5],
        "create_age": "young",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "resolution": 0.1,
        "expected_ucl": None,
        "expected_lcl": None,
        "expected_rule": "Non-Hard-Rule",
        "expected_k": 8.0,
        "notes": "一般資料 N=4、未滿一年：使用 8σ",
    },
    {
        "chart": "NonHardRule_N15_Young",
        "values": [round(9.3 + 0.1 * i, 1) for i in range(15)],
        "create_age": "young",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "resolution": 0.1,
        "expected_ucl": None,
        "expected_lcl": None,
        "expected_rule": "Non-Hard-Rule",
        "expected_k": 8.0,
        "notes": "一般資料 N=15、未滿一年：使用 8σ",
    },
    {
        "chart": "NonHardRule_N16",
        "values": [round(9.25 + 0.1 * i, 2) for i in range(16)],
        "create_age": "young",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "resolution": 0.1,
        "expected_ucl": None,
        "expected_lcl": None,
        "expected_rule": "Non-Hard-Rule",
        "expected_k": 5.0,
        "notes": "一般資料 N=16：固定使用 5σ",
    },
    {
        "chart": "NonHardRule_N29",
        "values": [round(8.6 + 0.1 * i, 1) for i in range(29)],
        "create_age": "old",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "resolution": 0.1,
        "expected_ucl": None,
        "expected_lcl": None,
        "expected_rule": "Non-Hard-Rule",
        "expected_k": 5.0,
        "notes": "一般資料 N=29：固定使用 5σ",
    },
    {
        "chart": "NonHardRule_N30",
        "values": [round(8.55 + 0.1 * i, 2) for i in range(30)],
        "create_age": "young",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "resolution": 0.1,
        "expected_ucl": None,
        "expected_lcl": None,
        "expected_rule": "Non-Hard-Rule",
        "expected_k": 3.0,
        "notes": "一般資料 N=30：回到既有 3σ（Normal 高峰度時才 4σ）",
    },
    {
        "chart": "NonHardRule_N60",
        "values": [
            round(9.05 + 0.1 * (i % 20), 2)
            for i in range(60)
        ],
        "create_age": "old",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "resolution": 0.1,
        "expected_ucl": None,
        "expected_lcl": None,
        "expected_rule": "Non-Hard-Rule",
        "expected_k": 3.0,
        "notes": "一般資料 N=60：維持既有 3σ",
    },
    {
        "chart": "HR2_N30_NoExpansion",
        "values": [10, 11] * 15,
        "create_age": "young",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": 11.0,
        "expected_lcl": 10.0,
        "expected_rule": "Hard Rule 2: Two Categories",
        "notes": "Hard Rule N=30：不套用 resolution 擴張",
    },
    {
        "chart": "HR2_N60_NoExpansion",
        "values": [10, 11] * 30,
        "create_age": "old",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "expected_ucl": 11.0,
        "expected_lcl": 10.0,
        "expected_rule": "Hard Rule 2: Two Categories",
        "notes": "Hard Rule N=60：不套用 resolution 擴張",
    },
    {
        "chart": "NullValues_20Rows_16Valid",
        "values": [
            None,
            9.2, 9.3, 9.4, 9.5, 9.6,
            "",
            9.7, 9.8, 9.9, 10.0, 10.1,
            "not-a-number",
            10.2, 10.3, 10.4, 10.5, 10.6, 10.7,
            None,
        ],
        "create_age": "old",
        "characteristic": "Nominal",
        "ucl": 100.0,
        "lcl": -100.0,
        "resolution": 0.1,
        "expected_ucl": None,
        "expected_lcl": None,
        "expected_rule": "Non-Hard-Rule",
        "expected_k": 5.0,
        "notes": "Raw 共20列，排除4列空白/非數字後，以16個有效點計算並繪圖",
    },
]


def main():
    RAW_DIR.mkdir(parents=True, exist_ok=True)

    today = pd.Timestamp.today().normalize()
    create_times = {
        "young": today - pd.DateOffset(months=6),
        "old": today - pd.DateOffset(years=2),
    }
    chart_rows = []

    for case_index, case in enumerate(CASES, start=1):
        create_time = create_times[case["create_age"]]
        chart_rows.append(
            {
                "GroupName": "ControlRule",
                "ChartName": case["chart"],
                "ChartID": f"CTRL_RULE_{case_index:02d}",
                "Material_no": "TEST",
                "Target": 11.0,
                "UCL": case["ucl"],
                "LCL": case["lcl"],
                "USL": 1000.0,
                "LSL": -1000.0,
                "Characteristics": case["characteristic"],
                "DetectionLimit": pd.NA,
                "ExpectedPattern": case["expected_rule"],
                "SampleCount": len(case["values"]),
                "Resolution": case.get("resolution", 1.0),
                "CHART_CREATE_TIME": create_time,
                "Expected_Suggest_UCL": case["expected_ucl"],
                "Expected_Suggest_LCL": case["expected_lcl"],
                "Expected_K": case.get("expected_k", pd.NA),
                "Test_Notes": case["notes"],
            }
        )

        point_start = today - pd.Timedelta(days=len(case["values"]))
        raw_rows = []
        for point_index, value in enumerate(case["values"], start=1):
            raw_rows.append(
                {
                    "GroupName": "ControlRule",
                    "ChartName": case["chart"],
                    "point_time": (
                        point_start + pd.Timedelta(days=point_index)
                    ).strftime("%Y/%m/%d %H:%M"),
                    "point_val": value,
                    "Batch_ID": f"BATCH_{point_index:03d}",
                    "Matching": "TestTool",
                    "Customer": "TSMC",
                }
            )

        raw_path = RAW_DIR / f"ControlRule_{case['chart']}.csv"
        pd.DataFrame(raw_rows).to_csv(raw_path, index=False, encoding="utf-8-sig")

    chart_path = OUTPUT_DIR / "All_Chart_Information_Control_Rules.xlsx"
    with pd.ExcelWriter(chart_path, engine="openpyxl") as writer:
        pd.DataFrame(chart_rows).to_excel(writer, sheet_name="Chart", index=False)

    print(f"Created {len(CASES)} charts in: {OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
