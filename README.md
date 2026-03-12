# SPC OOB Analysis Tool

A PyQt6-based desktop application for **Statistical Process Control (SPC)** Out-of-Bounds (OOB) analysis. It processes SPC chart data, detects rule violations (OOC / WE-Rules / OOB), calculates Cpk, and generates reports.

---

## Features

- SPC chart processing with OOC / WE-Rule / OOB detection
- Cpk calculation dashboard
- Control Limit (CL) tightening calculator using Johnson transformation
- Data health check for Excel & CSV input files
- Tool matching / sigma & mean comparison widget
- Bilingual UI (繁體中文 / English)

---

## Requirements

- Python 3.10 or higher
- Windows (uses Microsoft JhengHei font for CJK rendering; other platforms may require a CJK font installed)

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Input File Structure

Place your data files under an `input/` folder in the same directory as the scripts:

```
input/
├── All_Chart_Information.xlsx   # Chart configuration (GroupName, ChartName, UCL, LCL, USL, LSL, etc.)
└── <GroupName>_<ChartName>_*.csv  # Raw measurement data (must contain: point_time, point_val columns)
```

### Excel (`All_Chart_Information.xlsx`) — required columns

| Column | Description |
|---|---|
| GroupName | Group identifier |
| ChartName | Chart identifier |
| UCL / LCL | Upper / Lower Control Limit |
| USL / LSL | Upper / Lower Spec Limit |
| Target | Center / Target value |
| Characteristics | `Nominal` / `Smaller` / `Larger` |

### CSV files — required columns

| Column | Description |
|---|---|
| point_time | Timestamp, format: `YYYY/MM/DD HH:MM` |
| point_val  | Measured value |

---

## Running

```bash
python oob_module_NGK_nostatic.py
```

---

## License

MIT License — see [LICENSE](LICENSE) for details.
"# OOB_SourceCode_V3_With_CL" 
