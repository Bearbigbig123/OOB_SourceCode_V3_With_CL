# SPC OOB Analysis Tool

以 PyQt6 開發的 SPC（Statistical Process Control）桌面分析工具，用於處理管制圖資料、判斷 OOC／Western Electric Rules／OOB、計算 Cpk，以及產生圖表與 Excel 報告。

## 主要功能

- SPC 管制圖與每週 OOB 分析
- Western Electric Rules（WE Rules）異常判定
- Cpk 分析儀表板
- Control Limit（CL）收緊與 Johnson transformation
- Excel／CSV 輸入資料健康檢查
- Tool matching、Sigma 與 Mean 比較
- 客戶與圖表篩選
- 繁體中文、English、한국어、日本語介面

## 系統需求

- Windows 10／11
- Python 3.10 以上（使用原始碼執行時）
- 建議安裝支援中、日、韓文字型的系統字型

## 快速開始

### 使用原始碼

```powershell
git clone https://github.com/Bearbigbig123/OOB_SourceCode_V3_With_CL.git
cd OOB_SourceCode_V3_With_CL
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python oob_module_NGK_nostatic.py
```

### 使用打包版本

Repository 內提供 V4.0 打包檔：

- `dist_opensource_v4_20260806/Opensource V4.0.zip`
- `dist_opensource_v4_onefile_20260904/Opensource V4.0_Onefile.zip`

解壓縮後，請將輸入資料放在執行檔旁的 `input` 目錄。程式會從執行檔所在位置讀取客戶維護的資料檔。

## 輸入資料結構

```text
input/
├── All_Chart_Information.xlsx
└── raw_charts/
    ├── <GroupName>_<ChartName>.csv
    └── ...
```

### `All_Chart_Information.xlsx`

常用必要欄位如下：

| 欄位 | 說明 |
| --- | --- |
| `GroupName` | 群組名稱 |
| `ChartName` | 管制圖名稱 |
| `UCL` / `LCL` | 管制上限／下限 |
| `USL` / `LSL` | 規格上限／下限 |
| `Target` | 中心值或目標值 |
| `Characteristics` | 特性類型，例如 `Nominal`、`Smaller`、`Larger` |

### Raw chart CSV

每個 CSV 至少需要以下欄位：

| 欄位 | 說明 |
| --- | --- |
| `point_time` | 量測時間，建議格式為 `YYYY/MM/DD HH:MM` |
| `point_val` | 量測值 |

檔名應與 Excel 中的 `GroupName` 和 `ChartName` 對應，例如：

```text
ETCH_Temperature.csv
```

## 專案結構

| 檔案 | 用途 |
| --- | --- |
| `oob_module_NGK_nostatic.py` | 主程式與 OOB 分析介面 |
| `CL_limit_class.py` | Control Limit 計算與收緊邏輯 |
| `spc_cpk_dashboard.py` | SPC／Cpk 儀表板 |
| `data_health_check.py` | 輸入資料檢查 |
| `tool_matching_widget.py` | Tool matching 分析 |
| `customer_filter.py` | 客戶篩選 |
| `chart_filter.py` | 圖表篩選 |
| `modern_ui.py` | 共用介面樣式 |
| `translations.py` | 多語系文字 |

## 注意事項

- `All_Chart_Information.xlsx` 必須位於 `input/`。
- Raw chart CSV 必須位於 `input/raw_charts/`。
- 若圖表無法載入，請先確認檔名、必要欄位及 `point_time` 格式。
- 執行分析後產生的圖表與報告，請依程式畫面選擇的輸出位置查看。
