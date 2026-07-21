import os
import logging
import pandas as pd
from customer_filter import apply_customer_filter
from PyQt6 import QtWidgets, QtCore, QtGui
import pickle # Import pickle module for deep copying
# Translation System
from translations import get_translator, tr

logger = logging.getLogger(__name__)

# Check if openpyxl package is installed
try:
    import openpyxl
except ImportError:
    openpyxl = None

class FormulaExplanationDialog(QtWidgets.QDialog):
    """公式說明對話框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("formula_explanation"))
        self.setMinimumSize(900, 600)
        self.init_ui()
    
    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        
        # 標題
        title_label = QtWidgets.QLabel("<h2 style='text-align:center; color:#344CB7;'>📖 Formula Explanation / 公式說明</h2>")
        title_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # 滾動區域
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # 內容 widget - 使用水平佈局並排顯示中英文
        content_widget = QtWidgets.QWidget()
        content_layout = QtWidgets.QHBoxLayout(content_widget)
        content_layout.setContentsMargins(10, 10, 10, 10)
        content_layout.setSpacing(15)
        
        # ===== 左側：英文版 =====
        left_widget = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_widget)
        left_layout.setSpacing(15)
        
        # 英文 - 注意事項
        notice_label_en = QtWidgets.QLabel()
        notice_label_en.setWordWrap(True)
        notice_label_en.setTextFormat(QtCore.Qt.TextFormat.RichText)
        notice_label_en.setFont(QtGui.QFont("Arial", 10))
        notice_html_en = """
<div style='background-color:#f5f5f5; padding:12px; border-radius:6px; font-size:13px; margin-bottom:15px; font-family:Arial;'>
  <strong style='font-size:15px;'>⚠ Notice</strong>
  <p style='margin:8px 0;'>Table only shows items requiring attention (abnormal detection criteria):</p>
  <ul style='margin:8px 0 8px 20px; padding-left:0; line-height:1.8;'>
    <li><span style='color:#d9534f;'><strong>mean_matching_index ≥ 1</strong></span>: Mean not matched</li>
    <li><span style='color:#d9534f;'><strong>sigma_matching_index ≥ K</strong></span>: Sigma not matched</li>
    <li><span style='color:#8a6d3b;'><strong>Insufficient Data</strong></span>: Sample size less than 5</li>
  </ul>
</div>
        """
        notice_label_en.setText(notice_html_en)
        left_layout.addWidget(notice_label_en)
        
        # 英文 - 計算公式
        formula_label_en = QtWidgets.QLabel()
        formula_label_en.setWordWrap(True)
        formula_label_en.setTextFormat(QtCore.Qt.TextFormat.RichText)
        formula_label_en.setFont(QtGui.QFont("Arial", 10))
        formula_html_en = """
<div style="background-color:#e8f4f8; padding:15px; border-radius:6px; font-size:13px; line-height:1.6; font-family:Arial; border: 2px solid #344CB7;">
  <strong style='font-size:15px;'>📘 Calculation Formula</strong>
  <table style="font-size:12px; margin-top:12px; font-family:Arial; width:100%;">
    <tr>
      <td style="vertical-align:top; padding:8px; padding-right:12px; width:35%;"><strong>Mean Matching Index:</strong></td>
      <td style="padding:8px;">
        <u>Two-group comparison:</u><br>
        <code style='background:#fff; padding:2px 6px; border-radius:3px;'>|μ₁ − μ₂| / min(σ₁, σ₂)</code><br><br>
        <u>Multi-group comparison:</u><br>
        <code style='background:#fff; padding:2px 6px; border-radius:3px;'>|μ − median(μ)| / median(σ)</code>
      </td>
    </tr>
    <tr>
      <td style="vertical-align:top; padding:8px; padding-right:12px;"><strong>Sigma Matching Index:</strong></td>
      <td style="padding:8px;">
        <u>Two-group comparison:</u><br>
        <code style='background:#fff; padding:2px 6px; border-radius:3px;'>σ / min(σ₁, σ₂)</code><br><br>
        <u>Multi-group comparison:</u><br>
        <code style='background:#fff; padding:2px 6px; border-radius:3px;'>σ / median(σ)</code>
      </td>
    </tr>
    <tr>
      <td style="vertical-align:top; padding:8px; padding-right:12px;"><strong>K Value:</strong></td>
      <td style="padding:8px;">
        <code style='background:#fff; padding:6px 10px; border-radius:3px; display:block; line-height:1.8;'>
          n = Sample size<br>
          n ≤ 4: No comparison<br>
          5 ≤ n ≤ 10: K = 1.73<br>
          11 ≤ n ≤ 120: K = 1.414<br>
          n > 120: K = 1.15
        </code>
      </td>
    </tr>
  </table>
  <div style="margin-top:15px; padding:12px; background:#fff; border-radius:4px; font-size:12px; color:#344CB7;">
    <strong>【Calculation Description for Filter Mode】</strong><br>
    <ul style="margin:8px 0 8px 20px; padding-left:0; line-height:1.8;">
      <li><strong>Mean/Std/Sample Size:</strong> For each matching_group, take data within "one month before the specified date", fill to specified number if insufficient (default 5, adjustable via UI), then calculate mean/std/count.</li>
      <li><strong>Median(sigma):</strong> For each matching_group, take data within "six months before the specified date", fill to specified number if insufficient (default 5), calculate std for each group, then take median.</li>
      <li><strong>Chart Display:</strong> Based on data "within one month (filled to specified number)".</li>
      <li><strong>Multiple Data Points:</strong> If multiple data points exist for the same group at the same time, all are included in calculation.</li>
    </ul>
    <span style="color:#8a6d3b; font-style:italic;">• All Data mode directly uses all data for group calculation without filling</span>
  </div>
</div>
        """
        formula_label_en.setText(formula_html_en)
        left_layout.addWidget(formula_label_en)
        left_layout.addStretch()
        
        # ===== 右側：中文版 =====
        right_widget = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right_widget)
        right_layout.setSpacing(15)
        
        # 中文 - 注意事項
        notice_label_zh = QtWidgets.QLabel()
        notice_label_zh.setWordWrap(True)
        notice_label_zh.setTextFormat(QtCore.Qt.TextFormat.RichText)
        notice_label_zh.setFont(QtGui.QFont("Microsoft JhengHei, sans-serif", 10))
        notice_html_zh = """
<div style='background-color:#f5f5f5; padding:12px; border-radius:6px; font-size:13px; margin-bottom:15px; font-family:"Microsoft JhengHei", "PingFang TC", "Noto Sans CJK TC", sans-serif;'>
  <strong style='font-size:15px;'>⚠ 注意</strong>
  <p style='margin:8px 0;'>表格僅顯示需要關注的項目（異常檢測標準）：</p>
  <ul style='margin:8px 0 8px 20px; padding-left:0; line-height:1.8;'>
    <li><span style='color:#d9534f;'><strong>mean_matching_index ≥ 1</strong></span>：平均值不匹配</li>
    <li><span style='color:#d9534f;'><strong>sigma_matching_index ≥ K</strong></span>：標準差不匹配</li>
    <li><span style='color:#8a6d3b;'><strong>資料不足</strong></span>：樣本數少於5個</li>
  </ul>
</div>
        """
        notice_label_zh.setText(notice_html_zh)
        right_layout.addWidget(notice_label_zh)
        
        # 中文 - 計算公式
        formula_label_zh = QtWidgets.QLabel()
        formula_label_zh.setWordWrap(True)
        formula_label_zh.setTextFormat(QtCore.Qt.TextFormat.RichText)
        formula_label_zh.setFont(QtGui.QFont("Microsoft JhengHei, sans-serif", 10))
        formula_html_zh = """
<div style="background-color:#e8f4f8; padding:15px; border-radius:6px; font-size:13px; line-height:1.6; font-family:'Microsoft JhengHei', 'PingFang TC', 'Noto Sans CJK TC', sans-serif; border: 2px solid #344CB7;">
  <strong style='font-size:15px;'>📘 計算公式</strong>
  <table style="font-size:12px; margin-top:12px; font-family:'Microsoft JhengHei', 'PingFang TC', 'Noto Sans CJK TC', sans-serif; width:100%;">
    <tr>
      <td style="vertical-align:top; padding:8px; padding-right:12px; width:35%;"><strong>平均值匹配指數：</strong></td>
      <td style="padding:8px;">
        <u>兩組比較：</u><br>
        <code style='background:#fff; padding:2px 6px; border-radius:3px;'>|μ₁ − μ₂| / min(σ₁, σ₂)</code><br><br>
        <u>多組比較：</u><br>
        <code style='background:#fff; padding:2px 6px; border-radius:3px;'>|μ − median(μ)| / median(σ)</code>
      </td>
    </tr>
    <tr>
      <td style="vertical-align:top; padding:8px; padding-right:12px;"><strong>標準差匹配指數：</strong></td>
      <td style="padding:8px;">
        <u>兩組比較：</u><br>
        <code style='background:#fff; padding:2px 6px; border-radius:3px;'>σ / min(σ₁, σ₂)</code><br><br>
        <u>多組比較：</u><br>
        <code style='background:#fff; padding:2px 6px; border-radius:3px;'>σ / median(σ)</code>
      </td>
    </tr>
    <tr>
      <td style="vertical-align:top; padding:8px; padding-right:12px;"><strong>K 值：</strong></td>
      <td style="padding:8px;">
        <code style='background:#fff; padding:6px 10px; border-radius:3px; display:block; line-height:1.8;'>
          n = 樣本數<br>
          n ≤ 4：不進行比較<br>
          5 ≤ n ≤ 10：K = 1.73<br>
          11 ≤ n ≤ 120：K = 1.414<br>
          n > 120：K = 1.15
        </code>
      </td>
    </tr>
  </table>
  <div style="margin-top:15px; padding:12px; background:#fff; border-radius:4px; font-size:12px; color:#344CB7;">
    <strong>【篩選模式計算說明】</strong><br>
    <ul style="margin:8px 0 8px 20px; padding-left:0; line-height:1.8;">
      <li><strong>平均值/標準差/樣本數：</strong>對每個 matching_group，取「指定日期前一個月內」的資料，不足時補滿至指定筆數（預設5筆，可由UI調整），再計算 mean/std/count。</li>
      <li><strong>Median(sigma)：</strong>對每個 matching_group，取「指定日期前六個月內」的資料，不足時補滿至指定筆數（預設5筆），對每組計算 std 後取中位數。</li>
      <li><strong>圖表顯示：</strong>基於「一個月內（補滿至指定筆數）」的資料。</li>
      <li><strong>多筆資料：</strong>若同一組同一時間有多筆資料，皆納入計算。</li>
    </ul>
    <span style="color:#8a6d3b; font-style:italic;">• 全算模式直接使用所有資料進行分組計算，不進行補滿</span>
  </div>
</div>
        """
        formula_label_zh.setText(formula_html_zh)
        right_layout.addWidget(formula_label_zh)
        right_layout.addStretch()
        
        # 加入左右兩側到內容佈局
        content_layout.addWidget(left_widget)
        content_layout.addWidget(right_widget)
        
        scroll.setWidget(content_widget)
        layout.addWidget(scroll)
        
        # 關閉按鈕
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addStretch()
        close_btn = QtWidgets.QPushButton(tr("close"))
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #344CB7;
                color: white;
                border-radius: 8px;
                padding: 8px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #577BC1;
            }
            QPushButton:pressed {
                background-color: #000957;
            }
        """)
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        layout.addLayout(button_layout)

class ToolMatchingSettingsDialog(QtWidgets.QDialog):
    """Tool Matching 設定對話框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("settings"))
        self.setMinimumWidth(500)
        self.init_ui()
    
    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        
        # 分組：閾值設定
        threshold_group = QtWidgets.QGroupBox(tr("threshold_settings"))
        threshold_layout = QtWidgets.QVBoxLayout(threshold_group)
        
        # Mean Index 閾值
        mean_layout = QtWidgets.QHBoxLayout()
        self.mean_index_checkbox = QtWidgets.QCheckBox()
        self.mean_threshold_label = QtWidgets.QLabel(tr("mean_index_threshold"))
        self.mean_index_threshold_spin = QtWidgets.QDoubleSpinBox()
        self.mean_index_threshold_spin.setRange(0, 10)
        self.mean_index_threshold_spin.setValue(1.0)
        self.mean_index_threshold_spin.setSingleStep(0.1)
        self.mean_index_threshold_spin.setEnabled(False)
        mean_layout.addWidget(self.mean_index_checkbox)
        mean_layout.addWidget(self.mean_threshold_label)
        mean_layout.addWidget(self.mean_index_threshold_spin)
        mean_layout.addStretch()
        self.mean_index_checkbox.stateChanged.connect(
            lambda: self.mean_index_threshold_spin.setEnabled(self.mean_index_checkbox.isChecked())
        )
        threshold_layout.addLayout(mean_layout)
        
        # Sigma Index 閾值
        sigma_layout = QtWidgets.QHBoxLayout()
        self.sigma_index_checkbox = QtWidgets.QCheckBox()
        self.sigma_threshold_label = QtWidgets.QLabel(tr("sigma_index_threshold"))
        self.sigma_index_threshold_spin = QtWidgets.QDoubleSpinBox()
        self.sigma_index_threshold_spin.setRange(0, 10)
        self.sigma_index_threshold_spin.setValue(2.0)
        self.sigma_index_threshold_spin.setSingleStep(0.1)
        self.sigma_index_threshold_spin.setEnabled(False)
        sigma_layout.addWidget(self.sigma_index_checkbox)
        sigma_layout.addWidget(self.sigma_threshold_label)
        sigma_layout.addWidget(self.sigma_index_threshold_spin)
        sigma_layout.addStretch()
        self.sigma_index_checkbox.stateChanged.connect(
            lambda: self.sigma_index_threshold_spin.setEnabled(self.sigma_index_checkbox.isChecked())
        )
        threshold_layout.addLayout(sigma_layout)
        
        layout.addWidget(threshold_group)
        
        # 分組：數據處理設定
        data_group = QtWidgets.QGroupBox(tr("data_processing_settings"))
        data_layout = QtWidgets.QVBoxLayout(data_group)
        
        # 補滿樣本數
        fillnum_layout = QtWidgets.QHBoxLayout()
        self.fillnum_checkbox = QtWidgets.QCheckBox()
        self.fillnum_label = QtWidgets.QLabel(tr("fill_sample_size"))
        self.fillnum_spin = QtWidgets.QSpinBox()
        self.fillnum_spin.setMinimum(1)
        self.fillnum_spin.setMaximum(100)
        self.fillnum_spin.setValue(5)
        self.fillnum_spin.setEnabled(False)
        fillnum_layout.addWidget(self.fillnum_checkbox)
        fillnum_layout.addWidget(self.fillnum_label)
        fillnum_layout.addWidget(self.fillnum_spin)
        fillnum_layout.addStretch()
        self.fillnum_checkbox.stateChanged.connect(
            lambda: self.fillnum_spin.setEnabled(self.fillnum_checkbox.isChecked())
        )
        data_layout.addLayout(fillnum_layout)
        
        # 資料篩選模式
        filter_layout = QtWidgets.QHBoxLayout()
        self.filter_mode_label = QtWidgets.QLabel(tr("data_filter_mode"))
        self.filter_mode_combo = QtWidgets.QComboBox()
        self.filter_mode_combo.addItems([
            tr("all_data"),
            tr("specified_date"),
            tr("latest_entry")
        ])
        self.filter_mode_combo.setFixedWidth(260)
        filter_layout.addWidget(self.filter_mode_label)
        filter_layout.addWidget(self.filter_mode_combo)
        filter_layout.addStretch()
        data_layout.addLayout(filter_layout)
        
        # 指定日期
        date_layout = QtWidgets.QHBoxLayout()
        self.base_date_label = QtWidgets.QLabel(tr("specified_base_date"))
        self.date_edit = QtWidgets.QDateEdit(QtCore.QDate.currentDate())
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setEnabled(False)
        date_layout.addWidget(self.base_date_label)
        date_layout.addWidget(self.date_edit)
        date_layout.addStretch()
        self.filter_mode_combo.currentIndexChanged.connect(
            lambda idx: self.date_edit.setEnabled(idx == 1)
        )
        data_layout.addLayout(date_layout)
        
        layout.addWidget(data_group)
        
        # 按鈕
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addStretch()
        
        cancel_btn = QtWidgets.QPushButton(tr("cancel"))
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        save_btn = QtWidgets.QPushButton(tr("save"))
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #344CB7;
                color: white;
                border-radius: 8px;
                padding: 8px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #577BC1;
            }
            QPushButton:pressed {
                background-color: #000957;
            }
        """)
        save_btn.clicked.connect(self.accept)
        button_layout.addWidget(save_btn)
        
        layout.addLayout(button_layout)
    
    def get_settings(self):
        """獲取當前設定"""
        return {
            'mean_index_enabled': self.mean_index_checkbox.isChecked(),
            'mean_index_threshold': self.mean_index_threshold_spin.value(),
            'sigma_index_enabled': self.sigma_index_checkbox.isChecked(),
            'sigma_index_threshold': self.sigma_index_threshold_spin.value(),
            'fillnum_enabled': self.fillnum_checkbox.isChecked(),
            'fillnum_value': self.fillnum_spin.value(),
            'filter_mode': self.filter_mode_combo.currentIndex(),
            'base_date': self.date_edit.date()
        }
    
    def set_settings(self, settings):
        """設定當前值"""
        self.mean_index_checkbox.setChecked(settings.get('mean_index_enabled', False))
        self.mean_index_threshold_spin.setValue(settings.get('mean_index_threshold', 1.0))
        self.sigma_index_checkbox.setChecked(settings.get('sigma_index_enabled', False))
        self.sigma_index_threshold_spin.setValue(settings.get('sigma_index_threshold', 2.0))
        self.fillnum_checkbox.setChecked(settings.get('fillnum_enabled', False))
        self.fillnum_spin.setValue(settings.get('fillnum_value', 5))
        self.filter_mode_combo.setCurrentIndex(settings.get('filter_mode', 0))
        self.date_edit.setDate(settings.get('base_date', QtCore.QDate.currentDate()))

class ToolMatchingWidget(QtWidgets.QWidget):
    """
    Tool Matching Analysis Tool:
    - Read CSV files
    - Group by GroupName + ChartName
    - Perform mean/sigma matching checks based on characteristic
    - Display non-matching results
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_app = parent
        # 註冊翻譯觀察者
        self.translator = get_translator()
        self.translator.register_observer(self)
        # Set global font (CJK fallback: Microsoft JhengHei on Windows, system sans-serif elsewhere)
        font = QtGui.QFont()
        font.setFamilies(["Microsoft JhengHei", "PingFang TC", "Noto Sans CJK TC", "sans-serif"])
        font.setPointSize(10)
        self.setFont(font)
        
        # 初始化設定（預設值）
        self.tool_matching_settings = {
            'mean_index_enabled': False,
            'mean_index_threshold': 1.0,
            'sigma_index_enabled': False,
            'sigma_index_threshold': 2.0,
            'fillnum_enabled': False,
            'fillnum_value': 5,
            'filter_mode': 0,
            'base_date': QtCore.QDate.currentDate()
        }
        
        self.init_ui()

    def refresh_ui_texts(self):
        """刷新UI文字（當語言切換時）"""
        # 更新標題
        self.setWindowTitle(tr("tool_matching_title"))
        self.title_label.setText(f"<h2 style='color:#34495E;'>{tr('tool_matching_title')}</h2>")
        
        # 更新按鈕
        self.file_btn.setText(tr("browse_files_with_icon"))
        self.temp_btn.setText(tr("example_button"))
        self.settings_button.setText(f"⚙️ {tr('settings')}")
        self.formula_btn.setText(f"📊 {tr('formula_explanation')}")
        self.run_btn.setText(f"▶ {tr('run_analysis')}")
        
        # 更新輸入框 placeholder
        self.file_path_entry.setPlaceholderText(tr("please_select_csv"))
        
        # 更新狀態標籤
        self.status_label.setText(tr("select_file_prompt"))
        
        # 更新表格標題
        self.result_table.setHorizontalHeaderLabels([
            tr("group_name"), tr("chart_name"), tr("matching_group"), 
            tr("mean_index"), tr("sigma_index"), tr("k_value"),
            tr("mean"), tr("sigma"), tr("mean_median"), 
            tr("sigma_median"), tr("sample_size")
        ])
        
        print("ToolMatchingWidget UI texts refreshed")
    
    def init_ui(self):
        self.setWindowTitle(tr("tool_matching_title"))
        self.resize(1200, 800)

        # Main layout
        self.main_layout = QtWidgets.QVBoxLayout(self)
        self.setLayout(self.main_layout)

        # --- Top Control Area ---
        top_layout_widget = QtWidgets.QWidget()
        top_layout = QtWidgets.QVBoxLayout(top_layout_widget)

        self.title_label = QtWidgets.QLabel(f"<h2 style='color:#34495E;'>{tr('tool_matching_title')}</h2>")
        # Force apply CJK font to title
        title_font = QtGui.QFont()
        title_font.setFamilies(["Microsoft JhengHei", "PingFang TC", "Noto Sans CJK TC", "sans-serif"])
        title_font.setPointSize(16)
        self.title_label.setFont(title_font)
        top_layout.addWidget(self.title_label)

        file_layout = QtWidgets.QHBoxLayout()
        self.file_path_entry = QtWidgets.QLineEdit()
        self.file_path_entry.setPlaceholderText(tr("please_select_csv"))
        self.file_path_entry.setReadOnly(True)
        # 加入資料夾符號於「瀏覽檔案...」按鈕
        self.file_btn = QtWidgets.QPushButton()
        # 使用 emoji 📁 作為 icon，並將文字設為粗體，字體大小與執行按鈕一致
        self.file_btn.setText(tr("browse_files_with_icon"))
        self.file_btn.setIcon(QtGui.QIcon())  # 移除原本的 QStyle icon
        btn_font = QtGui.QFont()
        btn_font.setFamilies(["Microsoft JhengHei", "PingFang TC", "Noto Sans CJK TC", "sans-serif"])
        btn_font.setBold(True)
        btn_font.setPointSize(12)
        self.file_btn.setFont(btn_font)
        self.file_btn.setFixedWidth(180)
        self.file_btn.clicked.connect(self.select_file)
        file_layout.addWidget(self.file_path_entry)
        file_layout.addWidget(self.file_btn)

        # 新增 temp 按鈕
        self.temp_btn = QtWidgets.QPushButton(tr("example_button"))
        self.temp_btn.setFont(btn_font)
        self.temp_btn.setFixedWidth(140)
        self.temp_btn.clicked.connect(self.generate_temp_csv)
        file_layout.addWidget(self.temp_btn)
        
        top_layout.addLayout(file_layout)
        
        # 按鈕區域 - 水平佈局
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.setSpacing(10)
        
        # 設定按鈕
        self.settings_button = QtWidgets.QPushButton(f"⚙️ {tr('settings')}")
        self.settings_button.setMinimumHeight(45)
        self.settings_button.setMinimumWidth(120)
        btn_font = QtGui.QFont()
        btn_font.setFamilies(["Microsoft JhengHei", "PingFang TC", "Noto Sans CJK TC", "sans-serif"])
        btn_font.setBold(True)
        btn_font.setPointSize(11)
        self.settings_button.setFont(btn_font)
        self.settings_button.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                color: #333;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                padding: 8px 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #adb5bd;
            }
            QPushButton:pressed {
                background-color: #dee2e6;
            }
        """)
        self.settings_button.clicked.connect(self.open_tool_matching_settings)
        button_layout.addWidget(self.settings_button)
        
        # 公式說明按鈕
        self.formula_btn = QtWidgets.QPushButton(f"📊 {tr('formula_explanation')}")
        self.formula_btn.setMinimumHeight(45)
        self.formula_btn.setMinimumWidth(120)
        self.formula_btn.setFont(btn_font)
        self.formula_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                color: #333;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                padding: 8px 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #adb5bd;
            }
            QPushButton:pressed {
                background-color: #dee2e6;
            }
        """)
        self.formula_btn.clicked.connect(self.open_formula_explanation)
        button_layout.addWidget(self.formula_btn)
        
        # 執行按鈕
        self.run_btn = QtWidgets.QPushButton(f"▶ {tr('run_analysis')}")
        self.run_btn.setFont(btn_font)
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #344CB7;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #577BC1;
            }
            QPushButton:pressed {
                background-color: #000957;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.run_btn.clicked.connect(self.run_analysis)
        button_layout.addWidget(self.run_btn)
        
        button_layout.addStretch()
        top_layout.addLayout(button_layout)

        # 狀態標籤
        self.status_label = QtWidgets.QLabel(tr("select_file_prompt"))
        self.status_label.setFont(QtGui.QFont("Microsoft JhengHei, PingFang TC, Noto Sans CJK TC, sans-serif", 10))
        top_layout.addWidget(self.status_label)

        self.main_layout.addWidget(top_layout_widget)

        # --- 結果表格 ---
        self.result_table = QtWidgets.QTableWidget()
        self.result_table.setColumnCount(11) # 調整為 11 列，因為 Need_matching 不在 UI 顯示
        self.result_table.setHorizontalHeaderLabels([
            tr("group_name"), tr("chart_name"), tr("matching_group"), 
            tr("mean_index"), tr("sigma_index"), tr("k_value"),
            tr("mean"), tr("sigma"), tr("mean_median"), 
            tr("sigma_median"), tr("sample_size")
        ])
        self.result_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.result_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.result_table.setAlternatingRowColors(True)
        self.result_table.horizontalHeader().setStretchLastSection(True)
        # 選取整列時有淡藍底且字體顏色為深色，異常欄位紅字不會被蓋掉
        self.result_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #d0d0d0;
            }
            QHeaderView::section {
                background-color: #344CB7;
                color: white;
                padding: 4px;
                font-weight: bold;
            }
            QTableWidget::item {
                background: transparent;
            }
            QTableWidget::item:selected {
                background: #e6f0fa !important;
                color: #222 !important;
            }
        """)
        self.main_layout.addWidget(self.result_table, 1) # 表格佔用更多空間
    def open_tool_matching_settings(self):
        """打開 Tool Matching 設定對話框"""
        dialog = ToolMatchingSettingsDialog(self)
        # 載入當前設定
        dialog.set_settings(self.tool_matching_settings)
        
        # 顯示對話框並等待用戶操作
        if dialog.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            # 用戶點擊保存，獲取新設定
            self.tool_matching_settings = dialog.get_settings()
            print(f"Tool Matching 設定已更新: {self.tool_matching_settings}")
        else:
            print("Tool Matching 設定更改已取消")
    
    def open_formula_explanation(self):
        """打開公式說明對話框"""
        dialog = FormulaExplanationDialog(self)
        dialog.exec()
    
    def generate_temp_csv(self):
        # 預設範例資料
        data = {
            "GroupName": ["GroupA"],
            "ChartName": ["X"],
            "point_time": ["2023/5/15 14:39"],
            "matching_group": ["A"],
            "point_val": [99.88135943],
            "characteristic": ["Nominal"]
        }
        df = pd.DataFrame(data)
        # 彈出儲存檔案對話框
        save_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Example CSV File",
            "tool_matching_input_example.csv",
            "CSV Files (*.csv);;All Files (*.*)"
        )
        if not save_path:
            self.status_label.setText("Cancelled saving example file.")
            return
        try:
            df.to_csv(save_path, index=False, encoding="utf-8-sig")
            self.status_label.setText(f"Example CSV saved to: {save_path}")
        except Exception as e:
            self.status_label.setText(f"Example CSV generation failed: {e}")
    def select_file(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select CSV File", "", "CSV Files (*.csv);;All Files (*.*)"
        )
        if file_path:
            self.file_path_entry.setText(file_path)
            self.status_label.setText(f"Selected file: {os.path.basename(file_path)}")

    def get_k_value(self, n):
        """根據樣本數量 n 返回 K 值"""
        if n <= 4:  # 樣本數量太少，不進行比較
            return "No Comparison"  # Return special marker indicating no comparison
        elif 5 <= n <= 10:
            return 1.73
        elif 11 <= n <= 120:
            return 1.414
        else:
            return 1.15

    def calculate_mean_index(self, mean1, mean2, min_sigma, characteristic):
        """計算 mean matching index，考慮方向性"""
        if min_sigma <= 0:
            return float('inf')
        
        if characteristic == 'Bigger':  # Bigger is better
            return (mean2 - mean1) / min_sigma
        elif characteristic in ['Smaller', 'Sigma']:  # Smaller is better, Sigma 與 Smaller 邏輯相同
            return (mean1 - mean2) / min_sigma
        else:  # Nominal
            return abs(mean1 - mean2) / min_sigma

    def run_analysis(self):
        file_path = self.file_path_entry.text()
        if not file_path or not os.path.exists(file_path):
            self.status_label.setText("Please select a valid CSV file first!")
            return

        try:
            df = pd.read_csv(file_path)
            if hasattr(self.parent_app, 'filter_customer_data'):
                customer_result = self.parent_app.filter_customer_data(df, os.path.basename(file_path))
            else:
                customer_result = apply_customer_filter(df, set())
            if customer_result.missing_column:
                self.status_label.setText("Customer filter active, but the CSV has no Customer column.")
                return
            if customer_result.data.empty:
                self.status_label.setText("No data for the selected Customer(s).")
                return
            df = customer_result.data
        except Exception as e:
            self.status_label.setText(f"Failed to read file: {e}")
            return

        # 檢查必要欄位
        required_cols = ["GroupName", "ChartName", "matching_group", "point_val", "characteristic", "point_time"]
        for col in required_cols:
            if col not in df.columns:
                self.status_label.setText(f"Missing required column: {col}")
                return

        # 轉換 point_time 為 datetime
        try:
            df["point_time"] = pd.to_datetime(df["point_time"])
        except Exception as e:
            self.status_label.setText(f"point_time column conversion failed: {e}")
            return

        # 從設定字典獲取參數
        filter_mode = self.tool_matching_settings.get('filter_mode', 0)
        base_date = self.tool_matching_settings.get('base_date', QtCore.QDate.currentDate()).toPyDate() if filter_mode == 1 else None

        # 取得補滿筆數
        fill_num = self.tool_matching_settings.get('fillnum_value', 5)

        results = []

        if filter_mode == 0:
            # 全算
            grouped = df.groupby(["GroupName", "ChartName"])
            print("\n[DEBUG] All unique (GroupName, ChartName) pairs:")
            for pair in grouped.groups.keys():
                print("  ", pair)
            for (gname, cname), subdf in grouped:
                print(f"[DEBUG] Now processing group: GroupName='{gname}', ChartName='{cname}' | subdf.shape={subdf.shape}")
                characteristic = subdf["characteristic"].dropna().unique()
                if len(characteristic) != 1:
                    self.status_label.setText(f"Group: {gname}-{cname}  has non-unique or missing characteristic")
                    continue
                group_stats = subdf.groupby("matching_group")["point_val"].agg(['mean', 'std', 'count']).reset_index()
                n_groups = len(group_stats)
                if n_groups == 2:
                    self._analyze_two_groups(group_stats, gname, cname, characteristic[0], results)
                else:
                    self._analyze_multiple_groups(subdf, group_stats, gname, cname, characteristic[0], results)
            self._create_boxplots(grouped)
        elif filter_mode == 1:
            # 指定日期模式
            grouped = df.groupby(["GroupName", "ChartName"])
            print("\n[DEBUG] All unique (GroupName, ChartName) pairs:")
            for pair in grouped.groups.keys():
                print("  ", pair)
            sigma_df_all = []  # 收集所有半年資料
            mean_df_all = []   # 收集所有一個月(補到5筆)資料
            for (gname, cname), subdf in grouped:
                print(f"[DEBUG] Now processing group: GroupName='{gname}', ChartName='{cname}' | subdf.shape={subdf.shape}")
                characteristic = subdf["characteristic"].dropna().unique()
                if len(characteristic) != 1:
                    self.status_label.setText(f"Group: {gname}-{cname}  has non-unique or missing characteristic")
                    continue
                mean_end = pd.Timestamp(base_date)
                sigma_end = pd.Timestamp(base_date)
                mean_start = mean_end - pd.DateOffset(months=1)
                sigma_start = sigma_end - pd.DateOffset(months=6)
                # 先抓初始區間
                mean_df = subdf[(subdf["point_time"] > mean_start) & (subdf["point_time"] <= mean_end)].copy()
                sigma_df = subdf[(subdf["point_time"] > sigma_start) & (subdf["point_time"] <= sigma_end)].copy()
                # 針對每個 matching_group 補足 mean_df（只在不足時才補）
                min_time = subdf["point_time"].min()
                for mg in subdf["matching_group"].unique():
                    mg_mean = mean_df[mean_df["matching_group"] == mg]
                    if len(mg_mean) < fill_num:
                        all_mg = subdf[subdf["matching_group"] == mg].sort_values("point_time")
                        cur_start = mean_start
                        while len(mg_mean) < fill_num and cur_start > min_time:
                            cur_start = cur_start - pd.Timedelta(days=7)
                            mg_mean = all_mg[(all_mg["point_time"] > cur_start) & (all_mg["point_time"] <= mean_end)]
                        # 合併補足
                        mean_df = pd.concat([mean_df, mg_mean]).drop_duplicates()
                # sigma_df同理（只在不足時才補）
                for mg in subdf["matching_group"].unique():
                    mg_sigma = sigma_df[sigma_df["matching_group"] == mg]
                    if len(mg_sigma) < fill_num:
                        all_mg = subdf[subdf["matching_group"] == mg].sort_values("point_time")
                        cur_start = sigma_start
                        while len(mg_sigma) < fill_num and cur_start > min_time:
                            cur_start = cur_start - pd.Timedelta(days=14)
                            mg_sigma = all_mg[(all_mg["point_time"] > cur_start) & (all_mg["point_time"] <= sigma_end)]
                        sigma_df = pd.concat([sigma_df, mg_sigma]).drop_duplicates()
                mean_df_all.append(mean_df.assign(GroupName=gname, ChartName=cname))
                sigma_df_all.append(sigma_df.assign(GroupName=gname, ChartName=cname))
                mean_stats = mean_df.groupby("matching_group")["point_val"].agg(['mean', 'count']).reset_index()
                sigma_stats = sigma_df.groupby("matching_group")["point_val"].agg(['std']).reset_index()
                group_stats = pd.merge(mean_stats, sigma_stats, on="matching_group", how="outer")
                group_stats = group_stats.fillna({"mean": 0, "std": 0, "count": 0})
                n_groups = len(group_stats)
                if n_groups == 2:
                    self._analyze_two_groups(group_stats, gname, cname, characteristic[0], results)
                else:
                    self._analyze_multiple_groups_time(mean_df, sigma_df, group_stats, gname, cname, characteristic[0], results)
            if mean_df_all:
                mean_df_concat = pd.concat(mean_df_all, ignore_index=True)
                mean_grouped = mean_df_concat.groupby(["GroupName", "ChartName"])
                self._create_boxplots(mean_grouped)
            else:
                self._create_boxplots(grouped)
        elif filter_mode == 2:
            # 最新進點模式
            grouped = df.groupby(["GroupName", "ChartName"])
            sigma_df_all = []
            mean_df_all = []
            for (gname, cname), subdf in grouped:
                characteristic = subdf["characteristic"].dropna().unique()
                if len(characteristic) != 1:
                    self.status_label.setText(f"Group: {gname}-{cname}  has non-unique or missing characteristic")
                    continue
                latest_time = subdf["point_time"].max()
                mean_end = latest_time
                sigma_end = latest_time
                mean_start = mean_end - pd.DateOffset(months=1)
                sigma_start = sigma_end - pd.DateOffset(months=6)
                mean_df = subdf[(subdf["point_time"] > mean_start) & (subdf["point_time"] <= mean_end)].copy()
                sigma_df = subdf[(subdf["point_time"] > sigma_start) & (subdf["point_time"] <= sigma_end)].copy()
                min_time = subdf["point_time"].min()
                for mg in subdf["matching_group"].unique():
                    mg_mean = mean_df[mean_df["matching_group"] == mg]
                    if len(mg_mean) < fill_num:
                        all_mg = subdf[subdf["matching_group"] == mg].sort_values("point_time")
                        cur_start = mean_start
                        while len(mg_mean) < fill_num and cur_start > min_time:
                            cur_start = cur_start - pd.Timedelta(days=7)
                            mg_mean = all_mg[(all_mg["point_time"] > cur_start) & (all_mg["point_time"] <= mean_end)]
                        mean_df = pd.concat([mean_df, mg_mean]).drop_duplicates()
                for mg in subdf["matching_group"].unique():
                    mg_sigma = sigma_df[sigma_df["matching_group"] == mg]
                    if len(mg_sigma) < fill_num:
                        all_mg = subdf[subdf["matching_group"] == mg].sort_values("point_time")
                        cur_start = sigma_start
                        while len(mg_sigma) < fill_num and cur_start > min_time:
                            cur_start = cur_start - pd.Timedelta(days=14)
                            mg_sigma = all_mg[(all_mg["point_time"] > cur_start) & (all_mg["point_time"] <= sigma_end)]
                        sigma_df = pd.concat([sigma_df, mg_sigma]).drop_duplicates()
                mean_df_all.append(mean_df.assign(GroupName=gname, ChartName=cname))
                sigma_df_all.append(sigma_df.assign(GroupName=gname, ChartName=cname))
                mean_stats = mean_df.groupby("matching_group")["point_val"].agg(['mean', 'count']).reset_index()
                sigma_stats = sigma_df.groupby("matching_group")["point_val"].agg(['std']).reset_index()
                group_stats = pd.merge(mean_stats, sigma_stats, on="matching_group", how="outer")
                group_stats = group_stats.fillna({"mean": 0, "std": 0, "count": 0})
                n_groups = len(group_stats)
                if n_groups == 2:
                    self._analyze_two_groups(group_stats, gname, cname, characteristic[0], results)
                else:
                    self._analyze_multiple_groups_time(mean_df, sigma_df, group_stats, gname, cname, characteristic[0], results)
            if mean_df_all:
                mean_df_concat = pd.concat(mean_df_all, ignore_index=True)
                mean_grouped = mean_df_concat.groupby(["GroupName", "ChartName"])
                self._create_boxplots(mean_grouped)
            else:
                self._create_boxplots(grouped)

        self._display_results(results)

    def _analyze_multiple_groups_time(self, mean_df, sigma_df, group_stats, gname, cname, characteristic, results):
        """
        多組分析（mean/std/count 來自一個月 window，median(sigma) 來自半年 window）
        - mean, std, count: 來自 mean_df（一個月 window，補到5筆）
        - median_sigma: 來自 sigma_df（半年 window，補到5筆）
        """
        # 只納入樣本數 >= 5 的 group 計算 median
        valid_mean_df = mean_df.groupby("matching_group").filter(lambda x: len(x) >= 5)
        sigma_by_group = sigma_df.groupby("matching_group")["point_val"].std()
        valid_groups = group_stats[group_stats['count'] >= 5]['matching_group']
        valid_sigma = sigma_by_group[valid_groups] if not valid_groups.empty else pd.Series(dtype=float)
        
        # B. 多群自動切換邏輯：當只有2群樣本數足夠時，自動切換為兩群比較
        if len(valid_groups) == 2:
            print(f"[INFO] {gname}-{cname}: 多群時間模式自動切換為兩群比較 (有效群組數: {len(valid_groups)})")
            valid_stats = group_stats[group_stats['count'] >= 5]
            self._analyze_two_groups(valid_stats, gname, cname, characteristic, results)
            # 對於樣本數不足的群組，仍然添加到結果中標記為Insufficient Data
            insufficient_stats = group_stats[group_stats['count'] < 5]
            for i, row in insufficient_stats.iterrows():
                group = row["matching_group"]
                mean = row["mean"]
                std = row["std"]
                n = row["count"]
                results.append([
                    gname, cname, group, "group_all",
                    'Insufficient Data', 'Insufficient Data', 
                    self.get_k_value(n), mean, std, 
                    '-', '-', n, characteristic
                ])
            return
        
        # Failsafe: if there is only one or zero effective groups, mark all as insufficient data
        if len(valid_groups) <= 1:
            for i, row in group_stats.iterrows():
                group = row["matching_group"]
                mean = row["mean"]
                std = row["std"]
                n = row["count"]
                results.append([
                    gname, cname, group, "group_all",
                    'Insufficient Data', 'Insufficient Data', 
                    self.get_k_value(n), mean, std, 
                    '-', '-', n, characteristic
                ])
            return
        mean_median = valid_mean_df["point_val"].median() if not valid_mean_df.empty else 0
        median_sigma = valid_sigma.median() if not valid_sigma.empty else 0
        for i, row in group_stats.iterrows():
            group = row["matching_group"]
            mean = row["mean"]
            std = row["std"]  # 這是來自 mean_df（一個月 window）
            n = row["count"]
            if n < 5:
                results.append([
                    gname, cname, group, "group_all",
                    'Insufficient Data', 'Insufficient Data', 
                    self.get_k_value(n), mean, std, 
                    mean_median, median_sigma, n, characteristic
                ])
                continue
            if median_sigma > 0:
                # 使用 calculate_mean_index 方法計算方向性的 mean index
                mean_index = self.calculate_mean_index(mean, mean_median, median_sigma, characteristic)
                sigma_index = std / median_sigma
            else:
                # 分母為零時，判斷所有 mean 是否相等
                all_means = group_stats['mean'].tolist() if not group_stats.empty else [mean]
                if len(set([round(m, 8) for m in all_means])) == 1:
                    mean_index = 0
                    sigma_index = 0
                else:
                    mean_index = float('inf')
                    sigma_index = float('inf')
            K = self.get_k_value(n)
            if K == "No Comparison":
                results.append([
                    gname, cname, group, "group_all",
                    'Insufficient Data', 'Insufficient Data', 
                    'No Comparison', round(mean, 2), round(std, 2), 
                    round(mean_median, 2), round(median_sigma, 2), n, characteristic
                ])
            else:
                results.append([
                    gname, cname, group, "group_all",
                    round(mean_index, 2), round(sigma_index, 2), 
                    round(K, 2), round(mean, 2), round(std, 2), 
                    round(mean_median, 2), round(median_sigma, 2), n, characteristic
                ])

    def _analyze_two_groups(self, group_stats, gname, cname, characteristic, results):
        """分析兩台設備的匹配情況"""
        row1 = group_stats.iloc[0]
        row2 = group_stats.iloc[1]

        group1 = row1["matching_group"]
        group2 = row2["matching_group"]
        mean1, std1, n1 = row1["mean"], row1["std"], row1["count"]
        mean2, std2, n2 = row2["mean"], row2["std"], row2["count"]

        min_sigma = min(std1, std2)

        # 統一格式：第4欄都用 'group_all'，與多群分析一致
        # mean_median, sigma_median 欄位（兩組時用 mean2, min_sigma 或 mean1, min_sigma）
        # 這裡用 mean2, min_sigma for group1, mean1, min_sigma for group2

        if n1 < 5 or n2 < 5:
            results.append([
                gname, cname, group1, 'group_all',
                'Insufficient Data', 'Insufficient Data',
                self.get_k_value(n1), mean1, std1,
                mean2, min_sigma, n1, characteristic
            ])
            results.append([
                gname, cname, group2, 'group_all',
                'Insufficient Data', 'Insufficient Data',
                self.get_k_value(n2), mean2, std2,
                mean1, min_sigma, n2, characteristic
            ])
            return

        k1 = self.get_k_value(n1)
        k2 = self.get_k_value(n2)

        # 使用 calculate_mean_index 方法計算方向性的 mean index
        if min_sigma > 0:
            mean_index_1 = self.calculate_mean_index(mean1, mean2, min_sigma, characteristic)
            sigma_index_1 = std1 / min_sigma
        else:
            all_means = [mean1, mean2]
            if len(set([round(m, 8) for m in all_means])) == 1:
                mean_index_1 = 0
                sigma_index_1 = 0
            else:
                mean_index_1 = float('inf')
                sigma_index_1 = float('inf')

        if k1 == "No Comparison":
            results.append([
                gname, cname, group1, 'group_all',
                'Insufficient Data', 'Insufficient Data',
                'No Comparison', round(mean1, 2), round(std1, 2),
                round(mean2, 2), round(min_sigma, 2), n1, characteristic
            ])
        else:
            results.append([
                gname, cname, group1, 'group_all',
                round(mean_index_1, 2), round(sigma_index_1, 2),
                round(k1, 2), round(mean1, 2), round(std1, 2),
                round(mean2, 2), round(min_sigma, 2), n1, characteristic
            ])

        # 第二組
        if min_sigma > 0:
            mean_index_2 = self.calculate_mean_index(mean2, mean1, min_sigma, characteristic)
            sigma_index_2 = std2 / min_sigma
        else:
            all_means = [mean1, mean2]
            if len(set([round(m, 8) for m in all_means])) == 1:
                mean_index_2 = 0
                sigma_index_2 = 0
            else:
                mean_index_2 = float('inf')
                sigma_index_2 = float('inf')

        if k2 == "No Comparison":
            results.append([
                gname, cname, group2, 'group_all',
                'Insufficient Data', 'Insufficient Data',
                'No Comparison', round(mean2, 2), round(std2, 2),
                round(mean1, 2), round(min_sigma, 2), n2, characteristic
            ])
        else:
            results.append([
                gname, cname, group2, 'group_all',
                round(mean_index_2, 2), round(sigma_index_2, 2),
                round(k2, 2), round(mean2, 2), round(std2, 2),
                round(mean1, 2), round(min_sigma, 2), n2, characteristic
            ])

    def _analyze_multiple_groups(self, subdf, group_stats, gname, cname, characteristic, results):
        """分析多台設備的匹配情況 (mean matching index 分母都用 median_sigma)"""
        # 只納入樣本數 >= 5 的 group 計算 median
        valid_stats = group_stats[group_stats['count'] >= 5]
        
        # B. 多群自動切換邏輯：當只有2群樣本數足夠時，自動切換為兩群比較
        if valid_stats.shape[0] == 2:
            print(f"[INFO] {gname}-{cname}: 多群模式自動切換為兩群比較 (有效群組數: {valid_stats.shape[0]})")
            self._analyze_two_groups(valid_stats, gname, cname, characteristic, results)
            # 對於樣本數不足的群組，仍然添加到結果中標記為Insufficient Data
            insufficient_stats = group_stats[group_stats['count'] < 5]
            for i, row in insufficient_stats.iterrows():
                group = row["matching_group"]
                mean = row["mean"]
                std = row["std"]
                n = row["count"]
                results.append([
                    gname, cname, group, "group_all",
                    'Insufficient Data', 'Insufficient Data', 
                    self.get_k_value(n), mean, std, 
                    '-', '-', n, characteristic
                ])
            return
        
        if valid_stats.shape[0] <= 1:
            # 只有一個或零個有效群組，全部標記Insufficient Data
            for i, row in group_stats.iterrows():
                group = row["matching_group"]
                mean = row["mean"]
                std = row["std"]
                n = row["count"]
                results.append([
                    gname, cname, group, "group_all",
                    'Insufficient Data', 'Insufficient Data', 
                    self.get_k_value(n), mean, std, 
                    '-', '-', n, characteristic
                ])
            return

        mean_median = valid_stats['mean'].median() if not valid_stats.empty else 0
        median_sigma = valid_stats['std'].median() if not valid_stats.empty else 0

        for i, row in group_stats.iterrows():
            group = row["matching_group"]
            mean = row["mean"]
            std = row["std"]
            n = row["count"]

            # 計算 mean matching index（考慮方向性）
            if n < 5:  # 樣本數不足5個，不進行比較
                results.append([
                    gname, cname, group, "group_all",
                    'Insufficient Data', 'Insufficient Data', 
                    self.get_k_value(n), mean, std, 
                    mean_median, median_sigma, n, characteristic
                ])
                continue

            if median_sigma > 0:
                # 使用 calculate_mean_index 方法計算方向性的 mean index
                mean_index = self.calculate_mean_index(mean, mean_median, median_sigma, characteristic)
                sigma_index = std / median_sigma
            else:
                # 分母為零時，判斷所有 mean 是否相等
                all_means = group_stats['mean'].tolist() if not group_stats.empty else [mean]
                if len(set([round(m, 8) for m in all_means])) == 1:
                    mean_index = 0
                    sigma_index = 0
                else:
                    mean_index = float('inf')
                    sigma_index = float('inf')

            K = self.get_k_value(n)

            # Check if K value is the string "No Comparison"
            if K == "No Comparison":
                # 樣本數不足，使用 "Insufficient Data" 標記
                results.append([
                    gname, cname, group, "group_all",
                    'Insufficient Data', 'Insufficient Data', 
                    'No Comparison', round(mean, 2), round(std, 2), 
                    round(mean_median, 2), round(median_sigma, 2), n, characteristic
                ])
            else:
                # 正常比較情況
                # 無論是否匹配都添加結果，保證所有比較都出現在報表中
                results.append([
                    gname, cname, group, "group_all",
                    round(mean_index, 2), round(sigma_index, 2), 
                    round(K, 2), round(mean, 2), round(std, 2), 
                    round(mean_median, 2), round(median_sigma, 2), n, characteristic
                ])

    def _display_results(self, results):
        """以新格式顯示分析結果，並在表格中添加按鈕以查看詳情。"""
        # 儲存報告數據以供彈出視窗使用
        self.report_data = {}
        
        # 遍歷結果，整理報表資料
        for row in results:
            gname, cname = row[0], row[1]
            key = f"{gname}_{cname}"
            
            if key not in self.report_data:
                self.report_data[key] = {
                    "GroupName": gname,
                    "ChartName": cname,
                    "groups": {}
                }
            
            group1, group2 = row[2], row[3]
            mean_index = row[4]
            sigma_index = row[5]
            
            if len(row) >= 13:
                k_value, mean, sigma, mean_median, sigma_median, n, characteristic = row[6:13]
            else:
                k_value, mean, sigma, mean_median, sigma_median, n, characteristic = [""] * 6 + [row[6] if len(row) > 6 else ""]
            
            if group2 == "group_all":
                self.report_data[key]["groups"][group1] = {
                    "mean_matching_index": mean_index,
                    "sigma_matching_index": sigma_index,
                    "K": k_value,
                    "mean": mean,
                    "sigma": sigma,
                    "mean_median": mean_median,
                    "sigma_median": sigma_median,
                    "samplesize": n,
                    "characteristic": characteristic
                }
            else:
                if group1 not in self.report_data[key]["groups"]:
                    self.report_data[key]["groups"][group1] = {}
                self.report_data[key]["groups"][group1][group2] = {
                    "mean_matching_index": mean_index,
                    "sigma_matching_index": sigma_index,
                    "K": k_value,
                    "mean": mean,
                    "sigma": sigma,
                    "mean_median": mean_median,
                    "sigma_median": sigma_median,
                    "samplesize": n,
                    "characteristic": characteristic
                }

        all_table_rows = []
        abnormal_ui_rows = []
        
        for key, data in self.report_data.items():
            gname = data["GroupName"]
            cname = data["ChartName"]
            
            for group_id, stats in data["groups"].items():
                mean_index = stats.get("mean_matching_index", "")
                sigma_index = stats.get("sigma_matching_index", "")
                k_value = stats.get("K", "")
                
                is_abnormal = False
                is_data_insufficient = mean_index == 'Insufficient Data' or sigma_index == 'Insufficient Data' or k_value == 'No Comparison'
                abnormal_type = ""
                if not is_data_insufficient:
                    try:
                        # 從設定中讀取門檻值
                        mean_threshold = self.tool_matching_settings.get('mean_index_threshold', 1.0) if self.tool_matching_settings.get('mean_index_enabled', False) else 1.0
                        sigma_threshold = self.tool_matching_settings.get('sigma_index_threshold', 2.0) if self.tool_matching_settings.get('sigma_index_enabled', False) else (float(k_value) if k_value not in [None, '', 'No Comparison'] else 2.0)
                        mean_abn = float(mean_index) >= mean_threshold
                        sigma_abn = float(sigma_index) >= sigma_threshold
                        if mean_abn or sigma_abn:
                            is_abnormal = True
                            if mean_abn and sigma_abn:
                                abnormal_type = "Mean, Sigma"
                            elif mean_abn:
                                abnormal_type = "Mean"
                            elif sigma_abn:
                                abnormal_type = "Sigma"
                    except (ValueError, TypeError):
                        pass
                else:
                    abnormal_type = ""
                
                # 樣本數 n 強制轉為 int 顯示
                samplesize_val = stats.get("samplesize", "")
                try:
                    if samplesize_val != '' and samplesize_val is not None:
                        samplesize_val = int(float(samplesize_val))
                except Exception:
                    pass
                row_data = [
                    gname, cname, group_id,
                    stats.get("mean_matching_index", ""), stats.get("sigma_matching_index", ""),
                    stats.get("K", ""), stats.get("mean", ""), stats.get("sigma", ""),
                    stats.get("mean_median", ""), stats.get("sigma_median", ""),
                    samplesize_val, stats.get("characteristic", "")
                ]
                
                all_row_data = [is_abnormal, abnormal_type] + row_data
                all_table_rows.append(all_row_data)
                
                if is_abnormal or is_data_insufficient:
                    abnormal_ui_rows.append({
                        "key": (gname, cname),
                        "group_id": group_id,
                        "data": [abnormal_type] + row_data
                    })
        

        # 填充表格 (只顯示異常項目)
        self.result_table.setColumnCount(14)
        self.result_table.setHorizontalHeaderLabels([
            "View Details", "Abnormal Type", "Group Name", "Chart Name", "Matching Group", "Mean Index", "Sigma Index",
            "K", "Mean", "Sigma", "Mean Median", "Sigma Median", "Sample Size", "Characteristic"
        ])
        self.result_table.setRowCount(len(abnormal_ui_rows))


        for i, item_info in enumerate(abnormal_ui_rows):
            # 使用眼睛 icon 按鈕
            view_button = QtWidgets.QPushButton()
            eye_icon = self.style().standardIcon(QtWidgets.QStyle.StandardPixmap.SP_DesktopIcon)  # fallback 預設 icon
            # 嘗試用 PyQt6 內建的 eye icon，如果有
            try:
                eye_icon = self.style().standardIcon(QtWidgets.QStyle.StandardPixmap.SP_FileDialogContentsView)
            except Exception:
                pass
            view_button.setIcon(eye_icon)
            view_button.setToolTip("檢視詳細資訊")
            view_button.setFixedWidth(36)
            view_button.setFixedHeight(36)
            view_button.setIconSize(QtCore.QSize(22, 22))
            view_button.setStyleSheet("QPushButton { border: none; background: transparent; } QPushButton:hover { background: #e0e7ef; }")
            # 置中顯示
            cell_widget = QtWidgets.QWidget()
            layout = QtWidgets.QHBoxLayout(cell_widget)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(view_button)
            view_button.clicked.connect(
                lambda checked, key=item_info["key"], gid=item_info["group_id"]: self._show_details_dialog(key, gid)
            )
            self.result_table.setCellWidget(i, 0, cell_widget)

            # Fill other data (with additional abnormal type column)
            row_data = item_info["data"]
            for j, val in enumerate(row_data):
                # --- 格式化數值欄位為兩位小數 ---
                if j in [4,5,6,7,8,9,10]:  # Mean Index, Sigma Index, K, Mean, Sigma, Mean Median, Sigma Median
                    try:
                        if val != 'Insufficient Data' and val != 'No Comparison' and val != '' and val is not None:
                            val = float(val)
                            val = f"{val:.2f}"
                    except Exception:
                        pass
                item = QtWidgets.QTableWidgetItem(str(val))
                item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                # 標記異常值
                is_abnormal = False
                try:
                    mean_idx_val = float(row_data[4])
                    sigma_idx_val = float(row_data[5])
                    k_val = float(row_data[6])
                    if (j == 4 and mean_idx_val >= 1) or (j == 5 and sigma_idx_val >= k_val):
                        is_abnormal = True
                except (ValueError, TypeError):
                    pass
                # 只標記異常欄位為紅字，不設底色，避免 QSS 衝突
                if is_abnormal:
                    item.setForeground(QtGui.QColor("#D32F2F"))
                self.result_table.setItem(i, j + 1, item)

        self.result_table.resizeColumnsToContents()
        self.result_table.horizontalHeader().setStretchLastSection(True)

        # 匯出全部結果到 Excel 檔案
        if all_table_rows and hasattr(self, 'file_path_entry') and self.file_path_entry.text():
            self._export_to_excel(all_table_rows, self.file_path_entry.text())
        else:
            self.status_label.setText(f"Analysis completed, found {len(abnormal_ui_rows)} items requiring attention.")
            
        if len(abnormal_ui_rows) > 0:
            self.status_label.setText(f"Analysis completed, found {len(abnormal_ui_rows)} items requiring attention (total {len(all_table_rows)} items).")
        else:
            self.status_label.setText(f"Analysis completed, no items requiring attention found (total {len(all_table_rows)} items).")

    def _show_details_dialog(self, chart_key, group_id):
        """Pop up a window to display detailed information and charts, with data at the top and charts at the bottom."""
        try:
            from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
        except ImportError:
            QtWidgets.QMessageBox.warning(self, "Missing Package", "Matplotlib is required to display charts.")
            return

        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle(f"Detailed Information: {chart_key[0]} - {chart_key[1]} | Group: {group_id}")
        dialog.setMinimumSize(1400, 450) # 調整視窗大小以適應新佈局 (高度減少)

        main_layout = QtWidgets.QVBoxLayout(dialog)
        main_layout.setSpacing(10)

        # --- 上方：數據表格 (水平排列) ---
        try:
            stats = self.report_data[f"{chart_key[0]}_{chart_key[1]}"]["groups"][group_id]
        except KeyError:
            QtWidgets.QMessageBox.critical(self, "Error", "Cannot find detailed data for this item.")
            return


        info_group = QtWidgets.QGroupBox("Analysis Data")
        info_v_layout = QtWidgets.QVBoxLayout(info_group)

        # Get abnormal type
        # Need to recalculate abnormal type here, consistent with UI/Excel
        mean_index = stats.get("mean_matching_index", "")
        sigma_index = stats.get("sigma_matching_index", "")
        k_value = stats.get("K", "")
        abnormal_type = ""
        is_data_insufficient = mean_index == 'Insufficient Data' or sigma_index == 'Insufficient Data' or k_value == 'No Comparison'
        if not is_data_insufficient:
            try:
                mean_threshold = self.tool_matching_settings.get('mean_index_threshold', 1.0) if self.tool_matching_settings.get('mean_index_enabled', False) else 1.0
                sigma_threshold = self.tool_matching_settings.get('sigma_index_threshold', 2.0) if self.tool_matching_settings.get('sigma_index_enabled', False) else float(k_value) if k_value not in [None, '', 'No Comparison'] else 2.0
                mean_abn = float(mean_index) >= mean_threshold
                sigma_abn = float(sigma_index) >= sigma_threshold
                if mean_abn and sigma_abn:
                    abnormal_type = "Mean, Sigma"
                elif mean_abn:
                    abnormal_type = "Mean"
                elif sigma_abn:
                    abnormal_type = "Sigma"
            except (ValueError, TypeError):
                pass

        # 新增異常類型欄位
        headers = [
            "Abnormal Type", "Group Name", "Chart Name", "Matching Group", "Mean Index", "Sigma Index",
            "K", "Mean", "Sigma", "Mean Median", "Sigma Median", "Sample Size", "Characteristic"
        ]
        gname, cname = chart_key
        # 樣本數 n 強制轉為 int 顯示
        samplesize_val = stats.get("samplesize", "")
        try:
            if samplesize_val != '' and samplesize_val is not None:
                samplesize_val = int(float(samplesize_val))
        except Exception:
            pass
        row_values = [
            abnormal_type,
            gname, cname, group_id,
            mean_index, sigma_index,
            stats.get("K", ""), stats.get("mean", ""), stats.get("sigma", ""),
            stats.get("mean_median", ""), stats.get("sigma_median", ""),
            samplesize_val, stats.get("characteristic", "")
        ]

        info_table = QtWidgets.QTableWidget()
        info_table.setColumnCount(len(headers))
        info_table.setHorizontalHeaderLabels(headers)
        info_table.setRowCount(1)
        info_table.verticalHeader().setVisible(False)
        info_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)

        for j, value in enumerate(row_values):
            if j in [4,5,6,7,8,9,10]:  # Mean Index, Sigma Index, K, Mean, Sigma, Mean Median, Sigma Median
                try:
                    if value != 'Insufficient Data' and value != 'No Comparison' and value != '' and value is not None:
                        value = float(value)
                        value = f"{value:.2f}"
                except Exception:
                    pass
            item = QtWidgets.QTableWidgetItem(str(value))
            item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            info_table.setItem(0, j, item)

        info_table.resizeColumnsToContents()
        info_table.setFixedHeight(info_table.horizontalHeader().height() + info_table.rowHeight(0) + 5)
        info_v_layout.addWidget(info_table)
        main_layout.addWidget(info_group)

        # --- 下方：圖表區塊 ---
        charts_container_widget = QtWidgets.QWidget()
        charts_layout = QtWidgets.QHBoxLayout(charts_container_widget)

        if hasattr(self, 'chart_figures') and chart_key in self.chart_figures:
            figures = self.chart_figures[chart_key]
            
            if figures['scatter'] and figures['box']:
                # --- 解決圖表重複開啟變大問題 ---
                # 使用 pickle 進行深度複製，確保每次顯示都是全新的 Figure 物件
                scatter_fig_copy = pickle.loads(pickle.dumps(figures['scatter']))
                box_fig_copy = pickle.loads(pickle.dumps(figures['box']))

                scatter_canvas = FigureCanvas(scatter_fig_copy)
                box_canvas = FigureCanvas(box_fig_copy)
                # ------------------------------------
                
                scatter_canvas.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Expanding)
                box_canvas.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Expanding)

                charts_layout.addWidget(scatter_canvas)
                charts_layout.addWidget(box_canvas)
            else:
                charts_layout.addWidget(QtWidgets.QLabel("Charts for this item were not generated due to insufficient data."))
        else:
            charts_layout.addWidget(QtWidgets.QLabel("Cannot find corresponding charts."))
        
        main_layout.addWidget(charts_container_widget)

        # 設定佈局伸展因子，讓圖表區域佔用更多空間
        main_layout.setStretchFactor(info_group, 0) # 數據表格高度固定
        main_layout.setStretchFactor(charts_container_widget, 1) # 圖表區域填滿剩餘空間

        # --- 關閉按鈕 ---
        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.StandardButton.Close)
        button_box.rejected.connect(dialog.reject)
        main_layout.addWidget(button_box)

        dialog.exec()

    def _create_boxplots(self, grouped):
        """創建 SPC 圖和盒鬚圖，將 figure 物件保存在 self.chart_figures 中，不在 UI 上顯示。"""
        try:
            # 這些導入是必要的，因為 Matplotlib 在子線程或不同上下文中可能需要重新導入
            import matplotlib.pyplot as plt
            from matplotlib import cm
            import numpy as np
        except ImportError:
            print("[ERROR] Matplotlib is not installed.")
            return

        # 保存圖表與分組鍵的對應關係，用於後續的彈出視窗和 Excel 匯出
        self.chart_figures = {}
        
        # 為每個 (GroupName, ChartName) 組合創建圖表
        for (gname, cname), subdf in grouped:
            # 依 matching_group 字母順序排序
            unique_groups = sorted(subdf["matching_group"].unique(), key=lambda x: str(x))
            labels = [str(mg) for mg in unique_groups]

            # 檢查是否有數據可供繪圖
            if subdf.empty or not any(len(grp["point_val"]) > 0 for _, grp in subdf.groupby("matching_group")):
                print(f"[WARNING] Skipping chart creation for {gname} - {cname} due to empty data.")
                self.chart_figures[(gname, cname)] = {'scatter': None, 'box': None}
                continue

            # 依排序後 unique_groups 組裝 box_data，確保顏色/label/資料一致
            box_data = [subdf[subdf["matching_group"] == mg]["point_val"].values for mg in unique_groups]
            group_stats = subdf.groupby("matching_group")["point_val"].agg(['mean', 'std', 'count'])

            # 為不同的組設置顏色
            colors = cm.tab10(np.linspace(0, 1, len(unique_groups)))

            # 1. 創建 SPC 風格的圖表
            scatter_fig, scatter_ax = plt.subplots(figsize=(7, 4.5)) # 調整尺寸為較小的長方形
            
            # 計算整體統計量用於控制線
            all_values = subdf["point_val"].values
            # overall_mean = np.mean(all_values)
            # overall_std = np.std(all_values)
            

            # 為每個群組繪製數據點，按時間順序連線
            x_position = 0
            for i, mg in enumerate(unique_groups):
                group_data = subdf[subdf["matching_group"] == mg].sort_values("point_time")
                if not group_data.empty:
                    # 為每個群組創建連續的x位置
                    x_vals = np.arange(x_position, x_position + len(group_data))
                    y_vals = group_data["point_val"].values
                    
                    # 繪製數據點
                    scatter_ax.scatter(x_vals, y_vals, color=colors[i], alpha=0.8, s=40, label=f'{mg}', zorder=3)
                    
                    # 連接同組內的點
                    scatter_ax.plot(x_vals, y_vals, color=colors[i], alpha=0.5, linewidth=1, zorder=2)
                    
                    # 在群組間添加分隔線
                    if i < len(unique_groups) - 1:  # 不在最後一組後面加線
                        separator_x = x_position + len(group_data) - 0.5
                        scatter_ax.axvline(x=separator_x, color='gray', linestyle='-', alpha=0.3, zorder=1)
                    
                    x_position += len(group_data)
            
            # 設置圖表樣式
            scatter_ax.set_title(f"SPC Chart: {gname} - {cname}", fontsize=10)
            scatter_ax.set_xlabel("Sample Sequence (Grouped by Matching Group)")
            scatter_ax.set_ylabel("Point Value")
            scatter_ax.grid(True, linestyle='--', alpha=0.3, zorder=0)
            
            # 添加群組標籤在x軸上
            if unique_groups:
                group_positions = []
                x_pos = 0
                for mg in unique_groups:
                    group_size = len(subdf[subdf["matching_group"] == mg])
                    group_positions.append(x_pos + group_size/2 - 0.5)
                    x_pos += group_size
                
                # 設置x軸刻度和標籤
                scatter_ax.set_xticks(group_positions)
                scatter_ax.set_xticklabels(labels, rotation=0, ha='center')
                
                # 添加次要刻度顯示樣本序號
                scatter_ax.tick_params(axis='x', which='minor', bottom=True, top=False)
            
            # 調整圖例位置
            scatter_ax.legend(loc='upper left', bbox_to_anchor=(1.02, 1), fontsize='small')
            scatter_fig.tight_layout()

            # 2. 創建盒鬚圖
            box_fig, box_ax = plt.subplots(figsize=(7, 4.5)) # 調整尺寸為較小的長方形
            if box_data:
                bp = box_ax.boxplot(box_data, labels=labels, patch_artist=True, widths=0.6)
                for patch, color in zip(bp['boxes'], colors):
                    patch.set_facecolor(color)

                # legend 也照 unique_groups 順序
                legend_labels = [
                    f"{label}: μ={group_stats.loc[mg, 'mean']:.2f}, σ={group_stats.loc[mg, 'std']:.2f}, n={int(group_stats.loc[mg, 'count'])}"
                    for label, mg in zip(labels, unique_groups)
                ]
                box_ax.legend([bp["boxes"][i] for i in range(len(labels))], legend_labels, loc='upper left', bbox_to_anchor=(1.02, 1), fontsize='small')

            box_ax.set_title(f"Boxplot: {gname} - {cname}", fontsize=10)
            box_ax.set_xlabel("Matching Group")
            box_ax.set_ylabel("Point Value")
            box_ax.grid(True, linestyle='--', alpha=0.6)
            box_fig.subplots_adjust(right=0.7)
            box_fig.tight_layout()

            # 保存圖表與分組鍵的映射
            key = (gname, cname)
            self.chart_figures[key] = {'scatter': scatter_fig, 'box': box_fig}  # scatter實際上是SPC圖

            # 關鍵：關閉 figure 以釋放記憶體，因為我們已經將其保存在 self.chart_figures 中
            # FigureCanvas 會在需要時重新繪製它
            plt.close(scatter_fig)
            plt.close(box_fig)

    def _export_to_excel(self, all_results, source_path):
        """將分析結果匯出為 Excel 檔案，並在第一欄嵌入完整的盒鬚圖和散點圖。包含異常類型欄。"""
        try:
            # 檢查是否已安裝 openpyxl
            if openpyxl is None:
                QtWidgets.QMessageBox.warning(
                    self, "缺少套件", 
                    "請安裝 openpyxl 以匯出 Excel 檔案。\n可在終端執行: pip install openpyxl"
                )
                self.status_label.setText(f"分析完成。無法匯出 Excel：需要 openpyxl 套件。")
                return None

            # 嘗試導入所需的模組
            try:
                import matplotlib.pyplot as plt
                import numpy as np
                import io
                from PIL import Image
                import matplotlib.cm as cm
                from openpyxl.drawing.image import Image as XLImage
            except ImportError as e:
                QtWidgets.QMessageBox.warning(
                    self, "Missing Package", 
                    f"Embedding charts requires additional packages: {str(e)}\nPlease install the required packages."
                )
                print(f"[WARNING] Missing packages required for embedding charts: {e}")
                return None

            # Add abnormal type column, all_results: [is_abnormal, abnormal_type, ...]
            columns = [
                "Need_matching", "AbnormalType", "GroupName", "ChartName", "matching_group", "mean_matching_index", 
                "sigma_matching_index", "K", "mean", "sigma", "mean_median", "sigma_median", "samplesize", "characteristic"
            ]
            df = pd.DataFrame(all_results, columns=columns)

            # 打印資料框資訊以確認結構
            print(f"DataFrame info: {df.shape}")
            print(f"DataFrame columns: {df.columns.tolist()}")
            print(f"First row: {df.iloc[0].tolist() if len(df) > 0 else 'No data'}")

            # 生成輸出檔案路徑（與輸入檔案相同目錄）
            dir_path = os.path.dirname(source_path)
            file_name = os.path.splitext(os.path.basename(source_path))[0]
            output_path = os.path.join(dir_path, f"{file_name}_matching_results.xlsx")

            # 創建臨時目錄用於保存圖片
            import tempfile
            temp_dir = tempfile.mkdtemp()
            print(f"[INFO] 創建臨時目錄: {temp_dir}")

            # 先在 DataFrame 前添加兩個空白欄位，分別用於SPC圖和盒鬚圖
            df.insert(0, "SPC_Chart", "")    # 第一欄：SPC圖
            df.insert(1, "BoxPlot", "")      # 第二欄：盒鬚圖

            # 創建 Excel 文件
            writer = pd.ExcelWriter(output_path, engine='openpyxl')
            df.to_excel(writer, sheet_name='Tool Matching Results', index=False)

            # 獲取工作表
            workbook = writer.book
            worksheet = writer.sheets['Tool Matching Results']

            # 設定標題列格式
            header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
            header_fill = openpyxl.styles.PatternFill(start_color="344CB7", end_color="344CB7", fill_type="solid")
            header_alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")

            # 設置標題列格式
            for cell in worksheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment

            # 增加圖表欄寬度以容納圖片
            worksheet.column_dimensions['A'].width = 70  # 第一欄：SPC圖
            worksheet.column_dimensions['B'].width = 70  # 第二欄：盒鬚圖

            # 設定異常行的格式
            abnormal_fill = openpyxl.styles.PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

            # 定義圖表在 Excel 中顯示的尺寸 (單位：像素)
            img_display_width, img_display_height = 450, 250

            # 檢查是否有可用的圖表數據
            has_chart_figures = hasattr(self, 'chart_figures') and self.chart_figures
            if not has_chart_figures:
                print("[WARNING] 沒有可用的圖表數據，將使用簡單的狀態指示圖")

            # 從第二行開始遍歷（跳過標題行）
            for row_idx, row in enumerate(df.iterrows(), start=2):
                _, row_data = row

                # 檢查Need_matching欄位是否為True
                is_abnormal = row_data["Need_matching"]

                if is_abnormal:
                    # 將整行設為淺紅色
                    for cell in worksheet[row_idx]:
                        cell.fill = abnormal_fill

                # 創建並嵌入圖表到第一欄
                try:
                    # 獲取關鍵數據
                    group_name = str(row_data["GroupName"])
                    chart_name = str(row_data["ChartName"])
                    group_id = str(row_data["matching_group"])
                    mean_index = row_data["mean_matching_index"]
                    sigma_index = row_data["sigma_matching_index"]
                    k_value = row_data["K"]

                    # 檢查是否Insufficient Data
                    is_data_insufficient = (mean_index == 'Insufficient Data' or sigma_index == 'Insufficient Data' or k_value == 'No Comparison')

                    # 嘗試使用完整的SPC圖和盒鬚圖
                    chart_key = (group_name, chart_name)
                    if has_chart_figures and chart_key in self.chart_figures:
                        # 存在完整的分析圖表，使用實際的SPC圖和盒鬚圖
                        chart_data = self.chart_figures[chart_key]

                        # 1. 處理SPC圖 (放在第一欄)
                        try:
                            scatter_fig = chart_data['scatter']
                            temp_scatter_path = os.path.join(temp_dir, f"spc_{group_name}_{chart_name}_{row_idx}.png")
                            scatter_fig.savefig(temp_scatter_path, format='png', bbox_inches='tight', transparent=True, dpi=100)
                            try:
                                scatter_img = XLImage(temp_scatter_path)
                                scatter_img.width = img_display_width
                                scatter_img.height = img_display_height
                                scatter_position = f"A{row_idx}"
                                worksheet.add_image(scatter_img, scatter_position)
                                print(f"[INFO] 已添加SPC圖到單元格: {scatter_position}")
                            except Exception as img_e:
                                print(f"[ERROR] 添加SPC圖到 Excel 失敗: {img_e}")
                                worksheet.cell(row=row_idx, column=1).value = "SPC圖載入失敗"
                        except Exception as scatter_e:
                            print(f"[ERROR] Error occurred while processing SPC chart: {scatter_e}")
                            import traceback
                            traceback.print_exc()
                            worksheet.cell(row=row_idx, column=1).value = "SPC圖生成失敗"

                        # 2. 處理盒鬚圖 (放在第二欄)
                        try:
                            box_fig = chart_data['box']
                            temp_box_path = os.path.join(temp_dir, f"box_{group_name}_{chart_name}_{row_idx}.png")
                            box_fig.savefig(temp_box_path, format='png', bbox_inches='tight', transparent=True, dpi=100)
                            try:
                                box_img = XLImage(temp_box_path)
                                box_img.width = img_display_width
                                box_img.height = img_display_height
                                box_position = f"B{row_idx}"
                                worksheet.add_image(box_img, box_position)
                                print(f"[INFO] 已添加盒鬚圖到單元格: {box_position}")
                            except Exception as img_e:
                                print(f"[ERROR] 添加盒鬚圖到 Excel 失敗: {img_e}")
                                worksheet.cell(row=row_idx, column=2).value = "盒鬚圖載入失敗"
                        except Exception as box_e:
                            print(f"[ERROR] Error occurred while processing box plot: {box_e}")
                            import traceback
                            traceback.print_exc()
                            worksheet.cell(row=row_idx, column=2).value = "盒鬚圖生成失敗"

                    else:
                        # 沒有找到匹配的圖表，使用狀態指示器
                        print(f"[INFO] 未找到 {group_name}/{chart_name} 的分析圖表，使用狀態指示器")
                        fig, ax = plt.subplots(figsize=(6, 4), dpi=100)
                        title = f"{group_name}\n{chart_name}\n組別: {group_id}"
                        ax.set_title(title, fontsize=12)
                        if is_data_insufficient:
                            circle = plt.Circle((0.5, 0.5), 0.3, color='yellow', alpha=0.6, edgecolor='goldenrod', linewidth=2)
                            ax.add_patch(circle)
                            ax.text(0.5, 0.5, "Insufficient Data", ha='center', va='center', fontsize=14, color='black')
                            status_text = "Insufficient Data，無法進行分析"
                        elif is_abnormal:
                            circle = plt.Circle((0.5, 0.5), 0.3, color='red', alpha=0.6, edgecolor='darkred', linewidth=2)
                            ax.add_patch(circle)
                            ax.text(0.5, 0.5, "需要對齊", ha='center', va='center', fontsize=14, color='white', fontweight='bold')
                            status_text = f"均值差異指數: {mean_index}, 標準差差異指數: {sigma_index}, K值: {k_value}"
                        else:
                            circle = plt.Circle((0.5, 0.5), 0.3, color='green', alpha=0.6, edgecolor='darkgreen', linewidth=2)
                            ax.add_patch(circle)
                            ax.text(0.5, 0.5, "正常", ha='center', va='center', fontsize=14, color='white', fontweight='bold')
                            status_text = f"均值差異指數: {mean_index}, 標準差差異指數: {sigma_index}, K值: {k_value}"
                        ax.text(0.5, 0.2, status_text, ha='center', va='center', fontsize=10, 
                               bbox=dict(boxstyle='round,pad=0.5', facecolor='white', alpha=0.8))
                        ax.set_xticks([])
                        ax.set_yticks([])
                        ax.set_xlim(0, 1)
                        ax.set_ylim(0, 1)
                        ax.set_aspect('equal')
                        temp_img_path = os.path.join(temp_dir, f"status_chart_{row_idx}.png")
                        plt.savefig(temp_img_path, format='png', bbox_inches='tight', transparent=True, dpi=300)
                        plt.close(fig)
                        try:
                            # 使用 xlsxwriter 寫法 (insert_image) 取代 openpyxl 的 add_image
                            # 需先取得 xlsxwriter 的 worksheet 物件
                            # 但目前本程式是用 openpyxl，無法直接用 insert_image
                            # 所以這裡僅說明：如果你要用 insert_image，必須用 xlsxwriter 建立 writer
                            # 下面是 xlsxwriter 寫法範例：
                            # worksheet.insert_image(row_idx-1, 0, temp_img_path, {'x_scale': 1, 'y_scale': 1, 'x_offset': 0, 'y_offset': 0, 'object_position': 1})
                            # worksheet.insert_image(row_idx-1, 1, temp_img_path, {'x_scale': 1, 'y_scale': 1, 'x_offset': 0, 'y_offset': 0, 'object_position': 1})
                            # 但 openpyxl 不支援 insert_image，僅支援 add_image
                            # 若要完全改用 xlsxwriter，需重構整個 Excel 輸出流程。
                            # 這裡保留原本 openpyxl add_image 寫法，僅註明差異。
                            img1 = XLImage(temp_img_path)
                            img1.width = img_display_width
                            img1.height = img_display_height
                            cell_position_1 = f"A{row_idx}"
                            worksheet.add_image(img1, cell_position_1)
                            img2 = XLImage(temp_img_path)
                            img2.width = img_display_width
                            img2.height = img_display_height
                            cell_position_2 = f"B{row_idx}"
                            worksheet.add_image(img2, cell_position_2)
                            print(f"[INFO] 已添加狀態圖到單元格: {cell_position_1} 和 {cell_position_2}")
                        except Exception as img_e:
                            print(f"[ERROR] 添加圖片到 Excel 失敗: {img_e}")
                            worksheet.cell(row=row_idx, column=1).value = "圖片載入失敗"
                            worksheet.cell(row=row_idx, column=2).value = "圖片載入失敗"

                except Exception as img_e:
                    print(f"[ERROR] Error occurred while adding chart at row {row_idx}: {img_e}")
                    import traceback
                    traceback.print_exc()
                    worksheet.cell(row=row_idx, column=1).value = "圖片生成失敗"

            # 調整行高以適應圖表
            for i in range(2, worksheet.max_row + 1):
                worksheet.row_dimensions[i].height = 190

            # 調整其他列寬
            for col_idx, column in enumerate(worksheet.columns, start=1):
                if col_idx <= 2:  # 跳過圖表列 A 和 B，已手動設置寬度
                    continue
                max_length = 0
                column_letter = openpyxl.utils.get_column_letter(col_idx)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except Exception:
                        pass
                adjusted_width = (max_length + 4)
                worksheet.column_dimensions[column_letter].width = adjusted_width

            # 儲存 Excel 檔案
            try:
                writer.close()
                print(f"[INFO] Excel 檔案已儲存到: {output_path}")
            except Exception as save_e:
                print(f"[ERROR] 儲存 Excel 檔案失敗: {save_e}")
                import traceback
                traceback.print_exc()
            finally:
                try:
                    import shutil
                    shutil.rmtree(temp_dir)
                    print(f"[INFO] 已清理臨時目錄: {temp_dir}")
                except Exception as e:
                    print(f"[WARNING] Unable to clean temporary directory: {temp_dir}, Error: {e}")

            self.status_label.setText(f"Analysis completed. Results exported to: {output_path}")
            return output_path
        except Exception as e:
            self.status_label.setText(f"Excel export failed: {e}")
            import traceback
            traceback.print_exc()
            return None
