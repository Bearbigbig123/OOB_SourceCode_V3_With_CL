import sys
import os
import logging
os.environ["QT_API"] = "PyQt6" # 確認使用 PyQt6
import re

logger = logging.getLogger(__name__)
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import traceback
# CL Tighten Calculator
from CL_limit_class import CLTightenCalculator
# Translation System
from translations import TranslationManager, get_translator, tr
from customer_filter import apply_customer_filter, discover_customers
from chart_filter import filter_chart_information
# Excel 和圖片處理
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils.dataframe import dataframe_to_rows
import xlsxwriter
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas # 習慣上會將 FigureCanvasQTAgg 命名為 FigureCanvas
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
import mplcursors  # 新增 mplcursors 用於互動功能
# PyQt 和圖片顯示
from PyQt6 import QtWidgets, QtGui, QtCore
from PyQt6.QtCore import QDateTime
from PyQt6.QtWidgets import QMessageBox, QLabel, QVBoxLayout, QScrollArea, QGridLayout, QPushButton, QDateTimeEdit, QDateEdit, QHBoxLayout
from PIL import Image
from PIL.ImageQt import ImageQt
plt.rcParams['font.sans-serif'] = [
    'Microsoft JhengHei', 'Yu Gothic', 'Meiryo', 'Malgun Gothic',
    'Noto Sans CJK TC', 'Noto Sans CJK JP', 'Noto Sans CJK KR',
    'SimHei', 'Arial Unicode MS'
]
plt.rcParams['axes.unicode_minus'] = False  # 正確顯示負號

OOB_CALCULATED = 'CALCULATED'
OOB_INSUFFICIENT_DATA = 'INSUFFICIENT_DATA'
OOB_ERROR = 'ERROR'


def is_valid_number(value):
    """Return True only for a finite numeric value."""
    try:
        return value is not None and pd.notna(value) and np.isfinite(float(value))
    except (TypeError, ValueError):
        return False


def unavailable_kshift_result(status):
    """Build a K-shift result that cannot be mistaken for a normal result."""
    return pd.Series({
        'P95_k': np.nan,
        'P50_k': np.nan,
        'P05_k': np.nan,
        'P95_k_ori': np.nan,
        'P50_k_ori': np.nan,
        'P05_k_ori': np.nan,
        'P95_shift': status,
        'P50_shift': status,
        'P05_shift': status,
    })


def draw_horizontal_limit(plotter, value, label, x_min, x_max, fontsize=10):
    """Draw an optional control line without sending NaN/None to Matplotlib."""
    if not is_valid_number(value):
        return
    numeric_value = float(value)
    plotter.hlines(numeric_value, x_min, x_max, colors='#E83F6F', linestyles='--', linewidth=1)
    plotter.text(x=x_max, y=numeric_value, s=label, va='center', ha='left', fontsize=fontsize, color='#E83F6F')

# Toggle Switch 自定義類
class ToggleSwitch(QtWidgets.QWidget):
    """滑動開關 Widget"""
    stateChanged = QtCore.pyqtSignal(bool)
    
    def __init__(self, parent=None, label_text=""):
        super().__init__(parent)
        self._checked = False
        self.label_text = label_text
        
        # 創建主佈局
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        
        # 添加開關
        self.toggle = self._ToggleSwitchControl(self)
        self.toggle.clicked.connect(self._on_toggle_clicked)
        layout.addWidget(self.toggle)
        
        # 添加標籤
        self.label = QLabel(label_text)
        self.label.setFont(get_app_font(9))
        layout.addWidget(self.label)
        layout.addStretch()
        
    def _on_toggle_clicked(self):
        self._checked = not self._checked
        self.toggle.setChecked(self._checked)
        self.stateChanged.emit(self._checked)
    
    def isChecked(self):
        return self._checked
    
    def setChecked(self, checked):
        if self._checked != checked:
            self._checked = checked
            self.toggle.setChecked(checked)
            self.stateChanged.emit(checked)
    
    def setText(self, text):
        self.label_text = text
        self.label.setText(text)
    
    def text(self):
        return self.label_text
    
    class _ToggleSwitchControl(QtWidgets.QPushButton):
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setCheckable(True)
            self.setChecked(False)
            self.setFixedSize(50, 24)
            self.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
            self._update_style()
        
        def setChecked(self, checked):
            super().setChecked(checked)
            self._update_style()
        
        def _update_style(self):
            if self.isChecked():
                # 開啟狀態 - 藍色
                self.setStyleSheet("""
                    QPushButton {
                        background-color: #344CB7;
                        border: none;
                        border-radius: 12px;
                    }
                    QPushButton:hover {
                        background-color: #577BC1;
                    }
                """)
            else:
                # 關閉狀態 - 灰色
                self.setStyleSheet("""
                    QPushButton {
                        background-color: #ccc;
                        border: none;
                        border-radius: 12px;
                    }
                    QPushButton:hover {
                        background-color: #bbb;
                    }
                """)
        
        def paintEvent(self, event):
            super().paintEvent(event)
            painter = QtGui.QPainter(self)
            painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
            
            # 繪製白色圓形滑塊
            painter.setBrush(QtGui.QBrush(QtGui.QColor(255, 255, 255)))
            painter.setPen(QtCore.Qt.PenStyle.NoPen)
            
            if self.isChecked():
                # 開啟 - 圓形在右側
                painter.drawEllipse(28, 2, 20, 20)
            else:
                # 關閉 - 圓形在左側
                painter.drawEllipse(2, 2, 20, 20)

# OOB 設定對話框
class OOBSettingsDialog(QtWidgets.QDialog):
    """圖表處理設定對話框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_widget = parent
        from translations import tr
        
        self.setWindowTitle(tr("chart_processing_settings"))
        self.setMinimumWidth(720)
        self.setMinimumHeight(520)
        self.setObjectName("OOBSettingsDialog")
        
        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(24, 22, 24, 22)

        self.settings_intro_label = QtWidgets.QLabel(tr("oob_settings_intro"))
        self.settings_intro_label.setObjectName("SettingsIntro")
        self.settings_intro_label.setWordWrap(True)
        main_layout.addWidget(self.settings_intro_label)
        
        # === 圖表顯示設定 ===
        display_group = QtWidgets.QGroupBox(tr("display_settings"))
        display_group.setObjectName("SettingsCard")
        display_layout = QtWidgets.QVBoxLayout(display_group)
        display_layout.setSpacing(10)
        
        self.display_gui_checkbox = ToggleSwitch(label_text=tr("show_charts_gui"))
        self.display_gui_checkbox.setChecked(True)
        display_layout.addWidget(self.display_gui_checkbox)
        
        self.interactive_charts_checkbox = ToggleSwitch(label_text=tr("use_interactive_charts"))
        self.interactive_charts_checkbox.setChecked(True)
        display_layout.addWidget(self.interactive_charts_checkbox)
        
        self.use_batch_id_labels_checkbox = ToggleSwitch(label_text=tr("use_batch_id_labels"))
        self.use_batch_id_labels_checkbox.setChecked(False)
        display_layout.addWidget(self.use_batch_id_labels_checkbox)

        # 開啟後，既有兩張圖加上兩張 By Tool 時間圖，合計四張。
        self.show_by_tool_checkbox = ToggleSwitch(label_text=tr("show_by_tool_charts"))
        self.show_by_tool_checkbox.setChecked(False)
        display_layout.addWidget(self.show_by_tool_checkbox)
        self.show_by_tool_help_label = QtWidgets.QLabel(tr("show_by_tool_help"))
        self.show_by_tool_help_label.setObjectName("SettingsHelp")
        self.show_by_tool_help_label.setWordWrap(True)
        display_layout.addWidget(self.show_by_tool_help_label)
        
        main_layout.addWidget(display_group)

        self.oob_rule_group = QtWidgets.QGroupBox(tr("oob_rule_settings"))
        self.oob_rule_group.setObjectName("SettingsCard")
        oob_rule_layout = QtWidgets.QVBoxLayout(self.oob_rule_group)
        oob_rule_layout.setSpacing(8)

        self.record_high_low_checkbox = ToggleSwitch(
            label_text=tr("enable_record_high_low")
        )
        self.record_high_low_checkbox.setChecked(False)
        oob_rule_layout.addWidget(self.record_high_low_checkbox)

        self.record_high_low_help_label = QtWidgets.QLabel(tr("record_high_low_help"))
        self.record_high_low_help_label.setObjectName("SettingsHelp")
        self.record_high_low_help_label.setWordWrap(True)
        oob_rule_layout.addWidget(self.record_high_low_help_label)
        main_layout.addWidget(self.oob_rule_group)
        
        # === 時間範圍設定 ===
        time_range_group = QtWidgets.QGroupBox(tr("custom_time_range"))
        time_range_group.setObjectName("SettingsCard")
        time_range_layout = QtWidgets.QVBoxLayout(time_range_group)
        time_range_layout.setSpacing(10)
        
        # 啟用/停用自定義時間範圍
        self.custom_time_range_checkbox = QtWidgets.QCheckBox(tr("enable_custom_time_range"))
        self.custom_time_range_checkbox.setChecked(False)
        self.custom_time_range_checkbox.toggled.connect(self._toggle_time_range_controls)
        time_range_layout.addWidget(self.custom_time_range_checkbox)
        
        # 時間選擇器
        datetime_layout = QHBoxLayout()
        datetime_layout.setSpacing(10)
        
        # Start time
        self.start_time_label = QLabel(tr("start_time"))
        self.start_time_label.setMaximumWidth(80)
        self.start_datetime_edit = QDateEdit()
        self.start_datetime_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_datetime_edit.setDate(QtCore.QDate.currentDate().addDays(-30))
        self.start_datetime_edit.setEnabled(False)
        self.start_datetime_edit.setMaximumHeight(32)
        self.start_datetime_edit.setCalendarPopup(True)
        self.start_datetime_edit.setFixedWidth(130)
        self.start_datetime_edit.setDateRange(
            QtCore.QDate(2020, 1, 1),
            QtCore.QDate.currentDate()
        )
        
        # End time
        self.end_time_label = QLabel(tr("end_time"))
        self.end_time_label.setMaximumWidth(80)
        self.end_datetime_edit = QDateEdit()
        self.end_datetime_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_datetime_edit.setDate(QtCore.QDate.currentDate())
        self.end_datetime_edit.setEnabled(False)
        self.end_datetime_edit.setMaximumHeight(32)
        self.end_datetime_edit.setCalendarPopup(True)
        self.end_datetime_edit.setFixedWidth(130)
        self.end_datetime_edit.setDateRange(
            QtCore.QDate(2020, 1, 1),
            QtCore.QDate.currentDate().addYears(1)
        )
        
        datetime_layout.addWidget(self.start_time_label)
        datetime_layout.addWidget(self.start_datetime_edit)
        datetime_layout.addWidget(self.end_time_label)
        datetime_layout.addWidget(self.end_datetime_edit)
        datetime_layout.addStretch()
        
        time_range_layout.addLayout(datetime_layout)
        
        # 快速時間選擇按鈕
        quick_select_layout = QHBoxLayout()
        quick_select_layout.setSpacing(5)
        self.quick_select_label = QLabel(tr("quick_select"))
        self.quick_select_label.setMaximumWidth(85)
        quick_select_layout.addWidget(self.quick_select_label)
        
        self.last_7_days_btn = QtWidgets.QPushButton(tr("last_7_days"))
        self.last_30_days_btn = QtWidgets.QPushButton(tr("last_30_days"))
        self.last_90_days_btn = QtWidgets.QPushButton(tr("last_90_days"))
        self.this_month_btn = QtWidgets.QPushButton(tr("this_month"))
        self.last_month_btn = QtWidgets.QPushButton(tr("last_month"))
        
        for btn in [self.last_7_days_btn, self.last_30_days_btn, self.last_90_days_btn,
                   self.this_month_btn, self.last_month_btn]:
            btn.setMaximumWidth(90)
            btn.setMaximumHeight(28)
            btn.setEnabled(False)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #f0f0f0;
                    color: #666666;
                    border: 1px solid #cccccc;
                    border-radius: 6px;
                    padding: 5px 10px;
                    font-size: 12px;
                    font-weight: normal;
                }
                QPushButton:enabled {
                    background-color: #e8f4f8;
                    color: #344CB7;
                    border-color: #344CB7;
                    font-weight: bold;
                }
                QPushButton:enabled:hover {
                    background-color: #577BC1;
                    color: white;
                }
                QPushButton:enabled:pressed {
                    background-color: #000957;
                }
            """)
        
        quick_select_layout.addWidget(self.last_7_days_btn)
        quick_select_layout.addWidget(self.last_30_days_btn)
        quick_select_layout.addWidget(self.last_90_days_btn)
        quick_select_layout.addWidget(self.this_month_btn)
        quick_select_layout.addWidget(self.last_month_btn)
        quick_select_layout.addStretch()
        
        time_range_layout.addLayout(quick_select_layout)
        
        # 連接快速選擇按鈕
        self.last_7_days_btn.clicked.connect(lambda: self._set_quick_time_range(7))
        self.last_30_days_btn.clicked.connect(lambda: self._set_quick_time_range(30))
        self.last_90_days_btn.clicked.connect(lambda: self._set_quick_time_range(90))
        self.this_month_btn.clicked.connect(self._set_this_month_range)
        self.last_month_btn.clicked.connect(self._set_last_month_range)
        
        main_layout.addWidget(time_range_group)
        
        # 按鈕
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.cancel_btn = QtWidgets.QPushButton(tr("cancel"))
        self.cancel_btn.setObjectName("SecondaryDialogButton")
        self.cancel_btn.setMinimumWidth(100)
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)
        
        self.ok_btn = QtWidgets.QPushButton(tr("save"))
        self.ok_btn.setObjectName("PrimaryDialogButton")
        self.ok_btn.setMinimumWidth(100)
        self.ok_btn.setDefault(True)
        self.ok_btn.clicked.connect(self.accept)
        button_layout.addWidget(self.ok_btn)
        
        main_layout.addLayout(button_layout)

        self.setStyleSheet("""
            QDialog#OOBSettingsDialog { background-color: #F8FAFC; }
            QDialog#OOBSettingsDialog QLabel,
            QDialog#OOBSettingsDialog QCheckBox {
                background-color: transparent;
            }
            QLabel#SettingsIntro {
                color: #64748B; font-size: 13px; font-weight: 400;
                background: transparent; padding: 0 2px 4px 2px;
            }
            QGroupBox#SettingsCard {
                background-color: white;
                border: 1px solid #E2E8F0;
                border-radius: 14px;
                margin-top: 12px;
                padding: 14px;
                color: #0F172A;
                font-size: 14px;
                font-weight: 700;
            }
            QGroupBox#SettingsCard::title {
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 6px;
                background-color: #FFFFFF;
            }
            QLabel#SettingsHelp {
                color: #64748B;
                background: transparent;
                font-size: 11px;
                font-weight: 400;
                padding-left: 60px;
            }
            QPushButton#SecondaryDialogButton {
                background-color: white; color: #334155;
                border: 1px solid #CBD5E1; border-radius: 9px;
                padding: 9px 18px; font-weight: 650;
            }
            QPushButton#SecondaryDialogButton:hover { background-color: #F1F5F9; }
            QPushButton#PrimaryDialogButton {
                background-color: #4F46E5; color: white;
                border: none; border-radius: 9px;
                padding: 9px 20px; font-weight: 700;
            }
            QPushButton#PrimaryDialogButton:hover { background-color: #4338CA; }
        """)
    
    def _toggle_time_range_controls(self, checked):
        """ 啟用/停用時間範圍控制 """
        self.start_datetime_edit.setEnabled(checked)
        self.end_datetime_edit.setEnabled(checked)
        self.last_7_days_btn.setEnabled(checked)
        self.last_30_days_btn.setEnabled(checked)
        self.last_90_days_btn.setEnabled(checked)
        self.this_month_btn.setEnabled(checked)
        self.last_month_btn.setEnabled(checked)
    
    def _set_quick_time_range(self, days):
        """ 設定快速時間範圍 """
        end_date = QtCore.QDate.currentDate()
        start_date = end_date.addDays(-days)
        self.start_datetime_edit.setDate(start_date)
        self.end_datetime_edit.setDate(end_date)
    
    def _set_this_month_range(self):
        """ 設定本月範圍 """
        now = QtCore.QDate.currentDate()
        start_of_month = QtCore.QDate(now.year(), now.month(), 1)
        self.start_datetime_edit.setDate(start_of_month)
        self.end_datetime_edit.setDate(now)
    
    def _set_last_month_range(self):
        """ 設定上月範圍 """
        now = QtCore.QDate.currentDate()
        first_day_this_month = QtCore.QDate(now.year(), now.month(), 1)
        last_day_last_month = first_day_this_month.addDays(-1)
        first_day_last_month = QtCore.QDate(last_day_last_month.year(),
                                             last_day_last_month.month(), 1)
        self.start_datetime_edit.setDate(first_day_last_month)
        self.end_datetime_edit.setDate(last_day_last_month)
    
    def get_settings(self):
        """ 獲取所有設定 """
        return {
            'show_charts_gui': self.display_gui_checkbox.isChecked(),
            'use_interactive_charts': self.interactive_charts_checkbox.isChecked(),
            'use_batch_id_labels': self.use_batch_id_labels_checkbox.isChecked(),
            'show_by_tool_charts': self.show_by_tool_checkbox.isChecked(),
            'enable_record_high_low': self.record_high_low_checkbox.isChecked(),
            'custom_time_range_enabled': self.custom_time_range_checkbox.isChecked(),
            'start_time': self.start_datetime_edit.date(),
            'end_time': self.end_datetime_edit.date()
        }
    
    def set_settings(self, settings):
        """ 設定所有選項 """
        if 'show_charts_gui' in settings:
            self.display_gui_checkbox.setChecked(settings['show_charts_gui'])
        if 'use_interactive_charts' in settings:
            self.interactive_charts_checkbox.setChecked(settings['use_interactive_charts'])
        if 'use_batch_id_labels' in settings:
            self.use_batch_id_labels_checkbox.setChecked(settings['use_batch_id_labels'])
        if 'show_by_tool_charts' in settings:
            self.show_by_tool_checkbox.setChecked(settings['show_by_tool_charts'])
        self.record_high_low_checkbox.setChecked(
            settings.get('enable_record_high_low', False)
        )
        if 'custom_time_range_enabled' in settings:
            self.custom_time_range_checkbox.setChecked(settings['custom_time_range_enabled'])
        if 'start_time' in settings:
            # 支持 QDate 和 QDateTime 兼容
            start_val = settings['start_time']
            if isinstance(start_val, QtCore.QDateTime):
                self.start_datetime_edit.setDate(start_val.date())
            else:
                self.start_datetime_edit.setDate(start_val)
        if 'end_time' in settings:
            # 支持 QDate 和 QDateTime 兼容
            end_val = settings['end_time']
            if isinstance(end_val, QtCore.QDateTime):
                self.end_datetime_edit.setDate(end_val.date())
            else:
                self.end_datetime_edit.setDate(end_val)
    
    def refresh_ui_texts(self):
        """ 更新 UI 文字 """
        from translations import tr
        self.setWindowTitle(tr("chart_processing_settings"))
        self.display_gui_checkbox.setText(tr("show_charts_gui"))
        self.interactive_charts_checkbox.setText(tr("use_interactive_charts"))
        self.use_batch_id_labels_checkbox.setText(tr("use_batch_id_labels"))
        self.show_by_tool_checkbox.setText(tr("show_by_tool_charts"))
        self.show_by_tool_help_label.setText(tr("show_by_tool_help"))
        self.oob_rule_group.setTitle(tr("oob_rule_settings"))
        self.record_high_low_checkbox.setText(tr("enable_record_high_low"))
        self.record_high_low_help_label.setText(tr("record_high_low_help"))
        self.settings_intro_label.setText(tr("oob_settings_intro"))
        self.custom_time_range_checkbox.setText(tr("enable_custom_time_range"))
        self.start_time_label.setText(tr("start_time"))
        self.end_time_label.setText(tr("end_time"))
        self.quick_select_label.setText(tr("quick_select"))
        self.last_7_days_btn.setText(tr("last_7_days"))
        self.last_30_days_btn.setText(tr("last_30_days"))
        self.last_90_days_btn.setText(tr("last_90_days"))
        self.this_month_btn.setText(tr("this_month"))
        self.last_month_btn.setText(tr("last_month"))
        self.cancel_btn.setText(tr("cancel"))
        self.ok_btn.setText(tr("save"))

# 統一的字體設置函數 - 支持中文（微軟正黑體）
def get_app_font(size=9, weight=QtGui.QFont.Weight.Normal):
    """
    返回統一的應用程序字體
    字體優先級：Segoe UI -> 繁中 / 日文 / 韓文字型
    """
    font = QtGui.QFont()
    # 設置字體家族列表
    font.setFamilies([
        "Segoe UI", "Microsoft JhengHei", "Yu Gothic UI", "Malgun Gothic",
        "Meiryo UI", "Noto Sans CJK TC", "Noto Sans CJK JP",
        "Noto Sans CJK KR", "sans-serif"
    ])
    font.setPointSize(size)
    font.setWeight(weight)
    return font

def load_execution_time(raw_data_file):
    # 檢查檔案是否存在
    if not os.path.exists(raw_data_file):
        print(f" - load_execution_time: 檔案不存在: {raw_data_file}. 返回 None.")
        return None

    try:
        # 讀取 Excel 的 'Time' sheet
        # 加上 try/except 處理讀取 Sheet 可能發生的錯誤
        try:
            df = pd.read_excel(raw_data_file, sheet_name='Time', engine='openpyxl')
        except Exception as e:
            print(f" - load_execution_time: 無法讀取 'Time' Sheet 或檔案格式錯誤: {e}. 返回 None.")
            return None
      
        # === 修改點：檢查 DataFrame 是否為空 ===
        if df.empty:
            print(" - load_execution_time: 'Time' Sheet 是空的. 返回 None.")
            return None
        # ====================================

        # 確保有 'execTime' 這個欄位
        if 'execTime' not in df.columns:
            print(" - load_execution_time: 'Time' Sheet 中找不到 'execTime' 欄位. 返回 None.")
            return None

        # 嘗試從第一列取得 'execTime' 的值
        # 由於上面已經檢查 df.empty，這裡 df.iloc[0] 應該不會再報錯 SINGLE POSITIONAL INDEXER IS OUT OF BOUND
        execution_time_str = df.iloc[0]['execTime']

        # === 修改點：檢查讀取到的值是否為空或無效 ===
        if pd.isna(execution_time_str):
             print(" - load_execution_time: 'execTime' 儲存格是空的或無效. 返回 None.")
             return None
        # ========================================

        # 嘗試將字串轉換為時間格式
        # 加上 try/except 處理轉換失敗的錯誤
        try:
            execution_time = pd.to_datetime(execution_time_str, format='%Y-%m-%d %H:%M:%S')
            print(f" - load_execution_time: 成功讀取執行時間: {execution_time}")
            return execution_time
        except ValueError as e:
            print(f" - load_execution_time: 無法將 '{execution_time_str}' 轉換為日期時間: {e}. 返回 None.")
            return None # 轉換失敗也返回 None

    except Exception as e:
        # 捕捉讀取或處理過程中的其他未知錯誤
        print(f" - load_execution_time: 讀取執行時間時發生未知錯誤: {e}. 返回 None.")
        return None

def load_chart_information(raw_data_file):
    import pandas as pd
    print("載入圖表信息...")
    all_charts_info = pd.read_excel(raw_data_file, sheet_name='Chart', engine='openpyxl')
    
    # 必須欄位（不含 CHART_CREATE_TIME）
    required_columns = ['GroupName', 'ChartName', 'Material_no', 'USL', 'LSL', 'UCL', 'LCL', 'Target', 'ChartID', 'Characteristics']
    for col in required_columns:
        if col not in all_charts_info.columns:
            raise KeyError(f"欄位 '{col}' 不存在於圖表信息中")
    
    # 標準化 Characteristics 欄位（不分大小寫）
    if 'Characteristics' in all_charts_info.columns:
        all_charts_info['Characteristics'] = all_charts_info['Characteristics'].apply(normalize_characteristic)
        print(f"  已標準化 {len(all_charts_info)} 筆圖表的 Characteristics 欄位")
    
    return all_charts_info

def normalize_characteristic(value):
    """
    統一 Characteristics 欄位格式（不分大小寫）
    
    支援的值:
    - Nominal (預設)
    - Bigger
    - Smaller
    - Sigma
    
    Args:
        value: 原始 Characteristics 值
    
    Returns:
        str: 標準化後的 Characteristics 值
    """
    import pandas as pd
    
    if pd.isna(value):
        print("  [Warning] Characteristics 為空，使用預設值 'Nominal'")
        return 'Nominal'  # 預設值
    
    value_str = str(value).strip()
    
    # 不分大小寫對應表
    mapping = {
        'nominal': 'Nominal',
        'bigger': 'Bigger',
        'smaller': 'Smaller',
        'sigma': 'Sigma'
    }
    
    normalized = mapping.get(value_str.lower(), None)
    
    if normalized is None:
        print(f"  [Warning] 無效的 Characteristics 值: '{value_str}'，使用預設值 'Nominal'")
        return 'Nominal'
    
    # 如果原始值與標準值不同，輸出提示
    if value_str != normalized:
        print(f"  [Info] Characteristics 標準化: '{value_str}' → '{normalized}'")
    
    return normalized

def preprocess_raw_df(raw_df):
    import numpy as np
    import pandas as pd
    raw_df = raw_df.copy()
    raw_df.replace([np.inf, -np.inf, 'na', 'NA', 'NaN', 'nan'], np.nan, inplace=True)
    required_columns = ['GroupName', 'ChartName', 'point_val', 'Batch_ID', 'point_time']
    missing_columns = [col for col in required_columns if col not in raw_df.columns]
    if missing_columns:
        raise ValueError(f"原始數據缺少的欄位: {missing_columns}")
    raw_df['point_val'] = pd.to_numeric(raw_df['point_val'], errors='coerce')
    raw_df.loc[~np.isfinite(raw_df['point_val']), 'point_val'] = np.nan
    for column in ('GroupName', 'ChartName', 'Batch_ID', 'point_time'):
        raw_df[column] = raw_df[column].astype('str')
    return raw_df

def format_datetime(dt):
    import pandas as pd
    try:
        return pd.to_datetime(dt, format='%Y/%m/%d %H:%M', errors='coerce')
    except Exception as e:
        print(f"日期格式化錯誤: {e}")
        return pd.NaT

def format_and_clean_data(raw_df, chart_info):
    import pandas as pd
    raw_df = raw_df.copy()
    original_count = len(raw_df)

    # 計算、繪圖共用同一套有效值。空白、NaN、Inf、文字都排除。
    raw_df['point_val'] = pd.to_numeric(raw_df['point_val'], errors='coerce')
    raw_df.loc[~np.isfinite(raw_df['point_val']), 'point_val'] = np.nan

    # 性能優化：使用向量化操作代替 apply
    if not pd.api.types.is_datetime64_any_dtype(raw_df['point_time']):
        raw_df['point_time'] = pd.to_datetime(
            raw_df['point_time'],
            format='%Y/%m/%d %H:%M',
            errors='coerce'
        )
    
    # CHART_CREATE_TIME 為可選欄位
    if 'CHART_CREATE_TIME' in chart_info and pd.notna(chart_info['CHART_CREATE_TIME']):
        create_time = pd.to_datetime(chart_info['CHART_CREATE_TIME'], format="%m/%d/%Y %I:%M:%S %p", errors='coerce')
        if pd.notna(create_time):
            raw_df = raw_df[raw_df['point_time'] >= create_time]
    
    raw_df.dropna(subset=['point_val', 'point_time'], inplace=True)
    removed_count = original_count - len(raw_df)
    if removed_count:
        logger.warning(
            "OOB preprocessing removed %d invalid point_val/point_time rows; %d rows remain",
            removed_count,
            len(raw_df),
        )
    return raw_df
def update_chart_limits(raw_df, chart_info):
    import numpy as np
    # 排序並重設索引
    raw_df.sort_values(by='point_time', inplace=True)
    raw_df.reset_index(drop=True, inplace=True)
    
    # 確保必要的欄位存在，並初始化為 NaN
    required_columns = ['usl_val', 'lsl_val', 'ucl_val', 'lcl_val', 'target_val']
    for col in required_columns:
        if col not in raw_df.columns:
            raw_df[col] = np.nan  # 初始化欄位為 NaN
    
    # 使用向量化的方式更新欄位中的 NaN 值
    raw_df[required_columns] = raw_df[required_columns].fillna({
        'usl_val': chart_info['USL'],
        'lsl_val': chart_info['LSL'],
        'ucl_val': chart_info['UCL'],
        'lcl_val': chart_info['LCL'],
        'target_val': chart_info['Target']
    })
    
    # 四捨五入到 8 位小數
    # raw_df = raw_df.round(8)
    
    return raw_df, chart_info

def exclude_oos_data(raw_df, chart_info):
    import pandas as pd
    usl = raw_df['usl_val'].iat[0]
    lsl = raw_df['lsl_val'].iat[0]
    characteristic = normalize_characteristic(chart_info.get('Characteristics'))

    # 單邊特性只使用有意義的規格側；另一側允許留空。
    if characteristic in ('Smaller', 'Sigma'):
        return raw_df[raw_df['point_val'] <= usl] if is_valid_number(usl) else raw_df
    if characteristic == 'Bigger':
        return raw_df[raw_df['point_val'] >= lsl] if is_valid_number(lsl) else raw_df

    if is_valid_number(usl) and is_valid_number(lsl):
        return raw_df[(raw_df['point_val'] <= usl) & (raw_df['point_val'] >= lsl)]
    elif not is_valid_number(usl) and is_valid_number(lsl):
        return raw_df[raw_df['point_val'] >= lsl]
    elif not is_valid_number(lsl) and is_valid_number(usl):
        return raw_df[raw_df['point_val'] <= usl]
    return raw_df  # 如果都沒有符合條件，則直接回傳原始資料

# 優化後的 preprocess_data 函數
def preprocess_data(chart_info, raw_df):
    try:
        chart_info = chart_info.copy()
        for column in ('USL', 'LSL', 'UCL', 'LCL', 'Target', 'Resolution'):
            if column in chart_info:
                chart_info[column] = pd.to_numeric(chart_info[column], errors='coerce')
        chart_info['Characteristics'] = normalize_characteristic(chart_info.get('Characteristics'))

        raw_df = format_and_clean_data(raw_df, chart_info)  # 確保這個函數已經是最佳化的
        
        if raw_df.empty:
            return False, None, None
        
        raw_df, chart_info = update_chart_limits(raw_df, chart_info)  # 確保這個函數已經是最佳化的
        raw_df = exclude_oos_data(raw_df, chart_info)
        
        # 保留必要欄位，如果有 Batch_ID 則保留
        columns_to_keep = ['point_val', 'point_time']
        if 'Batch_ID' in raw_df.columns:
            columns_to_keep.append('Batch_ID')
        if 'Customer' in raw_df.columns:
            columns_to_keep.append('Customer')
        # By Tool 圖同時支援新舊 rawdata 欄位名稱。
        for tool_column in ('Matching', 'Tool'):
            if tool_column in raw_df.columns:
                columns_to_keep.append(tool_column)
        raw_df = raw_df[columns_to_keep]
        
        
        chart_info = chart_info.rename({
            'Material_no': 'material_no', 
            'GroupName': 'group_name',
            'ChartName': 'chart_name'
        })
        
        return True, raw_df, chart_info
    except ValueError as ve:
        print(f'跳過圖表處理，因爲缺少欄位: {ve}')
        return False, None, None
    except Exception as e:
        print(f'預處理過程中出錯: {e}')
        return False, None, None

# 優化後的 find_matching_file 函數
def find_matching_file(directory, group_name, chart_name):
    group_name = str(group_name)
    chart_name = str(chart_name)
    
    # 預編譯正則表達式
    pattern = re.compile(rf"{re.escape(group_name)}_{re.escape(chart_name)}(?:_\d+_\d+)?\.csv$")
    
    # 使用列表推導式來提高效率
    matching_files = [
        os.path.join(directory, filename)
        for filename in os.listdir(directory)
        if pattern.match(filename)
    ]
    
    return matching_files[0] if matching_files else None

# 優化後的 get_percentiles 函數
def get_percentiles(values):
    import numpy as np
    values = np.array(values)  # 確保數值是 NumPy 陣列，這樣計算會更快
    return {
        'P05': np.percentile(values, 5),
        'P50': np.percentile(values, 50),
        'P75': np.percentile(values, 75),
        'P25': np.percentile(values, 25),
        'P95': np.percentile(values, 95),
        'P99.865': np.percentile(values, 99.865),
        'P0.135': np.percentile(values, 0.135)
    }

# 優化後的 rolling_calculation 函數
def rolling_calculation(data_values, days_to_roll):
    # 確保數據是 NumPy 陣列，這樣切片會更高效
    data_values = np.array(data_values)
    
    # 滾動數據，取最後 'days_to_roll' 個元素
    return data_values[-days_to_roll:] if len(data_values) >= days_to_roll else data_values

def record_high_low_calculator(current_week_data, historical_data):
    """
    判斷當週數據是否創下歷史新高或新低
    
    Args:
        current_week_data: 當週數據的 point_val 值 (array-like)
        historical_data: 歷史數據的 point_val 值 (array-like)  
    
    Returns:
        dict: 包含 record_high, record_low, highlight_status 的字典
    """
    try:
        # 快速檢查：如果任一數據集為空，直接返回
        if len(current_week_data) == 0 or len(historical_data) == 0:
            return {
                'record_high': False,
                'record_low': False, 
                'highlight_status': 'NO_HIGHLIGHT'
            }
        
        # 性能優化：使用numpy操作，避免Python循環
        current_week_data = np.asarray(current_week_data)
        historical_data = np.asarray(historical_data)
        
        # DEBUG: 輸出數據詳細信息
        print(f"  DEBUG: 當週數據點數={len(current_week_data)}, 基線數據點數={len(historical_data)}")
        print(f"  DEBUG: 當週數據前5個值={current_week_data[:5] if len(current_week_data) >= 5 else current_week_data}")
        print(f"  DEBUG: 基線數據前5個值={historical_data[:5] if len(historical_data) >= 5 else historical_data}")
        print(f"  DEBUG: 基線數據後5個值={historical_data[-5:] if len(historical_data) >= 5 else historical_data}")
        
        # 計算當週最高值和最低值 - 使用numpy的快速操作
        current_max = np.max(current_week_data)
        current_min = np.min(current_week_data)
        
        # 計算歷史最高值和最低值 - 使用numpy的快速操作
        historical_max = np.max(historical_data)
        historical_min = np.min(historical_data)
        
        # DEBUG: 輸出詳細比較信息
        print(f"  DEBUG: 當週最高值={current_max:.8f}, 歷史最高值={historical_max:.8f}")
        print(f"  DEBUG: 當週最低值={current_min:.8f}, 歷史最低值={historical_min:.8f}")
        print(f"  DEBUG: 最高值差異={current_max - historical_max:.8f}")
        print(f"  DEBUG: 最低值差異={current_min - historical_min:.8f}")
        
        # 檢查當週數據是否包含歷史極值
        current_has_hist_max = np.any(current_week_data == historical_max)
        current_has_hist_min = np.any(current_week_data == historical_min)
        print(f"  DEBUG: 當週數據是否包含歷史最高值={current_has_hist_max}")
        print(f"  DEBUG: 當週數據是否包含歷史最低值={current_has_hist_min}")
        
        # 判斷是否創下新高或新低 - 簡單的數值比較，非常快速
        record_high = current_max > historical_max
        record_low = current_min < historical_min
        
        # 如果創下新高或新低，則需要高亮顯示
        highlight_status = 'HIGHLIGHT' if (record_high or record_low) else 'NO_HIGHLIGHT'
        
        print(f"  record_high_low_calculator: 當週最高={current_max:.4f}, 歷史最高={historical_max:.4f}, 創新高={record_high}")
        print(f"  record_high_low_calculator: 當週最低={current_min:.4f}, 歷史最低={historical_min:.4f}, 創新低={record_low}")
        print(f"  record_high_low_calculator: 高亮狀態={highlight_status}")
        
        return {
            'record_high': record_high,
            'record_low': record_low,
            'highlight_status': highlight_status
        }
        
    except Exception as e:
        print(f"  record_high_low_calculator: 計算過程中發生錯誤: {e}")
        return {
            'record_high': False,
            'record_low': False,
            'highlight_status': OOB_ERROR
        }


def optional_record_high_low_calculator(
    current_week_data,
    historical_data,
    enabled=False,
    data_available=True,
):
    """Run Record High/Low only when the optional OOB rule is enabled."""
    if not enabled:
        return {
            'record_high': False,
            'record_low': False,
            'highlight_status': 'NO_HIGHLIGHT',
        }
    if not data_available:
        return {
            'record_high': False,
            'record_low': False,
            'highlight_status': OOB_INSUFFICIENT_DATA,
        }
    return record_high_low_calculator(current_week_data, historical_data)
def review_kshift_results(results, resolution, characteristic, data_percentiles, base_percentiles):
    # 設定 highlight 的初始值
    highlight_conditions = {key: 'NO_HIGHLIGHT' for key in ['P95_shift', 'P50_shift', 'P05_shift']}

    # 檢查各個百分位數的 k 值是否需要高亮 (簡化版: 絕對差值 > resolution OR K絕對值 > 2)
    for percentile in ['P95', 'P50', 'P05']:
        k_value = results.get(f'{percentile}_k', np.nan) # 使用 .get 安全獲取 K 值 (絕對值)

        # 獲取當前和基線的百分位數，並計算絕對差值 (K 值計算的分子)
        data_p = data_percentiles.get(percentile, np.nan)
        base_p = base_percentiles.get(percentile, np.nan)
        abs_diff = np.nan # 預設絕對差值為 NaN

        if not pd.isna(data_p) and not pd.isna(base_p):
            abs_diff = abs(data_p - base_p)

        # --- 修改點開始：簡化高亮條件 ---

        # 判斷 resolution 是否有效（可選填）
        has_valid_resolution = not pd.isna(resolution) and resolution is not None and resolution > 0
        
        # 判斷絕對差值是否顯著
        if has_valid_resolution:
            # 有填寫 resolution: 使用 resolution 作為閾值
            is_significant_diff = not pd.isna(abs_diff) and abs_diff >= resolution
        else:
            # 沒填寫 resolution: 只要有差異就算顯著
            is_significant_diff = not pd.isna(abs_diff)

        # 判斷 K 絕對值是否超過 2 (且非NaN)
        is_significant_k = not pd.isna(k_value) and abs(k_value) > 2

        # 新增判斷 k_value 是否為無限值
        is_infinite_k = not pd.isna(k_value) and np.isinf(abs(k_value))

        # 設定初始高亮: 絕對差值 > resolution 且 (K絕對值 > 2 或 K絕對值為無限)
        if is_significant_diff and (is_significant_k or is_infinite_k):  # 使用 AND 和 OR 結合邏輯
            highlight_conditions[f'{percentile}_shift'] = 'HIGHLIGHT'   

        # --- 修改點結束 ---


    # 根據 characteristic 來更新 highlight (這部分邏輯不變，用於取消高亮)
    # 確保訪問 results 和 percentiles 時使用 .get() 和檢查 None/NaN
    if characteristic == 'Bigger':
        # 檢查 data_percentiles 和 base_percentiles 的鍵是否存在且值非空
        if data_percentiles.get('P95') is not None and base_percentiles.get('P05') is not None and data_percentiles['P95'] >= base_percentiles['P05']:
            highlight_conditions['P95_shift'] = 'NO_HIGHLIGHT'
        if data_percentiles.get('P50') is not None and base_percentiles.get('P25') is not None and data_percentiles['P50'] >= base_percentiles['P25']:
            highlight_conditions['P50_shift'] = 'NO_HIGHLIGHT'
        # 檢查 results 的鍵是否存在且值非空
        if results.get('P95_k_ori') is not None and results['P95_k_ori'] >= 0:
            highlight_conditions['P95_shift'] = 'NO_HIGHLIGHT'
        if results.get('P50_k_ori') is not None and results['P50_k_ori'] >= 0:
            highlight_conditions['P50_shift'] = 'NO_HIGHLIGHT'
        if results.get('P05_k_ori') is not None and results['P05_k_ori'] >= 0:
            highlight_conditions['P05_shift'] = 'NO_HIGHLIGHT'

    elif characteristic in ['Smaller', 'Sigma']:  # Sigma 使用與 Smaller 相同的邏輯
        if data_percentiles.get('P05') is not None and base_percentiles.get('P95') is not None and data_percentiles['P05'] <= base_percentiles['P95']:
            highlight_conditions['P05_shift'] = 'NO_HIGHLIGHT'
        if data_percentiles.get('P50') is not None and base_percentiles.get('P75') is not None and data_percentiles['P50'] <= base_percentiles['P75']:
            highlight_conditions['P50_shift'] = 'NO_HIGHLIGHT'
        if results.get('P95_k_ori') is not None and results['P95_k_ori'] <= 0:
            highlight_conditions['P95_shift'] = 'NO_HIGHLIGHT'
        if results.get('P50_k_ori') is not None and results['P50_k_ori'] <= 0:
            highlight_conditions['P50_shift'] = 'NO_HIGHLIGHT'
        if results.get('P05_k_ori') is not None and results['P05_k_ori'] <= 0:
            highlight_conditions['P05_shift'] = 'NO_HIGHLIGHT'

    elif characteristic == 'Nominal':
        if data_percentiles.get('P95') is not None and base_percentiles.get('P95') is not None and data_percentiles['P95'] <= base_percentiles['P95']:
            highlight_conditions['P95_shift'] = 'NO_HIGHLIGHT'
        if data_percentiles.get('P05') is not None and base_percentiles.get('P05') is not None and data_percentiles['P05'] >= base_percentiles['P05']:
            highlight_conditions['P05_shift'] = 'NO_HIGHLIGHT'
        # 檢查 P25, P50, P75 的鍵是否存在且值非空
        if (base_percentiles.get('P25') is not None and
            data_percentiles.get('P50') is not None and
            base_percentiles.get('P75') is not None and
            base_percentiles['P25'] <= data_percentiles['P50'] <= base_percentiles['P75']):
            highlight_conditions['P50_shift'] = 'NO_HIGHLIGHT'
        if results.get('P95_k_ori') is not None and results['P95_k_ori'] <= 0:
            highlight_conditions['P95_shift'] = 'NO_HIGHLIGHT'
        if results.get('P05_k_ori') is not None and results['P05_k_ori'] >= 0:
            highlight_conditions['P05_shift'] = 'NO_HIGHLIGHT'

    return highlight_conditions



def safe_division(numerator, denominator, epsilon=1e-9):
    """
    執行安全除法，避免除以零或極小值。
    如果分母接近零，返回 np.nan。
    """
    if abs(denominator) < epsilon:
        # print(f"  kshift: 警告：嘗試除以零或極小值 ({denominator})") # 可以在需要時打開這行
        return np.nan # 或者返回 float('inf'), 根據您希望在結果中如何表示這種情況
    return np.round(numerator, 8) / denominator


def kshift_sigma_ratio_calculator(base, data, characteristic, resolution, ucl, lcl):
    """
    計算 K-shift 和 Sigma 比例相關指標，並判斷高亮狀態。
    處理週數據點數為 1 時的滾動計算和數據填充。
    加入安全除法避免標準差為零導致的問題。
    """
    results = {
        'P95_k': np.nan,
        'P50_k': np.nan,
        'P05_k': np.nan,
        # 確保所有 results 的鍵都有初始值，包括 review_kshift_results 返回的
        'P95_k_ori': np.nan,
        'P50_k_ori': np.nan,
        'P05_k_ori': np.nan,
        'P95_shift': 'NO_HIGHLIGHT',
        'P50_shift': 'NO_HIGHLIGHT',
        'P05_shift': 'NO_HIGHLIGHT'
    }

    print("--- 進入 kshift_sigma_ratio_calculator 函數 ---")
    # 這裡 base 和 data 應該是字典，包含 'values' 和 'percentiles' (如果已計算)
    # 根據您之前的調試，analyze_chart 傳入的 base 和 data 是包含 'values', 'cnt', 'mean', 'sigma' 的字典
    # 但 kshift_sigma_ratio_calculator 內部使用了 base['values'] 和 data['values']
    # 這裡假設 base 和 data 是包含 'values' 鍵的字典
    if 'values' not in base or 'values' not in data:
         print("  kshift: 錯誤：輸入數據字典缺少 'values' 鍵。")
         return unavailable_kshift_result(OOB_ERROR)

    data_values = data['values']
    base_values = base['values']

    data_cnt = len(data_values)
    base_cnt = len(base_values) # 也獲取基線數據長度

    print(f"  kshift: 接收到的 data_values shape: {data_values.shape}, base_values shape: {base_values.shape}")
    print(f"  kshift: data_cnt: {data_cnt}, base_cnt: {base_cnt}")

    # 如果週數據少於 1 個點，直接返回預設結果
    if data_cnt < 1:
        print("  kshift: data_cnt < 1, 返回預設結果。")
        return unavailable_kshift_result(OOB_INSUFFICIENT_DATA)

    # 計算基線百分位數。請確保 get_percentiles 能處理 base_cnt = 3 的情況
    try:
        base_percentiles = get_percentiles(base_values)
        print(f"  kshift: 計算出的 base_percentiles (部分): P05={base_percentiles.get('P05')}, P50={base_percentiles.get('P50')}, P95={base_percentiles.get('P95')}")
        # 檢查計算分母所需的關鍵百分位數是否存在且不是 NaN
        if np.isnan(base_percentiles.get('P99.865', np.nan)) or np.isnan(base_percentiles.get('P0.135', np.nan)) or np.isnan(base_percentiles.get('P50', np.nan)):
             print("  kshift: 警告：基線百分位數計算結果無效 (包含 NaN)，可能基線數據不足。無法計算 K 值。")
             return unavailable_kshift_result(OOB_ERROR)

    except Exception as e:
         print(f"  kshift: 計算基線百分位數時發生錯誤: {e}")
         traceback.print_exc()
         return unavailable_kshift_result(OOB_ERROR)


    rolled_data = None # 預設沒有滾動數據
    data_percentiles = None # 預設沒有當前週數據的百分位數

    if data_cnt == 1:
        print("  kshift: 處理 data_cnt == 1 分支 (週數據只有 1 點)")
        days_to_roll = 1
        rolled_data_values = np.copy(data_values) # 從單點週數據開始

        # 修正無限迴圈：在 base_values 變空時跳出迴圈
        # 並且確保合併後的數據長度達到 5
        while len(rolled_data_values) < 5:
            print(f"  kshift: While 迴圈開始, days_to_roll: {days_to_roll}, rolled_data_values len: {len(rolled_data_values)}")

            # 如果 base_values 已經是空的，無法再滾動，跳出
            if len(base_values) == 0:
                print("  kshift: base_values 已經是空的，無法再滾動。跳出迴圈。")
                break

            # 呼叫 rolling_calculation 前
            print(f"  kshift: 呼叫 rolling_calculation 前, base_values shape: {base_values.shape}, days_to_roll: {days_to_roll}")
            try:
                rolled_base_values = rolling_calculation(base_values, days_to_roll)
                print(f"  kshift: rolling_calculation 返回 shape: {rolled_base_values.shape}")

                # 合併數據
                print(f"  kshift: 合併 rolled_data_values ({rolled_data_values.shape}) 與 rolled_base_values ({rolled_base_values.shape})")
                rolled_data_values = np.concatenate((rolled_data_values, rolled_base_values))
                print(f"  kshift: 合併後 rolled_data_values shape: {rolled_data_values.shape}")

                # 縮短 base_values，用於下一次迴圈
                base_values = base_values[:-days_to_roll]
                print(f"  kshift: 縮短後 base_values shape: {base_values.shape}")

            except Exception as e:
                 print(f"  kshift: rolling_calculation 或合併數據時發生錯誤: {e}")
                 traceback.print_exc()
                 # 如果發生錯誤，可能無法繼續，返回預設結果
                 return unavailable_kshift_result(OOB_ERROR)


            days_to_roll += 1
            # 保持安全跳出機制，防止意外情況
            if days_to_roll > base_cnt + 10: # 如果 rolling 天數遠超過原始基線數據，可能出錯
                print("  kshift: 警告：rolling 迴圈 days_to_roll 過大，可能邏輯有誤或數據問題。強制跳出。")
                break


        print(f"  kshift: While 迴圈結束。最終 rolled_data_values shape: {rolled_data_values.shape}")

        # 迴圈結束後，檢查是否湊滿了至少 5 個點用於滾動計算
        if len(rolled_data_values) < 5:
             print(f"  kshift: 警告：無法湊滿至少 5 個點用於滾動計算 (實際湊到 {len(rolled_data_values)} 點)。滾動結果將使用現有數據，計算可能不穩定。")
             # 您可以根據需求決定如果點數不足 5 時是否返回預設結果
             return unavailable_kshift_result(OOB_INSUFFICIENT_DATA)


        # 計算百分位數 (使用原始單點數據和滾動/填充後的數據)
        try:
            data_percentiles = get_percentiles(data_values) # 這是用原始單點週數據算的
            print(f"  kshift: 原始週數據 percentiles (data_cnt=1): {data_percentiles}")

            rolled_data_percentiles = get_percentiles(rolled_data_values) # 這是用滾動後的數據算的
            print(f"  kshift: 滾動數據 percentiles (shape={rolled_data_values.shape}): {rolled_data_percentiles}")

            rolled_data = {'values': rolled_data_values, 'percentiles': rolled_data_percentiles}

            # 檢查計算K值所需的當前和滾動數據百分位數是否存在且非NaN
            for p in ['P95', 'P50', 'P05']:
                if np.isnan(data_percentiles.get(p, np.nan)):
                     print(f"  kshift: 警告：原始週數據 {p} 百分位數為 NaN。無法計算 K 值。")
                     return unavailable_kshift_result(OOB_ERROR)
                if np.isnan(rolled_data_percentiles.get(p, np.nan)):
                     print(f"  kshift: 警告：滾動數據 {p} 百分位數為 NaN。影響滾動 K 值計算。") # 這裡只是警告，可能可以繼續

        except Exception as e:
            print(f"  kshift: 計算百分位數時發生錯誤 (data_cnt=1 分支): {e}")
            traceback.print_exc()
            return unavailable_kshift_result(OOB_ERROR)


    elif data_cnt >= 2:
        print(f"  kshift: 處理 data_cnt >= 2 分支, data_cnt: {data_cnt}")
        try:
             data_percentiles = get_percentiles(data_values)
             print(f"  kshift: 當前週數據 percentiles (data_cnt>1): {data_percentiles}")
             # 檢查計算K值所需的當前百分位數是否存在且非NaN
             for p in ['P95', 'P50', 'P05']:
                 if np.isnan(data_percentiles.get(p, np.nan)):
                      print(f"  kshift: 警告：當前週數據 {p} 百分位數為 NaN。無法計算 K 值。")
                      return unavailable_kshift_result(OOB_ERROR)

        except Exception as e:
            print(f"  kshift: 計算百分位數時發生錯誤 (data_cnt>=2 分支): {e}")
            traceback.print_exc()
            return unavailable_kshift_result(OOB_ERROR)


        rolled_data = None # data_cnt >= 2 時，沒有滾動數據的概念用於 highlight 判斷


    else: # 這個分支理論上不會走到，因為開頭已經處理 data_cnt < 1
        print(f"  kshift: Warning: 未預期的 data_cnt 情況: {data_cnt}")
        return unavailable_kshift_result(OOB_ERROR)

    # --- 計算分母 ---
    try:
        # 檢查 UCL/LCL 是否有效
        ucl_valid = not pd.isna(ucl) and ucl is not None
        lcl_valid = not pd.isna(lcl) and lcl is not None
        
        # 計算百分位數基礎分母值
        p95k_percentile = safe_division(base_percentiles.get('P99.865', np.nan) - base_percentiles.get('P50', np.nan), 3)
        p50k_percentile = safe_division(base_percentiles.get('P99.865', np.nan) - base_percentiles.get('P0.135', np.nan), 6)
        p05k_percentile = safe_division(base_percentiles.get('P50', np.nan) - base_percentiles.get('P0.135', np.nan), 3)
        
        if ucl_valid and not lcl_valid:
            # 單邊 UCL：P95/P50/P05 都以 (UCL - Base P50) / 6
            # 與各自的基線百分位分母取較大值。
            single_side_deno = safe_division(ucl - base_percentiles.get('P50', np.nan), 6)
            p95k_deno = np.round(np.max([p95k_percentile, single_side_deno]), 8)
            p50k_deno = np.round(np.max([p50k_percentile, single_side_deno]), 8)
            p05k_deno = np.round(np.max([p05k_percentile, single_side_deno]), 8)
            print(f"  kshift: 僅 UCL 有效，三個分母各自與 UCL 單邊分母 {single_side_deno} 取 max")
        elif lcl_valid and not ucl_valid:
            # 單邊 LCL：P95/P50/P05 都以 (Base P50 - LCL) / 6
            # 與各自的基線百分位分母取較大值。
            single_side_deno = safe_division(base_percentiles.get('P50', np.nan) - lcl, 6)
            p95k_deno = np.round(np.max([p95k_percentile, single_side_deno]), 8)
            p50k_deno = np.round(np.max([p50k_percentile, single_side_deno]), 8)
            p05k_deno = np.round(np.max([p05k_percentile, single_side_deno]), 8)
            print(f"  kshift: 僅 LCL 有效，三個分母各自與 LCL 單邊分母 {single_side_deno} 取 max")
        else:
            # 雙邊都有或雙邊都沒有時，維持原本的分母算法。
            if ucl_valid:
                p95k_ucl = safe_division(ucl - base_percentiles.get('P50', np.nan), 6)
                p95k_deno = np.round(np.max([p95k_percentile, p95k_ucl]), 8)
                print(f"  kshift: P95 分母使用 max(百分位數={p95k_percentile}, UCL計算={p95k_ucl}) = {p95k_deno}")
            else:
                p95k_deno = np.round(p95k_percentile, 8)
                print(f"  kshift: UCL 無效，P95 分母直接使用百分位數 = {p95k_deno}")

            if ucl_valid and lcl_valid:
                p50k_ucl_lcl = safe_division(ucl - lcl, 12)
                p50k_deno = np.round(np.max([p50k_percentile, p50k_ucl_lcl]), 8)
                print(f"  kshift: P50 分母使用 max(百分位數={p50k_percentile}, UCL-LCL計算={p50k_ucl_lcl}) = {p50k_deno}")
            else:
                p50k_deno = np.round(p50k_percentile, 8)
                print(f"  kshift: UCL/LCL 無效，P50 分母直接使用百分位數 = {p50k_deno}")

            if lcl_valid:
                p05k_lcl = safe_division(base_percentiles.get('P50', np.nan) - lcl, 6)
                p05k_deno = np.round(np.max([p05k_percentile, p05k_lcl]), 8)
                print(f"  kshift: P05 分母使用 max(百分位數={p05k_percentile}, LCL計算={p05k_lcl}) = {p05k_deno}")
            else:
                p05k_deno = np.round(p05k_percentile, 8)
                print(f"  kshift: LCL 無效，P05 分母直接使用百分位數 = {p05k_deno}")

        # YC edit：分母為 0 時的處理邏輯
        if p95k_deno == 0:
            if p05k_deno == 0:
                p95k_deno = p50k_deno
            elif p50k_deno == 0:
                p95k_deno = p05k_deno
            else:
                p95k_deno = min(p50k_deno, p05k_deno)
        if p05k_deno == 0:
            if p95k_deno == 0:
                p05k_deno = p50k_deno
            elif p50k_deno == 0:
                p05k_deno = p95k_deno
            else:
                p05k_deno = min(p50k_deno, p95k_deno)
        if p50k_deno == 0:
            if p95k_deno == 0:
                p50k_deno = p05k_deno
            elif p05k_deno == 0:
                p50k_deno = p95k_deno
            else:
                p50k_deno = min(p05k_deno, p95k_deno)

        denominators = {
            'p95k_deno': p95k_deno,
            'p50k_deno': p50k_deno,
            'p05k_deno': p05k_deno
        }
        print(f"  kshift: 計算出的分母: {denominators}")

        # 檢查分母是否有效 (非 NaN, 非 Inf)
        if np.isnan(p95k_deno) or np.isnan(p50k_deno) or np.isnan(p05k_deno) or np.isinf(p95k_deno) or np.isinf(p50k_deno) or np.isinf(p05k_deno):
            print("  kshift: 警告：計算出的分母無效 (包含 NaN 或 Inf)。無法計算 K 值。")
            return unavailable_kshift_result(OOB_ERROR)

    except Exception as e:
        print(f"  kshift: 計算分母時發生錯誤: {e}")
        traceback.print_exc()
        return unavailable_kshift_result(OOB_ERROR)


    # --- 計算 K 值 ---
    try:
        # 計算 K 值 (原始) - 使用安全除法
        results['P95_k_ori'] = safe_division(np.round(data_percentiles.get('P95', np.nan) - base_percentiles.get('P95', np.nan), 8), p95k_deno)
        results['P50_k_ori'] = safe_division(np.round(data_percentiles.get('P50', np.nan) - base_percentiles.get('P50', np.nan), 8), p50k_deno)
        results['P05_k_ori'] = safe_division(np.round(data_percentiles.get('P05', np.nan) - base_percentiles.get('P05', np.nan), 8), p05k_deno)

        # 計算 K 值 (絕對值) - 使用安全除法
        results['P95_k'] = safe_division(np.round(abs(data_percentiles.get('P95', np.nan) - base_percentiles.get('P95', np.nan)), 8), p95k_deno)
        results['P50_k'] = safe_division(np.round(abs(data_percentiles.get('P50', np.nan) - base_percentiles.get('P50', np.nan)), 8), p50k_deno)
        results['P05_k'] = safe_division(np.round(abs(data_percentiles.get('P05', np.nan) - base_percentiles.get('P05', np.nan)), 8), p05k_deno)

        if not all(is_valid_number(results[key]) for key in ('P95_k', 'P50_k', 'P05_k')):
            print("  kshift: K 值包含無效數值，標記為 ERROR。")
            return unavailable_kshift_result(OOB_ERROR)

        print(f"  kshift: 計算出的 K 值結果: {results}")

    except Exception as e:
        print(f"  kshift: 計算 K 值時發生錯誤: {e}")
        traceback.print_exc()
        return unavailable_kshift_result(OOB_ERROR)


    # --- 判斷當前高亮條件 ---
    try:
        # 確保傳給 review_kshift_results 的 percentiles 字典是完整的
        current_highlight_conditions = review_kshift_results(results, resolution, characteristic, data_percentiles, base_percentiles)
        print(f"  kshift: current_highlight_conditions: {current_highlight_conditions}")
    except Exception as e:
        print(f"  kshift: 判斷當前高亮條件時發生錯誤: {e}")
        traceback.print_exc()
        return unavailable_kshift_result(OOB_ERROR)


    # --- 計算滾動結果高亮條件 (如果存在滾動數據) ---
    rolling_highlight_conditions = {key: 'NO_HIGHLIGHT' for key in ['P95_shift', 'P50_shift', 'P05_shift']}

    if rolled_data is not None:
        print(f"  kshift: 處理 rolled_data != None 分支，rolled_data shape: {rolled_data['values'].shape}")
        print(f"  kshift: 滾動後 base_percentiles: {base_percentiles}")
        try:
            # 計算滾動結果 (K 值) - 使用安全除法
            rolling_results = {
                'P95_k': safe_division(np.round(abs(rolled_data['percentiles'].get('P95', np.nan) - base_percentiles.get('P95', np.nan)), 8), p95k_deno),
                'P50_k': safe_division(np.round(abs(rolled_data['percentiles'].get('P50', np.nan) - base_percentiles.get('P50', np.nan)), 8), p50k_deno),
                'P05_k': safe_division(np.round(abs(rolled_data['percentiles'].get('P05', np.nan) - base_percentiles.get('P05', np.nan)), 8), p05k_deno),
                'P95_k_ori': safe_division(np.round((rolled_data['percentiles'].get('P95', np.nan) - base_percentiles.get('P95', np.nan)), 8), p95k_deno),
                'P50_k_ori': safe_division(np.round((rolled_data['percentiles'].get('P50', np.nan) - base_percentiles.get('P50', np.nan)), 8), p50k_deno),
                'P05_k_ori': safe_division(np.round((rolled_data['percentiles'].get('P05', np.nan) - base_percentiles.get('P05', np.nan)), 8), p05k_deno),
            }
            print(f"  kshift: 計算出的 rolling_results: {rolling_results}")

            # 判斷滾動結果高亮條件
            # 確保傳給 review_kshift_results 的 percentiles 字典是完整的
            rolling_highlight_conditions = review_kshift_results(rolling_results, resolution, characteristic, rolled_data['percentiles'], base_percentiles)
            print(f"  kshift: rolling_highlight_conditions: {rolling_highlight_conditions}")

        except Exception as e:
            print(f"  kshift: 判斷滾動高亮條件時發生錯誤: {e}")
            traceback.print_exc()
            return unavailable_kshift_result(OOB_ERROR)


    # --- 最終的高亮條件 ---
    # 結合當前和滾動的高亮結果
    # 檢查 current_highlight_conditions 和 rolling_highlight_conditions 中的值是否是預期的 'HIGHLIGHT'/'NO_HIGHLIGHT'/'ERROR'
    results['P95_shift'] = 'HIGHLIGHT' if current_highlight_conditions.get('P95_shift') == 'HIGHLIGHT' and (rolled_data is None or rolling_highlight_conditions.get('P95_shift') == 'HIGHLIGHT') else 'NO_HIGHLIGHT'
    results['P50_shift'] = 'HIGHLIGHT' if current_highlight_conditions.get('P50_shift') == 'HIGHLIGHT' and (rolled_data is None or rolling_highlight_conditions.get('P50_shift') == 'HIGHLIGHT') else 'NO_HIGHLIGHT'
    results['P05_shift'] = 'HIGHLIGHT' if current_highlight_conditions.get('P05_shift') == 'HIGHLIGHT' and (rolled_data is None or rolling_highlight_conditions.get('P05_shift') == 'HIGHLIGHT') else 'NO_HIGHLIGHT'

    print(f"  kshift: 最終 shift 結果: P95={results['P95_shift']}, P50={results['P50_shift']}, P05={results['P05_shift']}")
    print("--- 退出 kshift_sigma_ratio_calculator 函數 ---")

    return pd.Series(results)

# 數據類型判斷
def determine_data_type(data_values):
    """
    判斷數據是離散型還是連續型
    
    判斷標準：
    1. (unique數值種類/總樣本數N < 1/3 且 unique數值種類 < 5) OR
    2. (總樣本數N >= 30 且 unique數值種類 <= 10)
    滿足以上任一條件即認定為離散型
    
    Parameters:
    - data_values: 數據值的 numpy array 或 pandas Series
    
    Returns:
    - 'discrete' 或 'continuous'
    """
    import numpy as np
    
    # 移除 NaN 值
    clean_values = data_values.dropna() if hasattr(data_values, 'dropna') else data_values[~np.isnan(data_values)]
    
    if len(clean_values) == 0:
        return 'continuous'  # 預設為連續型
    
    unique_values = np.unique(clean_values)
    unique_count = len(unique_values)
    total_count = len(clean_values)
    unique_ratio = unique_count / total_count
    
    print(f"  數據類型判斷: 唯一值數量={unique_count}, 總數量={total_count}, 比例={unique_ratio:.3f}")
    
    # 判斷邏輯：
    # 條件1: unique數值種類/總樣本數N < 1/3 且 unique數值種類 < 5
    condition1 = (unique_ratio <= 1/3) and (unique_count <= 5)
    
    # 條件2: 總樣本數N >= 30 且 unique數值種類 <= 10
    condition2 = (total_count >= 30) and (unique_count <= 10)
    
    if condition1 or condition2:
        print(f"    判定為離散型 - 條件1滿足: {condition1}, 條件2滿足: {condition2}")
        return 'discrete'
    else:
        print(f"    判定為連續型 - 條件1滿足: {condition1}, 條件2滿足: {condition2}")
        return 'continuous'

# OOC計算
def ooc_calculator(data, ucl, lcl, characteristic='Nominal'):
    values = pd.to_numeric(data['point_val'], errors='coerce').dropna()
    data_cnt = len(values)
    characteristic = normalize_characteristic(characteristic)

    if characteristic in ('Smaller', 'Sigma'):
        if not is_valid_number(ucl):
            raise ValueError('Smaller/Sigma OOC requires a numeric UCL')
        ooc_mask = values > float(ucl)
    elif characteristic == 'Bigger':
        if not is_valid_number(lcl):
            raise ValueError('Bigger OOC requires a numeric LCL')
        ooc_mask = values < float(lcl)
    else:
        if not is_valid_number(ucl) or not is_valid_number(lcl):
            raise ValueError('Nominal OOC requires numeric UCL and LCL')
        ooc_mask = (values > float(ucl)) | (values < float(lcl))

    ooc_cnt = int(ooc_mask.sum())
    ooc_ratio = ooc_cnt / data_cnt if data_cnt != 0 else 0
    return data_cnt, ooc_cnt, ooc_ratio

# OOC結果檢查
def review_ooc_results(ooc_cnt, ooc_ratio, threshold=0.05):
    return 'HIGHLIGHT' if ooc_ratio > threshold and ooc_cnt > 1 else 'NO_HIGHLIGHT'

# 計算Sticking Rate
def sticking_rate_calculator(baseline_data, weekly_data):
    def get_mode(data):
        return data.mode()[0]

    def get_percentage(data, value):
        return (data == value).sum() / len(data)

    # 如果週資料少於10筆，與基線資料進行合併
    if len(weekly_data) < 10:
        rolling_window_size = 20 if len(baseline_data) > 1000 else 10
        weekly_data = pd.concat([baseline_data.tail(rolling_window_size), weekly_data])

    threshold = 0.7
    baseline_mode = get_mode(baseline_data)
    weekly_mode = get_mode(weekly_data)

    baseline_mode_percentage_in_baseline = get_percentage(baseline_data, baseline_mode)
    baseline_mode_percentage_in_weekly = get_percentage(weekly_data, baseline_mode)
    weekly_mode_percentage_in_baseline = get_percentage(baseline_data, weekly_mode)
    weekly_mode_percentage_in_weekly = get_percentage(weekly_data, weekly_mode)

    baseline_mode_diff = abs(baseline_mode_percentage_in_baseline - baseline_mode_percentage_in_weekly)
    weekly_mode_diff = abs(weekly_mode_percentage_in_baseline - weekly_mode_percentage_in_weekly)

    highlight_needed = (baseline_mode_diff >= threshold) or (weekly_mode_diff >= threshold)
    highlight_status = 'HIGHLIGHT' if highlight_needed else 'NO_HIGHLIGHT'

    return {
        'baseline_mode': baseline_mode,
        'weekly_mode': weekly_mode,
        'baseline_mode_percentage_in_baseline': baseline_mode_percentage_in_baseline,
        'baseline_mode_percentage_in_weekly': baseline_mode_percentage_in_weekly,   
        'weekly_mode_percentage_in_baseline': weekly_mode_percentage_in_baseline,
        'weekly_mode_percentage_in_weekly': weekly_mode_percentage_in_weekly,
        'highlight_status': highlight_status
    }

# 趨勢檢查
def trending(raw_df, weekly_start_date, weekly_end_date, baseline_start_date, baseline_end_date):
    # 時間欄位轉換
    raw_df['point_time'] = pd.to_datetime(raw_df['point_time'])
    weekly_end_date = pd.to_datetime(weekly_end_date)
    baseline_start_date = pd.to_datetime(baseline_start_date)
    baseline_end_date = pd.to_datetime(baseline_end_date)

    # 每週資料的摘要
    weekly_summary = []
    current_end = weekly_end_date
    week_count = 0

    while week_count < 7:
        current_start = current_end - timedelta(days=6)
        week_data = raw_df[
            (raw_df['point_time'] >= current_start) &
            (raw_df['point_time'] <= current_end)
        ]['point_val']

        weekly_summary.append({
            'week_start': current_start,
            'week_end': current_end,
            'median': week_data.median() if not week_data.empty else np.nan,
            'count': len(week_data)
        })

        current_end = current_start - timedelta(days=1)
        week_count += 1

    weekly_data = pd.DataFrame(weekly_summary)

    if weekly_data.empty:
        return 'NO_HIGHLIGHT'

    weekly_medians = weekly_data['median'].tolist()
    weekly_counts = weekly_data['count'].tolist()

    # 檢查最近幾週的資料點數條件
    def check_weeks_condition(weeks_counts):
        if len(weeks_counts) >= 4 and sum(x >= 10 for x in weeks_counts[:4]) >= 3 and weeks_counts[0] >= 10:
            return 4
        elif len(weeks_counts) >= 5 and sum(x >= 6 for x in weeks_counts[:5]) >= 4 and weeks_counts[0] >= 6:
            return 5
        elif len(weeks_counts) >= 6 and sum(x >= 3 for x in weeks_counts[:6]) >= 5 and weeks_counts[0] >= 3:
            return 6
        elif len(weeks_counts) >= 7 and sum(x >= 1 for x in weeks_counts[:7]) >= 6 and weeks_counts[0] >= 1:
            return 7
        return 0

    num_weeks_to_check = check_weeks_condition(weekly_counts)

    if num_weeks_to_check == 0:
        return 'NO_HIGHLIGHT'

    # 趨勢檢查函式
    def is_trending_up(medians):
        return all(earlier > later for earlier, later in zip(medians, medians[1:]))

    def is_trending_down(medians):
        return all(earlier < later for earlier, later in zip(medians, medians[1:]))

    # 基準區間百分位
    baseline_df = raw_df[
        (raw_df['point_time'] >= baseline_start_date) &
        (raw_df['point_time'] <= baseline_end_date)
    ]
    baseline_values = baseline_df['point_val']

    if baseline_values.empty:
        return 'NO_HIGHLIGHT'

    p95 = np.percentile(baseline_values, 95)
    p05 = np.percentile(baseline_values, 5)

    # 檢查是否上升或下降
    check_medians = [m for m in weekly_medians[:num_weeks_to_check] if not np.isnan(m)]

    if len(check_medians) < 2:
        return 'NO_HIGHLIGHT'  # 資料不夠比趨勢

    if is_trending_up(check_medians) and check_medians[0] > p95:
        return 'HIGHLIGHT'
    elif is_trending_down(check_medians) and check_medians[0] < p05:
        return 'HIGHLIGHT'
    return 'NO_HIGHLIGHT'

# 離散型 OOB 處理函數
def discrete_oob_calculator(base_data, weekly_data, chart_info, raw_df=None, 
                            weekly_start_date=None, weekly_end_date=None, 
                            baseline_start_date=None, baseline_end_date=None):
    """
    離散型數據的 OOB 計算方法
    包含修改後的 k-shift、新增的 category_LT_Shift 和 trending
    
    Parameters:
    - base_data: 基線數據字典 (包含 'values', 'cnt', 'mean', 'sigma')
    - weekly_data: 週數據字典 (包含 'values', 'cnt', 'mean', 'sigma')
    - chart_info: 圖表信息
    - raw_df: 原始數據 DataFrame (用於 trending 計算)
    - weekly_start_date: 週開始日期
    - weekly_end_date: 週結束日期
    - baseline_start_date: 基線開始日期
    - baseline_end_date: 基線結束日期

    
    Returns:
    - dict: 包含 OOB 結果的字典
    """
    import numpy as np
    
    print(f"  離散型 OOB 計算: 基線數據點數={base_data['cnt']}, 週數據點數={weekly_data['cnt']}")
    
    results = {
        'HL_P95_shift': 'NO_HIGHLIGHT',
        'HL_P50_shift': 'NO_HIGHLIGHT', 
        'HL_P05_shift': 'NO_HIGHLIGHT',
        'HL_sticking_shift': 'NO_HIGHLIGHT',
        'HL_trending': 'NO_HIGHLIGHT',
        'HL_high_OOC': 'NO_HIGHLIGHT',
        'HL_category_LT_shift': 'NO_HIGHLIGHT',  # 新增的離散型專用項目
        'discrete_method': True
    }
    
    try:
        # 1. 使用與連續型相同的 sticking_rate_calculator
        print("  離散型 OOB: 計算 sticking rate...")
        sticking_rate_results = sticking_rate_calculator(
            pd.Series(base_data['values']), 
            pd.Series(weekly_data['values'])
        )
        results['HL_sticking_shift'] = sticking_rate_results.get('highlight_status', 'NO_HIGHLIGHT')
        
        # 2. Trending 計算
        print("  離散型 OOB: 計算 trending...")
        if (raw_df is not None and weekly_start_date is not None and 
            weekly_end_date is not None and baseline_start_date is not None and 
            baseline_end_date is not None):
            trending_result = discrete_trending_calculator(
                raw_df, weekly_start_date, weekly_end_date, 
                baseline_start_date, baseline_end_date
            )
            results['HL_trending'] = trending_result
        else:
            results['HL_trending'] = OOB_INSUFFICIENT_DATA
        
        # 3. 使用與連續型相同的 high_OOC 檢查
        print("  離散型 OOB: 計算 OOC...")
        weekly_df = pd.DataFrame({'point_val': weekly_data['values']})
        ooc_results = ooc_calculator(
            weekly_df,
            chart_info.get('UCL'),
            chart_info.get('LCL'),
            chart_info.get('Characteristics'),
        )
        ooc_highlight = review_ooc_results(ooc_results[1], ooc_results[2])
        results['HL_high_OOC'] = ooc_highlight
        
        # 4. 修改後的 k-shift 計算（加入 capping rule）
        print("  離散型 OOB: 計算修改後的 K-shift...")
        kshift_results = discrete_kshift_calculator(
            base_data, weekly_data, 
            chart_info.get('Characteristics'), 
            chart_info.get('Resolution'), 
            chart_info.get('UCL'), 
            chart_info.get('LCL')
        )
        results['HL_P95_shift'] = kshift_results.get('P95_shift', 'NO_HIGHLIGHT')
        results['HL_P50_shift'] = kshift_results.get('P50_shift', 'NO_HIGHLIGHT')
        results['HL_P05_shift'] = kshift_results.get('P05_shift', 'NO_HIGHLIGHT')
        
        # 5. 新增的 category_LT_Shift 計算
        print("  離散型 OOB: 計算 category_LT_Shift...")
        category_lt_results = category_lt_shift_calculator(base_data, weekly_data)
        results['HL_category_LT_shift'] = category_lt_results.get('highlight_status', 'NO_HIGHLIGHT')
        
        print(f"  離散型 OOB 計算完成: {results}")
        
    except Exception as e:
        print(f"  離散型 OOB 計算錯誤: {e}")
        import traceback
        traceback.print_exc()
        for key in (
            'HL_P95_shift',
            'HL_P50_shift',
            'HL_P05_shift',
            'HL_sticking_shift',
            'HL_trending',
            'HL_category_LT_shift',
        ):
            results[key] = OOB_ERROR
    
    return results

def discrete_trending_calculator(raw_df, weekly_start_date, weekly_end_date, baseline_start_date, baseline_end_date):
    """
    離散型數據的 trending 計算（移植自原 trending 函數）
    """
    from datetime import timedelta
    import numpy as np
    import pandas as pd
    
    # 時間欄位轉換
    raw_df['point_time'] = pd.to_datetime(raw_df['point_time'])
    weekly_end_date = pd.to_datetime(weekly_end_date)
    baseline_start_date = pd.to_datetime(baseline_start_date)
    baseline_end_date = pd.to_datetime(baseline_end_date)

    # 每週資料的摘要
    weekly_summary = []
    current_end = weekly_end_date
    week_count = 0

    while week_count < 7:
        current_start = current_end - timedelta(days=6)
        week_data = raw_df[
            (raw_df['point_time'] >= current_start) &
            (raw_df['point_time'] <= current_end)
        ]['point_val']

        weekly_summary.append({
            'week_start': current_start,
            'week_end': current_end,
            'median': week_data.median() if not week_data.empty else np.nan,
            'count': len(week_data)
        })

        current_end = current_start - timedelta(days=1)
        week_count += 1

    weekly_data = pd.DataFrame(weekly_summary)

    if weekly_data.empty:
        return 'NO_HIGHLIGHT'

    weekly_medians = weekly_data['median'].tolist()
    weekly_counts = weekly_data['count'].tolist()

    # 檢查最近幾週的資料點數條件
    def check_weeks_condition(weeks_counts):
        if len(weeks_counts) >= 4 and sum(x >= 10 for x in weeks_counts[:4]) >= 3 and weeks_counts[0] >= 10:
            return 4
        elif len(weeks_counts) >= 5 and sum(x >= 6 for x in weeks_counts[:5]) >= 4 and weeks_counts[0] >= 6:
            return 5
        elif len(weeks_counts) >= 6 and sum(x >= 3 for x in weeks_counts[:6]) >= 5 and weeks_counts[0] >= 3:
            return 6
        elif len(weeks_counts) >= 7 and sum(x >= 1 for x in weeks_counts[:7]) >= 6 and weeks_counts[0] >= 1:
            return 7
        return 0

    num_weeks_to_check = check_weeks_condition(weekly_counts)

    if num_weeks_to_check == 0:
        return 'NO_HIGHLIGHT'

    # 趨勢檢查函式
    def is_trending_up(medians):
        return all(earlier > later for earlier, later in zip(medians, medians[1:]))

    def is_trending_down(medians):
        return all(earlier < later for earlier, later in zip(medians, medians[1:]))

    # 基準區間百分位
    baseline_df = raw_df[
        (raw_df['point_time'] >= baseline_start_date) &
        (raw_df['point_time'] <= baseline_end_date)
    ]
    baseline_values = baseline_df['point_val']

    if baseline_values.empty:
        return 'NO_HIGHLIGHT'

    p95 = np.percentile(baseline_values, 95)
    p05 = np.percentile(baseline_values, 5)

    # 檢查是否上升或下降
    check_medians = [m for m in weekly_medians[:num_weeks_to_check] if not np.isnan(m)]

    if len(check_medians) < 2:
        return 'NO_HIGHLIGHT'  # 資料不夠比趨勢

    if is_trending_up(check_medians) and check_medians[0] > p95:
        return 'HIGHLIGHT'
    elif is_trending_down(check_medians) and check_medians[0] < p05:
        return 'HIGHLIGHT'
    return 'NO_HIGHLIGHT'

# 修改後的 K-shift 函數（加入 capping rule）
def discrete_kshift_calculator(base_data, weekly_data, characteristic, resolution, ucl, lcl):
    """
    離散型數據的 K-shift 計算，加入 capping rule
    
    Capping rule: 如果當周點數<=10 且 當周P95/P50/P05沒有超過 baseline的P05和P95範圍外，就不HL
    """
    import numpy as np
    
    print("  discrete_kshift: 開始計算離散型 K-shift")
    
    # 先使用原本的 kshift_sigma_ratio_calculator 獲取結果
    kshift_results = kshift_sigma_ratio_calculator(
        base_data, weekly_data, characteristic, resolution, ucl, lcl
    )
    
    weekly_cnt = weekly_data['cnt']
    print(f"  discrete_kshift: 當周點數 = {weekly_cnt}")
    
    # 應用 capping rule
    if weekly_cnt <= 10:
        print("  discrete_kshift: 當周點數 <= 10，檢查 capping rule")
        
        try:
            # 計算當周和基線的百分位數
            weekly_percentiles = get_percentiles(weekly_data['values'])
            base_percentiles = get_percentiles(base_data['values'])
            
            weekly_p95 = weekly_percentiles.get('P95')
            weekly_p50 = weekly_percentiles.get('P50') 
            weekly_p05 = weekly_percentiles.get('P05')
            base_p95 = base_percentiles.get('P95')
            base_p05 = base_percentiles.get('P05')
            
            print(f"  discrete_kshift: 當周百分位數 - P95:{weekly_p95}, P50:{weekly_p50}, P05:{weekly_p05}")
            print(f"  discrete_kshift: 基線範圍 - P05:{base_p05}, P95:{base_p95}")
            
            # 檢查當周百分位數是否都在基線 P05-P95 範圍內
            if (not pd.isna(weekly_p95) and not pd.isna(base_p05) and not pd.isna(base_p95) and
                not pd.isna(weekly_p50) and not pd.isna(weekly_p05)):
                
                p95_in_range = base_p05 <= weekly_p95 <= base_p95
                p50_in_range = base_p05 <= weekly_p50 <= base_p95  
                p05_in_range = base_p05 <= weekly_p05 <= base_p95
                
                print(f"  discrete_kshift: 範圍檢查 - P95 in range:{p95_in_range}, P50 in range:{p50_in_range}, P05 in range:{p05_in_range}")
                
                if p95_in_range and p50_in_range and p05_in_range:
                    print("  discrete_kshift: Capping rule 觸發 - 所有百分位數都在基線範圍內，設為 NO_HIGHLIGHT")
                    kshift_results['P95_shift'] = 'NO_HIGHLIGHT'
                    kshift_results['P50_shift'] = 'NO_HIGHLIGHT' 
                    kshift_results['P05_shift'] = 'NO_HIGHLIGHT'
                else:
                    print("  discrete_kshift: 有百分位數超出基線範圍，維持原始 K-shift 結果")
            else:
                print("  discrete_kshift: 百分位數計算有 NaN 值，維持原始 K-shift 結果")
                
        except Exception as e:
            print(f"  discrete_kshift: Capping rule 檢查時發生錯誤: {e}")
            # 發生錯誤時維持原始結果
    else:
        print("  discrete_kshift: 當周點數 > 10，不適用 capping rule")
    
    print(f"  discrete_kshift: 最終結果 - P95:{kshift_results.get('P95_shift')}, P50:{kshift_results.get('P50_shift')}, P05:{kshift_results.get('P05_shift')}")
    
    return kshift_results

# 新的 category_LT_Shift 函數
def category_lt_shift_calculator(base_data, weekly_data, threshold=0.7):
    """
    計算 category_LT_Shift
    
    邏輯：
    1. 當周<20則rolling to 20筆
    2. 拿當周data範圍去對應baseline同樣data範圍
    3. 檢查data所佔比例是否超過70%
    
    Parameters:
    - base_data: 基線數據字典
    - weekly_data: 週數據字典  
    - threshold: 佔比差異閾值，預設0.7 (70%)
    
    Returns:
    - dict: 包含highlight_status的結果字典
    """
    import numpy as np
    import pandas as pd
    
    print("  category_LT_shift: 開始計算")
    
    result = {
        'highlight_status': 'NO_HIGHLIGHT',
        'weekly_range': None,
        'baseline_ratio_in_range': None,
        'weekly_ratio_in_range': None, 
        'ratio_diff': None
    }
    
    try:
        weekly_values = weekly_data['values'].copy()
        base_values = base_data['values'].copy()
        
        print(f"  category_LT_shift: 原始當周點數 = {len(weekly_values)}")
        
        # 1. 如果當周 < 20 則 rolling to 20筆
        if len(weekly_values) < 20:
            print(f"  category_LT_shift: 當周點數 < 20，rolling 到 20 筆")
            
            # 需要從基線數據中補充
            needed_points = 20 - len(weekly_values)
            if len(base_values) >= needed_points:
                # 取基線最後的點來補充
                additional_points = base_values[-needed_points:]
                weekly_values = np.concatenate([additional_points, weekly_values])
                print(f"  category_LT_shift: 補充後當周點數 = {len(weekly_values)}")
            else:
                print(f"  category_LT_shift: 基線數據不足以補充到20筆，使用現有數據")
                weekly_values = np.concatenate([base_values, weekly_values])
        
        # 2. 計算當周數據範圍
        weekly_min = np.min(weekly_values)
        weekly_max = np.max(weekly_values)
        result['weekly_range'] = (weekly_min, weekly_max)
        
        print(f"  category_LT_shift: 當周數據範圍 = [{weekly_min:.3f}, {weekly_max:.3f}]")
        
        # 3. 計算基線數據在此範圍內的比例
        baseline_in_range = base_values[(base_values >= weekly_min) & (base_values <= weekly_max)]
        baseline_ratio = len(baseline_in_range) / len(base_values) if len(base_values) > 0 else 0
        result['baseline_ratio_in_range'] = baseline_ratio
        
        # 4. 計算當周數據在此範圍內的比例（應該是100%，因為就是用當周數據定義的範圍）
        weekly_ratio = 1
        result['weekly_ratio_in_range'] = weekly_ratio
        
        # 5. 計算比例差異
        ratio_diff = abs(weekly_ratio - baseline_ratio)
        result['ratio_diff'] = ratio_diff
        
        print(f"  category_LT_shift: 基線在範圍內比例 = {baseline_ratio:.3f}")
        print(f"  category_LT_shift: 當周在範圍內比例 = {weekly_ratio:.3f}")
        print(f"  category_LT_shift: 比例差異 = {ratio_diff:.3f}")
        
        # 6. 判斷是否需要高亮
        if ratio_diff > threshold:
            result['highlight_status'] = 'HIGHLIGHT'
            print(f"  category_LT_shift: 比例差異 {ratio_diff:.3f} > {threshold}，需要 HIGHLIGHT")
        else:
            result['highlight_status'] = 'NO_HIGHLIGHT' 
            print(f"  category_LT_shift: 比例差異 {ratio_diff:.3f} <= {threshold}，NO_HIGHLIGHT")
            
    except Exception as e:
        print(f"  category_LT_shift: 計算時發生錯誤: {e}")
        import traceback
        traceback.print_exc()
        result['highlight_status'] = OOB_ERROR
    
    return result


def process_single_chart(
    chart_info,
    raw_df,
    initial_baseline_start_date,
    baseline_end_date,
    weekly_start_date,
    weekly_end_date,
    enable_record_high_low=False,
):
    print("--- 進入外部 process_single_chart 函數 ---")
    print(f"  接收到的 raw_df shape: {raw_df.shape}")
    print(f"  週數據範圍: {weekly_start_date} 至 {weekly_end_date}")
    # 注意：這裡接收的是 initial_baseline_start_date (通常是往前一年)
    print(f"  初始基線數據範圍 (往前一年): {initial_baseline_start_date} 至 {baseline_end_date}")

    if raw_df is None or raw_df.empty:
        print("  raw_df 是空的或 None, 返回 None")
        return None

    try:
        print("  正在篩選週數據...")
        weekly_data = raw_df[(raw_df['point_time'] >= weekly_start_date) & (raw_df['point_time'] <= weekly_end_date)].copy() # Use copy()
        print(f"  篩選後 weekly_data shape: {weekly_data.shape}")

        if weekly_data.empty:
             print(f'未找到週數據, GroupName: {chart_info.get("group_name", "N/A")}, ChartName: {chart_info.get("chart_name", "N/A")}, 返回 None')
             return None

        # --- 基線數據範圍選擇邏輯開始 ---

        # 步驟 1: 使用初始的一年基線範圍過濾數據並計數
        print("  正在篩選初始一年基線數據...")
        baseline_data_one_year = raw_df[(raw_df['point_time'] >= initial_baseline_start_date) & (raw_df['point_time'] <= baseline_end_date)].copy() # Use copy()
        baseline_count_one_year = len(baseline_data_one_year)
        print(f"  初始一年基線數據點數量: {baseline_count_one_year}")

        # 步驟 2: 根據計數決定最終使用的基線開始日期
        # === 新增：基線數據不足標記 ===
        baseline_insufficient = False
        
        if baseline_count_one_year < 10:
            # 如果少於 10 點，將基線期擴展到兩年
            actual_baseline_start_date = baseline_end_date - pd.Timedelta(days=365 * 2)
            print(f"  基線數據點數量 ({baseline_count_one_year}) < 10，將基線期擴展至兩年: {actual_baseline_start_date} 至 {baseline_end_date}")
                
            # 檢查擴展後的數量
            baseline_data_two_year = raw_df[(raw_df['point_time'] >= actual_baseline_start_date) & (raw_df['point_time'] <= baseline_end_date)].copy()
            baseline_count_two_year = len(baseline_data_two_year)
            print(f"  擴展至兩年後基線數據點數量: {baseline_count_two_year}")
            
            if baseline_count_two_year < 10:
                print(f"  ⚠️  擴展至兩年後仍少於10點 ({baseline_count_two_year})，將跳過 OOB 分析但繼續處理其他功能")
                baseline_insufficient = True
            print(f"  基線數據點數量 ({baseline_count_one_year}) < 10，將基線期擴展至兩年: {actual_baseline_start_date} 至 {baseline_end_date}")
        else:
            # 如果大於等於 10 點，使用一年的基線期
            actual_baseline_start_date = initial_baseline_start_date
            print(f"  基線數據點數量 ({baseline_count_one_year}) >= 10，使用一年基線期: {actual_baseline_start_date} 至 {baseline_end_date}")

        # 步驟 3: 使用最終確定的基線範圍過濾數據
        print("  正在篩選最終基線數據...")
        baseline_data = raw_df[(raw_df['point_time'] >= actual_baseline_start_date) & (raw_df['point_time'] <= baseline_end_date)].copy() # Use copy()
        print(f"  篩選後 baseline_data shape (使用 {len(baseline_data)} 點從 {actual_baseline_start_date} 至 {baseline_end_date}): {baseline_data.shape}")


        # === 修改：基線為空時仍然繼續處理 ===
        baseline_empty = baseline_data.empty
        if baseline_empty:
             print(f'基線數據為空，但仍繼續處理 WE Rule 和圖表生成, GroupName: {chart_info.get("group_name", "N/A")}, ChartName: {chart_info.get("chart_name", "N/A")}')
             baseline_insufficient = True

        # --- 基線數據範圍選擇邏輯結束 ---


        # 計算統計數據（週數據與基線數據）
        def calculate_statistics(data):
             # 新增檢查，避免對只有一個點的數據計算標準差產生 NaN (ddof=1 時)
             if data.shape[0] <= 1:
                  sigma = 0.0 if data.shape[0] == 1 else 0.0 # 單點或零點標準差視為 0
             else:
                  sigma = data['point_val'].std() # ddof=1 是 pandas 預設，計算樣本標準差

             # 如果 sigma 是 NaN (例如，所有值都相同，但數據點多於 1 且少於某個閾值，或計算出問題)
             if np.isnan(sigma):
                 print(f"  calculate_statistics 警告: 計算 sigma 得到 NaN. Data shape: {data.shape}")
                 sigma = 0.0 # 將無效的標準差視為 0

             return {
                 'values': data['point_val'].values,
                 'cnt': data.shape[0],
                 'mean': data['point_val'].mean(),
                 'sigma': sigma # 使用處理過的 sigma
                 }

        print("  正在計算週數據統計...")
        weekly_data_dict = calculate_statistics(weekly_data)
        print(f"  週數據統計結果 (部分): cnt={weekly_data_dict['cnt']}, mean={weekly_data_dict['mean']}, sigma={weekly_data_dict['sigma']}")


        # IMPORTANT: 這裡的 baseline_data_dict 現在是使用 *實際確定* 的基線範圍數據計算的
        print("  正在計算基線數據統計...")
        baseline_data_dict = calculate_statistics(baseline_data) if not baseline_empty else None
        if baseline_data_dict is not None:
            print(f"  基線數據統計結果 (部分): cnt={baseline_data_dict['cnt']}, mean={baseline_data_dict['mean']}, sigma={baseline_data_dict['sigma']}")
        else:
            print("  基線數據為空，跳過基線統計輸出")

        # 確保基線統計數據的標準差不會導致後續計算問題
        if baseline_data_dict and 'sigma' in baseline_data_dict and (baseline_data_dict['sigma'] == 0 or np.isnan(baseline_data_dict['sigma'])):
             print("  警告: 基線標準差為零或無效，可能影響 K 值計算和需要標準差的其他指標。")
             # 您可以選擇在這裡返回 None，或讓後續函數自行處理 NaN/inf

        print("  正在呼叫 kshift_sigma_ratio_calculator...")
        # 傳入使用實際基線範圍計算出的 baseline_data_dict
        # kshift_sigma_ratio_calculator 需要處理 sigma=0 或其他分母為 0 的情況 (已在 safe_division 中處理)
        if not baseline_insufficient and not baseline_empty:
            kshift_results = kshift_sigma_ratio_calculator(baseline_data_dict, weekly_data_dict, chart_info.get('Characteristics'), chart_info.get('Resolution'), chart_info.get('UCL'), chart_info.get('LCL')) # 使用 .get 防止 key 錯誤
        else:
            kshift_results = unavailable_kshift_result(OOB_INSUFFICIENT_DATA)

        print(f"  kshift_sigma_ratio_calculator 返回: {kshift_results}")

        print("  正在呼叫 ooc_calculator...")
        # ooc_calculator 使用週數據計算 OOC 點數
        try:
            ooc_results = ooc_calculator(
                weekly_data,
                chart_info.get('UCL'),
                chart_info.get('LCL'),
                chart_info.get('Characteristics'),
            )
            ooc_highlight = review_ooc_results(ooc_results[1], ooc_results[2])
        except Exception as e:
            logger.exception("OOC calculation failed: %s", e)
            ooc_results = (len(weekly_data), 0, 0)
            ooc_highlight = OOB_ERROR
        print(f"  ooc_calculator 返回: {ooc_results}")

        print("  正在呼叫 review_ooc_results...")
        print(f"  review_ooc_results 返回: {ooc_highlight}")

        print("  正在呼叫 sticking_rate_calculator...")
        # sticking_rate_calculator 需要週數據和基線數據的 Series
        # IMPORTANT: 這裡傳入的 baseline_data['point_val'] 是使用 *實際確定* 的基線範圍數據
        sticking_rate_results = sticking_rate_calculator(baseline_data['point_val'], weekly_data['point_val']) if not baseline_insufficient and not baseline_empty else {'highlight_status': OOB_INSUFFICIENT_DATA}
        print(f"  sticking_rate_calculator 返回: {sticking_rate_results}")

        print("  正在呼叫 trending...")
        # trending 也需要使用實際確定後的基線範圍
        trending_results = trending(raw_df, weekly_start_date, weekly_end_date, actual_baseline_start_date, baseline_end_date) if not baseline_insufficient and not baseline_empty else OOB_INSUFFICIENT_DATA
        print(f"  trending 返回: {trending_results}")

        if enable_record_high_low:
            print("  正在呼叫 record_high_low_calculator...")
            logger.debug("基線時間範圍 - 從 %s 到 %s", actual_baseline_start_date, baseline_end_date)
            logger.debug("當週時間範圍 - 從 %s 到 %s", weekly_start_date, weekly_end_date)
            logger.debug("基線結束與當週開始間隔 = %s", weekly_start_date - baseline_end_date)
        else:
            print("  Record High/Low 判定已關閉。")
        record_results = optional_record_high_low_calculator(
            weekly_data['point_val'].values,
            baseline_data['point_val'].values,
            enabled=enable_record_high_low,
            data_available=not baseline_insufficient and not baseline_empty,
        )
        print(f"  record_high_low_calculator 返回: {record_results}")

        # 判斷是否需要 highlight (任何一個子指標需要高亮，則總體高亮)
        highlight_status = 'HIGHLIGHT' if (
             kshift_results.get('P95_shift') == 'HIGHLIGHT' or
             kshift_results.get('P50_shift') == 'HIGHLIGHT' or
             kshift_results.get('P05_shift') == 'HIGHLIGHT' or
             sticking_rate_results.get('highlight_status') == 'HIGHLIGHT' or
             trending_results == 'HIGHLIGHT' or
             ooc_highlight == 'HIGHLIGHT' or # 應該也要考慮 ooc_highlight
             record_results.get('highlight_status') == 'HIGHLIGHT' # 新增 record high/low 判斷
        ) else 'NO_HIGHLIGHT'
        print(f"  計算出的 highlight_status: {highlight_status}")


        oob_values = [
            kshift_results.get('P95_shift'),
            kshift_results.get('P50_shift'),
            kshift_results.get('P05_shift'),
            sticking_rate_results.get('highlight_status'),
            trending_results,
            ooc_highlight,
            record_results.get('highlight_status'),
        ]
        if OOB_ERROR in oob_values:
            oob_status = OOB_ERROR
        elif OOB_INSUFFICIENT_DATA in oob_values:
            oob_status = OOB_INSUFFICIENT_DATA
        else:
            oob_status = OOB_CALCULATED

        # 組織結果
        # 注意使用 .get(key, default_value) 來安全存取字典鍵，防止 KeyError
        result = {
            'data_cnt': ooc_results[0], # 週數據點數
            'ooc_cnt': ooc_results[1], # 週數據 OOC 點數
            'WE_Rule': '', # 這個欄位在 GUI 類的 build_result 中填充
            'OOB_Rule': '' if not baseline_empty else 'N/A - No Baseline', # 基線為空時標記
            'HL_P95_shift': kshift_results.get('P95_shift', 'N/A'), # 使用 get 並提供預設值，避免 key 錯誤
            'HL_P50_shift': kshift_results.get('P50_shift', 'N/A'),
            'HL_P05_shift': kshift_results.get('P05_shift', 'N/A'),
            'HL_sticking_shift': sticking_rate_results.get('highlight_status', 'N/A'),
            'HL_trending': trending_results, # trending_results 本身就是 HIGHLIGHT/NO_HIGHLIGHT
            'HL_high_OOC': ooc_highlight, # ooc_highlight 本身就是 HIGHLIGHT/NO_HIGHLIGHT
            'HL_record_high_low': record_results.get('highlight_status', 'N/A'), # 新增 record high/low 欄位
            'record_high': record_results.get('record_high', False), # 是否創新高
            'record_low': record_results.get('record_low', False), # 是否創新低
            'Material_no': chart_info.get('material_no', 'N/A'),
            'group_name': chart_info.get('group_name', 'N/A'),
            'chart_name': chart_info.get('chart_name', 'N/A'),
            'chart_ID': chart_info.get('ChartID', 'N/A'),
            'Characteristics': chart_info.get('Characteristics', 'N/A'),
            'USL': chart_info.get('USL', 'N/A'),
            'LSL': chart_info.get('LSL', 'N/A'),
            'UCL': chart_info.get('UCL', 'N/A'),
            'LCL': chart_info.get('LCL', 'N/A'),
            'Target': chart_info.get('Target', 'N/A'),
            'Resolution': chart_info.get('Resolution', 'N/A'),
            'OOB_Status': oob_status,
            'baseline_insufficient': baseline_insufficient,  # 新增標記，供後續使用
            'baseline_empty': baseline_empty  # 新增標記，記錄基線是否為空
            # 可以考慮添加 actual_baseline_start_date 到結果中，用於記錄實際使用的基線範圍
            # 'Actual_Baseline_Start': actual_baseline_start_date
        }
        print("--- 外部 process_single_chart 函數成功退出 ---")
        return result

    except Exception as e:
        # 在外部函數中捕獲異常並印出 traceback
        print(f'處理圖表時出錯 (外部函數) {chart_info.get("group_name", "N/A")} - {chart_info.get("chart_name", "N/A")}: {e}')
        traceback.print_exc()
        return None

def calculate_sigma(UCL, LCL, mean):
    """
    計算 sigma_upper 和 sigma_lower
    如果 UCL/LCL/mean 任一為 None 或 NaN，返回 (NaN, NaN)
    """
    import pandas as pd
    import numpy as np
    
    # 檢查參數有效性
    if pd.isna(UCL) or UCL is None or pd.isna(LCL) or LCL is None or pd.isna(mean) or mean is None:
        print("  [Warning] calculate_sigma: UCL/LCL/Target 有缺失，返回 NaN")
        return np.nan, np.nan
    
    try:
        sigma_upper = (UCL - mean) / 3
        sigma_lower = (mean - LCL) / 3
        return sigma_upper, sigma_lower
    except (TypeError, ValueError) as e:
        print(f"  [Warning] calculate_sigma: 計算失敗 - {e}")
        return np.nan, np.nan

def check_rules(raw_df, chart_info):
    import pandas as pd
    import numpy as np
    
    mean = chart_info.get('Target')
    UCL = chart_info.get('UCL')
    LCL = chart_info.get('LCL')
    characteristics = chart_info['Characteristics']
    
    # 計算 sigma（可能返回 NaN）
    sigma_upper, sigma_lower = calculate_sigma(UCL, LCL, mean)
    
    # 檢查 sigma 是否有效
    sigma_valid = not pd.isna(sigma_upper) and not pd.isna(sigma_lower) and not pd.isna(mean)
    
    if sigma_valid:
        UWL = mean + 2 * sigma_upper
        LWL = mean - 2 * sigma_lower
    else:
        UWL = np.nan
        LWL = np.nan
        print("  [Warning] check_rules: Sigma 無效，WE2-WE10 將設為 False")

    rules = {
        "WE2": False,
        "WE3": False,
        "WE4": False,
        "WE6": False,
        "WE7": False,
        "WE8": False,
        "WE9": False,
        "WE10": False,
        "CU1": False,
        "CU2": False
    }

    # WE1/WE5: 只需要 UCL/LCL，不依賴 sigma
    if not pd.isna(UCL) and UCL is not None:
        rules["WE1"] = raw_df['point_val'].iloc[-1] > UCL
    
    if not pd.isna(LCL) and LCL is not None:
        rules["WE5"] = raw_df['point_val'].iloc[-1] < LCL
    
    # CU1/CU2: 趨勢規則，不依賴 sigma
    if chart_info.get('CU1', 'N') == 'Y' and len(raw_df) >= 7:
        tail_7 = raw_df['point_val'].tail(7)
        diffs = tail_7.diff().dropna()
        rules["CU1"] = (diffs > 0).all()
    
    if chart_info.get('CU2', 'N') == 'Y' and len(raw_df) >= 7:
        tail_7 = raw_df['point_val'].tail(7)
        diffs = tail_7.diff().dropna()
        rules["CU2"] = (diffs < 0).all()

    # WE2-WE10 需要 sigma 有效才能判斷
    if not sigma_valid:
        return rules  # Sigma 無效，直接返回（WE2-WE10 保持 False，但 WE1/WE5/CU1/CU2 已判斷）
    
    if chart_info.get('WE2', 'N') == 'Y' and len(raw_df) >= 3:
        rules["WE2"] = (raw_df['point_val'].tail(3) > UWL).sum() >= 2 if characteristics not in ['Bigger', 'Smaller', 'Sigma'] else False
    if chart_info.get('WE3', 'N') == 'Y' and len(raw_df) >= 5:
        threshold = mean + sigma_upper  # 修正：使用標準的 1σ 線
        rules["WE3"] = (raw_df['point_val'].tail(5) > threshold).sum() >= 4
    if chart_info.get('WE4', 'N') == 'Y' and len(raw_df) >= 8:
        rules["WE4"] = (raw_df['point_val'].tail(8) > mean).all()
    if chart_info.get('WE6', 'N') == 'Y' and len(raw_df) >= 3:
        rules["WE6"] = (raw_df['point_val'].tail(3) < LWL).sum() >= 2 if characteristics not in ['Bigger', 'Smaller', 'Sigma'] else False
    if chart_info.get('WE7', 'N') == 'Y' and len(raw_df) >= 5:
        threshold = mean - sigma_lower  # 修正：使用標準的 1σ 線
        rules["WE7"] = (raw_df['point_val'].tail(5) < threshold).sum() >= 4
    if chart_info.get('WE8', 'N') == 'Y' and len(raw_df) >= 8:
        rules["WE8"] = (raw_df['point_val'].tail(8) < mean).all()
    if chart_info.get('WE9', 'N') == 'Y' and len(raw_df) >= 15:
        # 取得最後 15 筆資料
        tail_points = raw_df['point_val'].tail(15)
        
        # 如果所有資料點報定值（唯一值數量為 1），則直接返回 False
        if tail_points.nunique() == 1:  # 檢查唯一值數量是否為 1
            rules["WE9"] = False
        else:
            # 修正：使用 >= 和 <= 包含邊界值
            condition_result = (tail_points >= (mean - sigma_lower)) & \
                            (tail_points <= (mean + sigma_upper))
            rules["WE9"] = condition_result.all()
    if chart_info.get('WE10', 'N') == 'Y' and len(raw_df) >= 8:                   
        rules["WE10"] = ((raw_df['point_val'].tail(8) < (mean - sigma_lower) ) | 
                        (raw_df['point_val'].tail(8) > (mean + sigma_upper) )).all() if characteristics not in ['Bigger', 'Smaller', 'Sigma'] else False
    
    return rules
def calculate_cpk(raw_df, chart_info):
    mean = raw_df['point_val'].mean()
    std = raw_df['point_val'].std()
    characteristic = chart_info['Characteristics']
    usl = chart_info.get('USL', None)
    lsl = chart_info.get('LSL', None)

    cpk = None

    if std > 0:
        if characteristic == 'Nominal':
            if usl is not None and lsl is not None:
                cpu = (usl - mean) / (3 * std)
                cpl = (mean - lsl) / (3 * std)
                cpk = min(cpu, cpl)
        elif characteristic in ['Smaller', 'Sigma']:  # Sigma 使用與 Smaller 相同的邏輯
            if usl is not None:
                cpk = (usl - mean) / (3 * std)
        elif characteristic == 'Bigger':
            if lsl is not None:
                cpk = (mean - lsl) / (3 * std)

    if cpk is not None:
        cpk = round(cpk, 3)  # 統一四捨五入到小數第三位

    return {'Cpk': cpk}
def precompute_weekly_rule_results(raw_df, chart_info, weekly_start_date, weekly_end_date):
    """Compute each weekly point's WE/CU rules once for all chart renderers."""
    prepared = raw_df.copy()
    prepared['point_time'] = pd.to_datetime(prepared['point_time'])
    prepared = prepared.sort_values('point_time').reset_index(drop=True)
    ws = pd.to_datetime(weekly_start_date)
    we = pd.to_datetime(weekly_end_date)
    weekly_indices = prepared[
        (prepared['point_time'] >= ws) & (prepared['point_time'] <= we)
    ].index
    return {
        int(index): check_rules(prepared.iloc[:index + 1].tail(15).copy(), chart_info)
        for index in weekly_indices
    }


def plot_spc_chart(raw_df, chart_info, weekly_start_date, weekly_end_date, debug=False,
                   precomputed_rule_results=None):
    import os
    import numpy as np
    import matplotlib.pyplot as plt
    import pandas as pd

    plt.figure(figsize=(14, 6))

    group_name = chart_info['group_name']
    display_group_name = "" if group_name == "Default" else f"Group: [{group_name}]"
    title = (f"{display_group_name}[{chart_info['chart_name']}][{chart_info['Characteristics']}]\n"
             f"UCL: [{chart_info['UCL']}] | Target: [{chart_info['Target']}] | LCL: [{chart_info['LCL']}]")
    plt.title(title, loc='left', fontsize=12)

    # === 先排序 + reset_index，確保 index = 0..N-1 ===
    raw_df = raw_df.copy()
    raw_df['point_time'] = pd.to_datetime(raw_df['point_time'])
    raw_df = raw_df.sort_values('point_time').reset_index(drop=True)

    # === 移除點數限制，使用全部數據 ===
    print(f"  plot_spc_chart: 使用全部數據點數 {len(raw_df)}")

    points_num = len(raw_df)
    x_values = np.arange(points_num)

    # === 控制線 ===
    draw_horizontal_limit(plt, chart_info.get('UCL'), 'UCL', -0.8, points_num + 2)
    plt.hlines(chart_info['Target'], -0.8, points_num + 2, colors='#087E8B', linestyles='--', linewidth=1)
    draw_horizontal_limit(plt, chart_info.get('LCL'), 'LCL', -0.8, points_num + 2)

    plt.text(x=points_num + 2, y=chart_info['Target'], s='Target', va='center', ha='left', fontsize=10, color='#087E8B')

    # === 畫數據線 ===
    plt.plot(x_values, raw_df['point_val'], color='#5863F8', marker='o', linestyle='-')

    # === 找當週的 index ===
    ws = pd.to_datetime(weekly_start_date)
    we = pd.to_datetime(weekly_end_date)

    start_index = raw_df[raw_df['point_time'] >= ws].index.min()
    end_index   = raw_df[raw_df['point_time'] <= we].index.max()

    if debug:
        logger.debug("weekly_start_date=%s, weekly_end_date=%s", ws, we)
        logger.debug("start_index=%s, time=%s", start_index, raw_df.loc[start_index, 'point_time'])
        logger.debug("end_index=%s, time=%s", end_index, raw_df.loc[end_index, 'point_time'])

    # === 檢查 rule，標紅點 ===
    violated_rules = {rule: False for rule in chart_info.get('rule_list', [])}

    for i in range(start_index, end_index + 1):
        data_subset = raw_df.iloc[:i+1].tail(15)
        if not data_subset.empty:
            rules = (
                precomputed_rule_results.get(i, {})
                if precomputed_rule_results is not None
                else check_rules(data_subset.copy(), chart_info)
            )
            for rule, violated in rules.items():
                if violated:
                    violated_rules[rule] = True
                    plt.plot(i, raw_df['point_val'].iloc[i], 'ro', markersize=10)

    # === X 軸 ===
    interval = max(1, len(raw_df) // 30)
    plt.xticks(x_values[::interval], raw_df['point_time'].dt.strftime("%Y-%m-%d")[::interval], rotation=90)

    # === 區間上色 ===
    plt.axvspan(start_index, end_index, color='#E83F6F', alpha=0.1, label='Weekly Data')
    if start_index > 0:
        plt.axvspan(0, start_index - 1, color='#3772FF', alpha=0.1, label='Baseline Data')

    plt.xlim([x_values[0] - 1, None])
    plt.legend()

    # === 美化 ===
    ax = plt.gca()
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)

    plt.tight_layout()

    output_path = 'output'
    os.makedirs(output_path, exist_ok=True)
    safe_group_name = "" if group_name == "Default" else group_name
    image_path = f"{output_path}/SPC_{safe_group_name}_{chart_info['chart_name']}.png"

    plt.savefig(image_path, bbox_inches='tight')
    plt.close()

    return image_path, violated_rules
def plot_spc_chart_interactive(raw_df, chart_info, weekly_start_date, weekly_end_date,
                               record_results=None, debug=False, use_batch_id_labels=False,
                               precomputed_rule_results=None):
    """
    建立互動式 SPC 圖表，返回 FigureCanvas 而不是儲存圖片
    支援滑鼠懸停 tooltip 顯示時間、數值、WE rule 資訊、Record High/Low 資訊
    支援 Batch_ID 作為 X 軸標籤
    
    Args:
        record_results: record_high_low_calculator 的返回結果，包含 record_high, record_low, highlight_status
    """
    import numpy as np
    import pandas as pd
    from matplotlib.figure import Figure
    
    # 創建 Figure 物件
    fig = Figure(figsize=(13, 2.5))
    ax = fig.add_subplot(111)

    group_name = chart_info['group_name']
    display_group_name = "" if group_name == "Default" else f"Group: [{group_name}]"
    title = (f"{display_group_name}[{chart_info['chart_name']}][{chart_info['Characteristics']}]\n"
             f"UCL: [{chart_info['UCL']}] | Target: [{chart_info['Target']}] | LCL: [{chart_info['LCL']}]")
    ax.set_title(title, loc='left', fontsize=9)  # 原12 → 9

    # === 先排序 + reset_index，確保 index = 0..N-1 ===
    raw_df = raw_df.copy()
    raw_df['point_time'] = pd.to_datetime(raw_df['point_time'])
    raw_df = raw_df.sort_values('point_time').reset_index(drop=True)

    # === 移除點數限制，使用全部數據 ===
    print(f"  plot_spc_chart_interactive: 使用全部數據點數 {len(raw_df)}")

    points_num = len(raw_df)
    x_values = np.arange(points_num)

    # === 控制線 ===
    draw_horizontal_limit(ax, chart_info.get('UCL'), 'UCL', -0.8, points_num + 2, fontsize=7)
    ax.hlines(chart_info['Target'], -0.8, points_num + 2, colors='#087E8B', linestyles='--', linewidth=1)
    draw_horizontal_limit(ax, chart_info.get('LCL'), 'LCL', -0.8, points_num + 2, fontsize=7)

    # 標籤字體縮小
    ax.text(x=points_num + 2, y=chart_info['Target'], s='Target', va='center', ha='left', fontsize=7, color='#087E8B')

    # === 畫數據線 ===
    line = ax.plot(x_values, raw_df['point_val'], color='#5863F8', marker='o', linestyle='-',markersize=4)[0]

    # === 找當週的 index ===
    ws = pd.to_datetime(weekly_start_date)
    we = pd.to_datetime(weekly_end_date)

    start_index = raw_df[raw_df['point_time'] >= ws].index.min()
    end_index   = raw_df[raw_df['point_time'] <= we].index.max()

    if debug:
        logger.debug("weekly_start_date=%s, weekly_end_date=%s", ws, we)
        logger.debug("start_index=%s, time=%s", start_index, raw_df.loc[start_index, 'point_time'])
        logger.debug("end_index=%s, time=%s", end_index, raw_df.loc[end_index, 'point_time'])

    # === 檢查 rule，標紅點並收集資訊 ===
    violated_rules = {rule: False for rule in chart_info.get('rule_list', [])}
    violation_info = {}  # 存儲每個點的違規資訊
    record_high_low_info = {}  # 存儲 record high/low 的點
    
    # 如果有 record_results，檢查哪些點創了 record high/low
    if record_results and (record_results.get('record_high', False) or record_results.get('record_low', False)):
        ws = pd.to_datetime(weekly_start_date)
        we = pd.to_datetime(weekly_end_date)
        weekly_mask = (raw_df['point_time'] >= ws) & (raw_df['point_time'] <= we)
        weekly_indices = raw_df[weekly_mask].index
        
        if len(weekly_indices) > 0:
            # 獲取歷史數據（weekly 之前的所有數據）
            historical_mask = raw_df['point_time'] < ws
            historical_data = raw_df[historical_mask]['point_val'].values
            
            if len(historical_data) > 0:
                historical_max = np.max(historical_data)
                historical_min = np.min(historical_data)
                
                # 只標記第一個創記錄的點
                record_high_marked = False
                record_low_marked = False
                
                # 檢查每個 weekly 點是否創新高/新低
                for idx in weekly_indices:
                    val = raw_df.loc[idx, 'point_val']
                    record_types = []
                    
                    if record_results.get('record_high', False) and not record_high_marked and val > historical_max:
                        record_types.append('Record High')
                        record_high_marked = True  # 標記已找到第一個創新高的點
                    
                    if record_results.get('record_low', False) and not record_low_marked and val < historical_min:
                        record_types.append('Record Low')
                        record_low_marked = True  # 標記已找到第一個創新低的點
                    
                    if record_types:
                        record_high_low_info[idx] = record_types

    for i in range(start_index, end_index + 1):
        data_subset = raw_df.iloc[:i+1].tail(15)
        if not data_subset.empty:
            rules = (
                precomputed_rule_results.get(i, {})
                if precomputed_rule_results is not None
                else check_rules(data_subset.copy(), chart_info)
            )
            violated_rule_names = []
            for rule, violated in rules.items():
                if violated:
                    violated_rules[rule] = True
                    violated_rule_names.append(rule)
            
            if violated_rule_names:
                ax.plot(i, raw_df['point_val'].iloc[i], 'ro', markersize=4)
                violation_info[i] = violated_rule_names
    
    # === 標記 Record High/Low 點（用紫紅色三角形） ===
    # *** 功能已關閉 - 不在圖表上顯示 record high/low 標記 ***
    # for idx, record_types in record_high_low_info.items():
    #     if 'Record High' in record_types:
    #         ax.plot(idx, raw_df['point_val'].iloc[idx], marker='^', color='#FF1493', markersize=6, zorder=5)  # 上三角
    #     if 'Record Low' in record_types:
    #         ax.plot(idx, raw_df['point_val'].iloc[idx], marker='v', color='#FF1493', markersize=6, zorder=5)  # 下三角

    # === X 軸 ===
    interval = max(1, len(raw_df) // 30)
    ax.set_xticks(x_values[::interval])
    
    # 根據設定使用 Batch_ID 或日期作為 X 軸標籤
    if use_batch_id_labels and 'Batch_ID' in raw_df.columns:
        # 使用 Batch_ID 作為 X 軸標籤
        ax.set_xticklabels(
            raw_df['Batch_ID'].astype(str)[::interval],
            rotation=90,
            fontsize=7
        )
    else:
        # 使用日期作為 X 軸標籤
        ax.set_xticklabels(
            raw_df['point_time'].dt.strftime("%Y-%m-%d")[::interval],
            rotation=90,
            fontsize=7  # 原預設10 → 7
        )

    # === Y 軸 ===
    ax.tick_params(axis='y', labelsize=7)  # 原預設10 → 7

    # === 區間上色 ===
    ax.axvspan(start_index, end_index, color='#E83F6F', alpha=0.1, label='Weekly Data')
    if start_index > 0:
        ax.axvspan(0, start_index - 1, color='#3772FF', alpha=0.1, label='Baseline Data')

    ax.set_xlim([x_values[0] - 1, None])
    ax.legend(fontsize=7)  # 原預設10 → 7

    # === 美化 ===
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)

    fig.tight_layout()
    fig.subplots_adjust(left=0.08, right=0.95)  # 讓y軸數字和右側標籤不會被裁切



    # === 創建 FigureCanvas ===
    canvas = FigureCanvas(fig)
    # 保存繪圖參數到 canvas，放大視窗時可重建獨立 Figure，避免共用 Figure 導致縮圖跟著改變
    try:
        canvas._plot_args = (raw_df.copy(), chart_info.copy(), weekly_start_date, weekly_end_date, use_batch_id_labels)
        canvas._plot_kind = 'spc'
    except Exception:
        canvas._plot_args = None
        canvas._plot_kind = None
    
    # === 添加 mplcursors 互動功能 ===
    # 使用預設設定，mplcursors 會自動在接近資料點時觸發
    cursor = mplcursors.cursor(
        line, 
        hover=mplcursors.HoverMode.Transient,
        highlight=False  # 啟用點高亮
    )
    
    # 用字典追蹤垂直柱，避免 sel.extras 屬性問題
    active_bars = {}
    
    def on_add(sel):
        # 1. 改變游標為手型，表示可互動
        canvas.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        
        # 2. 獲取點擊的點的索引
        if hasattr(sel.target, 'index'):
            index = int(sel.target.index)
        else:
            # 如果沒有 index 屬性，嘗試從座標推算
            x_coord = sel.target[0]
            index = int(round(x_coord))
        
        # 3. 確保索引在有效範圍內
        if 0 <= index < len(raw_df):
            # 4. 判斷該點是否在當週區域，決定柱體顏色
            bar_color = 'lightblue'  # 預設顏色
            bar_alpha = 0.3
            
            # 嘗試獲取當週範圍信息
            try:
                if hasattr(canvas, '_plot_args') and canvas._plot_args:
                    # 支援舊格式（4個參數）和新格式（5個參數）
                    if len(canvas._plot_args) >= 5:
                        _, _, weekly_start, weekly_end, _ = canvas._plot_args
                    else:
                        _, _, weekly_start, weekly_end = canvas._plot_args
                    ws = pd.to_datetime(weekly_start)
                    we = pd.to_datetime(weekly_end)
                    start_idx = raw_df[raw_df['point_time'] >= ws].index.min()
                    end_idx = raw_df[raw_df['point_time'] <= we].index.max()
                    
                    # 如果點在當週區域，使用淺紅色；否則使用藍色
                    if pd.notna(start_idx) and pd.notna(end_idx) and start_idx <= index <= end_idx:
                        bar_color = 'lightcoral'  # 當週數據點用淺紅色
                        bar_alpha = 0.3
            except Exception as e:
                # 如果出錯，使用預設顏色
                pass
            
            # 在懸停點後面繪製垂直透明柱
            vertical_bar = ax.axvline(
                x=index, 
                ymin=0, ymax=1,  # 從底部到頂部
                color=bar_color,
                alpha=bar_alpha,  # 透明度
                linewidth=8,  # 柱子寬度
                zorder=1  # 在背景層，不遮蓋資料點
            )
            canvas.draw_idle()
            
            # 5. 儲存垂直柱供後續移除，使用 selection 物件 id 作為 key
            active_bars[id(sel)] = vertical_bar
            
            # 6. 設置 tooltip
            time_str = raw_df['point_time'].iloc[index].strftime("%Y-%m-%d %H:%M:%S")
            value = raw_df['point_val'].iloc[index]
            
            # 總是顯示 Batch_ID（如果有的話），不論 use_batch_id_labels 設定如何
            if 'Batch_ID' in raw_df.columns:
                batch_id = raw_df['Batch_ID'].iloc[index]
                tooltip_text = f"Batch ID: {batch_id}\nTime: {time_str}\nValue: {value:.3f}"
            else:
                tooltip_text = f"Time: {time_str}\nValue: {value:.3f}"

            if 'Customer' in raw_df.columns:
                customer = raw_df['Customer'].iloc[index]
                if pd.notna(customer) and str(customer).strip():
                    tooltip_text = f"Customer: {str(customer).strip()}\n{tooltip_text}"
            
            # 如果這個點有違規，添加 WE rule 資訊
            if index in violation_info:
                we_rules = ", ".join(violation_info[index])
                tooltip_text += f"\nWE Rules: {we_rules}"
            else:
                tooltip_text += "\nWE Rules: None"
            
            # 如果這個點創了 Record High/Low，添加資訊
            # *** 功能已關閉 - 不在 tooltip 中顯示 record high/low 資訊 ***
            # if index in record_high_low_info:
            #     record_info = ", ".join(record_high_low_info[index])
            #     tooltip_text += f"\n{record_info}"
            
            sel.annotation.set_text(tooltip_text)
    
    def on_remove(sel):
        # 1. 恢復預設箭頭游標
        canvas.setCursor(QtCore.Qt.CursorShape.ArrowCursor)
        
        # 2. 移除垂直柱
        sel_id = id(sel)
        if sel_id in active_bars:
            vertical_bar = active_bars[sel_id]
            if hasattr(vertical_bar, 'remove'):
                vertical_bar.remove()
            del active_bars[sel_id]
            canvas.draw_idle()
    
    cursor.connect("add", on_add)
    cursor.connect("remove", on_remove)

    return canvas, violated_rules


    

def plot_weekly_spc_chart(raw_df, chart_info, weekly_start_date, weekly_end_date,
                          debug=False, precomputed_rule_results=None):
    import os
    import numpy as np
    import matplotlib.pyplot as plt
    import pandas as pd

    # work on a local copy，確保 point_time 是 datetime，並依時間排序、重設 index（得到連續的 global index）
    df = raw_df.copy()
    df['point_time'] = pd.to_datetime(df['point_time'])
    df = df.sort_values('point_time').reset_index(drop=True)

    ws = pd.to_datetime(weekly_start_date)
    we = pd.to_datetime(weekly_end_date)

    # 取當週資料（index 保留為 global index）
    df_weekly = df[(df['point_time'] >= ws) & (df['point_time'] <= we)].copy()

    if df_weekly.empty:
        # 若當週沒有資料，仍畫一張空圖以避免例外
        plt.figure(figsize=(14, 6))
        plt.title("No weekly data", loc='left')
        plt.tight_layout()
        output_path = 'output'
        os.makedirs(output_path, exist_ok=True)
        image_path = f'{output_path}/Weekly_SPC_empty.png'
        plt.savefig(image_path, bbox_inches='tight')
        plt.close()
        return image_path

    # 移除點數限制，使用全部週數據
    print(f"  plot_weekly_spc_chart: 使用週數據點數 {len(df_weekly)}")

    points_num = len(df_weekly)
    x_values = np.arange(points_num)

    plt.figure(figsize=(14, 6))

    group_name = chart_info.get('group_name', '')
    display_group_name = "" if group_name == "Default" else f"Group: [{group_name}]"
    title = (f"{display_group_name}[{chart_info['chart_name']}][{chart_info['Characteristics']}]\n"
             f"UCL: [{chart_info['UCL']}] | Target: [{chart_info['Target']}] | LCL: [{chart_info['LCL']}]")
    plt.title(title, loc='left', fontsize=12)

    # 繪製控制線（使用 weekly 範圍作為長度參考）
    draw_horizontal_limit(plt, chart_info.get('UCL'), 'UCL', -0.8, points_num + 2)
    plt.hlines(chart_info['Target'], -0.8, points_num + 2, colors='#087E8B', linestyles='--', linewidth=1)
    draw_horizontal_limit(plt, chart_info.get('LCL'), 'LCL', -0.8, points_num + 2)
    plt.text(x=points_num + 2, y=chart_info['Target'], s='Target', va='center', ha='left', fontsize=10, color='#087E8B')

    # 畫 weekly 的折線（x 軸使用 0..N-1）
    plt.plot(x_values, df_weekly['point_val'].values, color='#5863F8', marker='o', linestyle='-')

    # 檢查每一個 weekly 點：用 global index (df_weekly.index) 去取 global 的前 idx+1 筆資料來檢查 rules
    violated_points = []  # 收集觸發的點 (pos_in_weekly, global_index, time, value, rules)
    for pos_in_weekly, (global_idx, row) in enumerate(df_weekly.iterrows()):
        # full_data_subset = 全部原始資料從頭到這個 global index 的最後 15 筆
        full_data_subset = df.iloc[:global_idx + 1].tail(15)
        if full_data_subset.empty:
            continue
        rules = (
            precomputed_rule_results.get(int(global_idx), {})
            if precomputed_rule_results is not None
            else check_rules(full_data_subset.copy(), chart_info)
        )
        if any(rules.values()):
            # 在 weekly plot 的位置畫紅點（x = pos_in_weekly）
            plt.plot(pos_in_weekly, row['point_val'], 'ro', markersize=10)
            violated_points.append((pos_in_weekly, global_idx, row['point_time'], row['point_val'], rules))
            if debug:
                print(f'[VIOL] weekly_pos={pos_in_weekly} global_idx={global_idx} time={row["point_time"]} '
                      f'value={row["point_val"]} rules={ {k:v for k,v in rules.items() if v} }')

    # X axis labels
    interval = max(1, points_num // 30)
    plt.xticks(x_values[::interval], df_weekly['point_time'].dt.strftime("%Y-%m-%d")[::interval], rotation=90)

    plt.axvspan(0, points_num - 1, color='#E83F6F', alpha=0.1, label='Weekly Data')
    plt.xlim([x_values[0] - 1, None])
    plt.legend()

    ax = plt.gca()
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)

    plt.tight_layout()

    output_path = 'output'
    os.makedirs(output_path, exist_ok=True)
    safe_group_name = "" if group_name == "Default" else group_name
    image_path = f'{output_path}/Weekly_SPC_{safe_group_name}_{chart_info["chart_name"]}.png'
    plt.savefig(image_path, bbox_inches='tight')
    plt.close()

    # 回傳圖片路徑（如需，也可以回傳 violated_points 供 debug 使用）
    return image_path


def _get_tool_group_column(dataframe):
    """Return the supported By Tool column, preferring Matching when both exist."""
    for column in ('Matching', 'Tool'):
        if column in dataframe.columns:
            return column
    return None


def _plot_spc_by_tool_time_base(raw_df, chart_info, weekly_start_date, weekly_end_date, grouped=False):
    """Create a time-ordered or tool-grouped SPC canvas colored by Tool/Matching."""
    import matplotlib.colors as mcolors

    figure = Figure(figsize=(13, 3.2))
    axis = figure.add_subplot(111)
    data = raw_df.copy()
    tool_column = _get_tool_group_column(data)

    if tool_column is None:
        axis.text(0.5, 0.5, tr("by_tool_missing_column"), ha='center', va='center', transform=axis.transAxes)
        axis.set_axis_off()
        return FigureCanvas(figure)

    data['point_time'] = pd.to_datetime(data['point_time'], errors='coerce')
    data['point_val'] = pd.to_numeric(data['point_val'], errors='coerce')
    data = data.dropna(subset=['point_time', 'point_val']).copy()
    data[tool_column] = data[tool_column].fillna('Unassigned').astype(str)

    if data.empty:
        axis.text(0.5, 0.5, "No valid By Tool data", ha='center', va='center', transform=axis.transAxes)
        axis.set_axis_off()
        return FigureCanvas(figure)

    sort_columns = [tool_column, 'point_time'] if grouped else ['point_time']
    data = data.sort_values(sort_columns).reset_index(drop=True)
    tools = sorted(data[tool_column].unique(), key=lambda value: str(value).casefold())
    palette = list(mcolors.TABLEAU_COLORS.values())
    if len(tools) > len(palette):
        palette = [plt.get_cmap('tab20')(i / max(len(tools) - 1, 1)) for i in range(len(tools))]
    tool_colors = {tool: palette[index % len(palette)] for index, tool in enumerate(tools)}

    x_values = np.arange(len(data))
    if not grouped:
        axis.plot(x_values, data['point_val'], color='#94A3B8', linewidth=0.8, alpha=0.45, zorder=1)

    artist_rows = {}
    for tool_index, tool in enumerate(tools):
        subset = data[data[tool_column] == tool]
        line_style = '-' if grouped else 'None'
        artist, = axis.plot(
            subset.index,
            subset['point_val'],
            marker='o',
            linestyle=line_style,
            linewidth=0.9,
            markersize=4,
            alpha=0.88,
            color=tool_colors[tool],
            label=str(tool),
            zorder=3,
        )
        artist_rows[id(artist)] = subset.reset_index(drop=True)
        if grouped and tool_index > 0:
            axis.axvline(subset.index.min() - 0.5, color='#CBD5E1', linestyle=':', linewidth=1)

    if not grouped and weekly_start_date is not None and weekly_end_date is not None:
        weekly_mask = (
            (data['point_time'] >= pd.to_datetime(weekly_start_date))
            & (data['point_time'] <= pd.to_datetime(weekly_end_date))
        )
        weekly_indices = data.index[weekly_mask]
        if len(weekly_indices):
            axis.axvspan(weekly_indices.min() - 0.5, weekly_indices.max() + 0.5,
                         color='#FDE68A', alpha=0.22, label='Weekly range', zorder=0)

    x_min, x_max = -0.8, len(data) + 1
    draw_horizontal_limit(axis, chart_info.get('UCL'), 'UCL', x_min, x_max, fontsize=7)
    draw_horizontal_limit(axis, chart_info.get('LCL'), 'LCL', x_min, x_max, fontsize=7)
    if is_valid_number(chart_info.get('Target')):
        target = float(chart_info.get('Target'))
        axis.hlines(target, x_min, x_max, colors='#0F766E', linestyles='--', linewidth=1)
        axis.text(x=x_max, y=target, s='Target', va='center', ha='left', fontsize=7, color='#0F766E')

    tick_interval = max(1, len(data) // 24)
    tick_indices = data.index[::tick_interval]
    axis.set_xticks(tick_indices)
    axis.set_xticklabels(
        data.loc[tick_indices, 'point_time'].dt.strftime('%m/%d %H:%M'),
        rotation=90,
        fontsize=7,
    )
    axis.tick_params(axis='y', labelsize=8)
    axis.grid(axis='y', color='#E2E8F0', linewidth=0.7, alpha=0.8)
    axis.spines['right'].set_visible(False)
    axis.spines['top'].set_visible(False)

    group_name = chart_info.get('group_name', chart_info.get('GroupName', ''))
    chart_name = chart_info.get('chart_name', chart_info.get('ChartName', ''))
    title_key = 'show_by_tool_group_title' if grouped else 'show_by_tool_time_title'
    axis.set_title(f"{group_name} / {chart_name}\n{tr(title_key)}", loc='left', fontsize=9, fontweight='bold')
    axis.legend(
        loc='upper left',
        bbox_to_anchor=(1.10, 1.0),
        fontsize=8.5,
        frameon=False,
        labelspacing=0.65,
        handletextpad=0.7,
        borderaxespad=0,
    )
    # Legend 與控制線標籤都保留在 Figure 內，避免被相鄰圖表裁切。
    figure.subplots_adjust(left=0.07, right=0.72, top=0.80, bottom=0.30)

    canvas = FigureCanvas(figure)
    artists = [artist for artist in axis.lines if id(artist) in artist_rows]
    if artists:
        cursor = mplcursors.cursor(artists, hover=True)

        @cursor.connect('add')
        def _show_tool_point(selection):
            rows = artist_rows.get(id(selection.artist))
            point_index = int(round(float(selection.index)))
            if rows is None or point_index < 0 or point_index >= len(rows):
                return
            point = rows.iloc[point_index]
            selection.annotation.set_text(
                f"{tool_column}: {point[tool_column]}\n"
                f"Time: {point['point_time']:%Y-%m-%d %H:%M}\n"
                f"Value: {point['point_val']:.4g}"
            )
        canvas._mpl_cursor = cursor

    canvas._plot_args = (raw_df.copy(), chart_info.copy(), weekly_start_date, weekly_end_date)
    canvas._plot_kind = 'spc_by_tool_group' if grouped else 'spc_by_tool_time'
    return canvas


def plot_spc_by_tool_time(raw_df, chart_info, weekly_start_date=None, weekly_end_date=None):
    return _plot_spc_by_tool_time_base(
        raw_df, chart_info, weekly_start_date, weekly_end_date, grouped=False
    )


def plot_spc_by_tool_time_grouping(raw_df, chart_info, weekly_start_date=None, weekly_end_date=None):
    return _plot_spc_by_tool_time_base(
        raw_df, chart_info, weekly_start_date, weekly_end_date, grouped=True
    )


def plot_weekly_spc_chart_interactive(raw_df, chart_info, weekly_start_date, weekly_end_date,
                                      record_results=None, debug=False,
                                      use_batch_id_labels=False,
                                      precomputed_rule_results=None):
    """
    建立互動式 Weekly SPC 圖表，返回 FigureCanvas 而不是儲存圖片
    支援滑鼠懸停 tooltip 顯示時間、數值、WE rule 資訊
    支援 Batch_ID 作為 X 軸標籤
    支援標記 record high/low 點
    """
    import numpy as np
    import pandas as pd
    from matplotlib.figure import Figure

    # 創建 Figure 物件
    fig = Figure(figsize=(13, 2.5))
    ax = fig.add_subplot(111)

    # work on a local copy，確保 point_time 是 datetime，並依時間排序、重設 index（得到連續的 global index）
    df = raw_df.copy()
    df['point_time'] = pd.to_datetime(df['point_time'])
    df = df.sort_values('point_time').reset_index(drop=True)

    ws = pd.to_datetime(weekly_start_date)
    we = pd.to_datetime(weekly_end_date)

    # 取當週資料（index 保留為 global index）
    df_weekly = df[(df['point_time'] >= ws) & (df['point_time'] <= we)].copy()

    if df_weekly.empty:
        # 若當週沒有資料，仍畫一張空圖以避免例外
        ax.set_title("No weekly data", loc='left')
        fig.tight_layout()
        canvas = FigureCanvas(fig)
        return canvas

    # 移除點數限制，使用全部週數據
    print(f"  plot_weekly_spc_chart_interactive: 使用週數據點數 {len(df_weekly)}")

    points_num = len(df_weekly)
    x_values = np.arange(points_num)

    group_name = chart_info.get('group_name', '')
    display_group_name = "" if group_name == "Default" else f"Group: [{group_name}]"
    title = (f"{display_group_name}[{chart_info['chart_name']}][{chart_info['Characteristics']}]\n"
             f"UCL: [{chart_info['UCL']}] | Target: [{chart_info['Target']}] | LCL: [{chart_info['LCL']}]")
    ax.set_title(title, loc='left', fontsize=9)

    # 繪製控制線（使用 weekly 範圍作為長度參考）
    draw_horizontal_limit(ax, chart_info.get('UCL'), 'UCL', -0.8, points_num + 2, fontsize=7)
    ax.hlines(chart_info['Target'], -0.8, points_num + 2, colors='#087E8B', linestyles='--', linewidth=1)
    draw_horizontal_limit(ax, chart_info.get('LCL'), 'LCL', -0.8, points_num + 2, fontsize=7)
    ax.text(x=points_num + 2, y=chart_info['Target'], s='Target', va='center', ha='left', fontsize=7, color='#087E8B')

    # 畫 weekly 的折線（x 軸使用 0..N-1）
    line = ax.plot(x_values, df_weekly['point_val'].values, color='#5863F8', marker='o', linestyle='-',markersize=4)[0]

    # 檢查每一個 weekly 點：用 global index (df_weekly.index) 去取 global 的前 idx+1 筆資料來檢查 rules
    violated_info = {}  # 存儲每個 weekly 位置的違規資訊 {weekly_pos: [rule_names]}
    record_high_low_info_weekly = {}  # 存儲每個 weekly 位置的 record high/low 資訊
    
    for pos_in_weekly, (global_idx, row) in enumerate(df_weekly.iterrows()):
        # full_data_subset = 全部原始資料從頭到這個 global index 的最後 15 筆
        full_data_subset = df.iloc[:global_idx + 1].tail(15)
        if full_data_subset.empty:
            continue
        rules = (
            precomputed_rule_results.get(int(global_idx), {})
            if precomputed_rule_results is not None
            else check_rules(full_data_subset.copy(), chart_info)
        )
        violated_rule_names = []
        for rule, violated in rules.items():
            if violated:
                violated_rule_names.append(rule)
        
        if violated_rule_names:
            # 在 weekly plot 的位置畫紅點（x = pos_in_weekly）
            ax.plot(pos_in_weekly, row['point_val'], 'ro', markersize=4)
            violated_info[pos_in_weekly] = {
                'rules': violated_rule_names,
                'global_idx': global_idx,
                'time': row['point_time'],
                'value': row['point_val']
            }
            if debug:
                print(f'[VIOL] weekly_pos={pos_in_weekly} global_idx={global_idx} time={row["point_time"]} '
                      f'value={row["point_val"]} rules={ {k:v for k,v in rules.items() if v} }')
    
    # === 標記 Record High/Low 點（用粉紅色三角形）===
    if record_results and (record_results.get('record_high', False) or record_results.get('record_low', False)):
        # 獲取歷史數據（weekly 之前的所有數據）
        historical_mask = df['point_time'] < ws
        historical_data = df[historical_mask]['point_val'].values
        
        if len(historical_data) > 0:
            historical_max = np.max(historical_data)
            historical_min = np.min(historical_data)
            
            # 只標記第一個創記錄的點
            record_high_marked = False
            record_low_marked = False
            
            # 檢查每個 weekly 點是否創新高/新低
            for pos_in_weekly, (global_idx, row) in enumerate(df_weekly.iterrows()):
                val = row['point_val']
                record_types = []
                
                if record_results.get('record_high', False) and not record_high_marked and val > historical_max:
                    record_types.append('Record High')
                    record_high_marked = True  # 標記已找到第一個創新高的點
                
                if record_results.get('record_low', False) and not record_low_marked and val < historical_min:
                    record_types.append('Record Low')
                    record_low_marked = True  # 標記已找到第一個創新低的點
                
                if record_types:
                    record_high_low_info_weekly[pos_in_weekly] = record_types
                    # 繪製三角形標記
                    # *** 功能已關閉 - 不在週圖表上顯示 record high/low 標記 ***
                    # if 'Record High' in record_types:
                    #     ax.plot(pos_in_weekly, val, marker='^', color='#FF1493', markersize=6, zorder=5)
                    # if 'Record Low' in record_types:
                    #     ax.plot(pos_in_weekly, val, marker='v', color='#FF1493', markersize=6, zorder=5)

    # X axis labels
    interval = max(1, points_num // 30)
    ax.set_xticks(x_values[::interval])
    
    # 根據設定使用 Batch_ID 或日期作為 X 軸標籤
    if use_batch_id_labels and 'Batch_ID' in df_weekly.columns:
        # 使用 Batch_ID 作為 X 軸標籤
        ax.set_xticklabels(df_weekly['Batch_ID'].astype(str)[::interval], rotation=90, fontsize=7)
    else:
        # 使用日期作為 X 軸標籤
        ax.set_xticklabels(df_weekly['point_time'].dt.strftime("%Y-%m-%d")[::interval], rotation=90, fontsize=7)
    ax.tick_params(axis='y', labelsize=7)  # Y 軸字體縮小
    ax.axvspan(0, points_num - 1, color='#E83F6F', alpha=0.1, label='Weekly Data')
    ax.set_xlim([x_values[0] - 1, None])
    ax.legend(fontsize=7)

    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)

    fig.tight_layout()
    fig.subplots_adjust(left=0.08, right=0.95)  # 讓y軸數字和右側標籤不會被裁切

    # === 創建 FigureCanvas ===
    canvas = FigureCanvas(fig)
    # 保存繪圖參數到 canvas，放大視窗時可重建獨立 Figure
    try:
        canvas._plot_args = (df_weekly.copy(), chart_info.copy(), weekly_start_date, weekly_end_date, use_batch_id_labels)
        canvas._plot_kind = 'weekly'
    except Exception:
        canvas._plot_args = None
        canvas._plot_kind = None
    
    # === 添加 mplcursors 互動功能 ===
    cursor = mplcursors.cursor(line, hover=mplcursors.HoverMode.Transient)
    
    # 用字典追蹤垂直柱，避免 sel.extras 屬性問題
    active_weekly_bars = {}
    
    def on_weekly_add(sel):
        # 1. 改變游標為手型，表示可互動
        canvas.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        
        # 2. 獲取點擊的點的索引
        if hasattr(sel.target, 'index'):
            index = int(sel.target.index)
        else:
            # 如果沒有 index 屬性，嘗試從座標推算
            x_coord = sel.target[0]
            index = int(round(x_coord))
        
        # 3. 確保索引在有效範圍內
        if 0 <= index < len(df_weekly):
            row = df_weekly.iloc[index]
            
            # 4. 在懸停點後面繪製垂直透明柱
            vertical_bar = ax.axvline(
                x=index, 
                ymin=0, ymax=1,  # 從底部到頂部
                color='lightcoral',  # 淺珊瑚色
                alpha=0.3,  # 透明度
                linewidth=8,  # 柱子寬度
                zorder=1  # 在背景層，不遮蓋資料點
            )
            canvas.draw_idle()
            
            # 5. 儲存垂直柱供後續移除，使用 selection 物件 id 作為 key
            active_weekly_bars[id(sel)] = vertical_bar
            
            # 6. 設置 tooltip
            time_str = row['point_time'].strftime("%Y-%m-%d %H:%M:%S")
            value = row['point_val']
            
            # 總是顯示 Batch_ID（如果有的話），不論 use_batch_id_labels 設定如何
            if 'Batch_ID' in df_weekly.columns:
                batch_id = row['Batch_ID']
                tooltip_text = f"Batch ID: {batch_id}\nTime: {time_str}\nValue: {value:.3f}"
            else:
                tooltip_text = f"Time: {time_str}\nValue: {value:.3f}"

            if 'Customer' in df_weekly.columns:
                customer = row['Customer']
                if pd.notna(customer) and str(customer).strip():
                    tooltip_text = f"Customer: {str(customer).strip()}\n{tooltip_text}"
            
            # 如果這個點有違規，添加 WE rule 資訊
            if index in violated_info:
                we_rules = ", ".join(violated_info[index]['rules'])
                tooltip_text += f"\nWE Rules: {we_rules}"
            else:
                tooltip_text += "\nWE Rules: None"
            
            # 如果這個點創了 Record High/Low，添加資訊
            # *** 功能已關閉 - 不在週圖表 tooltip 中顯示 record high/low 資訊 ***
            # if index in record_high_low_info_weekly:
            #     record_info = ", ".join(record_high_low_info_weekly[index])
            #     tooltip_text += f"\n{record_info}"
            
            sel.annotation.set_text(tooltip_text)
    
    def on_weekly_remove(sel):
        # 1. 恢復預設箭頭游標
        canvas.setCursor(QtCore.Qt.CursorShape.ArrowCursor)
        
        # 2. 移除垂直柱
        sel_id = id(sel)
        if sel_id in active_weekly_bars:
            vertical_bar = active_weekly_bars[sel_id]
            if hasattr(vertical_bar, 'remove'):
                vertical_bar.remove()
            del active_weekly_bars[sel_id]
            canvas.draw_idle()
    
    cursor.connect("add", on_weekly_add)
    cursor.connect("remove", on_weekly_remove)

    return canvas


def save_results_to_excel(results_df, scale_factor=0.3):
    results_df['group_name'] = results_df['group_name'].replace("Default", "")  # 替換 Default 為空白

    workbook = xlsxwriter.Workbook('result_with_images.xlsx')
    worksheet = workbook.add_worksheet()

    cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10})
    header_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 12, 'bold': True})

    col_widths = {}

    for col_idx, header in enumerate(results_df.columns):
        worksheet.write(0, col_idx + 2, header, header_format)
        col_widths[col_idx + 2] = max(len(header), col_widths.get(col_idx + 2, 0))

    max_image_height = 0
    image_column_width = 0

    for row_idx, row in enumerate(results_df.itertuples(index=False), start=1):
        img_path = row.chart_path
        weekly_spc_chart_path = row.weekly_chart_path

        x_offset = 0
        y_offset = 10
        options = {
            'x_scale': scale_factor,
            'y_scale': scale_factor,
            'x_offset': x_offset,
            'y_offset': y_offset,
            'object_position': 1
        }

        worksheet.insert_image(row_idx, 0, img_path, options)
        worksheet.insert_image(row_idx, 1, weekly_spc_chart_path, options)

        img = Image.open(img_path)
        image_width, image_height = img.size
        scaled_width = image_width * scale_factor
        scaled_height = image_height * scale_factor

        if scaled_height > max_image_height:
            max_image_height = scaled_height
        if scaled_width > image_column_width:
            image_column_width = scaled_width

        for col_idx, value in enumerate(row, start=1):
            worksheet.write(row_idx, col_idx + 1, value, cell_format)
            col_widths[col_idx + 1] = max(col_widths.get(col_idx + 1, 0), len(str(value)))

    worksheet.set_column(0, 0, image_column_width / 7)
    worksheet.set_column(1, 1, image_column_width / 7)

    for col_idx, width in col_widths.items():
        worksheet.set_column(col_idx, col_idx, width + 5)

    for row_idx in range(1, len(results_df) + 1):
        worksheet.set_row(row_idx, max_image_height)

    workbook.close()


# 🔧 封裝路徑處理函式
def resource_path(relative_path):
    if getattr(sys, 'frozen', False):  # 如果是打包環境
        base_path = os.path.dirname(sys.executable)
    else:  # 開發環境
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# 常數定義
HEADERS = ["Total Chart", "Weekly Chart", "Chart Info."]
OOB_KEYS = ['HL_P95_shift', 'HL_P50_shift', 'HL_P05_shift', 'HL_sticking_shift', 'HL_trending', 'HL_high_OOC', 'HL_record_high_low', 'HL_category_LT_shift']


class TriangleButton(QtWidgets.QPushButton):
    """三角形形狀的摺疊按鈕"""
    def __init__(self, direction='left', parent=None):
        super().__init__(parent)
        self.direction = direction  # 'left' 或 'right'
        self.setFixedSize(20, 40)
        self.is_hovered = False
        
    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        
        # 獲取按鈕的實際寬高
        width = self.width()
        height = self.height()
        
        # 設定顏色
        if self.isDown():
            color = QtGui.QColor('#000957')
            border_color = QtGui.QColor('#FFEB00')
        elif self.is_hovered:
            color = QtGui.QColor('#577BC1')
            border_color = QtGui.QColor('#7B9FD3')
        else:
            color = QtGui.QColor('#344CB7')
            border_color = QtGui.QColor('#577BC1')
        
        # 計算三角形座標(留一些邊距)
        margin = 3
        
        # 繪製三角形
        path = QtGui.QPainterPath()
        if self.direction == 'left':
            # 向左的三角形 ◀
            path.moveTo(width - margin, margin)           # 右上
            path.lineTo(width - margin, height - margin)  # 右下
            path.lineTo(margin, height / 2)               # 左中
            path.closeSubpath()
        else:
            # 向右的三角形 ▶
            path.moveTo(margin, margin)                   # 左上
            path.lineTo(margin, height - margin)          # 左下
            path.lineTo(width - margin, height / 2)       # 右中
            path.closeSubpath()
        
        # 繪製陰影效果
        if self.is_hovered:
            painter.setPen(QtCore.Qt.PenStyle.NoPen)
            painter.setBrush(QtGui.QColor(87, 123, 193, 80))
            shadow_path = QtGui.QPainterPath(path)
            painter.drawPath(shadow_path.translated(2, 2))
        
        # 繪製填充
        painter.setBrush(color)
        painter.setPen(QtGui.QPen(border_color, 2))
        painter.drawPath(path)
        
    def enterEvent(self, event):
        self.is_hovered = True
        self.update()
        super().enterEvent(event)
        
    def leaveEvent(self, event):
        self.is_hovered = False
        self.update()
        super().leaveEvent(event)


class OOBAnalysisCancelled(Exception):
    """Raised at safe checkpoints when the user cancels OOB processing."""


class OOBProgressDialog(QtWidgets.QProgressDialog):
    """Popup progress dialog compatible with the former QProgressBar calls."""

    def __init__(self, parent=None):
        super().__init__("", "", 0, 100, parent)
        self.setWindowTitle(tr("oob_page_title"))
        self.setCancelButtonText(tr("cancel_processing"))
        self.setWindowModality(QtCore.Qt.WindowModality.WindowModal)
        self.setMinimumDuration(0)
        self.setAutoClose(False)
        self.setAutoReset(False)
        self.setMinimumWidth(520)
        self.setMinimumHeight(150)
        self.setStyleSheet("""
            QProgressDialog {
                background-color: white;
                color: #0F172A;
            }
            QLabel {
                color: #334155;
                font-size: 13px;
                font-weight: 600;
                padding: 8px 4px;
            }
            QProgressBar {
                min-height: 24px;
                border: none;
                border-radius: 12px;
                background-color: #E2E8F0;
                color: white;
                text-align: center;
                font-weight: 700;
            }
            QProgressBar::chunk {
                border-radius: 12px;
                background-color: #4F46E5;
            }
            QPushButton {
                background-color: white;
                color: #475569;
                border: 1px solid #CBD5E1;
                border-radius: 8px;
                padding: 7px 16px;
                font-weight: 650;
            }
            QPushButton:hover { background-color: #F1F5F9; }
        """)

    def setFormat(self, text):
        self.setLabelText(str(text))
        progress_widget = self.findChild(QtWidgets.QProgressBar)
        if progress_widget is not None:
            progress_widget.setFormat("%p%")

class SPCApp(QtWidgets.QMainWindow): # 將 QTabWidget 改為 QMainWindow
    def __init__(self):
        super().__init__()

        self.filepath = resource_path('input/All_Chart_Information.xlsx')
        self.raw_data_directory = resource_path('input/raw_charts/')
        self.image_path = resource_path('image.png')
        self.results = []
        self.available_customers = []
        self.selected_customers = set()  # Empty means all customers.
        self.selected_chart_keys = None  # None means health check has not limited the scope.
        self._customer_missing_warning_shown = False

        # 翻譯系統
        self.translator = get_translator()
        self.translator.register_observer(self)

        # 性能優化：添加快取
        self.csv_cache = {}  # CSV 文件快取
        self.chart_types_cache = {}  # 數據類型快取
        
        self.filter_type_combo = None
        self.filter_value_combo = None
        self.header_container = None
        # 新增用於兩個圓餅圖和一個長條圖的 Canvas 屬性
        self.status_pie_canvas = None
        self.processed_violation_pie_canvas = None
        self.anomaly_bar_canvas = None
        # Summary Tab 中圖表相關的屬性
        self.charts_main_layout = None
        self.charts_horizontal_layout = None

        # Summary Tab 中 TableWidget 屬性
        self.violation_table_label = None
        self.violation_table = None

        # QMainWindow 需要一個中央小部件
        self.central_widget = QtWidgets.QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_horizontal_layout = QtWidgets.QHBoxLayout(self.central_widget) # 主要的水平佈局
        self.main_horizontal_layout.setContentsMargins(0, 0, 0, 0)
        self.main_horizontal_layout.setSpacing(0)

        self.init_ui()
    
    def open_oob_settings(self):
        """打開 OOB 設定對話框"""
        dialog = OOBSettingsDialog(self)
        # 載入當前設定
        dialog.set_settings(self.oob_settings)
        
        # 顯示對話框並等待用戶操作
        if dialog.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            # 用戶點擊保存，獲取新設定
            self.oob_settings = dialog.get_settings()
            print(f"OOB 設定已更新: {self.oob_settings}")
        else:
            print("OOB 設定更改已取消")
    
    def toggle_custom_time_range(self, checked):
        """切換自定義時間範圍的啟用/停用狀態"""
        self.start_datetime_edit.setEnabled(checked)
        self.end_datetime_edit.setEnabled(checked)
        
        # 同時啟用/停用快速選擇按鈕
        for btn in [self.last_7_days_btn, self.last_30_days_btn, self.last_90_days_btn, 
                   self.this_month_btn, self.last_month_btn]:
            btn.setEnabled(checked)
        
        if checked:
            print(" - 已啟用自定義時間範圍")
        else:
            print(" - 已停用自定義時間範圍")

    def set_quick_time_range(self, days):
        """設定快速時間範圍（最近N天）"""
        from datetime import datetime, timedelta
        
        end_time = QtCore.QDateTime.currentDateTime()
        start_time = end_time.addDays(-days)
        
        self.start_datetime_edit.setDateTime(start_time)
        self.end_datetime_edit.setDateTime(end_time)
        
        print(f" - 已設定為最近{days}天的時間範圍")

    def set_this_month_range(self):
        """設定本月時間範圍"""
        from datetime import datetime
        
        now = datetime.now()
        # 本月第一天
        start_of_month = datetime(now.year, now.month, 1)
        # 現在時間
        end_time = now
        
        self.start_datetime_edit.setDateTime(QtCore.QDateTime(start_of_month))
        self.end_datetime_edit.setDateTime(QtCore.QDateTime(end_time))
        
        print(" - 已設定為本月時間範圍")

    def set_last_month_range(self):
        """設定上月時間範圍"""
        from datetime import datetime, timedelta
        from calendar import monthrange
        
        now = datetime.now()
        
        # 上個月的第一天
        if now.month == 1:
            last_month = 12
            last_year = now.year - 1
        else:
            last_month = now.month - 1
            last_year = now.year
            
        start_of_last_month = datetime(last_year, last_month, 1)
        
        # 上個月的最後一天
        days_in_last_month = monthrange(last_year, last_month)[1]
        end_of_last_month = datetime(last_year, last_month, days_in_last_month, 23, 59, 59)
        
        self.start_datetime_edit.setDateTime(QtCore.QDateTime(start_of_last_month))
        self.end_datetime_edit.setDateTime(QtCore.QDateTime(end_of_last_month))
        
        print(" - 已設定為上月時間範圍")

    def init_ui(self):
        self.setWindowTitle(tr("app_title"))
        screen = QtWidgets.QApplication.primaryScreen()
        available_geometry = screen.availableGeometry()  # 使用 availableGeometry() 避開工作列
        w = int(available_geometry.width() * 0.9)  # 改為螢幕寬度的85%
        h = int(available_geometry.height() * 0.80)  # 改為螢幕高度的80%
        self.setGeometry(0, 0, w, h)
        self.setStyleSheet("""
            * {
                color: #000957;
                font-weight: bold;
            }
            /* TabWidget related styles (for oob_system_tabs) */
            QTabWidget::pane {
                border: 1px solid #c4c4c3;
                top: -1px;
                background: #f4f6f9;
            }

            QTabWidget::tab-bar {
                left: 5px;
            }

            QTabBar::tab {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                            stop: 0 #E1E1E1, stop: 0.4 #DDDDDD,
                                            stop: 0.5 #D8D8D8, stop: 1.0 #D3D3D3);
                border: 1px solid #c4ccff;
                border-bottom-color: #c2c7cb;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                min-width: 8ex;
                padding: 8px;
                font-weight: bold;
                color: #000957;
            }

            QTabBar::tab:selected, QTabBar::tab:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                            stop: 0 #fafafa, stop: 0.4 #f4f4f4,
                                            stop: 0.5 #e7e7e7, stop: 1.0 #fafafa);
            }

            QTabBar::tab:selected {
                border-color: #c2c7cb;
                border-bottom-color: #f4f6f9;
            }

            QTabBar::tab:!selected {
                margin-top: 2px;
            }

            QWidget {
                font-family: 'Segoe UI', 'Microsoft JhengHei', 'Yu Gothic UI', 'Malgun Gothic', 'Meiryo UI', sans-serif;
                background-color: #f4f6f9;
                color: #000957;
            }
            QLabel {
                font-size: 14px;
                color: #000957;
            }
            QProgressBar {
                border-radius: 12px;
                height: 25px;
                background: #e0e0e0;
                text-align: center;
                color: #000957;
                margin: 5px;
            }
            QProgressBar::chunk {
                background: linear-gradient(45deg, #344CB7, #577BC1);
                border-radius: 12px;
            }
            QPushButton {
                background-color: #344CB7;
                color: white;
                border-radius: 12px;
                padding: 12px 25px;
                font-size: 16px;
                border: none;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #577BC1;
            }
            QComboBox {
                background-color: white;
                border: 1px solid #344CB7;
                border-radius: 8px;
                padding: 8px;
                color: #000957;
                font-weight: bold;
            }
            QComboBox QAbstractItemView {
                background-color: white;
                selection-background-color: #FFEB00;
                selection-color: #000957;
            }
            QDateTimeEdit {
                background-color: white;
                border: 2px solid #344CB7;
                border-radius: 8px;
                padding: 8px;
                color: #000957;
                font-weight: bold;
                font-size: 14px;
            }
            QDateTimeEdit:focus {
                border-color: #577BC1;
            }
            QDateTimeEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 1px;
                border-left-color: #344CB7;
                border-left-style: solid;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
                background-color: #344CB7;
            }
            QDateTimeEdit::down-arrow {
                image: none;
                border: none;
                width: 0px;
                height: 0px;
                border-left: 6px solid transparent;
                border-right: 6px solid transparent;
                border-top: 8px solid white;
            }
            QCalendarWidget QToolButton {
                background-color: #344CB7;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px;
            }
            QCalendarWidget QToolButton:hover {
                background-color: #577BC1;
            }
            QCalendarWidget QMenu {
                background-color: white;
                border: 1px solid #344CB7;
            }
            QScrollArea {
                border: none;
            }
            QVBoxLayout {
                spacing: 10px;
            }
            /* Styles for the left menu buttons */
            QPushButton.menu_button { /* Added a class for menu buttons */
                background-color: #344CB7;
                color: white;
                border-radius: 8px; /* Slightly smaller radius for menu buttons */
                padding: 10px 15px;
                font-size: 14px;
                text-align: left; /* Align text to left */
                border: none;
                font-weight: bold;
            }
            QPushButton.menu_button:hover {
                background-color: #577BC1;
            }
            QPushButton.menu_button:checked { /* Style for selected button */
                background-color: #000957; /* Darker blue when selected */
                border-left: 5px solid #FFEB00; /* Yellow accent on left */
                padding-left: 10px; /* Adjust padding due to border */
            }
        """)

        # --- 左側選單區域 ---
        self.left_menu_widget = QtWidgets.QWidget()
        self.left_menu_widget.setFixedWidth(224) # 現代側欄保留較舒適的文字空間
        self.left_menu_widget.setStyleSheet("background-color: #111827;")
        self.left_menu_layout = QtWidgets.QVBoxLayout(self.left_menu_widget)
        self.left_menu_layout.setAlignment(QtCore.Qt.AlignmentFlag.AlignTop) # 按鈕靠頂部對齊
        self.left_menu_layout.setContentsMargins(14, 20, 14, 14)
        self.left_menu_layout.setSpacing(7)

        # 語言切換按鈕（在最上方）
        self.lang_button = QtWidgets.QPushButton("中 / EN")
        self.lang_button.setFixedHeight(32)
        self.lang_button.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.lang_button.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 255, 255, 0.2);  /* 半透明白色 */
                color: white;
                border-radius: 16px;
                padding: 0 12px;
                font-size: 13px;
                font-weight: 600;
                border: 1.5px solid rgba(255, 255, 255, 0.5);
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.3);
                border-color: white;
            }
            QPushButton:pressed {
                background-color: rgba(255, 255, 255, 0.15);
            }
        """)
        self.lang_button.clicked.connect(self.toggle_language)
        self.left_menu_layout.addWidget(self.lang_button, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        self.customer_filter_button = QtWidgets.QPushButton()
        self.customer_filter_button.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.customer_filter_button.setToolTip("Select one or more Customers for all analyses")
        self.customer_filter_button.setMinimumHeight(58)
        self.customer_filter_button.setStyleSheet("""
            QPushButton {
                background-color: white; color: #172554;
                border: 1px solid rgba(255, 255, 255, 0.65);
                border-radius: 10px; padding: 8px 11px;
                text-align: left; font-size: 12px; font-weight: 600;
            }
            QPushButton:hover { background-color: #eef2ff; border-color: #c7d2fe; }
            QPushButton:pressed { background-color: #e0e7ff; }
        """)
        self.customer_filter_button.clicked.connect(self.open_customer_filter_dialog)
        self.left_menu_layout.addWidget(self.customer_filter_button)
        self.refresh_customer_filter(scan=True, show_messages=False)

        # 分隔線優化 (使用半透明白色，融合感更好)
        separator = QtWidgets.QFrame()
        separator.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        # 使用 RGBA 透明度，而不是純白，這樣看起來更精緻
        separator.setStyleSheet("background-color: rgba(255, 255, 255, 0.3); max-height: 1px;")
        self.left_menu_layout.addWidget(separator)

        # 選單按鈕
        self.home_button = self._create_menu_button(tr("home"))
        self.split_data_button = self._create_menu_button(tr("split_data"))
        self.oob_system_button = self._create_menu_button(tr("oob_spc_system"))
        self.cpk_calculation_button = self._create_menu_button(tr("cpk_calculator"))
        # --- 新增 Tool Matching 按鈕 ---
        self.tool_matching_button = self._create_menu_button(tr("tool_matching"))
        # --- 新增 CL Tighten Calculator 按鈕 ---
        self.cl_tighten_button = self._create_menu_button(tr("cl_tighten"))
        self.data_check_button = self._create_menu_button(tr("data_health_monitor"))
        self.left_menu_layout.addWidget(self.home_button)
        self.left_menu_layout.addWidget(self.data_check_button)
        self.left_menu_layout.addWidget(self.split_data_button)
        self.left_menu_layout.addWidget(self.oob_system_button)
        self.left_menu_layout.addWidget(self.cpk_calculation_button)
        self.left_menu_layout.addWidget(self.tool_matching_button)
        self.left_menu_layout.addWidget(self.cl_tighten_button)
        self.left_menu_layout.addStretch()
        # 將左側選單添加到主水平佈局
        self.main_horizontal_layout.addWidget(self.left_menu_widget)
        
        # --- 摺疊/展開按鈕 (使用三角形按鈕) ---
        self.toggle_menu_button = TriangleButton(direction='left')
        self.toggle_menu_button.clicked.connect(self.toggle_left_menu)
        self.main_horizontal_layout.addWidget(self.toggle_menu_button)

        # --- 右側內容區域 (QStackedWidget) ---
        self.content_stacked_widget = QtWidgets.QStackedWidget()
        self.main_horizontal_layout.addWidget(self.content_stacked_widget)

        # 1. 首頁內容
        self.home_page = self._create_home_page()
        self.content_stacked_widget.addWidget(self.home_page)

        # 2. 拆分資料頁面內容
        self.split_data_page = self._create_split_data_page()
        self.content_stacked_widget.addWidget(self.split_data_page)

        # 3. OOB System 頁面內容 (包含 Chart Processing 和 Summary Dashboard)
        self.oob_system_tabs = QtWidgets.QTabWidget()

        # 4. Tool Matching 頁面內容
        self.tool_matching_page = self._create_tool_matching_page()
        self.content_stacked_widget.addWidget(self.tool_matching_page)
        
        # 5. CL Tighten Calculator 頁面內容
        self.cl_tighten_page = self._create_cl_tighten_page()
        self.content_stacked_widget.addWidget(self.cl_tighten_page)
        
        self.oob_system_tabs.setObjectName("OOBSystemTabs") # 給予一個名稱以便於樣式控制
        # [NEW] 6. Data Health Check 頁面
        self.data_check_page = self._create_data_check_page()
        self.content_stacked_widget.addWidget(self.data_check_page)

        # --- 建立第一個分頁 (圖表處理與顯示) ---
        self.processing_tab_widget = QtWidgets.QWidget()
        self.processing_tab_widget.setObjectName("OOBProcessingPage")
        processing_layout = QtWidgets.QVBoxLayout(self.processing_tab_widget)
        processing_layout.setSpacing(18)
        processing_layout.setContentsMargins(28, 24, 28, 28)

        page_header = QtWidgets.QWidget()
        page_header.setObjectName("TransparentWidget")
        page_header_layout = QtWidgets.QVBoxLayout(page_header)
        page_header_layout.setContentsMargins(0, 0, 0, 0)
        page_header_layout.setSpacing(4)
        self.oob_page_title_label = QtWidgets.QLabel(tr("oob_page_title"))
        self.oob_page_title_label.setObjectName("OOBPageTitle")
        self.oob_page_subtitle_label = QtWidgets.QLabel(tr("oob_page_subtitle"))
        self.oob_page_subtitle_label.setObjectName("OOBPageSubtitle")
        self.oob_page_subtitle_label.setWordWrap(True)
        page_header_layout.addWidget(self.oob_page_title_label)
        page_header_layout.addWidget(self.oob_page_subtitle_label)
        processing_layout.addWidget(page_header)

        # 儲存設定（預設值）
        self.oob_settings = {
            'show_charts_gui': True,
            'use_interactive_charts': True,
            'show_by_tool_charts': False,
            'enable_record_high_low': False,
            'custom_time_range_enabled': False,
            'start_time': QtCore.QDateTime.currentDateTime().addDays(-30),
            'end_time': QtCore.QDateTime.currentDateTime()
        }
        
        # 按鈕區域 - 水平佈局
        toolbar = QtWidgets.QFrame()
        toolbar.setObjectName("OOBToolbar")
        button_layout = QHBoxLayout(toolbar)
        button_layout.setContentsMargins(18, 14, 18, 14)
        button_layout.setSpacing(10)

        ready_widget = QtWidgets.QWidget()
        ready_widget.setObjectName("TransparentWidget")
        ready_widget.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        ready_layout = QtWidgets.QVBoxLayout(ready_widget)
        ready_layout.setContentsMargins(0, 0, 0, 0)
        ready_layout.setSpacing(2)
        self.oob_ready_title_label = QtWidgets.QLabel(tr("oob_ready_title"))
        self.oob_ready_title_label.setObjectName("OOBReadyTitle")
        self.oob_ready_desc_label = QtWidgets.QLabel(tr("oob_ready_desc"))
        self.oob_ready_desc_label.setObjectName("OOBReadyDescription")
        self.oob_ready_desc_label.setWordWrap(True)
        self.oob_ready_desc_label.setMinimumWidth(0)
        self.oob_ready_desc_label.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        ready_layout.addWidget(self.oob_ready_title_label)
        ready_layout.addWidget(self.oob_ready_desc_label)
        button_layout.addWidget(ready_widget, 1)
        
        # 設定按鈕
        self.settings_button = QtWidgets.QPushButton(f"⚙️ {tr('settings')}")
        self.settings_button.setMinimumHeight(45)
        self.settings_button.setMinimumWidth(120)
        self.settings_button.setFont(get_app_font(11))
        self.settings_button.setStyleSheet("""
            QPushButton {
                background-color: white;
                color: #334155;
                border: 1px solid #CBD5E1;
                border-radius: 10px;
                padding: 9px 16px;
                font-weight: 650;
            }
            QPushButton:hover {
                background-color: #F8FAFC;
                border-color: #94A3B8;
            }
            QPushButton:pressed {
                background-color: #F1F5F9;
            }
        """)
        self.settings_button.clicked.connect(self.open_oob_settings)
        button_layout.addWidget(
            self.settings_button,
            0,
            QtCore.Qt.AlignmentFlag.AlignVCenter,
        )
        
        # 啟動按鈕
        self.start_button = QtWidgets.QPushButton(f"▶ {tr('start_process')}")
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #4F46E5;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 10px 20px;
                font-weight: 700;
            }
            QPushButton:hover {
                background-color: #4338CA;
            }
            QPushButton:pressed {
                background-color: #3730A3;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.start_button.clicked.connect(self.process_charts)
        button_layout.addWidget(
            self.start_button,
            0,
            QtCore.Qt.AlignmentFlag.AlignVCenter,
        )
        
        # 按下開始處理後才建立，避免開啟設定視窗時被 QProgressDialog 自動叫出。
        self.progress_bar = None
        
        processing_layout.addWidget(toolbar)
        
        # 圖表顯示區域
        self.image_container = QtWidgets.QScrollArea(self)
        self.image_container.setObjectName("OOBResultsArea")
        self.image_container.setWidgetResizable(True)
        self.image_container.setMinimumHeight(400)
        processing_layout.addWidget(self.image_container, 1)

        self.image_grid_widget = QtWidgets.QWidget()
        self.image_grid_widget.setObjectName("TransparentWidget")
        self.image_grid_layout = QtWidgets.QGridLayout(self.image_grid_widget)
        self.image_grid_layout.setContentsMargins(18, 18, 18, 18)
        self.image_grid_layout.setSpacing(20)
        self.image_container.setWidget(self.image_grid_widget)
        self.oob_empty_state_widget = None
        self._show_oob_empty_state()
        
        # 將 Chart Processing Tab 添加到 oob_system_tabs
        self.oob_system_tabs.addTab(self.processing_tab_widget, tr("chart_processing"))


        # --- 建立第二個分頁 (Summary Dashboard) ---
        # 正確建立 summary dashboard 與其屬性
        self.setup_summary_dashboard_tab()
        self.oob_system_tabs.addTab(self.summary_tab_widget, tr("summary_dashboard_tab"))

        # 將 OOB System 整個 QTabWidget 添加到 content_stacked_widget
        self.content_stacked_widget.addWidget(self.oob_system_tabs)

        # --- 新增 Cpk Calculator 頁面內容到 stacked widget ---
        self.cpk_calculation_page = self._create_cpk_calculation_page()
        self.content_stacked_widget.addWidget(self.cpk_calculation_page)


        # 連接選單按鈕到 QStackedWidget 的頁面切換（只切換頁面，不手動 setChecked）
        self.home_button.clicked.connect(lambda: (self.content_stacked_widget.setCurrentWidget(self.home_page), self.toggle_menu_button.hide()))
        self.split_data_button.clicked.connect(lambda: (self.content_stacked_widget.setCurrentWidget(self.split_data_page), self.toggle_menu_button.hide()))
        self.oob_system_button.clicked.connect(lambda: (self.content_stacked_widget.setCurrentWidget(self.oob_system_tabs), self.toggle_menu_button.show()))
        self.cpk_calculation_button.clicked.connect(lambda: (self.content_stacked_widget.setCurrentWidget(self.cpk_calculation_page), self.toggle_menu_button.hide()))
        self.tool_matching_button.clicked.connect(lambda: (self.content_stacked_widget.setCurrentWidget(self.tool_matching_page), self.toggle_menu_button.hide()))
        self.cl_tighten_button.clicked.connect(lambda: (self.content_stacked_widget.setCurrentWidget(self.cl_tighten_page), self.toggle_menu_button.hide()))
        self.data_check_button.clicked.connect(lambda: (self.content_stacked_widget.setCurrentWidget(self.data_check_page), self.toggle_menu_button.hide()))
        # === QButtonGroup 互斥設定與預設選中 ===
        self.menu_button_group = QtWidgets.QButtonGroup(self)
        self.menu_button_group.addButton(self.home_button)
        self.menu_button_group.addButton(self.split_data_button)
        self.menu_button_group.addButton(self.oob_system_button)
        self.menu_button_group.addButton(self.cpk_calculation_button)
        self.menu_button_group.addButton(self.tool_matching_button)
        self.menu_button_group.addButton(self.cl_tighten_button)
        self.menu_button_group.setExclusive(True)
        self.home_button.setChecked(True)
        self.menu_button_group.addButton(self.data_check_button)

        self._apply_home_oob_styles()
        
        # 預設隱藏摺疊按鈕（因為首頁不需要）
        self.toggle_menu_button.hide()
    def _create_tool_matching_page(self):
        """
        建立 Tool Matching 頁面 (Widget)。
        """
        from tool_matching_widget import ToolMatchingWidget
        widget = ToolMatchingWidget(self)
        return widget

    # --- 新增輔助方法用於建立選單按鈕 ---
    def get_selected_customers(self):
        return set(self.selected_customers)

    def get_selected_chart_keys(self):
        return None if self.selected_chart_keys is None else set(self.selected_chart_keys)

    def set_selected_chart_keys(self, selected_chart_keys):
        self.selected_chart_keys = set(selected_chart_keys)

    def filter_chart_information(self, chart_info):
        return filter_chart_information(chart_info, self.selected_chart_keys)

    def ensure_charts_selected(self):
        if self.selected_chart_keys is not None and not self.selected_chart_keys:
            QtWidgets.QMessageBox.warning(
                self, tr("chart_scope", "Analysis Scope"),
                tr("no_charts_selected", "No charts are selected. Select at least one chart in Data Health Monitor."))
            return False
        return True

    def filter_customer_data(self, dataframe, source_name="data"):
        result = apply_customer_filter(dataframe, self.selected_customers)
        if result.missing_column:
            print(f"[Customer Filter] Skipped {source_name}: missing Customer column")
            if not self._customer_missing_warning_shown:
                self._customer_missing_warning_shown = True
                QtWidgets.QMessageBox.warning(
                    self, tr("customer_filter", "Customer Filter"),
                    f"{source_name}: Customer filter is active, but this CSV has no Customer column. Files like this will be skipped.")
        return result

    def _update_customer_filter_button(self):
        if not self.selected_customers:
            summary = tr("all_customers", "All Customers")
        elif len(self.selected_customers) == 1:
            summary = next(iter(self.selected_customers))
        else:
            summary = tr("customers_selected", "{count} selected").format(count=len(self.selected_customers))
        self.customer_filter_button.setText(f"Customer\n{summary}")

    def refresh_customer_filter(self, scan=True, show_messages=True):
        previous = set(self.selected_customers)
        errors = []
        if scan:
            self.available_customers, errors = discover_customers(self.raw_data_directory)
        available = set(self.available_customers)
        removed = previous - available
        if previous:
            self.selected_customers = previous & available
        self._update_customer_filter_button()
        if show_messages and removed:
            QtWidgets.QMessageBox.warning(
                self, tr("customer_filter", "Customer Filter"),
                tr("customers_removed", "Customers no longer present were removed: {names}").format(
                    names=", ".join(sorted(removed))))
        if show_messages and errors:
            QtWidgets.QMessageBox.warning(
                self, tr("customer_filter", "Customer Filter"),
                tr("customer_scan_errors", "Some CSV files could not be scanned ({count}).").format(count=len(errors)))

    def open_customer_filter_dialog(self):
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle(tr("customer_filter", "Customer Filter"))
        dialog.setMinimumSize(480, 590)
        dialog.resize(500, 640)
        dialog.setModal(True)
        dialog.setStyleSheet("""
            QDialog { background: #f6f8fc; }
            QFrame#customerHeader {
                background: #344CB7; border-radius: 14px;
            }
            QLabel#customerTitle {
                color: #ffffff; background: transparent; border: none;
                font-size: 20px; font-weight: 700; padding: 0px;
            }
            QLabel#customerSubtitle {
                color: #eef2ff; background: transparent; border: none;
                font-size: 12px; font-weight: 500; padding: 0px;
            }
            QLabel#selectionCount {
                color: #344CB7; background: #e8edff; border-radius: 10px;
                padding: 4px 10px; font-size: 12px; font-weight: 700;
            }
            QLineEdit {
                min-height: 40px; padding: 0 14px; background: white;
                color: #172554; border: 1px solid #d8deea; border-radius: 10px;
                selection-background-color: #344CB7;
            }
            QLineEdit:focus { border: 2px solid #577BC1; }
            QListWidget {
                background: white; color: #1f2937; border: 1px solid #d8deea;
                border-radius: 12px; padding: 6px; outline: none;
            }
            QListWidget::item { min-height: 36px; padding: 2px 8px; border-radius: 7px; }
            QListWidget::item:hover { background: #f0f3ff; }
            QListWidget::item:selected { background: #e8edff; color: #172554; }
            QPushButton#secondaryButton {
                min-height: 34px; padding: 0 13px; color: #344CB7;
                background: white; border: 1px solid #cbd5e1; border-radius: 8px;
                font-weight: 600;
            }
            QPushButton#secondaryButton:hover { background: #eef2ff; border-color: #818cf8; }
            QPushButton#applyButton {
                min-width: 100px; min-height: 36px; padding: 0 18px;
                color: white; background: #344CB7; border: none;
                border-radius: 9px; font-weight: 700;
            }
            QPushButton#applyButton:hover { background: #293b99; }
            QPushButton#applyButton:disabled { background: #aeb8cf; }
            QPushButton#cancelButton {
                min-width: 88px; min-height: 36px; color: #475569;
                background: transparent; border: none; font-weight: 600;
            }
            QPushButton#cancelButton:hover { background: #e9edf5; border-radius: 9px; }
        """)
        layout = QtWidgets.QVBoxLayout(dialog)
        layout.setContentsMargins(22, 22, 22, 20)
        layout.setSpacing(14)

        header = QtWidgets.QFrame()
        header.setObjectName("customerHeader")
        header_layout = QtWidgets.QVBoxLayout(header)
        header_layout.setContentsMargins(20, 17, 20, 17)
        header_layout.setSpacing(4)
        title = QtWidgets.QLabel(tr("customer_filter", "Customer Filter"))
        title.setObjectName("customerTitle")
        subtitle = QtWidgets.QLabel(tr(
            "customer_filter_subtitle",
            "Choose the customers to include across all analyses",
        ))
        subtitle.setObjectName("customerSubtitle")
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        layout.addWidget(header)

        summary_row = QtWidgets.QHBoxLayout()
        section_label = QtWidgets.QLabel(tr("customer_list", "Customer list"))
        section_label.setStyleSheet("color: #64748b; font-size: 11px; font-weight: 700;")
        count_label = QtWidgets.QLabel()
        count_label.setObjectName("selectionCount")
        summary_row.addWidget(section_label)
        summary_row.addStretch()
        summary_row.addWidget(count_label)
        layout.addLayout(summary_row)

        search = QtWidgets.QLineEdit()
        search.setPlaceholderText(tr("search_customers", "Search customers..."))
        search.setClearButtonEnabled(True)
        layout.addWidget(search)
        customer_list = QtWidgets.QListWidget()
        customer_list.setAlternatingRowColors(False)
        layout.addWidget(customer_list)

        # An empty stored selection means "all customers" throughout the app.
        initial_selection = set(self.selected_customers or self.available_customers)

        def populate(customers, checked_names):
            customer_list.blockSignals(True)
            customer_list.clear()
            for customer in customers:
                item = QtWidgets.QListWidgetItem(customer)
                item.setFlags(item.flags() | QtCore.Qt.ItemFlag.ItemIsUserCheckable)
                state = (QtCore.Qt.CheckState.Checked if customer in checked_names
                         else QtCore.Qt.CheckState.Unchecked)
                item.setCheckState(state)
                customer_list.addItem(item)
            customer_list.blockSignals(False)

        def filter_items(text):
            query = text.casefold()
            visible = 0
            for index in range(customer_list.count()):
                item = customer_list.item(index)
                matches = query in item.text().casefold()
                item.setHidden(not matches)
                visible += int(matches)
            section_label.setText(
                tr("customer_list_results").format(count=visible)
                if query else tr("customer_list")
            )

        search.textChanged.connect(filter_items)
        actions = QtWidgets.QHBoxLayout()
        actions.setSpacing(8)
        all_button = QtWidgets.QPushButton(tr("select_all", "Select All"))
        clear_button = QtWidgets.QPushButton(tr("clear", "Clear"))
        refresh_button = QtWidgets.QPushButton(tr("refresh", "Refresh"))
        for button in (all_button, clear_button, refresh_button):
            button.setObjectName("secondaryButton")
        actions.addWidget(all_button)
        actions.addWidget(clear_button)
        actions.addStretch()
        actions.addWidget(refresh_button)
        layout.addLayout(actions)

        def set_all(check_state):
            for index in range(customer_list.count()):
                customer_list.item(index).setCheckState(check_state)

        def checked_names():
            return {customer_list.item(i).text() for i in range(customer_list.count())
                    if customer_list.item(i).checkState() == QtCore.Qt.CheckState.Checked}

        def update_count():
            selected = len(checked_names())
            total = customer_list.count()
            count_label.setText(
                tr("customer_selection_count").format(selected=selected, total=total)
            )
            apply_button.setEnabled(selected > 0 or total == 0)

        def refresh_list():
            current = checked_names()
            self.available_customers, errors = discover_customers(self.raw_data_directory)
            available = set(self.available_customers)
            populate(self.available_customers, current & available)
            filter_items(search.text())
            update_count()
            if errors:
                QtWidgets.QMessageBox.warning(
                    dialog, tr("customer_filter", "Customer Filter"),
                    tr("customer_scan_errors", "Some CSV files could not be scanned ({count}).").format(count=len(errors)))

        all_button.clicked.connect(lambda: set_all(QtCore.Qt.CheckState.Checked))
        clear_button.clicked.connect(lambda: set_all(QtCore.Qt.CheckState.Unchecked))
        refresh_button.clicked.connect(refresh_list)
        customer_list.itemChanged.connect(update_count)
        buttons = QtWidgets.QDialogButtonBox()
        apply_button = buttons.addButton(
            tr("apply_filter", "Apply filter"),
            QtWidgets.QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = buttons.addButton(
            tr("cancel", "Cancel"), QtWidgets.QDialogButtonBox.ButtonRole.RejectRole)
        apply_button.setObjectName("applyButton")
        cancel_button.setObjectName("cancelButton")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        populate(self.available_customers, initial_selection)
        update_count()
        search.setFocus()
        if dialog.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            checked = checked_names()
            self.selected_customers = set() if checked == set(self.available_customers) else checked
            self._update_customer_filter_button()

    def _create_menu_button(self, text):
        button = QtWidgets.QPushButton(text)
        button.setCheckable(True) # Make button checkable for selection feedback
        button.setFont(get_app_font(14, QtGui.QFont.Weight.Bold))
        button.setMinimumHeight(44)
        button.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        button.setStyleSheet(
            "QPushButton.menu_button { background-color: transparent; color: #CBD5E1; "
            "border-radius: 10px; padding: 11px 14px; font-size: 14px; text-align: left; "
            "border: none; font-weight: 600; }"
            "QPushButton.menu_button:hover { background-color: #1F2937; color: white; }"
            "QPushButton.menu_button:checked { background-color: #312E81; color: white; "
            "border-left: 3px solid #818CF8; padding-left: 11px; }"
        )
        # Apply the class for styling
        button.setProperty("class", "menu_button")
        return button

    def _apply_home_oob_styles(self):
        """首頁與 OOB 示範頁的現代化樣式；不改動其他分析頁結構。"""
        self.setStyleSheet(self.styleSheet() + """
            QWidget#HomePage,
            QWidget#OOBProcessingPage,
            QWidget#SummaryTabWidget {
                background-color: #F6F8FC;
            }
            QWidget#TransparentWidget { background: transparent; }

            QFrame#HomeHero {
                background-color: #111827;
                border: none;
                border-radius: 24px;
            }
            QLabel#PageEyebrow {
                background: transparent;
                color: #A5B4FC;
                font-size: 11px;
                font-weight: 800;
            }
            QLabel#HomeTitle {
                background: transparent;
                color: white;
                font-size: 30px;
                font-weight: 750;
            }
            QLabel#HomeSubtitle {
                background: transparent;
                color: #CBD5E1;
                font-size: 14px;
                font-weight: 400;
            }
            QLabel#SectionTitle, QLabel#OOBPageTitle {
                background: transparent;
                color: #0F172A;
                font-size: 22px;
                font-weight: 750;
            }
            QLabel#SectionSubtitle, QLabel#OOBPageSubtitle {
                background: transparent;
                color: #64748B;
                font-size: 13px;
                font-weight: 400;
            }
            QFrame#HomeActionCard, QFrame#SummaryMetricCard {
                background-color: white;
                border: 1px solid #E2E8F0;
                border-radius: 18px;
            }
            QFrame#HomeActionCard:hover {
                border: 1px solid #A5B4FC;
                background-color: #FCFCFF;
            }
            QLabel#HomeActionTitle {
                background: transparent;
                color: #0F172A;
                font-size: 16px;
                font-weight: 700;
            }
            QLabel#HomeActionDescription {
                background: transparent;
                color: #64748B;
                font-size: 12px;
                font-weight: 400;
            }
            QPushButton#HomeActionButton {
                background: transparent;
                color: #4F46E5;
                border: none;
                padding: 4px 0;
                font-size: 12px;
                font-weight: 700;
                text-align: left;
            }
            QPushButton#HomeActionButton:hover { color: #3730A3; }

            QTabWidget#OOBSystemTabs::pane {
                border: none;
                background-color: #F6F8FC;
                top: -1px;
            }
            QTabWidget#OOBSystemTabs QTabBar::tab {
                background-color: transparent;
                color: #64748B;
                border: none;
                border-bottom: 3px solid transparent;
                padding: 13px 22px;
                margin: 0;
                font-size: 13px;
                font-weight: 650;
            }
            QTabWidget#OOBSystemTabs QTabBar::tab:hover {
                color: #312E81;
                background-color: #EEF2FF;
            }
            QTabWidget#OOBSystemTabs QTabBar::tab:selected {
                color: #3730A3;
                background-color: white;
                border-bottom: 3px solid #4F46E5;
            }
            QFrame#OOBToolbar {
                background-color: white;
                border: 1px solid #E2E8F0;
                border-radius: 16px;
            }
            QLabel#OOBReadyTitle {
                background: transparent;
                color: #0F172A;
                font-size: 14px;
                font-weight: 700;
            }
            QLabel#OOBReadyDescription {
                background: transparent;
                color: #64748B;
                font-size: 12px;
                font-weight: 400;
            }
            QScrollArea#OOBResultsArea {
                background-color: white;
                border: 1px solid #E2E8F0;
                border-radius: 16px;
            }
            QFrame#OOBResultCard {
                background-color: white;
                border: 1px solid #E2E8F0;
                border-radius: 14px;
            }
            QFrame#OOBEmptyState {
                background-color: white;
                border: 1px dashed #CBD5E1;
                border-radius: 16px;
            }
            QLabel#OOBEmptyBadge {
                background-color: #EEF2FF;
                color: #4F46E5;
                border-radius: 18px;
                font-size: 13px;
                font-weight: 800;
            }
            QLabel#OOBEmptyTitle {
                background: transparent;
                color: #0F172A;
                font-size: 18px;
                font-weight: 750;
            }
            QLabel#OOBEmptyDescription {
                background: transparent;
                color: #64748B;
                font-size: 13px;
                font-weight: 400;
            }
            QLabel#SummaryMetricValue {
                background: transparent;
                color: #0F172A;
                font-size: 14px;
                font-weight: 700;
            }
            QFrame#SummaryChartCard {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 16px;
            }
            QLabel#SummarySectionTitle {
                background: transparent;
                color: #0F172A;
                font-size: 13px;
                font-weight: 750;
                padding-top: 4px;
            }
            QTableWidget#OOBViolationTable {
                background-color: white;
                alternate-background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
                gridline-color: #E2E8F0;
                selection-background-color: #EEF2FF;
                selection-color: #1E1B4B;
                font-size: 12px;
            }
            QTableWidget#OOBViolationTable QHeaderView::section {
                background-color: #F8FAFC;
                color: #475569;
                border: none;
                border-bottom: 1px solid #E2E8F0;
                padding: 9px;
                font-size: 12px;
                font-weight: 700;
            }
        """)

    def toggle_left_menu(self):
        """摺疊/展開左側選單"""
        if self.left_menu_widget.isVisible():
            # 隱藏左側選單
            self.left_menu_widget.hide()
            self.toggle_menu_button.direction = 'right'
            self.toggle_menu_button.update()
        else:
            # 顯示左側選單
            self.left_menu_widget.show()
            self.toggle_menu_button.direction = 'left'
            self.toggle_menu_button.update()
    
    def toggle_language(self):
        """切換語言"""
        new_lang = self.translator.toggle_language()
        print(f"Language switched to: {new_lang}")
        # 更新語言按鈕文字
        self.lang_button.setText(tr("lang_button"))
    
    def refresh_ui_texts(self):
        """刷新所有UI文字（當語言切換時被調用）"""
        # 更新視窗標題
        self.setWindowTitle(tr("app_title"))
        
        # 更新左側選單按鈕
        self.home_button.setText(tr("home"))
        self.data_check_button.setText(tr("data_health_monitor"))
        self.split_data_button.setText(tr("split_data"))
        self.oob_system_button.setText(tr("oob_spc_system"))
        self.cpk_calculation_button.setText(tr("cpk_calculator"))
        self.tool_matching_button.setText(tr("tool_matching"))
        self.cl_tighten_button.setText(tr("cl_tighten"))
        self._update_customer_filter_button()

        if hasattr(self, 'home_text_bindings'):
            for widget, key, template in self.home_text_bindings:
                widget.setText(template.format(tr(key)))
        if hasattr(self, 'home_tooltip_bindings'):
            for widget, title_key, description_key in self.home_tooltip_bindings:
                widget.setAccessibleName(tr(title_key))
                widget.setToolTip(tr(description_key))

        if hasattr(self, 'oob_page_title_label'):
            self.oob_page_title_label.setText(tr("oob_page_title"))
            self.oob_page_subtitle_label.setText(tr("oob_page_subtitle"))
            self.oob_ready_title_label.setText(tr("oob_ready_title"))
            self.oob_ready_desc_label.setText(tr("oob_ready_desc"))
        if getattr(self, 'oob_empty_state_widget', None) is not None:
            self.oob_empty_title_label.setText(tr("oob_empty_title"))
            self.oob_empty_desc_label.setText(tr("oob_empty_desc"))
        
        # 更新按鈕文字
        if hasattr(self, 'settings_button'):
            self.settings_button.setText(f"⚙️ {tr('settings')}")
        
        if hasattr(self, 'start_button'):
            self.start_button.setText(f"▶ {tr('start_process')}")
        
        # 更新 OOB System 標籤頁標題
        if hasattr(self, 'oob_system_tabs'):
            self.oob_system_tabs.setTabText(0, tr("chart_processing"))
            self.oob_system_tabs.setTabText(1, tr("summary_dashboard_tab"))
        
        # 更新 Summary Dashboard 標籤
        if hasattr(self, 'summary_title_label'):
            self.summary_title_label.setText(f"<b>{tr('summary_dashboard')}</b>")
            self.summary_subtitle_label.setText(tr("summary_overview_subtitle"))
            self.violation_table_label.setText(f"<b>{tr('charts_with_anomalies_details')}</b>")
            # 更新表頭
            headers = [tr('group_name'), tr('chart_name'), tr('ooc_count'), tr('we_rules'), tr('oob_rules')]
            self.violation_table.setHorizontalHeaderLabels(headers)
            
            # 更新統計數字標籤的前綴文字（保留數字部分）
            if hasattr(self, 'total_charts_label_summary'):
                # 提取當前數字（如果存在）
                current_text = self.total_charts_label_summary.text()
                # 檢查是否包含數字
                if ':' in current_text:
                    number_part = current_text.split(':', 1)[1].strip() if ':' in current_text else current_text.split(' ')[-1]
                    self.total_charts_label_summary.setText(f"{tr('total_charts')} {number_part}")
                elif ' ' in current_text and current_text.split(' ')[-1].replace('N/A', '').replace(',', '').isdigit():
                    number_part = current_text.split(' ')[-1]
                    self.total_charts_label_summary.setText(f"{tr('total_charts')} {number_part}")
                else:
                    self.total_charts_label_summary.setText(f"{tr('total_charts')} N/A")
            
            if hasattr(self, 'processed_charts_label_summary'):
                current_text = self.processed_charts_label_summary.text()
                if ':' in current_text:
                    number_part = current_text.split(':', 1)[1].strip() if ':' in current_text else current_text.split(' ')[-1]
                    self.processed_charts_label_summary.setText(f"{tr('processed_successfully')} {number_part}")
                elif ' ' in current_text and current_text.split(' ')[-1].replace('N/A', '').replace(',', '').isdigit():
                    number_part = current_text.split(' ')[-1]
                    self.processed_charts_label_summary.setText(f"{tr('processed_successfully')} {number_part}")
                else:
                    self.processed_charts_label_summary.setText(f"{tr('processed_successfully')} N/A")
            
            if hasattr(self, 'skipped_charts_label_summary'):
                current_text = self.skipped_charts_label_summary.text()
                if ':' in current_text:
                    number_part = current_text.split(':', 1)[1].strip() if ':' in current_text else current_text.split(' ')[-1]
                    self.skipped_charts_label_summary.setText(f"{tr('no_data_charts')} {number_part}")
                elif ' ' in current_text and current_text.split(' ')[-1].replace('N/A', '').replace(',', '').isdigit():
                    number_part = current_text.split(' ')[-1]
                    self.skipped_charts_label_summary.setText(f"{tr('no_data_charts')} {number_part}")
                else:
                    self.skipped_charts_label_summary.setText(f"{tr('no_data_charts')} N/A")
            
            if hasattr(self, 'ooc_charts_label_summary'):
                current_text = self.ooc_charts_label_summary.text()
                if ':' in current_text:
                    number_part = current_text.split(':', 1)[1].strip() if ':' in current_text else current_text.split(' ')[-1]
                    self.ooc_charts_label_summary.setText(f"{tr('charts_with_ooc')} {number_part}")
                elif ' ' in current_text and current_text.split(' ')[-1].replace('N/A', '').replace(',', '').isdigit():
                    number_part = current_text.split(' ')[-1]
                    self.ooc_charts_label_summary.setText(f"{tr('charts_with_ooc')} {number_part}")
                else:
                    self.ooc_charts_label_summary.setText(f"{tr('charts_with_ooc')} N/A")
            
            if hasattr(self, 'we_count_charts_label_summary'):
                current_text = self.we_count_charts_label_summary.text()
                if ':' in current_text:
                    number_part = current_text.split(':', 1)[1].strip() if ':' in current_text else current_text.split(' ')[-1]
                    self.we_count_charts_label_summary.setText(f"{tr('charts_with_we_rule')} {number_part}")
                elif ' ' in current_text and current_text.split(' ')[-1].replace('N/A', '').replace(',', '').isdigit():
                    number_part = current_text.split(' ')[-1]
                    self.we_count_charts_label_summary.setText(f"{tr('charts_with_we_rule')} {number_part}")
                else:
                    self.we_count_charts_label_summary.setText(f"{tr('charts_with_we_rule')} N/A")
            
            if hasattr(self, 'oob_charts_label_summary'):
                current_text = self.oob_charts_label_summary.text()
                if ':' in current_text:
                    number_part = current_text.split(':', 1)[1].strip() if ':' in current_text else current_text.split(' ')[-1]
                    self.oob_charts_label_summary.setText(f"{tr('charts_with_oob')} {number_part}")
                elif ' ' in current_text and current_text.split(' ')[-1].replace('N/A', '').replace(',', '').isdigit():
                    number_part = current_text.split(' ')[-1]
                    self.oob_charts_label_summary.setText(f"{tr('charts_with_oob')} {number_part}")
                else:
                    self.oob_charts_label_summary.setText(f"{tr('charts_with_oob')} N/A")
            
            metric_translations = [
                (self.total_charts_label_summary, 'total_charts'),
                (self.processed_charts_label_summary, 'processed_successfully'),
                (self.skipped_charts_label_summary, 'no_data_charts'),
                (self.ooc_charts_label_summary, 'charts_with_ooc'),
                (self.we_count_charts_label_summary, 'charts_with_we_rule'),
                (self.oob_charts_label_summary, 'charts_with_oob'),
            ]
            for metric_label, translation_key in metric_translations:
                value = metric_label.property("summaryValue")
                self._set_summary_metric(
                    metric_label,
                    tr(translation_key),
                    None if value in (None, "N/A") else value,
                )

            # 重新繪製圖表以使用新語言
            self.refresh_summary_charts()
        
        # 更新子 Widget
        if hasattr(self, 'split_data_page') and hasattr(self.split_data_page, 'refresh_ui_texts'):
            self.split_data_page.refresh_ui_texts()
        
        if hasattr(self, 'cl_tighten_page') and hasattr(self.cl_tighten_page, 'refresh_ui_texts'):
            self.cl_tighten_page.refresh_ui_texts()
        
        if hasattr(self, 'tool_matching_page') and hasattr(self.tool_matching_page, 'refresh_ui_texts'):
            self.tool_matching_page.refresh_ui_texts()
        
        print(f"UI texts refreshed to {self.translator.current_lang}")


    # --- 新增首頁和拆分資料頁面的佔位符方法 ---
    def _create_home_page(self):
        widget = QtWidgets.QWidget()
        widget.setObjectName("HomePage")
        self.home_text_bindings = []
        self.home_tooltip_bindings = []

        page_layout = QtWidgets.QVBoxLayout(widget)
        page_layout.setContentsMargins(32, 24, 32, 24)
        page_layout.setSpacing(16)

        hero = QtWidgets.QFrame()
        hero.setObjectName("HomeHero")
        hero.setMinimumHeight(138)
        hero.setMaximumHeight(154)
        hero_layout = QtWidgets.QVBoxLayout(hero)
        hero_layout.setContentsMargins(34, 20, 34, 20)
        hero_layout.setSpacing(6)

        self.home_eyebrow_label = QtWidgets.QLabel(tr("home_eyebrow"))
        self.home_eyebrow_label.setObjectName("PageEyebrow")
        self.home_title_label = QtWidgets.QLabel(tr("home_welcome"))
        self.home_title_label.setObjectName("HomeTitle")
        self.home_subtitle_label = QtWidgets.QLabel(tr("home_subtitle"))
        self.home_subtitle_label.setObjectName("HomeSubtitle")
        self.home_subtitle_label.setWordWrap(True)
        self.home_subtitle_label.setMaximumWidth(760)
        hero_layout.addWidget(self.home_eyebrow_label)
        hero_layout.addWidget(self.home_title_label)
        hero_layout.addWidget(self.home_subtitle_label)
        page_layout.addWidget(hero)

        section_header = QtWidgets.QWidget()
        section_header.setObjectName("TransparentWidget")
        section_header_layout = QtWidgets.QVBoxLayout(section_header)
        section_header_layout.setContentsMargins(2, 0, 0, 0)
        section_header_layout.setSpacing(3)
        self.home_actions_title_label = QtWidgets.QLabel(tr("home_quick_actions"))
        self.home_actions_title_label.setObjectName("SectionTitle")
        self.home_actions_subtitle_label = QtWidgets.QLabel(tr("home_quick_actions_subtitle"))
        self.home_actions_subtitle_label.setObjectName("SectionSubtitle")
        section_header_layout.addWidget(self.home_actions_title_label)
        section_header_layout.addWidget(self.home_actions_subtitle_label)
        page_layout.addWidget(section_header)

        self.home_text_bindings.extend([
            (self.home_eyebrow_label, "home_eyebrow", "{}"),
            (self.home_title_label, "home_welcome", "{}"),
            (self.home_subtitle_label, "home_subtitle", "{}"),
            (self.home_actions_title_label, "home_quick_actions", "{}"),
            (self.home_actions_subtitle_label, "home_quick_actions_subtitle", "{}"),
        ])

        action_grid = QtWidgets.QGridLayout()
        action_grid.setHorizontalSpacing(18)
        action_grid.setVerticalSpacing(14)
        action_grid.setColumnStretch(0, 1)
        action_grid.setColumnStretch(1, 1)
        cards = [
            ("DQ", "home_health_title", "home_health_desc", self.data_check_button, "#0891B2"),
            ("DS", "home_split_title", "home_split_desc", self.split_data_button, "#D97706"),
            ("OOB", "home_oob_title", "home_oob_desc", self.oob_system_button, "#4F46E5"),
            ("CPK", "home_cpk_title", "home_cpk_desc", self.cpk_calculation_button, "#059669"),
            ("TM", "home_matching_title", "home_matching_desc", self.tool_matching_button, "#7C3AED"),
            ("CL", "home_cl_title", "home_cl_desc", self.cl_tighten_button, "#DC2626"),
        ]
        for index, card_config in enumerate(cards):
            action_grid.addWidget(
                self._create_home_action_card(*card_config), index // 2, index % 2
            )
        page_layout.addLayout(action_grid)
        page_layout.addStretch(1)

        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setObjectName("HomeScrollArea")
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(
            QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        scroll_area.setStyleSheet(
            "QScrollArea#HomeScrollArea { background: #F6F8FC; border: none; }"
            "QScrollArea#HomeScrollArea > QWidget > QWidget { background: #F6F8FC; }"
        )
        scroll_area.setWidget(widget)
        return scroll_area

    def _create_home_action_card(self, badge, title_key, description_key, target_button, accent):
        card = QtWidgets.QFrame()
        card.setObjectName("HomeActionCard")
        card.setMinimumHeight(126)
        card.setMaximumHeight(140)
        card_layout = QtWidgets.QHBoxLayout(card)
        card_layout.setContentsMargins(20, 14, 18, 14)
        card_layout.setSpacing(14)

        badge_label = QtWidgets.QLabel(badge)
        badge_label.setObjectName("HomeActionBadge")
        badge_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        badge_label.setFixedSize(46, 46)
        badge_label.setStyleSheet(
            f"background-color: {accent}; color: white; border-radius: 14px; "
            "font-size: 12px; font-weight: 800;"
        )
        card_layout.addWidget(badge_label, alignment=QtCore.Qt.AlignmentFlag.AlignTop)

        content_layout = QtWidgets.QVBoxLayout()
        content_layout.setSpacing(5)
        title_label = QtWidgets.QLabel(tr(title_key))
        title_label.setObjectName("HomeActionTitle")
        description_label = QtWidgets.QLabel(tr(description_key))
        description_label.setObjectName("HomeActionDescription")
        description_label.setWordWrap(True)
        open_button = QtWidgets.QPushButton(f"{tr('home_open')}  →")
        open_button.setObjectName("HomeActionButton")
        open_button.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        open_button.setToolTip(tr(description_key))
        open_button.setAccessibleName(tr(title_key))
        open_button.clicked.connect(target_button.click)
        card.setFocusProxy(open_button)
        content_layout.addWidget(title_label)
        content_layout.addWidget(description_label)
        content_layout.addStretch(1)
        content_layout.addWidget(open_button, alignment=QtCore.Qt.AlignmentFlag.AlignLeft)
        card_layout.addLayout(content_layout, 1)

        self.home_text_bindings.extend([
            (title_label, title_key, "{}"),
            (description_label, description_key, "{}"),
            (open_button, "home_open", "{}  →"),
        ])
        self.home_tooltip_bindings.append((open_button, title_key, description_key))
        return card
    def _create_split_data_page(self):
        """
        這個方法現在會創建並返回你的 SplitDataWidget 實例。
        """
        widget = SplitDataWidget(self) # 創建 SplitDataWidget 的實例，並將 MainWindow 作為其父物件
        return widget
    
    def _create_cpk_calculation_page(self):
        from spc_cpk_dashboard import SPCCpkDashboard
        widget = SPCCpkDashboard(self)
        return widget
    
    def _create_cl_tighten_page(self):  
        """創建 CL Tighten Calculator 頁面"""
        widget = CLTightenWidget(self)
        return widget
    
    def setup_summary_dashboard_tab(self):
        from translations import tr
        
        self.summary_tab_widget = QtWidgets.QWidget()
        self.summary_tab_widget.setObjectName("SummaryTabWidget")
        summary_layout = QtWidgets.QVBoxLayout(self.summary_tab_widget)

        # === Summary Dashboard UI 元素 ===
        self.summary_title_label = QtWidgets.QLabel(f"<b>{tr('summary_dashboard')}</b>")
        self.summary_title_label.setFont(get_app_font(16, QtGui.QFont.Weight.Bold))
        self.summary_title_label.setObjectName("OOBPageTitle")
        self.summary_title_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
        summary_layout.addWidget(self.summary_title_label)
        self.summary_subtitle_label = QtWidgets.QLabel(tr("summary_overview_subtitle"))
        self.summary_subtitle_label.setObjectName("OOBPageSubtitle")
        summary_layout.addWidget(self.summary_subtitle_label)

        # 統計數字 Label 的網格佈局
        summary_stats_layout = QtWidgets.QGridLayout()
        summary_stats_layout.setObjectName("SummaryStatsGrid")

        # Label 定義保持不變
        self.total_charts_label_summary = QtWidgets.QLabel(f"{tr('total_charts')} N/A")
        self.processed_charts_label_summary = QtWidgets.QLabel(f"{tr('processed_successfully')} N/A")
        self.skipped_charts_label_summary = QtWidgets.QLabel(f"{tr('no_data_charts')} N/A")
        self.ooc_charts_label_summary = QtWidgets.QLabel(f"{tr('charts_with_ooc')} N/A")
        self.we_count_charts_label_summary = QtWidgets.QLabel(f"{tr('charts_with_we_rule')} N/A")
        self.oob_charts_label_summary = QtWidgets.QLabel(f"{tr('charts_with_oob')} N/A")

        initial_metrics = [
            (self.total_charts_label_summary, tr('total_charts')),
            (self.processed_charts_label_summary, tr('processed_successfully')),
            (self.skipped_charts_label_summary, tr('no_data_charts')),
            (self.ooc_charts_label_summary, tr('charts_with_ooc')),
            (self.we_count_charts_label_summary, tr('charts_with_we_rule')),
            (self.oob_charts_label_summary, tr('charts_with_oob')),
        ]
        for metric_label, metric_title in initial_metrics:
            self._set_summary_metric(metric_label, metric_title, None)

        metric_items = [
            (self.total_charts_label_summary, "#4F46E5"),
            (self.processed_charts_label_summary, "#059669"),
            (self.skipped_charts_label_summary, "#64748B"),
            (self.ooc_charts_label_summary, "#DC2626"),
            (self.we_count_charts_label_summary, "#D97706"),
            (self.oob_charts_label_summary, "#7C3AED"),
        ]
        for index, (metric_label, accent) in enumerate(metric_items):
            summary_stats_layout.addWidget(
                self._create_summary_metric_card(metric_label, accent),
                index // 3,
                index % 3,
            )

        summary_stats_layout.setSpacing(15)

        summary_layout.addLayout(summary_stats_layout)

        # --- 圖表顯示區域 ---
        self.charts_main_layout = QtWidgets.QVBoxLayout()
        self.charts_horizontal_layout = QtWidgets.QHBoxLayout()
        self.charts_horizontal_layout.setSpacing(16)
        self.charts_main_layout.addLayout(self.charts_horizontal_layout)
        summary_layout.addLayout(self.charts_main_layout)

        # --- 違規圖表詳細列表區域 (保持不變) ---
        self.violation_table_label = QtWidgets.QLabel(f"<b>{tr('charts_with_anomalies_details')}</b>")
        self.violation_table_label.setObjectName("SummarySectionTitle")
        self.violation_table_label.setFont(get_app_font(12, QtGui.QFont.Weight.Bold))
        self.violation_table_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
        summary_layout.addWidget(self.violation_table_label)

        self.violation_table = QtWidgets.QTableWidget()
        self.violation_table.setObjectName("OOBViolationTable")
        self.violation_table.setAlternatingRowColors(True)
        self.violation_table.setColumnCount(5)
        headers = [tr('group_name'), tr('chart_name'), tr('ooc_count'), tr('we_rules'), tr('oob_rules')]
        self.violation_table.setHorizontalHeaderLabels(headers)
        self.violation_table.horizontalHeader().setStretchLastSection(True)
        self.violation_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.violation_table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.violation_table.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.violation_table.verticalHeader().setVisible(False)
        self.violation_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.violation_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.violation_table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.violation_table.setMinimumHeight(150)
        summary_layout.addWidget(self.violation_table)

        summary_layout.addStretch()

        summary_layout.setContentsMargins(28, 24, 28, 28)
        summary_layout.setSpacing(14)
        
        # 注意: setup_summary_dashboard_tab 不再呼叫 self.addTab，
        # 因為它會被 oob_system_tabs 呼叫
    # --- 繪製圖表的輔助方法 ---

    def _set_summary_metric(self, label, title, value):
        """Render a compact metric title with a prominent value."""
        clean_title = str(title).rstrip("：:").strip()
        display_value = "N/A" if value is None else str(value)
        label.setProperty("summaryValue", display_value)
        label.setText(
            f"<span style='font-size:12px; color:#64748B; font-weight:600;'>"
            f"{clean_title}</span><br>"
            f"<span style='font-size:24px; color:#0F172A; font-weight:750;'>"
            f"{display_value}</span>"
        )

    def _create_summary_metric_card(self, metric_label, accent):
        card = QtWidgets.QFrame()
        card.setObjectName("SummaryMetricCard")
        card.setMinimumHeight(76)
        card_layout = QtWidgets.QHBoxLayout(card)
        card_layout.setContentsMargins(14, 12, 16, 12)
        card_layout.setSpacing(12)

        accent_bar = QtWidgets.QFrame()
        accent_bar.setFixedWidth(4)
        accent_bar.setStyleSheet(f"background-color: {accent}; border-radius: 2px;")
        metric_label.setObjectName("SummaryMetricValue")
        metric_label.setWordWrap(True)
        card_layout.addWidget(accent_bar)
        card_layout.addWidget(metric_label, 1)
        shadow = QtWidgets.QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(16)
        shadow.setOffset(0, 3)
        shadow.setColor(QtGui.QColor(15, 23, 42, 20))
        card.setGraphicsEffect(shadow)
        return card

    def _add_summary_chart(self, canvas):
        """Place a summary chart in a consistent responsive card."""
        card = QtWidgets.QFrame()
        card.setObjectName("SummaryChartCard")
        card.setMinimumHeight(285)
        card.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(10, 10, 10, 8)
        canvas.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        layout.addWidget(canvas)
        shadow = QtWidgets.QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(18)
        shadow.setOffset(0, 4)
        shadow.setColor(QtGui.QColor(15, 23, 42, 22))
        card.setGraphicsEffect(shadow)
        self.charts_horizontal_layout.addWidget(card, 1)
        return card

    def create_status_pie_chart(self, processed, skipped):
        from translations import tr
        fig = Figure(figsize=(4.2, 3.2), facecolor="#FFFFFF")
        ax = fig.add_subplot(111)
        labels = [tr('processed'), tr('no_data')]
        sizes = [processed, skipped]
        if sum(sizes) <= 0:
            sizes = [0, 1]

        wedges, _ = ax.pie(
            sizes,
            colors=['#4F46E5', '#CBD5E1'],
            startangle=90,
            counterclock=False,
            wedgeprops={'width': 0.24, 'edgecolor': '#FFFFFF', 'linewidth': 3},
        )
        ax.axis('equal')
        ax.text(0, 0.08, str(processed), ha='center', va='center',
                fontsize=22, fontweight='bold', color='#0F172A')
        ax.text(0, -0.16, tr('processed'), ha='center', va='center',
                fontsize=9, color='#64748B')
        ax.set_title(tr('overall_processing_status'), loc='left', fontsize=12,
                     fontweight='bold', color='#0F172A', pad=12)
        ax.legend(wedges, labels, loc='lower center', bbox_to_anchor=(0.5, -0.10),
                  ncol=2, frameon=False, fontsize=9, handlelength=1.0)
        fig.subplots_adjust(left=0.08, right=0.92, top=0.84, bottom=0.18)
        return FigureCanvas(fig)
    def create_processed_violation_pie_chart(self, processed_count, violating_count):
        from translations import tr
        
        fig = Figure(figsize=(4.2, 3.2), facecolor="#FFFFFF")
        ax = fig.add_subplot(111)

        # 計算未違規的已處理圖表數量
        non_violating_count = max(0, processed_count - violating_count)

        labels = [tr('violating'), tr('normal')]
        sizes = [violating_count, non_violating_count]
        colors = ['#EF4444', '#10B981']

        # 如果沒有成功處理的圖表，或者所有都已處理但都未違規
        if processed_count == 0 or (processed_count > 0 and violating_count == 0):
             labels = [tr('all_normal')]
             sizes = [processed_count if processed_count > 0 else 1] # 如果 processed_count=0，給個非零大小繪圖
             colors = ['#10B981']
             if processed_count == 0: # 如果 processed_count=0，顯示 N/A 或無數據
                  labels = ['N/A']
                  sizes = [1]
                  colors = ['#CBD5E1']


        # 甜甜圈圖設定
        wedgeprops = {'width': 0.24, 'edgecolor': '#FFFFFF', 'linewidth': 3}

        wedges, _ = ax.pie(
            sizes, colors=colors,
            startangle=90, counterclock=False, wedgeprops=wedgeprops,
        )
        ax.axis('equal')
        violation_rate = (violating_count / processed_count * 100) if processed_count else 0
        ax.text(0, 0.08, f'{violation_rate:.0f}%', ha='center', va='center',
                fontsize=22, fontweight='bold', color='#0F172A')
        ax.text(0, -0.16, tr('violating'), ha='center', va='center',
                fontsize=9, color='#64748B')
        ax.set_title(tr('violation_rate'), loc='left', fontsize=12,
                     fontweight='bold', color='#0F172A', pad=12)
        ax.legend(wedges, labels, loc='lower center', bbox_to_anchor=(0.5, -0.10),
                  ncol=2, frameon=False, fontsize=9, handlelength=1.0)

        # 確保圖表邊界有足夠空間
        fig.subplots_adjust(left=0.08, right=0.92, top=0.84, bottom=0.18)
        return FigureCanvas(fig)
    def create_anomaly_bar_chart(self, ooc_count, we_count, oob_count):
        from translations import tr
        
        fig = Figure(figsize=(4.8, 3.2), facecolor="#FFFFFF")
        ax = fig.add_subplot(111)

        categories = [tr('ooc'), tr('we_rule'), tr('oob')]
        counts = [ooc_count, we_count, oob_count]
        colors = ['#EF4444', '#F59E0B', '#8B5CF6']

        bars = ax.bar(categories, counts, color=colors, width=0.56, zorder=3)

        for bar in bars:
            yval = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2.0, yval, int(yval),
                    va='bottom', ha='center', fontsize=9, fontweight='bold',
                    color='#334155')

        ax.set_ylabel(tr('number_of_charts'), fontsize=9, color='#64748B')
        ax.set_title(tr('charts_with_anomalies'), loc='left', fontsize=12,
                     fontweight='bold', color='#0F172A', pad=12)
        ax.set_ylim(0, max(counts) * 1.3 or 1)
        ax.grid(axis='y', color='#E2E8F0', linewidth=0.8, zorder=0)
        ax.tick_params(axis='both', colors='#64748B', labelsize=9, length=0)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['bottom'].set_color('#CBD5E1')
        ax.set_axisbelow(True)

        fig.tight_layout(pad=1.6)

        canvas = FigureCanvas(fig)
        return canvas
    # --- 清理 Summary Tab 圖表和表格的方法 ---
    def clear_summary_charts(self):
        print("Clearing summary charts and table...")
        # 清理 charts_horizontal_layout 中的所有項目
        if self.charts_horizontal_layout is not None:
             while self.charts_horizontal_layout.count():
                 item = self.charts_horizontal_layout.takeAt(0)
                 if item.widget():
                     item.widget().deleteLater()
                 elif item.spacerItem():
                      self.charts_horizontal_layout.removeItem(item.spacerItem())

        # 將 Canvas 屬性設為 None，雖然 deleteLater() 已經標記刪除，但好習慣是解除引用
        self.status_pie_canvas = None
        self.processed_violation_pie_canvas = None # <--- 清理第二個圓餅圖 Canvas
        self.anomaly_bar_canvas = None


        # 清理表格內容
        if self.violation_table:
             self.violation_table.setRowCount(0) # 將行數設為 0 清空表格
        print("Summary charts and table cleared.")
    
    def refresh_summary_charts(self):
        """重新繪製 Summary Dashboard 的圖表以使用新語言"""
        from translations import tr
        
        # 如果沒有結果數據，則不需要重新繪製
        if not hasattr(self, 'results') or not self.results:
            return
        
        # 提取當前的統計數據
        try:
            # 從標籤文字中提取數字
            total_text = self.total_charts_label_summary.text()
            processed_text = self.processed_charts_label_summary.text()
            skipped_text = self.skipped_charts_label_summary.text()
            ooc_text = self.ooc_charts_label_summary.text()
            we_text = self.we_count_charts_label_summary.text()
            oob_text = self.oob_charts_label_summary.text()
            
            # 提取數字部分
            def extract_number(text):
                parts = text.split()
                if parts:
                    last_part = parts[-1]
                    if last_part.isdigit():
                        return int(last_part)
                    elif last_part.replace(',', '').isdigit():
                        return int(last_part.replace(',', ''))
                return 0
            
            def metric_value(label, fallback_text):
                value = label.property("summaryValue")
                try:
                    return int(value)
                except (TypeError, ValueError):
                    return extract_number(fallback_text)

            total = metric_value(self.total_charts_label_summary, total_text)
            processed = metric_value(self.processed_charts_label_summary, processed_text)
            skipped = metric_value(self.skipped_charts_label_summary, skipped_text)
            ooc_count = metric_value(self.ooc_charts_label_summary, ooc_text)
            we_count = metric_value(self.we_count_charts_label_summary, we_text)
            oob_count = metric_value(self.oob_charts_label_summary, oob_text)
            
            # 清理舊圖表
            self.clear_summary_charts()
            
            # 重新繪製圖表
            if total > 0:
                self.status_pie_canvas = self.create_status_pie_chart(processed, skipped)
                self._add_summary_chart(self.status_pie_canvas)
            
            if processed > 0:
                # 計算違規圖表數量
                violating_count = 0
                for result in self.results:
                    has_ooc = result.get('ooc_cnt', 0) > 0
                    has_we = result.get('WE_Rule', '') and result.get('WE_Rule', '') != 'N/A'
                    has_oob = result.get('OOB_Rule', '') and result.get('OOB_Rule', '') != 'N/A'
                    if has_ooc or has_we or has_oob:
                        violating_count += 1
                
                self.processed_violation_pie_canvas = self.create_processed_violation_pie_chart(processed, violating_count)
                self._add_summary_chart(self.processed_violation_pie_canvas)
            
            if ooc_count > 0 or we_count > 0 or oob_count > 0:
                self.anomaly_bar_canvas = self.create_anomaly_bar_chart(ooc_count, we_count, oob_count)
                self._add_summary_chart(self.anomaly_bar_canvas)
            
            print("Summary charts refreshed with new language.")
        except Exception as e:
            print(f"Error refreshing summary charts: {e}")

    def update_summary_dashboard(self, total, processed, skipped):
        from translations import tr
        
        print("\nUpdating Summary Dashboard...")
        ooc_count = 0
        we_count = 0
        oob_count = 0

        violating_charts = []

        for result in self.results: # self.results 已經是成功處理的圖表結果列表
            has_ooc = result.get('ooc_cnt', 0) > 0
            has_we = result.get('WE_Rule', '') and result.get('WE_Rule', '') != 'N/A'
            has_oob = result.get('OOB_Rule', '') and result.get('OOB_Rule', '') != 'N/A'

            if has_ooc:
                 ooc_count += 1
            if has_we:
                 we_count += 1
            if has_oob:
                 oob_count += 1

            if has_ooc or has_we or has_oob:
                violating_charts.append(result)


        logger.debug("Calculated OOC chart count: %d", ooc_count)
        logger.debug("Calculated WE_Rule chart count: %d", we_count)
        logger.debug("Calculated OOB chart count: %d", oob_count)
        logger.debug("Number of violating charts: %d", len(violating_charts))


        # 更新統計數字 Label 文本
        self._set_summary_metric(self.total_charts_label_summary, tr('total_charts'), total)
        self._set_summary_metric(self.processed_charts_label_summary, tr('processed_successfully'), processed)
        self._set_summary_metric(self.skipped_charts_label_summary, tr('no_data_charts'), skipped)
        self._set_summary_metric(self.ooc_charts_label_summary, tr('charts_with_ooc'), ooc_count)
        self._set_summary_metric(self.we_count_charts_label_summary, tr('charts_with_we_rule'), we_count)
        self._set_summary_metric(self.oob_charts_label_summary, tr('charts_with_oob'), oob_count)


        # --- 清理舊圖表並添加新圖表 ---
        self.clear_summary_charts() # 清理 Summary Tab 中的舊圖表和表格

        if total > 0:
             self.status_pie_canvas = self.create_status_pie_chart(processed, skipped)
             self._add_summary_chart(self.status_pie_canvas)


        # 添加成功處理圖表違規比例甜甜圈圖 (中間圖)
        # 這個圖只需要在有成功處理的圖表時顯示
        if processed > 0:
             # violating_charts 列表已經是從 self.results (已處理圖表) 中篩選的
             violating_count_in_processed = len(violating_charts)
             self.processed_violation_pie_canvas = self.create_processed_violation_pie_chart(processed, violating_count_in_processed)
             self._add_summary_chart(self.processed_violation_pie_canvas)

        if ooc_count > 0 or we_count > 0 or oob_count > 0:
             self.anomaly_bar_canvas = self.create_anomaly_bar_chart(ooc_count, we_count, oob_count)
             self._add_summary_chart(self.anomaly_bar_canvas)

        self.violation_table.setRowCount(len(violating_charts))


        for row_index, result in enumerate(violating_charts):
            # === 修改點：強制轉換為字符串並處理數字型名稱 ===
            group_name_raw = result.get('group_name', 'N/A')
            chart_name_raw = result.get('chart_name', 'N/A')
            
            # 確保轉換為字符串，即使原始值是數字、浮點數或其他類型
            if group_name_raw is None or group_name_raw == '':
                group_name = 'N/A'
            elif str(group_name_raw) == 'Default':
                group_name = ''  # 將 Default 顯示為空白
            else:
                group_name = str(group_name_raw)  # 強制轉換為字符串
                
            if chart_name_raw is None or chart_name_raw == '':
                chart_name = 'N/A'
            else:
                chart_name = str(chart_name_raw)  # 強制轉換為字符串
            
            print(f"DEBUG: 表格第 {row_index} 行 - 原始值: group_name={group_name_raw} (類型:{type(group_name_raw)}), chart_name={chart_name_raw} (類型:{type(chart_name_raw)})")
            print(f"DEBUG: 表格第 {row_index} 行 - 轉換後: group_name='{group_name}', chart_name='{chart_name}'")

            ooc_cnt = result.get('ooc_cnt', 0)
            we_rules = result.get('WE_Rule', 'N/A')
            oob_rules = result.get('OOB_Rule', 'N/A')

            # 在表格中插入新的一行
            self.violation_table.insertRow(row_index)

            # 創建 QTableWidgetItem 並設定 alignment
            item_group_name = QtWidgets.QTableWidgetItem(group_name)
            item_chart_name = QtWidgets.QTableWidgetItem(chart_name)
            item_ooc_cnt = QtWidgets.QTableWidgetItem(str(ooc_cnt))
            item_we_rules = QtWidgets.QTableWidgetItem(str(we_rules) if we_rules is not None else 'N/A')
            item_oob_rules = QtWidgets.QTableWidgetItem(str(oob_rules) if oob_rules is not None else 'N/A')

            # 將所有欄位內容都置中顯示
            for item in [item_group_name, item_chart_name, item_ooc_cnt, item_we_rules, item_oob_rules]:
                item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

            # 將項目設定到正確的欄位
            self.violation_table.setItem(row_index, 0, item_group_name)
            self.violation_table.setItem(row_index, 1, item_chart_name)
            self.violation_table.setItem(row_index, 2, item_ooc_cnt)
            self.violation_table.setItem(row_index, 3, item_we_rules)
            self.violation_table.setItem(row_index, 4, item_oob_rules)

        print("Summary Dashboard updated.")

    # --- UI部件 (需要確保這些方法在類別定義內) ---
    def create_start_button(self):
        button = QtWidgets.QPushButton("Start Processing", self)
        button.setFont(get_app_font(14))
        button.clicked.connect(self.process_charts)
        return button

    def create_progress_bar(self):
        progress_bar = QtWidgets.QProgressBar(self)
        progress_bar.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        return progress_bar

    def _open_oob_progress_dialog(self):
        """Create the OOB progress UI only for an explicitly started analysis."""
        self._close_oob_progress_dialog()
        self.start_button.setEnabled(False)
        self.progress_bar = OOBProgressDialog(self)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%")
        self.progress_bar.show()
        QtWidgets.QApplication.processEvents()

    def _close_oob_progress_dialog(self):
        dialog = getattr(self, 'progress_bar', None)
        if dialog is not None:
            dialog.hide()
            dialog.deleteLater()
            self.progress_bar = None
        if hasattr(self, 'start_button'):
            self.start_button.setEnabled(True)

    def _check_oob_cancelled(self):
        dialog = getattr(self, 'progress_bar', None)
        if dialog is not None and dialog.wasCanceled():
            raise OOBAnalysisCancelled()

    def create_image_label(self, image_path: str, max_width=450, max_height=350, keep_original_size=False):
        try:
            image = Image.open(image_path)
            qt_image = ImageQt(image)
            pixmap = QtGui.QPixmap.fromImage(qt_image)
        except Exception as e:
            print(f"Error loading image {image_path}: {e}")
            label = QtWidgets.QLabel("Image Not Found", self)
            label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            label.setStyleSheet("color: red;")
            return label

        label = QtWidgets.QLabel(self)
        label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        if keep_original_size:
            label.setPixmap(pixmap)
            label.setMaximumSize(pixmap.width(), pixmap.height())
            label.setMinimumSize(pixmap.width(), pixmap.height())
        else:
            scaled_pixmap = pixmap.scaled(
                max_width,
                max_height,
                QtCore.Qt.AspectRatioMode.KeepAspectRatio,
                QtCore.Qt.TransformationMode.SmoothTransformation
            )
            label.setPixmap(scaled_pixmap)
            label.setMaximumSize(max_width, max_height)

        return label

    # def create_canvas_widget(self, canvas, max_width=650, max_height=450):
    #     """
    #     將 FigureCanvas 轉換為適合顯示在 UI 中的 widget
    #     """
    #     try:
    #         # 不再使用 Expanding，改用固定大小限制
    #         # canvas.setMinimumSize(300, 200)
    #         canvas.setMaximumSize(max_width, max_height)

    #         # 創建一個容器 widget
    #         container = QtWidgets.QWidget(self)
    #         layout = QtWidgets.QVBoxLayout(container)
    #         layout.addWidget(canvas)
    #         layout.setContentsMargins(0, 0, 0, 0)

    #         # 可選：容器也限制大小
    #         container.setMinimumSize(300, 200)
    #         container.setMaximumSize(max_width, max_height)

    #         return container
    #     except Exception as e:
    #         print(f"Error creating canvas widget: {e}")
    #         label = QtWidgets.QLabel("Canvas Error", self)
    #         label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
    #         label.setStyleSheet("color: red;")
    #         label.setMinimumSize(300, 200)
    #         return label


    def get_cached_csv(self, filepath):
        """使用快取讀取 CSV 檔案，提升性能"""
        try:
            if filepath not in self.csv_cache:
                print(f"  - 讀取並快取 CSV: {os.path.basename(filepath)}")
                df = pd.read_csv(filepath)
                self.csv_cache[filepath] = df
            else:
                print(f"  - 使用快取的 CSV: {os.path.basename(filepath)}")
            
            # 返回副本避免修改快取的資料
            return self.csv_cache[filepath]
        except Exception as e:
            print(f"[Error] 讀取檔案 {filepath} 失敗: {str(e)}")
            return None

    def process_charts(self):
        import time
        self._customer_missing_warning_shown = False
        self.results = []
        total_charts_count = 0
        skipped_charts_count = 0
        processed_charts_count = 0

        try:
            if not self.ensure_charts_selected():
                return
            self.validate_files_and_directories()
            self._open_oob_progress_dialog()
            self.clear_image_grid()
            self.image_grid_widget.setUpdatesEnabled(False)

            all_charts_info = load_chart_information(self.filepath)
            all_charts_info = self.filter_chart_information(all_charts_info)
            total_charts_count = len(all_charts_info)
            self.progress_bar.setMaximum(100)

            # 每次執行重建檔案索引與 CSV 快取。
            self.csv_cache.clear()
            raw_file_map = self._build_raw_file_map(all_charts_info)

            # 性能優化：預處理所有圖表的數據類型
            print("=== 性能優化：開始預處理數據類型 ===")
            self.preprocess_chart_types(all_charts_info, raw_file_map)

            print("=== 預處理完成，開始處理圖表 ===")

            # if self.display_gui_checkbox.isChecked():
            #     self.add_column_headers()

            execution_time = load_execution_time(self.filepath)
            
            # 從設定中獲取自定義時間範圍
            custom_weekly_start = None
            custom_weekly_end = None
            if self.oob_settings.get('custom_time_range_enabled', False):
                # 將 QDate 或 QDateTime 轉換為 pandas datetime
                qt_start = self.oob_settings['start_time']
                qt_end = self.oob_settings['end_time']
                
                # 支持 QDate 和 QDateTime 兼容
                if isinstance(qt_start, QtCore.QDateTime):
                    custom_weekly_start = pd.to_datetime(f"{qt_start.date().year()}-{qt_start.date().month():02d}-{qt_start.date().day():02d} 00:00:00")
                else:  # QDate
                    custom_weekly_start = pd.to_datetime(f"{qt_start.year()}-{qt_start.month():02d}-{qt_start.day():02d} 00:00:00")
                
                if isinstance(qt_end, QtCore.QDateTime):
                    custom_weekly_end = pd.to_datetime(f"{qt_end.date().year()}-{qt_end.date().month():02d}-{qt_end.date().day():02d} 23:59:59")
                else:  # QDate
                    custom_weekly_end = pd.to_datetime(f"{qt_end.year()}-{qt_end.month():02d}-{qt_end.day():02d} 23:59:59")
                
                print(f" - 自定義時間範圍: {custom_weekly_start} to {custom_weekly_end}")

            for i, (_, chart_info) in enumerate(all_charts_info.iterrows()):
                self._check_oob_cancelled()
                group_name = str(chart_info['GroupName'])
                chart_name = str(chart_info['ChartName'])
                chart_key = f"{group_name}_{chart_name}"
                print(f"\n正在處理圖表: GroupName={group_name}, ChartName={chart_name}")

                try:
                    filepath = raw_file_map.get(chart_key)
                    
                    if filepath and os.path.exists(filepath):
                        # 性能優化：使用快取讀取 CSV
                        raw_df = self.get_cached_csv(filepath)
                        
                        if raw_df is not None:
                            customer_result = self.filter_customer_data(raw_df, os.path.basename(filepath))
                            if customer_result.missing_column or customer_result.data.empty:
                                skipped_charts_count += 1
                                print(f"[Customer Filter] No selected Customer data: {group_name}/{chart_name}")
                                continue
                            raw_df = customer_result.data
                            print(f" - 原始資料 shape: {raw_df.shape}")

                            # 性能優化：使用預處理的數據類型
                            data_type = self.chart_types_cache.get(chart_key, 'continuous')
                            chart_info = chart_info.copy()  # 避免修改原始數據
                            chart_info['data_type'] = data_type
                            print(f" - 使用快取的數據類型: {data_type}")

                            if 'point_time' in raw_df.columns:
                                raw_df['point_time'] = pd.to_datetime(raw_df['point_time'], errors='coerce')
                                raw_df.dropna(subset=['point_time'], inplace=True)

                            is_successful, processed_df, updated_chart_info = preprocess_data(chart_info, raw_df)

                            if not is_successful or processed_df is None or processed_df.empty:
                                print(f"[Info] 圖表 {group_name}/{chart_name} 預處理失敗或資料為空，跳過。")
                                skipped_charts_count += 1
                            else:
                                print(f" - 預處理後資料 shape: {processed_df.shape}")
                                print(f" - 準備分析圖表: {group_name}/{chart_name}")

                                # 性能優化：減少假進度條的步數，降低 GUI 更新頻率
                                fake_steps = 5  # 從 10 減少到 5
                                for fake_step in range(1):
                                    # 調整進度計算：20%已用於預處理，剩餘80%用於圖表處理
                                    progress_base = 20 + int(((i / total_charts_count) * 80))
                                    progress_step = int((fake_step / fake_steps) * (80 / total_charts_count))
                                    percent = min(100, progress_base + progress_step)
                                    self.progress_bar.setValue(percent)
                                    self.progress_bar.setFormat(f"{percent}% - {tr('processing')} {group_name}/{chart_name}")
                                    if fake_step % 2 == 0:  # 只在偶數步驟更新 GUI
                                        QtWidgets.QApplication.processEvents()
                                        self._check_oob_cancelled()
                                    time.sleep(0.005)  # 從 0.01 減少到 0.005

                                # 從設定中檢查是否使用互動式圖表和 Batch_ID 標籤
                                use_interactive = self.oob_settings.get('use_interactive_charts', True)
                                use_batch_id = self.oob_settings.get('use_batch_id_labels', False)
                                result = self.analyze_chart(execution_time, processed_df, updated_chart_info, use_interactive, use_batch_id, custom_weekly_start, custom_weekly_end)

                                if result:
                                    self.results.append(result)
                                    processed_charts_count += 1

                                    if self.oob_settings.get('show_charts_gui', True):
                                        # 檢查是否有互動圖表或靜態圖表
                                        has_interactive = 'spc_canvas' in result and 'weekly_canvas' in result
                                        has_static = 'chart_path' in result and 'weekly_chart_path' in result
                                        
                                        if has_interactive or has_static:
                                            self.display_image(result, len(self.results) - 1)
                                            chart_type = "互動式" if has_interactive else "靜態"
                                            print(f" - 顯示{chart_type}圖表完成: {group_name}/{chart_name}")
                                        else:
                                            print(f"[Warning] 圖表 {group_name}/{chart_name} 缺少圖表資料，無法顯示。")
                                    else:
                                        print(f" - GUI 顯示已禁用，跳過顯示圖表: {group_name}/{chart_name}")
                                else:
                                    print(f"[Info] 圖表 {group_name}/{chart_name} 分析返回 None，跳過結果記錄。")
                                    skipped_charts_count += 1
                        else:
                            print(f"[Error] 無法讀取檔案: {filepath}")
                            skipped_charts_count += 1
                    else:
                        print(f"[Info] 圖表 {group_name}/{chart_name} 對應檔案 {filepath} 不存在，跳過處理。")
                        skipped_charts_count += 1

                except OOBAnalysisCancelled:
                    raise
                except FileNotFoundError:
                    print(f"[Warning] 檔案未找到，跳過圖表: {group_name}/{chart_name}")
                    skipped_charts_count += 1
                except Exception as e:
                    print(f"[Error] 處理圖表 {group_name}/{chart_name} 時發生錯誤: {str(e)}")
                    traceback.print_exc()
                    skipped_charts_count += 1

                # 性能優化：減少 GUI 更新頻率
                if i % 2 == 0:  # 每3個圖表更新一次進度條
                    # 調整進度計算：20%已用於預處理，剩餘80%用於圖表處理
                    percent = min(100, 20 + int(((i + 1) / total_charts_count) * 80))
                    self.progress_bar.setValue(percent)
                    self.progress_bar.setFormat(f"{percent}% - {processed_charts_count}/{total_charts_count} {tr('processed')}")
                    QtWidgets.QApplication.processEvents()
                    self._check_oob_cancelled()

            # 圖表處理完成後仍需整理摘要與寫入 Excel。
            self.image_grid_widget.setUpdatesEnabled(True)
            self.image_grid_widget.update()
            self.progress_bar.setValue(98)
            self.progress_bar.setFormat(f"98% - {tr('finalizing_results')}")
            QtWidgets.QApplication.processEvents()

            self.update_summary_dashboard(total_charts_count, processed_charts_count, skipped_charts_count)

            if self.results:
                self.save_results()
                self.progress_bar.setValue(100)
                self.progress_bar.setFormat(f"100% - {tr('complete')}!")
                QtWidgets.QApplication.processEvents()
                self._close_oob_progress_dialog()
                QtWidgets.QMessageBox.information(self, "Processing Complete", "Results have been saved to result_with_images.xlsx")
            else:
                self.progress_bar.setValue(100)
                self.progress_bar.setFormat(f"100% - {tr('complete')}!")
                QtWidgets.QApplication.processEvents()
                self._close_oob_progress_dialog()
                QtWidgets.QMessageBox.information(self, "Processing Complete", "No charts were processed successfully to save.")

            # 清理快取（可選）
            print(f"處理完成，清理快取。CSV 快取大小: {len(self.csv_cache)}")
            # 如果記憶體有限，可以清空快取
            # self.csv_cache.clear()

        except OOBAnalysisCancelled:
            self.results = []
            self.clear_image_grid()
            self._close_oob_progress_dialog()
            QtWidgets.QMessageBox.information(
                self,
                tr("processing_cancelled_title"),
                tr("processing_cancelled_message"),
            )
        except FileNotFoundError as e:
            self._close_oob_progress_dialog()
            self.show_error("File Error", str(e))
        except NotADirectoryError as e:
            self._close_oob_progress_dialog()
            self.show_error("Directory Error", str(e))
        except Exception as e:
            self._close_oob_progress_dialog()
            self.show_error("Processing Error", str(e))
            traceback.print_exc()
        finally:
            if hasattr(self, 'image_grid_widget'):
                self.image_grid_widget.setUpdatesEnabled(True)

    # --- 新增清理 Grid Layout 的方法 (針對第一個分頁) ---
    def _show_oob_empty_state(self):
        if self.oob_empty_state_widget is not None:
            return
        empty_card = QtWidgets.QFrame()
        empty_card.setObjectName("OOBEmptyState")
        empty_card.setMinimumHeight(300)
        empty_layout = QtWidgets.QVBoxLayout(empty_card)
        empty_layout.setContentsMargins(36, 36, 36, 36)
        empty_layout.setSpacing(8)
        empty_layout.addStretch()

        badge = QtWidgets.QLabel("OOB")
        badge.setObjectName("OOBEmptyBadge")
        badge.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        badge.setFixedSize(58, 58)
        self.oob_empty_title_label = QtWidgets.QLabel(tr("oob_empty_title"))
        self.oob_empty_title_label.setObjectName("OOBEmptyTitle")
        self.oob_empty_title_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.oob_empty_desc_label = QtWidgets.QLabel(tr("oob_empty_desc"))
        self.oob_empty_desc_label.setObjectName("OOBEmptyDescription")
        self.oob_empty_desc_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.oob_empty_desc_label.setWordWrap(True)
        self.oob_empty_desc_label.setMinimumSize(420, 48)
        self.oob_empty_desc_label.setMaximumWidth(620)
        self.oob_empty_desc_label.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.MinimumExpanding,
        )

        empty_layout.addWidget(badge, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)
        empty_layout.addWidget(self.oob_empty_title_label)
        empty_layout.addWidget(self.oob_empty_desc_label, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)
        empty_layout.addStretch()
        self.oob_empty_state_widget = empty_card
        self.image_grid_layout.addWidget(empty_card, 0, 0, 1, 3)

    def _remove_oob_empty_state(self):
        if self.oob_empty_state_widget is None:
            return
        self.image_grid_layout.removeWidget(self.oob_empty_state_widget)
        self.oob_empty_state_widget.deleteLater()
        self.oob_empty_state_widget = None

    def clear_image_grid(self):
        print("Clearing image grid...")
        if self.header_container and self.processing_tab_widget.layout() and self.processing_tab_widget.layout().indexOf(self.header_container) != -1:
             while self.header_container.count():
                 item = self.header_container.takeAt(0)
                 if item.widget():
                     item.widget().deleteLater()
             if self.processing_tab_widget.layout():
                  self.processing_tab_widget.layout().removeItem(self.header_container)
             self.header_container = None

        if self.image_grid_layout:
            while self.image_grid_layout.count():
                item = self.image_grid_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
                elif item.layout():
                     self.clear_layout(item.layout())
        self.oob_empty_state_widget = None
        self._show_oob_empty_state()
        print("Image grid cleared.")

    def clear_layout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
                elif item.layout():
                    self.clear_layout(item.layout())

    def validate_files_and_directories(self):
        if not os.path.isdir(self.raw_data_directory):
            print(f"Creating directory: {self.raw_data_directory}")
            os.makedirs(self.raw_data_directory, exist_ok=True)

        if not os.path.exists(self.filepath):
            print(f"[Error] 缺少必要的 All_Chart_Information.xlsx 檔案於 {self.filepath}，請先準備好檔案再執行。")
            raise FileNotFoundError(f"{self.filepath} does not exist. Please provide the required Excel file.")

        print("Files and directories validated.")

    def add_column_headers(self):
        if self.header_container and self.processing_tab_widget.layout() and self.processing_tab_widget.layout().indexOf(self.header_container) != -1:
            return

        self.header_container = QtWidgets.QHBoxLayout()
        for header in HEADERS:
            label = QtWidgets.QLabel(f"<b>{header}</b>", self)
            label.setFont(get_app_font(12))
            label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            self.header_container.addWidget(label)

        if self.processing_tab_widget.layout():
             self.processing_tab_widget.layout().insertLayout(4, self.header_container)
        else:
             print("[Warning] Processing tab layout is not initialized when trying to add headers.")

    def update_filter_values(self, filter_type):
        self.filter_value_combo.clear()
        if filter_type == "FabPhase":
            self.filter_value_combo.addItems(["T14P5", "T14P6", "T12P4"])
        elif filter_type == "Production Line":
            self.filter_value_combo.addItems(["7", "11", "5", "8", "6", "9", "4", "10"])

    def _build_raw_file_map(self, all_charts_info):
        """Resolve all chart CSV paths from one directory listing."""
        filenames = os.listdir(self.raw_data_directory)
        file_map = {}
        for chart_info in all_charts_info.itertuples(index=False):
            group_name = str(getattr(chart_info, 'GroupName'))
            chart_name = str(getattr(chart_info, 'ChartName'))
            chart_key = f"{group_name}_{chart_name}"
            pattern = re.compile(
                rf"{re.escape(group_name)}_{re.escape(chart_name)}(?:_\d+_\d+)?\.csv$"
            )
            filename = next((name for name in filenames if pattern.match(name)), None)
            file_map[chart_key] = (
                os.path.join(self.raw_data_directory, filename) if filename else None
            )
        return file_map

    def preprocess_chart_types(self, all_charts_info, file_map=None):
        """
        性能優化：一次性預處理所有圖表的數據類型，避免重複計算
        """
        print("正在預處理圖表數據類型...")
        chart_types = {}
        processed_files = set()
        total_charts = len(all_charts_info)
        
        for i, (_, chart_info) in enumerate(all_charts_info.iterrows()):
            group_name = str(chart_info.get('GroupName', 'Unknown'))
            chart_name = str(chart_info.get('ChartName', 'Unknown'))
            chart_key = f"{group_name}_{chart_name}"
            
            # 更新預處理進度
            progress = int((i / total_charts) * 20)  # 預處理占總進度的20%
            self.progress_bar.setValue(progress)
            self.progress_bar.setFormat(f"{tr('preprocessing_chart_types')} {i+1}/{total_charts}")
            if i % 5 == 0 or i + 1 == total_charts:
                QtWidgets.QApplication.processEvents()
                self._check_oob_cancelled()
            
            # 找到對應的 CSV 文件
            filepath = (
                file_map.get(chart_key)
                if file_map is not None
                else find_matching_file(self.raw_data_directory, group_name, chart_name)
            )
            
            if filepath and os.path.exists(filepath) and filepath not in processed_files:
                try:
                    # 完整 CSV 僅讀取一次並留在快取；類型判斷仍只使用前 1000 筆。
                    full_df = pd.read_csv(filepath)
                    self.csv_cache[filepath] = full_df
                    raw_df = full_df.head(1000)
                    customer_result = self.filter_customer_data(raw_df, os.path.basename(filepath))
                    if customer_result.missing_column or customer_result.data.empty:
                        chart_types[chart_key] = 'continuous'
                        processed_files.add(filepath)
                        continue
                    raw_df = customer_result.data
                    
                    if 'point_val' in raw_df.columns:
                        data_type = determine_data_type(raw_df['point_val'].dropna())
                        chart_types[chart_key] = data_type
                        processed_files.add(filepath)
                        print(f"  預處理完成: {chart_key} -> {data_type}")
                    else:
                        chart_types[chart_key] = 'continuous'  # 預設值
                        
                except Exception as e:
                    print(f"  預處理錯誤 {chart_key}: {e}")
                    chart_types[chart_key] = 'continuous'  # 預設值
            else:
                chart_types[chart_key] = 'continuous'  # 預設值
        
        # 預處理完成，進度條顯示20%
        self.progress_bar.setValue(20)
        self.progress_bar.setFormat(tr('preprocessing_complete_starting_charts'))
        QtWidgets.QApplication.processEvents()
        
        self.chart_types_cache = chart_types
        print(f"數據類型預處理完成，共處理 {len(chart_types)} 個圖表")
        return chart_types

    def get_cached_csv(self, filepath):
        """
        性能優化：使用快取讀取 CSV 文件，避免重複讀取
        """
        if filepath not in self.csv_cache:
            try:
                self.csv_cache[filepath] = pd.read_csv(filepath)
                print(f"  CSV 文件已快取: {os.path.basename(filepath)}")
            except Exception as e:
                print(f"  CSV 讀取錯誤 {filepath}: {e}")
                return None
        
        return self.csv_cache[filepath] if self.csv_cache[filepath] is not None else None


    def analyze_chart(self, execution_time, raw_df, chart_info, use_interactive_charts=False, use_batch_id_labels=False, custom_weekly_start=None, custom_weekly_end=None):
        # 補齊 rule_list，確保每個 chart 都有正確的 WE 規則清單以及 CU1/CU2 趨勢規則
        if 'rule_list' not in chart_info or not chart_info['rule_list']:
            rule_list = []
            for rule in ['WE1','WE2','WE3','WE4','WE5','WE6','WE7','WE8','WE9','WE10','CU1','CU2']:
                if chart_info.get(rule, 'N') == 'Y':
                    rule_list.append(rule)
        chart_info['rule_list'] = rule_list
        group_name = str(chart_info.get('group_name', chart_info.get('GroupName', 'Unknown')))
        chart_name = str(chart_info.get('chart_name', chart_info.get('ChartName', 'Unknown')))
        print(f" - analyze_chart 開始處理 {group_name}/{chart_name}")
        print(f" - analyze_chart: 接收到的 raw_df shape: {raw_df.shape}")

        if 'point_time' not in raw_df.columns or not pd.api.types.is_datetime64_any_dtype(raw_df['point_time']):
                print(f" - analyze_chart: 'point_time' column missing or not datetime type for {group_name}/{chart_name}. Skipping analysis.")
                return None

        latest_raw_data_time = raw_df['point_time'].max()

        # 優先使用自定義週期時間範圍
        if custom_weekly_start is not None and custom_weekly_end is not None:
            print(f" - analyze_chart: 使用自定義週期時間範圍: {custom_weekly_start} to {custom_weekly_end}")
            weekly_start_date = custom_weekly_start
            weekly_end_date = custom_weekly_end
        else:
            # 如果沒有自定義時間範圍，使用原本的邏輯
            if execution_time is None or pd.isna(execution_time):
                print(" - analyze_chart: execution_time is None or NaT, using latest data time as weekly end date.")
                weekly_end_date = latest_raw_data_time
            else:
                print(f" - analyze_chart: execution_time is provided ({execution_time}), using it as weekly end date.")
                weekly_end_date = execution_time

            if pd.isna(weekly_end_date):
                print(f" - analyze_chart: Unable to determine weekly end date (latest_raw_data_time is also invalid). Skipping analysis.")
                return None

            weekly_start_date = weekly_end_date - pd.Timedelta(days=6)

        # baseline 邏輯保持不變：以週期開始時間的前一秒作為基線結束
        baseline_end_date = weekly_start_date - pd.Timedelta(seconds=1)
        # 這裡使用初始的一年基線範圍
        initial_baseline_start_date = baseline_end_date - pd.Timedelta(days=365)

        print(f" - analyze_chart: 計算出的時間範圍")
        print(f"   Weekly 週期: {weekly_start_date} to {weekly_end_date}")
        print(f"   Initial Baseline 基線: {initial_baseline_start_date} to {baseline_end_date}")
        print(f"   Baseline 時間長度: {(baseline_end_date - initial_baseline_start_date).days} 天")

        try:
            # === 提前進行數據類型判斷 ===
            if raw_df is None or raw_df.empty or 'point_val' not in raw_df.columns:
                print(" - analyze_chart: raw_df 無效或為空，預設為連續型")
                data_type = 'continuous'
            else:
                # 使用全部 point_val（移除 NaN）來判斷是否為離散
                data_type = determine_data_type(raw_df['point_val'].dropna())
                print(f" - analyze_chart: 數據類型判斷結果: {data_type}")
            
            chart_info['data_type'] = data_type

            # === 根據數據類型分流處理 ===
            if data_type == 'discrete':
                print(f" - analyze_chart: 執行離散型專用流程 for {group_name}/{chart_name}")
                result = self._process_discrete_chart(raw_df, chart_info, weekly_start_date, weekly_end_date, 
                                                    initial_baseline_start_date, baseline_end_date)
            else:
                print(f" - analyze_chart: 執行連續型流程 for {group_name}/{chart_name}")
                result = process_single_chart(
                    chart_info.copy(),
                    raw_df,
                    initial_baseline_start_date,
                    baseline_end_date,
                    weekly_start_date,
                    weekly_end_date,
                    enable_record_high_low=self.oob_settings.get(
                        'enable_record_high_low', False
                    ),
                )
                if result:
                    result['data_type'] = 'continuous'

            if result is None:
                print(f" - analyze_chart: 處理返回 None for {group_name}/{chart_name}")
                return None

            # === 共同的後處理步驟 ===
            print(f" - analyze_chart: 準備生成圖表 for {group_name}/{chart_name}")
            
            # 提取 record_results 供繪圖使用
            record_results = {
                'record_high': result.get('record_high', False),
                'record_low': result.get('record_low', False)
            }
            rule_results = precompute_weekly_rule_results(
                raw_df, chart_info, weekly_start_date, weekly_end_date
            )
            
            if use_interactive_charts:
                # 生成互動式 SPC 圖表（返回 FigureCanvas）
                spc_canvas, violated_rules = plot_spc_chart_interactive(
                    raw_df, chart_info, weekly_start_date, weekly_end_date,
                    record_results=record_results,
                    use_batch_id_labels=use_batch_id_labels,
                    precomputed_rule_results=rule_results,
                )
                print(f" - analyze_chart: plot_spc_chart_interactive 完成")

                # 生成互動式週圖表（返回 FigureCanvas）
                weekly_canvas = plot_weekly_spc_chart_interactive(
                    raw_df, chart_info, weekly_start_date, weekly_end_date,
                    record_results=record_results,
                    use_batch_id_labels=use_batch_id_labels,
                    precomputed_rule_results=rule_results,
                )
                print(f" - analyze_chart: plot_weekly_spc_chart_interactive 完成")
                
                # 為了保持相容性，仍然保存靜態圖片版本供 Excel 使用
                image_path, _ = plot_spc_chart(
                    raw_df, chart_info, weekly_start_date, weekly_end_date,
                    precomputed_rule_results=rule_results,
                )
                weekly_image_path = plot_weekly_spc_chart(
                    raw_df, chart_info, weekly_start_date, weekly_end_date,
                    precomputed_rule_results=rule_results,
                )
                
                # 儲存 canvas 供 UI 使用
                result['spc_canvas'] = spc_canvas
                result['weekly_canvas'] = weekly_canvas
            else:
                # 生成靜態 SPC 圖表
                image_path, violated_rules = plot_spc_chart(
                    raw_df, chart_info, weekly_start_date, weekly_end_date,
                    precomputed_rule_results=rule_results,
                )
                print(f" - analyze_chart: plot_spc_chart 完成，image_path: {image_path}")

                # 生成靜態週圖表
                weekly_image_path = plot_weekly_spc_chart(
                    raw_df, chart_info, weekly_start_date, weekly_end_date,
                    precomputed_rule_results=rule_results,
                )
                print(f" - analyze_chart: plot_weekly_spc_chart 完成，weekly_image_path: {weekly_image_path}")

            # Cpk 計算
            weekly_data = raw_df[(raw_df['point_time'] >= weekly_start_date) & 
                               (raw_df['point_time'] <= weekly_end_date)].copy()
            cpk_result = calculate_cpk(weekly_data, chart_info)
            result['Cpk'] = cpk_result.get('Cpk', np.nan) if cpk_result else np.nan

            # 更新結果
            result['violated_rules'] = violated_rules if violated_rules is not None else {}
            self.build_result(result, image_path, weekly_image_path)

            if self.oob_settings.get('show_by_tool_charts', False):
                # By Tool 圖僅影響顯示，不改變既有 OOB／SPC 計算結果。
                try:
                    result['by_tool_time_canvas'] = plot_spc_by_tool_time(
                        raw_df, chart_info, weekly_start_date, weekly_end_date
                    )
                    result['by_tool_group_canvas'] = plot_spc_by_tool_time_grouping(
                        raw_df, chart_info, weekly_start_date, weekly_end_date
                    )
                except Exception as chart_error:
                    logger.exception("Failed to create By Tool charts: %s", chart_error)
                    result['by_tool_chart_error'] = str(chart_error)

            print(f" - analyze_chart 處理完成並返回結果 for {group_name}/{chart_name}")
            return result

        except Exception as e:
                print(f"[Error] analyze_chart 處理圖表 {group_name}/{chart_name} 時發生錯誤: {str(e)}")
                traceback.print_exc()
                return None

    def _process_discrete_chart(self, raw_df, chart_info, weekly_start_date, weekly_end_date, 
                              initial_baseline_start_date, baseline_end_date):
        """
        離散型數據的專用處理流程，包含 record high low 判斷
        """
        group_name = chart_info.get('group_name', 'Unknown')
        chart_name = chart_info.get('chart_name', 'Unknown')
        record_high_low_enabled = bool(
            getattr(self, 'oob_settings', {}).get('enable_record_high_low', False)
        )
        
        print(f" - _process_discrete_chart: 開始離散型專用處理 {group_name}/{chart_name}")
        
        try:
            # === 基線範圍選擇邏輯 ===
            baseline_data_one_year = raw_df[(raw_df['point_time'] >= initial_baseline_start_date) & 
                                          (raw_df['point_time'] <= baseline_end_date)].copy()
            baseline_count_one_year = len(baseline_data_one_year)
            print(f" - _process_discrete_chart: 初始一年基線數據點數量: {baseline_count_one_year}")

            baseline_insufficient = False
            if baseline_count_one_year < 10:
                actual_baseline_start_date = baseline_end_date - pd.Timedelta(days=365 * 2)
                print(f" - _process_discrete_chart: 基線數據不足，擴展至兩年: {actual_baseline_start_date}")
                
                baseline_data_two_year = raw_df[(raw_df['point_time'] >= actual_baseline_start_date) & 
                                              (raw_df['point_time'] <= baseline_end_date)].copy()
                baseline_count_two_year = len(baseline_data_two_year)
                
                if baseline_count_two_year < 10:
                    print(f" - _process_discrete_chart: 擴展至兩年後仍少於10點，標記為基線不足")
                    baseline_insufficient = True
            else:
                actual_baseline_start_date = initial_baseline_start_date

            # 篩選最終數據
            baseline_data = raw_df[(raw_df['point_time'] >= actual_baseline_start_date) & 
                                 (raw_df['point_time'] <= baseline_end_date)].copy()
            weekly_data = raw_df[(raw_df['point_time'] >= weekly_start_date) & 
                               (raw_df['point_time'] <= weekly_end_date)].copy()

            baseline_empty = baseline_data.empty
            if baseline_empty:
                print(f" - _process_discrete_chart: 基線數據為空，但仍繼續處理 WE Rule 和圖表生成")
                baseline_insufficient = True
                
            if weekly_data.empty:
                print(f" - _process_discrete_chart: 週數據為空，跳過處理")
                return None

            # === 計算統計數據 ===
            def calculate_statistics(data):
                if data.shape[0] <= 1:
                    sigma = 0.0
                else:
                    sigma = data['point_val'].std()
                if np.isnan(sigma):
                    sigma = 0.0
                return {
                    'values': data['point_val'].values,
                    'cnt': data.shape[0],
                    'mean': data['point_val'].mean(),
                    'sigma': sigma
                }

            base_data_dict = calculate_statistics(baseline_data) if not baseline_empty else None
            weekly_data_dict = calculate_statistics(weekly_data)

            if not baseline_empty:
                print(f" - _process_discrete_chart: 基線統計 - cnt={base_data_dict['cnt']}, mean={base_data_dict['mean']}")
            else:
                print(f" - _process_discrete_chart: 基線數據為空，跳過基線統計輸出")
            print(f" - _process_discrete_chart: 週統計 - cnt={weekly_data_dict['cnt']}, mean={weekly_data_dict['mean']}")

            # === 初始化結果字典 ===
            result = {
                'data_cnt': weekly_data_dict['cnt'],
                'ooc_cnt': 0,
                'WE_Rule': '',
                'OOB_Rule': '',
                'Material_no': chart_info.get('material_no', 'N/A'),
                'group_name': chart_info.get('group_name', 'N/A'),
                'chart_name': chart_info.get('chart_name', 'N/A'),
                'chart_ID': chart_info.get('ChartID', 'N/A'),
                'Characteristics': chart_info.get('Characteristics', 'N/A'),
                'USL': chart_info.get('USL', 'N/A'),
                'LSL': chart_info.get('LSL', 'N/A'),
                'UCL': chart_info.get('UCL', 'N/A'),
                'LCL': chart_info.get('LCL', 'N/A'),
                'Target': chart_info.get('Target', 'N/A'),
                'Resolution': chart_info.get('Resolution', 'N/A'),
                'OOB_Status': OOB_CALCULATED,
                'baseline_insufficient': baseline_insufficient,
                'baseline_empty': baseline_empty,  # 新增標記
                'data_type': 'discrete'
            }

            # OOC 只依賴當週資料與控制線，不應被 Baseline 是否足夠影響。
            print(" - _process_discrete_chart: 計算 OOC...")
            weekly_df = pd.DataFrame({'point_val': weekly_data['point_val']})
            try:
                ooc_results = ooc_calculator(
                    weekly_df,
                    chart_info.get('UCL'),
                    chart_info.get('LCL'),
                    chart_info.get('Characteristics'),
                )
                ooc_highlight = review_ooc_results(ooc_results[1], ooc_results[2])
                result['ooc_cnt'] = ooc_results[1]
            except Exception as e:
                logger.exception("Discrete OOC calculation failed: %s", e)
                ooc_highlight = OOB_ERROR
                result['OOB_Status'] = OOB_ERROR

            if not baseline_insufficient and not baseline_empty:
                
                # === 離散型 OOB 計算 ===
                print(" - _process_discrete_chart: 計算離散型 OOB...")
                discrete_oob_result = discrete_oob_calculator(
                    base_data_dict, weekly_data_dict, chart_info,
                    raw_df, weekly_start_date, weekly_end_date,
                    actual_baseline_start_date, baseline_end_date
                )
                
                # === Record High Low 計算（可選，預設關閉）===
                if record_high_low_enabled:
                    logger.debug("基線時間範圍 - 從 %s 到 %s", actual_baseline_start_date, baseline_end_date)
                    logger.debug("當週時間範圍 - 從 %s 到 %s", weekly_start_date, weekly_end_date)
                    logger.debug("基線結束與當週開始間隔 = %s", weekly_start_date - baseline_end_date)
                record_results = optional_record_high_low_calculator(
                    weekly_data['point_val'].values,
                    baseline_data['point_val'].values,
                    enabled=record_high_low_enabled,
                )
                
                # === 更新結果 ===
                result.update({
                    'HL_P95_shift': discrete_oob_result.get('HL_P95_shift', 'NO_HIGHLIGHT'),
                    'HL_P50_shift': discrete_oob_result.get('HL_P50_shift', 'NO_HIGHLIGHT'),
                    'HL_P05_shift': discrete_oob_result.get('HL_P05_shift', 'NO_HIGHLIGHT'),
                    'HL_sticking_shift': discrete_oob_result.get('HL_sticking_shift', 'NO_HIGHLIGHT'),
                    'HL_trending': discrete_oob_result.get('HL_trending', 'NO_HIGHLIGHT'),
                    'HL_high_OOC': ooc_highlight,
                    'HL_category_LT_shift': discrete_oob_result.get('HL_category_LT_shift', 'NO_HIGHLIGHT'),
                    'HL_record_high_low': record_results.get('highlight_status', 'NO_HIGHLIGHT'),
                    'record_high': record_results.get('record_high', False),
                    'record_low': record_results.get('record_low', False)
                })
                
                print(f" - _process_discrete_chart: 離散型 OOB 計算完成")
                
            else:
                # Baseline 指標明確標示未計算；OOC 保留獨立計算結果。
                result.update({
                    'HL_P95_shift': OOB_INSUFFICIENT_DATA,
                    'HL_P50_shift': OOB_INSUFFICIENT_DATA,
                    'HL_P05_shift': OOB_INSUFFICIENT_DATA,
                    'HL_sticking_shift': OOB_INSUFFICIENT_DATA,
                    'HL_trending': OOB_INSUFFICIENT_DATA,
                    'HL_high_OOC': ooc_highlight,
                    'HL_category_LT_shift': OOB_INSUFFICIENT_DATA,
                    'HL_record_high_low': (
                        OOB_INSUFFICIENT_DATA
                        if record_high_low_enabled
                        else 'NO_HIGHLIGHT'
                    ),
                    'record_high': False,
                    'record_low': False,
                    'OOB_Status': OOB_ERROR if ooc_highlight == OOB_ERROR else OOB_INSUFFICIENT_DATA,
                })
                print(f" - _process_discrete_chart: 基線數據不足，Baseline 指標設為 INSUFFICIENT_DATA，OOC 已獨立計算")

            if result.get('OOB_Status') != OOB_ERROR:
                metric_states = [result.get(key) for key in OOB_KEYS]
                if OOB_ERROR in metric_states:
                    result['OOB_Status'] = OOB_ERROR
                elif OOB_INSUFFICIENT_DATA in metric_states:
                    result['OOB_Status'] = OOB_INSUFFICIENT_DATA

            print(f" - _process_discrete_chart: 離散型處理完成 {group_name}/{chart_name}")
            return result

        except Exception as e:
            print(f" - _process_discrete_chart: 處理錯誤 {group_name}/{chart_name}: {e}")
            traceback.print_exc()
            return None

    def build_result(self, result, image_path, weekly_image_path):
        violated_rules = result.get('violated_rules', {})
        we_true_keys = [k for k, v in violated_rules.items() if v]
        result['WE_Rule'] = ', '.join(we_true_keys) if we_true_keys else 'N/A'
        # 只要當週有任何點違規就亮 HL
        result['HL_WE'] = 'HIGHLIGHT' if we_true_keys else 'NO_HIGHLIGHT'

        oob_true_keys = [k for k in OOB_KEYS if result.get(k) == 'HIGHLIGHT']
        metric_states = [result.get(k) for k in OOB_KEYS]
        if result.get('OOB_Status') not in (OOB_CALCULATED, OOB_INSUFFICIENT_DATA, OOB_ERROR):
            if OOB_ERROR in metric_states:
                result['OOB_Status'] = OOB_ERROR
            elif OOB_INSUFFICIENT_DATA in metric_states:
                result['OOB_Status'] = OOB_INSUFFICIENT_DATA
            else:
                result['OOB_Status'] = OOB_CALCULATED

        if oob_true_keys:
            result['OOB_Rule'] = ', '.join(oob_true_keys)
        elif result['OOB_Status'] == OOB_ERROR:
            result['OOB_Rule'] = OOB_ERROR
        elif result['OOB_Status'] == OOB_INSUFFICIENT_DATA:
            result['OOB_Rule'] = OOB_INSUFFICIENT_DATA
        else:
            result['OOB_Rule'] = 'N/A'

        for key in OOB_KEYS:
            result.pop(key, None)
        result.pop('violated_rules', None) # 移除原始的 violated_rules 字典

        result['chart_path'] = image_path
        result['weekly_chart_path'] = weekly_image_path

        # --- 修改點：更強健的 group_name/chart_name 處理 ---
        # 確保 group_name 和 chart_name 都是字符串格式
        raw_group_name = result.get('group_name') or result.get('GroupName')
        raw_chart_name = result.get('chart_name') or result.get('ChartName')
        
        # 處理 group_name
        if raw_group_name is None or raw_group_name == '':
            result['group_name'] = 'N/A'
        else:
            result['group_name'] = str(raw_group_name)  # 強制轉換為字符串
        
        # 處理 chart_name
        if raw_chart_name is None or raw_chart_name == '':
            result['chart_name'] = 'N/A'
        else:
            result['chart_name'] = str(raw_chart_name)  # 強制轉換為字符串

        logger.debug("build_result - 原始 group_name: %s (%s)", raw_group_name, type(raw_group_name))
        logger.debug("build_result - 原始 chart_name: %s (%s)", raw_chart_name, type(raw_chart_name))
        logger.debug("build_result - 處理後 group_name: '%s'", result['group_name'])
        logger.debug("build_result - 處理後 chart_name: '%s'", result['chart_name'])

        # 確保 Cpk 在 result 中，即使是 N/A 或 NaT
        if 'Cpk' not in result:
            result['Cpk'] = np.nan

        print(f" - build_result 完成更新 result for {result.get('group_name', 'Unknown')}/{result.get('chart_name', 'Unknown')}")

    def save_results(self):
        results_df = pd.DataFrame(self.results)

        # 確保所有預期的列都存在，包括新增的數據類型欄位
        expected_cols = ['data_cnt', 'ooc_cnt', 'WE_Rule', 'OOB_Rule', 'OOB_Status', 'data_type', 'Material_no',
                         'group_name', 'chart_name', 'chart_ID', 'Characteristics',
                         'USL', 'LSL', 'UCL', 'LCL', 'Target', 'Cpk', 'Resolution',
                         'baseline_insufficient', 'baseline_empty',
                         'HL_record_high_low', 'record_high', 'record_low',
                         'chart_path', 'weekly_chart_path']
        for col in expected_cols:
            if col not in results_df.columns:
                results_df[col] = np.nan

        cols_to_order = [col for col in expected_cols if col in results_df.columns]
        results_df = results_df[cols_to_order]

        results_df = results_df.replace([np.nan, np.inf, -np.inf], 'N/A')

        try:
             save_results_to_excel(results_df)
             print("Results saved to Excel.")
        except Exception as e:
             print(f"Error saving results to Excel: {e}")
             self.show_error("Save Error", f"Failed to save results to Excel: {e}")

    def display_image(self, result, index):
        if index == 0:
            self._remove_oob_empty_state()
        # 檢查是否有互動式圖表 canvas
        spc_canvas = result.get('spc_canvas')
        weekly_canvas = result.get('weekly_canvas')
        show_by_tool = self.oob_settings.get('show_by_tool_charts', False)
        
        if spc_canvas:
            # 使用互動式 SPC 圖表
            spc_chart_layout = self.create_canvas_layout(spc_canvas, min_width=350, min_height=150)
        else:
            # 使用靜態 SPC 圖表
            spc_image_path = result.get('chart_path')
            if not spc_image_path or not os.path.exists(spc_image_path):
                print(f"[Warning] SPC chart image path invalid: {spc_image_path}")
                spc_chart_layout = QtWidgets.QVBoxLayout()
                spc_chart_layout.addWidget(QtWidgets.QLabel("SPC Chart Not Available"))
            else:
                spc_chart_layout = self.create_image_layout(spc_image_path)

        if weekly_canvas:
            # 使用互動式週圖表
            weekly_chart_layout = self.create_canvas_layout(weekly_canvas, min_width=350, min_height=150)
        else:
            # 使用靜態週圖表
            weekly_image_path = result.get('weekly_chart_path')
            if not weekly_image_path or not os.path.exists(weekly_image_path):
                print(f"[Warning] Weekly chart image path invalid: {weekly_image_path}")
                weekly_chart_layout = QtWidgets.QVBoxLayout()
                weekly_chart_layout.addWidget(QtWidgets.QLabel("Weekly Chart Not Available"))
            else:
                weekly_chart_layout = self.create_image_layout(weekly_image_path)

        info_layout = QtWidgets.QVBoxLayout()
        info_layout.addWidget(self.create_info_label(result))

        by_tool_time_layout = None
        by_tool_group_layout = None
        if show_by_tool:
            by_tool_time_canvas = result.get('by_tool_time_canvas')
            by_tool_group_canvas = result.get('by_tool_group_canvas')
            unavailable_message = result.get('by_tool_chart_error', tr('by_tool_missing_column'))
            if by_tool_time_canvas is not None:
                by_tool_time_layout = self.create_canvas_layout(
                    by_tool_time_canvas, min_width=350, min_height=150
                )
            else:
                by_tool_time_layout = self._create_chart_unavailable_layout(
                    tr('show_by_tool_time_title'), unavailable_message
                )
            if by_tool_group_canvas is not None:
                by_tool_group_layout = self.create_canvas_layout(
                    by_tool_group_canvas, min_width=350, min_height=150
                )
            else:
                by_tool_group_layout = self._create_chart_unavailable_layout(
                    tr('show_by_tool_group_title'), unavailable_message
                )

        result_card = QtWidgets.QFrame()
        result_card.setObjectName("OOBResultCard")
        result_layout = QtWidgets.QGridLayout(result_card)
        result_layout.setContentsMargins(16, 16, 16, 16)
        result_layout.setHorizontalSpacing(16)
        result_layout.setVerticalSpacing(14)
        result_layout.addLayout(spc_chart_layout, 0, 0)
        result_layout.addLayout(weekly_chart_layout, 0, 1)
        if show_by_tool:
            result_layout.addLayout(by_tool_time_layout, 1, 0)
            result_layout.addLayout(by_tool_group_layout, 1, 1)
            result_layout.addLayout(info_layout, 0, 2, 2, 1)
        else:
            result_layout.addLayout(info_layout, 0, 2)
        result_layout.setColumnStretch(0, 4)
        result_layout.setColumnStretch(1, 4)
        result_layout.setColumnStretch(2, 2)
        self.image_grid_layout.addWidget(result_card, index, 0, 1, 3)

    def _create_chart_unavailable_layout(self, title, message):
        layout = QtWidgets.QVBoxLayout()
        label = QtWidgets.QLabel(f"{title}\n{message}")
        label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        label.setWordWrap(True)
        label.setMinimumSize(350, 150)
        label.setStyleSheet(
            "color: #64748B; background-color: #F8FAFC; border: 1px dashed #CBD5E1; "
            "border-radius: 10px; padding: 14px; font-weight: 500;"
        )
        layout.addWidget(label)
        return layout

    def create_image_layout(self, image_path):
        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(30, 0, 0, 0)

        image_label = self.create_image_label(image_path, max_width=450, max_height=550)
        view_button = self.create_view_button(image_path)

        layout.addWidget(image_label)
        layout.addWidget(view_button, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)
        return layout
    def create_canvas_layout(self, canvas, min_width=600, min_height=200):
        """
        創建 canvas 佈局 + Zoom 按鈕
        不使用 ScrollArea，避免出現捲軸
        """
        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        # Canvas widget
        canvas_widget = self.create_canvas_widget(canvas, min_width=min_width, min_height=min_height)

        # 直接把 canvas 放入 layout，不用 ScrollArea
        layout.addWidget(canvas_widget)

        # Zoom 按鈕
        zoom_button = QtWidgets.QPushButton("🔍 Zoom +")
        zoom_button.setMaximumWidth(120)
        zoom_button.setStyleSheet("font-size: 12px;")  # 這裡改字體大小
        zoom_button.clicked.connect(lambda: self.show_full_canvas(canvas))
        layout.addWidget(zoom_button, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)


        return layout


    def create_canvas_widget(self, canvas, min_width=
    600, min_height=200):
        """
        將 FigureCanvas 包裝成 Qt Widget，可自動縮放
        """
        # 讓 canvas 可隨 layout 撐開
        canvas.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Expanding
        )
        canvas.updateGeometry()

        # container 放 canvas
        container = QtWidgets.QWidget(self)
        layout = QtWidgets.QVBoxLayout(container)
        layout.addWidget(canvas)
        layout.setContentsMargins(0, 0, 0, 0)

        # 設保底最小尺寸，避免太小
        container.setMinimumSize(min_width, min_height)
        return container

   
    def show_full_canvas(self, canvas):
        """
        在新視窗中顯示完整的互動圖表
        """
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Interactive Chart")
        dialog.setModal(True)
        dialog.resize(1400, 600)

        layout = QtWidgets.QVBoxLayout(dialog)
        
        # 嘗試使用儲存在 canvas 的參數，呼叫原來的 interactive plotting 函式以確保樣式與互動一致
        new_canvas = None
        try:
            if getattr(canvas, '_plot_args', None) and getattr(canvas, '_plot_kind', None):
                args = canvas._plot_args
                kind = canvas._plot_kind
                
                # 支援舊格式（4個參數）和新格式（5個參數）
                if len(args) >= 5:
                    raw_df_arg, chart_info_arg, ws_arg, we_arg, use_batch_id_arg = args
                else:
                    raw_df_arg, chart_info_arg, ws_arg, we_arg = args
                    use_batch_id_arg = False

                if kind == 'spc':
                    # plot_spc_chart_interactive 返回 (canvas, violated_rules)
                    try:
                        new_canvas, _ = plot_spc_chart_interactive(raw_df_arg, chart_info_arg, ws_arg, we_arg, use_batch_id_labels=use_batch_id_arg)
                    except Exception:
                        new_canvas = None
                elif kind == 'weekly':
                    try:
                        new_canvas = plot_weekly_spc_chart_interactive(raw_df_arg, chart_info_arg, ws_arg, we_arg, use_batch_id_labels=use_batch_id_arg)
                    except Exception:
                        new_canvas = None
                elif kind == 'spc_by_tool_time':
                    new_canvas = plot_spc_by_tool_time(raw_df_arg, chart_info_arg, ws_arg, we_arg)
                elif kind == 'spc_by_tool_group':
                    new_canvas = plot_spc_by_tool_time_grouping(raw_df_arg, chart_info_arg, ws_arg, we_arg)
        except Exception:
            new_canvas = None

        # fallback: 如果無法透過原函式重建，使用原始 canvas 的 figure 複本（較弱但可用）
        if new_canvas is None:
            try:
                new_canvas = FigureCanvas(canvas.figure)
                new_canvas.setSizePolicy(
                    QtWidgets.QSizePolicy.Policy.Expanding,
                    QtWidgets.QSizePolicy.Policy.Expanding
                )
                new_canvas.setMinimumSize(1350, 500)  # 更大的最小尺寸
            except Exception:
                # 最後退回在 dialog 顯示文字
                fallback_label = QtWidgets.QLabel('Unable to create full view of the chart.', dialog)
                fallback_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                layout.addWidget(fallback_label)

        if new_canvas is not None:
            # 設定 canvas 的尺寸策略，讓它能夠填滿整個視窗
            new_canvas.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Expanding,
                QtWidgets.QSizePolicy.Policy.Expanding
            )
            new_canvas.setMinimumSize(1350, 500)  # 設定更大的最小尺寸以配合視窗
            layout.addWidget(new_canvas)
            # 注意：新建的 canvas 通過重新調用原始的 plotting 函數已經包含了 mplcursors 互動功能
            # 因此不需要重複添加，避免出現重疊的提示框
        
        # 添加關閉按鈕
        close_button = QtWidgets.QPushButton("Close")
        close_button.clicked.connect(dialog.close)
        layout.addWidget(close_button, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)
        
        dialog.exec()

    def create_info_label(self, result):
        table_style = """
            <style>
                .info-container {
                    margin-left: 45px;
                    margin-top: 5px;
                }
                table {
                    border-collapse: collapse;
                    width: 450px;
                }
                td, th {
                    padding: 2.5px;
                    text-align: left;
                    border: 1px solid #ddd;
                }
                th {
                    background-color: #344CB7;
                    color: white;
                    font-weight: bold;
                }
                tr:nth-child(even) {
                    background-color: #f4f4f4;
                }
                tr:hover {
                    background-color: #e0e0e0;
                }
                td {
                    color: #000957;
                }
                .title {
                    font-size: 16px;
                    font-weight: bold;
                    margin-bottom: 6px;
                    color: #344CB7;
                }
            </style>
        """

        info_text = f"""
            <html>
                {table_style}
                <div class="info-container">
                    <div class="title"> </div>
                    <table>
                        <thead>
                            <tr>
                                <th>Property</th>
                                <th>Value</th>
                            </tr>
                        </thead>
                        <tbody>
                            {''.join(self.create_table_row(key, result) for key in [
                                'data_cnt', 'ooc_cnt', 'WE_Rule', 'OOB_Rule', 'OOB_Status', 'data_type', 'Material_no',
                                'group_name', 'chart_name', 'Cpk'
                            ])}
                        </tbody>
                    </table>
                </div>
            </html>
        """

        label = QtWidgets.QLabel(info_text, self)
        label.setFont(get_app_font(10))
        label.setTextFormat(QtCore.Qt.TextFormat.RichText)
        label.setWordWrap(True)
        label.setAlignment(QtCore.Qt.AlignmentFlag.AlignTop)

        return label

    def create_table_row(self, key, result):
            value = result.get(key, 'N/A')

            if key == 'group_name' and value == 'Default':
                value = ''

            display_key = key.replace('_', ' ').title()

            if key in ['WE_Rule', 'OOB_Rule']:
                 value = value.replace(', ', '<br>')

            return f"<tr><td>{display_key}:</td><td>{value}</td></tr>"

    def create_view_button(self, image_path):
        button = QtWidgets.QPushButton("Zoom +", self)
        button.setFont(get_app_font(7))
        button.setFixedWidth(85)
        button.setStyleSheet("""
            QPushButton {
                padding: 4px 8px;
                background-color: #344CB7;
                color: white;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #577BC1;
            }
        """)
        button.clicked.connect(lambda checked, path=image_path: self.show_full_image(path))
        return button

    def show_full_image(self, image_path):
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("View Details")
        dialog.setGeometry(100, 100, 1400, 600)

        layout = QtWidgets.QVBoxLayout(dialog)

        if not os.path.exists(image_path):
             error_label = QtWidgets.QLabel("Image file not found.", dialog)
             error_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
             error_label.setStyleSheet("color: red;")
             layout.addWidget(error_label)
        else:
             scroll_area = QtWidgets.QScrollArea(dialog)
             scroll_area.setWidgetResizable(True)

             image_label = self.create_image_label(image_path, keep_original_size=True)

             scroll_area.setWidget(image_label)

             layout.addWidget(scroll_area)

        dialog.setLayout(layout)
        dialog.exec()

    def show_error(self, title, message, warning=False):
        if warning:
            QtWidgets.QMessageBox.warning(self, title, message)
        else:
            QtWidgets.QMessageBox.critical(self, title, message)
    
    def _create_data_check_page(self):
        """建立 Data Health Check 頁面"""
        from data_health_check import DataHealthCheckWidget
        widget = DataHealthCheckWidget(self)
        
        # 傳入目前的 Excel 和 raw data 路徑
        widget.update_paths(self.filepath, self.raw_data_directory)
        
        return widget

class SplitDataWidget(QtWidgets.QWidget):
    """
    A PyQt interface for selecting input/output paths and processing CSV files.
    This Widget will be added to the main window's QStackedWidget.
    Uses modern UI design.
    """
    
    COMMON_ENCODINGS = ["utf-8-sig", "utf-8", "big5", "cp950", "latin1", "cp1252"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        self.apply_styles() # Apply style sheet
        TranslationManager().register_observer(self)
    
    def refresh_ui_texts(self):
        """刷新UI文字（當語言切換時）"""
        from translations import tr
        
        # 更新描述標籤
        self.description_label.setText(
            f"<h2 style='color:#34495E;'>{tr('split_data_title')}</h2>"
            f"<p style='color:#5D6D7E;'>{tr('split_data_description')}</p>"
            f"<p style='color:#5D6D7E;'>{tr('split_data_type2_desc')}</p>"
            f"<p style='color:#5D6D7E;'>{tr('split_data_type3_desc')}</p>"
        )
        
        # 更新 GroupBox 標題
        self.input_group_box.setTitle(tr('select_input_files'))
        self.output_group_box.setTitle(tr('select_output_folder_title'))
        self.mode_group_box.setTitle(tr('select_processing_mode'))
        
        # 更新輸入框提示
        self.input_path_entry.setPlaceholderText(tr('select_csv_files'))
        
        # 更新按鈕
        self.input_button.setText(tr('browse'))
        self.output_button.setText(tr('browse'))
        self.mode_label.setText(tr('select_file_type'))
        
        # 更新下拉選單
        current_index = self.processing_mode_combo.currentIndex()
        self.processing_mode_combo.clear()
        self.processing_mode_combo.addItems([tr('type3_horizontal'), tr('type2_vertical')])
        self.processing_mode_combo.setCurrentIndex(current_index)
        
        # 更新範例按鈕
        self.download_example_button.setText(tr('type3_example'))
        self.download_type2_example_button.setText(tr('type2_example'))
        
        # 更新處理按鈕
        self.process_button.setText(tr('start_processing'))
        
        # 更新進度條
        if self.progress_bar is not None:
            self.progress_bar.setFormat(tr('processing_progress'))
        
        # 更新狀態標籤
        if self.status_label.text() == "Ready." or self.status_label.text() == "準備就緒。":
            self.status_label.setText(tr('ready'))

    def init_ui(self):
        """Initialize user interface elements."""
        from translations import tr
        
        # Main layout uses grid layout for more control
        main_layout = QtWidgets.QGridLayout(self)
        main_layout.setContentsMargins(30, 30, 30, 30) # Add margins
        main_layout.setSpacing(15) # Add spacing between components

        # --- Description text (Using concise title and guidance) ---
        self.description_label = QtWidgets.QLabel(
            f"<h2 style='color:#34495E;'>{tr('split_data_title')}</h2>"
            f"<p style='color:#5D6D7E;'>{tr('split_data_description')}</p>"
            f"<p style='color:#5D6D7E;'>{tr('split_data_type2_desc')}</p>"
            f"<p style='color:#5D6D7E;'>{tr('split_data_type3_desc')}</p>"
        )
        self.description_label.setWordWrap(True)
        self.description_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignTop)
        main_layout.addWidget(self.description_label, 0, 0, 1, 2) # Row 0, Col 0, Span 1 row, 2 columns

        # ...existing code...

        # --- Input file selection block ---
        # Use QFormLayout to organize labels and input/button pairs within QGroupBox
        self.input_group_box = QtWidgets.QGroupBox(tr('select_input_files'))
        input_layout = QtWidgets.QFormLayout(self.input_group_box)
        input_layout.setContentsMargins(15, 20, 15, 15) # 調整 GroupBox 內部邊距
        input_layout.setHorizontalSpacing(10)

        self.input_path_entry = QtWidgets.QLineEdit()
        self.input_path_entry.setPlaceholderText(tr('select_csv_files'))
        self.input_path_entry.setReadOnly(True) # Set as read-only, can only be selected through button
        
        self.input_button = QtWidgets.QPushButton(tr('browse'))
        # *** PyQt6 兼容性修正 START ***
        # input_button.setIcon(self.style().standardIcon(QtWidgets.QStyle.StandardPixmap.SP_DialogOpenButton))
        # *** PyQt6 兼容性修正 END ***
        self.input_button.clicked.connect(self.select_input_files)

        input_row_layout = QtWidgets.QHBoxLayout()
        input_row_layout.addWidget(self.input_path_entry)
        input_row_layout.addWidget(self.input_button)
        input_layout.addRow(input_row_layout)
        main_layout.addWidget(self.input_group_box, 1, 0, 1, 2) # Row 2, Col 0, Span 1 row, 2 columns

        # --- Output folder selection block ---
        self.output_group_box = QtWidgets.QGroupBox(tr('select_output_folder_title'))
        output_layout = QtWidgets.QFormLayout(self.output_group_box)
        output_layout.setContentsMargins(15, 20, 15, 15)
        output_layout.setHorizontalSpacing(10)

        self.output_folder_entry = QtWidgets.QLineEdit()
        # self.output_folder_entry.setPlaceholderText("Split files will be saved in the 'raw_charts' subfolder under this directory...")
        self.output_folder_entry.setReadOnly(True) # Set as read-only
        
        self.output_button = QtWidgets.QPushButton(tr('browse'))
        # *** PyQt6 兼容性修正 START ***
        # output_button.setIcon(self.style().standardIcon(QtWidgets.QStyle.StandardPixmap.SP_DialogOpenButton))
        # *** PyQt6 兼容性修正 END ***
        self.output_button.clicked.connect(self.select_output_folder)

        output_row_layout = QtWidgets.QHBoxLayout()
        output_row_layout.addWidget(self.output_folder_entry)
        output_row_layout.addWidget(self.output_button)
        output_layout.addRow(output_row_layout)
        main_layout.addWidget(self.output_group_box, 2, 0, 1, 2) # Row 2, Col 0, Span 1 row, 2 columns

        # --- Processing mode selection block ---
        self.mode_group_box = QtWidgets.QGroupBox(tr('select_processing_mode'))
        mode_layout = QtWidgets.QHBoxLayout(self.mode_group_box)
        mode_layout.setContentsMargins(15, 20, 15, 15)

        self.mode_label = QtWidgets.QLabel(tr('select_file_type'))
        self.processing_mode_combo = QtWidgets.QComboBox()
        self.processing_mode_combo.addItems([tr('type3_horizontal'), tr('type2_vertical')])
        self.processing_mode_combo.setFixedWidth(250)  # 設定較窄寬度
        self.processing_mode_combo.currentIndexChanged.connect(self._update_processing_mode)
        self._current_processing_mode = "Type3_Horizontal" # 預設內部模式

        mode_layout.addWidget(self.mode_label)
        mode_layout.addWidget(self.processing_mode_combo)


        # --- Download example buttons (Vertical arrangement, pushed to the right) ---
        mode_layout.addStretch(1)
        example_buttons_layout = QtWidgets.QVBoxLayout()
        self.download_example_button = QtWidgets.QPushButton(tr('type3_example'))
        self.download_example_button.setFixedSize(150, 36)
        self.download_example_button.clicked.connect(self.download_type3_example)
        example_buttons_layout.addWidget(self.download_example_button)

        self.download_type2_example_button = QtWidgets.QPushButton(tr('type2_example'))
        self.download_type2_example_button.setFixedSize(150, 36)
        self.download_type2_example_button.clicked.connect(self.download_type2_example)
        example_buttons_layout.addWidget(self.download_type2_example_button)

        mode_layout.addLayout(example_buttons_layout)

        main_layout.addWidget(self.mode_group_box, 3, 0, 1, 2) # Row 3, Col 0, Span 1 row, 2 columns

        # --- Process button ---
        self.process_button = QtWidgets.QPushButton(tr('start_processing'))
        self.process_button.setFixedSize(200, 50) # 更大的按鈕
        self.process_button.clicked.connect(self.run_processing)
        self.process_button.setObjectName("processButton") # 設定物件名稱，用於QSS

        # 將按鈕放在佈局中央
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addStretch(1)
        button_layout.addWidget(self.process_button)
        button_layout.addStretch(1)
        main_layout.addLayout(button_layout, 4, 0, 1, 2) # Row 4, Col 0, Span 1 row, 2 columns

        # --- Progress bar and status message ---
        self.progress_bar = None

        self.status_label = QtWidgets.QLabel(tr('ready'))
        self.status_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("color: #607D8B; font-style: italic;")

        main_layout.addWidget(self.status_label, 6, 0, 1, 2) # Row 6
        main_layout.setRowStretch(7, 1) # 將所有內容推到頂部

    def download_type2_example(self):
        import csv
        from PyQt6 import QtWidgets
        columns = ["GroupName", "ChartName", "point_time", "Batch_ID", "point_val"]
        data = [
            ["Group1", "A", "2025/3/10 00:45", 123, 56.5],
            ["Group1", "A", "2025/3/11 00:45", 123, 56.6],
            ["Group1", "A", "2025/3/12 00:45", 123, 56.5],
            ["Group1", "B", "2025/3/10 00:45", 123, 84],
            ["Group1", "B", "2025/3/11 00:45", 123, 84.2],
            ["Group1", "B", "2025/3/12 00:45", 123, 83.8],
        ]
        save_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save type2 example file", "type2_example.csv", "CSV Files (*.csv)")
        if save_path:
            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(columns)
                writer.writerows(data)
            QtWidgets.QMessageBox.information(self, "Complete", f"Saved to {save_path}")
    def download_type3_example(self):
        import csv
        # 範例資料
        data = [
            ["2025/3/10 00:45", 123, "", 56.5, 84, 123.3, 140, 0.0065, 16820, 16811, -0.11, -0.07, -0.06, 9044],
            ["2025/3/11 00:45", 123, "", 56.6, 84.2, 124, 140, 0.0065, 16748, 16813, -0.11, -0.06, -0.03, 9065],
            ["2025/3/12 00:45", 123, "", 56.5, 83.8, 123, 139.7, 0.0065, 16822, 16822, -0.1, -0.05, -0.13, 9030],
        ]
        columns1 = ["point_time", "Batch_ID", "GroupName", "Group1", "Group1", "Group1", "Group1", "Group1", "Group1", "Group1", "Group1", "Group1", "Group1", "Group1"]
        columns2 = ["", "", "ChartName", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]

        save_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save type3 example file", "type3_example.csv", "CSV Files (*.csv)")
        if save_path:
            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=",")
                writer.writerow(columns1)
                writer.writerow(columns2)
                for row in data:
                    writer.writerow(row)
            QtWidgets.QMessageBox.information(self, "Complete", f"Saved to {save_path}")


    def apply_styles(self):
        """應用 QSS 樣式表。"""
        self.setStyleSheet("""
            QWidget {
                font-family: "Segoe UI", "Microsoft JhengHei", "Yu Gothic UI", "Malgun Gothic", "Meiryo UI", sans-serif;
                font-size: 14px;
                color: #333;
            }
            QLabel {
                color: #333;
            }
            QGroupBox {
                font-size: 15px;
                font-weight: bold;
                color: #2C3E50;
                margin-top: 10px;
                border: 1px solid #D1D1D1;
                border-radius: 8px;
                padding-top: 20px;
                padding-bottom: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                left: 10px;
                margin-left: 5px;
                color: #2C3E50;
            }
            QLineEdit {
                border: 1px solid #BDC3C7;
                border-radius: 5px;
                padding: 8px;
                background-color: #ECF0F1;
                selection-background-color: #3498DB;
            }
            QPushButton {
                background-color: #344CB7;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 20px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #2980B9;
            }
            QPushButton:pressed {
                background-color: #1F618D;
            }
            #processButton {
                background-color: #344CB7;
                font-size: 16px;
                padding: 12px 25px;
                border-radius: 8px;
            }
            #processButton:hover {
                background-color: #2980B9;
                color: #fff;
            }
            #processButton:pressed {
                background-color: #1F618D;
            }
            QComboBox {
                border: 1px solid #BDC3C7;
                border-radius: 5px;
                padding: 8px;
                background-color: #ECF0F1;
                selection-background-color: #3498DB;
            }


            QProgressBar {
                border: 1px solid #BDC3C7;
                border-radius: 5px;
                text-align: center;
                background-color: #ECF0F1;
            }
            QProgressBar::chunk {
                background-color: #344CB7;
                border-radius: 5px;
            }
        """)

        from modern_ui import apply_modern_feature_page
        apply_modern_feature_page(
            self,
            primary_buttons=(self.process_button,),
            secondary_buttons=(
                self.input_button,
                self.output_button,
                self.download_example_button,
                self.download_type2_example_button,
            ),
            cards=(self.input_group_box, self.output_group_box, self.mode_group_box),
            status_labels=(self.status_label,),
        )

    def _update_processing_mode(self, index):
        """根據下拉選單的選擇更新內部處理模式。"""
        # 使用 index 而非文字比對，避免語言切換後字串不符合導致模式無法更新
        if index == 0:
            self._current_processing_mode = "Type3_Horizontal"
        elif index == 1:
            self._current_processing_mode = "Type2_Vertical"

    def select_input_files(self):
        """開啟檔案對話框，選擇多個輸入 CSV 檔案。"""
        file_paths, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Select Input CSV Files (Multiple Selection)", "", "CSV Files (*.csv);;All Files (*.*)"
        )
        if file_paths:
            self.input_path_entry.setText(";".join(file_paths))
            self.status_label.setText(f"Selected {len(file_paths)} files.")

    def select_output_folder(self):
        """開啟資料夾對話框，選擇輸出資料夾。"""
        folder_path = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select Output Folder"
        )
        if folder_path:
            self.output_folder_entry.setText(folder_path)
            self.status_label.setText(f"Selected output folder: {os.path.basename(folder_path)}")

    def sanitize_filename(self, name):
        """輔助函數：清理字串，使其適用於檔案名稱。"""
        invalid_chars = '<>:"/\\|?*\''
        for char in invalid_chars:
            name = name.replace(char, '')
        return name.strip() 

    def _read_csv_with_encoding_fallback(self, filepath, header_val=None):
        """
        嘗試用多種編碼讀取 CSV 檔案，直到成功為止。
        """
        for enc in self.COMMON_ENCODINGS:
            try:
                df = pd.read_csv(filepath, header=header_val, encoding=enc)
                print(f"檔案 '{os.path.basename(filepath)}' 已成功使用 '{enc}' 編碼讀取。")
                return df
            except (UnicodeDecodeError, pd.errors.ParserError) as e:
                print(f"嘗試使用 '{enc}' 編碼讀取 '{os.path.basename(filepath)}' 失敗: {e}")
                continue
            except Exception as e:
                print(f"讀取檔案 '{os.path.basename(filepath)}' 時發生意外錯誤: {e}")
                raise 

        raise ValueError(f"無法使用任何嘗試的編碼讀取檔案 '{os.path.basename(filepath)}'。")

    def _process_type3_horizontal_csv(self, input_path, final_output_folder):
        """處理 Type3 (水平展開) 的 CSV 檔案。"""
        try:
            self.status_label.setText(f"Processing: Type3 file {os.path.basename(input_path)}...")
            print(f"\n--- Processing Type3 (Horizontal Layout) file: {os.path.basename(input_path)} ---")
            df = self._read_csv_with_encoding_fallback(input_path, header_val=None)

            new_columns = []
            for col1, col2 in zip(df.iloc[0], df.iloc[1]):
                if pd.isna(col2):
                    new_columns.append(str(col1))
                elif pd.isna(col1):
                    new_columns.append(str(col2))
                else:
                    new_columns.append(f"{col1}_{col2}")

            df = df.iloc[2:].copy()
            df.columns = new_columns

            chartname_col_name = None
            for col in df.columns:
                if 'GroupName' in col and 'ChartName' in col:
                    chartname_col_name = col
                    break
            
            if chartname_col_name is None:
                QtWidgets.QMessageBox.critical(self, "Error", f"File {os.path.basename(input_path)}: Cannot find 'GroupName' and 'ChartName' combined field. This file will be skipped.")
                return False 
            
            chartname_idx = df.columns.get_loc(chartname_col_name)

            universal_info_columns = df.columns[:chartname_idx + 1].tolist()
            chart_columns = df.columns[(chartname_idx + 1):]

            for chart_col in chart_columns:
                temp_df = df[universal_info_columns].copy()
                temp_df['point_val'] = df[chart_col]
                
                if '_' in chart_col:
                    groupname, chartname = chart_col.split('_', 1)
                else:
                    groupname = ''
                    chartname = chart_col

                temp_df['GroupName'] = groupname
                temp_df['ChartName'] = chartname

                if 'point_time' in temp_df.columns:
                    try:
                        temp_df['point_time'] = pd.to_datetime(temp_df['point_time'], errors='coerce') 
                        temp_df['point_time'] = temp_df['point_time'].dt.strftime('%Y/%#m/%#d %H:%M') 
                    except Exception as time_e:
                        print(f"Warning: Error occurred while processing 'point_time' field in file {os.path.basename(input_path)}: {time_e}")
                        print("This field will be output in original format or contain NaT values.")
                        
                final_columns_order = ['GroupName', 'ChartName', 'point_time', 'point_val']
                for col in universal_info_columns:
                    if col not in final_columns_order and col != chartname_col_name:
                        final_columns_order.append(col)
                
                existing_final_columns_order = [col for col in final_columns_order if col in temp_df.columns]
                temp_df = temp_df[existing_final_columns_order]

                safe_groupname = self.sanitize_filename(groupname)
                safe_chartname = self.sanitize_filename(chartname)
                
                output_file = os.path.join(final_output_folder, f"{safe_groupname}_{safe_chartname}.csv")
                
                if not temp_df.empty: 
                    temp_df.to_csv(output_file, index=False, encoding='utf-8-sig')
                    print(f"已輸出：{os.path.basename(output_file)}")
            return True 

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Error occurred while processing file {os.path.basename(input_path)}: {e}")
            return False 

    def _process_type2_vertical_csv(self, input_path, final_output_folder):
        """處理 Type2 (垂直堆疊 / long format) 的 CSV 檔案。"""
        try:
            self.status_label.setText(f"Processing: Type2 file {os.path.basename(input_path)}...")
            print(f"\n--- Processing Type2 (Vertical Layout) file: {os.path.basename(input_path)} ---")
            df = self._read_csv_with_encoding_fallback(input_path, header_val='infer') 
            
            required_cols = ['GroupName', 'ChartName', 'point_time', 'point_val']
            if not all(col in df.columns for col in required_cols):
                missing_cols = [col for col in required_cols if col not in df.columns]
                QtWidgets.QMessageBox.critical(self, "Error", 
                                     f"Type2 (Vertical Layout) file {os.path.basename(input_path)} is missing required fields: {', '.join(missing_cols)}.\n"
                                     "Please ensure the file contains 'GroupName', 'ChartName', 'point_time', 'point_val' fields.")
                return False

            if 'point_time' in df.columns:
                try:
                    df['point_time'] = pd.to_datetime(df['point_time'], errors='coerce') 
                    df['point_time'] = df['point_time'].dt.strftime('%Y/%#m/%#d %H:%M')
                except Exception as time_e:
                    print(f"Warning: Error occurred while processing 'point_time' field in file {os.path.basename(input_path)}: {time_e}")
                    print("This field will be output in original format or contain NaT values.")

            unique_combinations = df[['GroupName', 'ChartName']].drop_duplicates()

            processed_any_sub_file = False
            for i, (index, row) in enumerate(unique_combinations.iterrows()):
                groupname = row['GroupName']
                chartname = row['ChartName']

                temp_df = df[(df['GroupName'] == groupname) & (df['ChartName'] == chartname)].copy()

                other_cols = [col for col in temp_df.columns if col not in ['GroupName', 'ChartName', 'point_time', 'point_val']]
                final_cols_order = ['GroupName', 'ChartName', 'point_time', 'point_val'] + other_cols
                
                existing_final_cols = [col for col in final_cols_order if col in temp_df.columns]
                temp_df = temp_df[existing_final_cols]

                safe_groupname = self.sanitize_filename(str(groupname)) 
                safe_chartname = self.sanitize_filename(str(chartname)) 

                output_file = os.path.join(final_output_folder, f"{safe_groupname}_{safe_chartname}.csv")
                
                if not temp_df.empty: 
                    temp_df.to_csv(output_file, index=False, encoding='utf-8-sig')
                    print(f"已輸出：{os.path.basename(output_file)}")
                    processed_any_sub_file = True
                
                # Update progress bar (for multiple sub-files within a single Type2 file)
                progress = int((i + 1) / len(unique_combinations) * 100)
                self.progress_bar.setValue(progress)
                self.status_label.setText(f"Processing: {os.path.basename(input_path)} - {i+1}/{len(unique_combinations)} sub-charts")
                QtWidgets.QApplication.processEvents() # Force UI update

            if not processed_any_sub_file:
                QtWidgets.QMessageBox.warning(self, "Warning", f"File {os.path.basename(input_path)} did not generate any split files. Please check its content.")
                return False

            return True 

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Error occurred while processing file {os.path.basename(input_path)}: {e}")
            return False 

    def run_processing(self):
        """
        執行 CSV 處理邏輯，處理多個 CSV 檔案。
        """
        input_paths_str = self.input_path_entry.text()
        input_paths = [path.strip() for path in input_paths_str.split(';') if path.strip()]
        
        base_output_folder = self.output_folder_entry.text()
        # 從內部模式變數獲取，而不是直接從下拉選單獲取顯示文本
        processing_mode = self._current_processing_mode 

        if not input_paths or not base_output_folder:
            QtWidgets.QMessageBox.warning(self, "Warning", "Please select at least one input file and output folder!")
            return

        final_output_folder = os.path.join(base_output_folder, "raw_charts")
        
        try:
            os.makedirs(final_output_folder, exist_ok=True)
            self.status_label.setText(f"Created output folder: {os.path.basename(final_output_folder)}")
            print(f"Created output folder: {final_output_folder}")
        except OSError as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Unable to create 'raw_charts' folder: {final_output_folder}\nError message: {e}")
            return
        
        from modern_ui import ModernProgressDialog
        self.progress_bar = ModernProgressDialog(
            tr("split_progress_title", "Splitting CSV Data"),
            tr("split_progress_preparing", "Preparing input files..."),
            0,
            100,
            self,
        )
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        self.process_button.setEnabled(False)
        self.status_label.setText("Starting file processing...")
        QtWidgets.QApplication.processEvents() # Force UI update

        processed_count = 0
        failed_files = []

        total_files = len(input_paths)
        for i, input_path in enumerate(input_paths):
            success = False
            self.status_label.setText(f"Processing file {i+1}/{total_files}: {os.path.basename(input_path)}")
            self.progress_bar.setLabelText(
                f"{tr('processing', 'Processing')} {i+1}/{total_files}: {os.path.basename(input_path)}"
            )
            # Update overall progress bar (evenly distribute progress for each file)
            overall_progress = int((i / total_files) * 100) 
            self.progress_bar.setValue(overall_progress)
            QtWidgets.QApplication.processEvents() # Force UI update

            try: 
                if processing_mode == "Type3_Horizontal":
                    success = self._process_type3_horizontal_csv(input_path, final_output_folder)
                elif processing_mode == "Type2_Vertical":
                    # Type2 內部會自己更新進度條，這裡只需確保它的起始值正確
                    success = self._process_type2_vertical_csv(input_path, final_output_folder)
            except ValueError as ve: 
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to read file {os.path.basename(input_path)}: {ve}\nPlease ensure the file is in valid CSV format.")
                failed_files.append(os.path.basename(input_path))
                continue
            except Exception as e: 
                QtWidgets.QMessageBox.critical(self, "Error", f"Unexpected error occurred while processing file {os.path.basename(input_path)}: {e}")
                failed_files.append(os.path.basename(input_path))
                continue
            
            if success:
                processed_count += 1
            else:
                failed_files.append(os.path.basename(input_path))
        
        # Processing complete, set progress bar to 100%
        self.progress_bar.setValue(100)
        self.progress_bar.setLabelText(tr("complete", "Complete"))
        QtWidgets.QApplication.processEvents()
        self.progress_bar.close()
        self.progress_bar.deleteLater()
        self.progress_bar = None
        self.process_button.setEnabled(True)

        if processed_count > 0:
            if not failed_files:
                QtWidgets.QMessageBox.information(self, "Complete", f"Successfully processed all {processed_count} files!")
                self.status_label.setText("All files processing completed.")
            else:
                QtWidgets.QMessageBox.warning(self, "Partially Complete", f"Processed {processed_count} files. The following files failed to process:\n{', '.join(failed_files)}")
                self.status_label.setText("Some files failed to process, please check the message.")
        else:
            QtWidgets.QMessageBox.critical(self, "Error", "No files were successfully processed.")
            self.status_label.setText("Processing failed: No files were successfully processed.")

# ============================================================================
# Modern Progress Bar
# ============================================================================

class ModernProgressBar(QtWidgets.QWidget):
    """
    自定義的現代化進度條
    風格：Style 2 - 動態運算 (Active Stripe)
    顏色：#344CB7 (Royal Blue)
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(35)
        self.setFixedWidth(300)
        
        self._value = 0
        self._maximum = 100
        self._text = ""
        
        self._animation_offset = 0
        self._timer = QtCore.QTimer(self)
        self._timer.timeout.connect(self._animate)
        self._timer.start(30)
        
        self.bar_color = QtGui.QColor("#344CB7")
        self.track_color = QtGui.QColor("#DBE4F3")
        self.text_color = QtGui.QColor("#FFFFFF")
        self.stripe_color = QtGui.QColor(255, 255, 255, 40)

    def _animate(self):
        self._animation_offset += 1
        if self._animation_offset >= 40: 
            self._animation_offset = 0
        self.update()

    def setValue(self, value):
        self._value = max(0, min(value, self._maximum))
        self.update()

    def value(self):
        return self._value
    
    def setMaximum(self, value):
        self._maximum = value
        self.update()
        
    def maximum(self):
        return self._maximum

    def setFormat(self, text):
        self._text = text
        self.update()

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)

        rect = self.rect()
        
        # 1. 繪製軌道 (背景)
        painter.setBrush(self.track_color)
        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        painter.drawRoundedRect(rect, 8, 8)

        # 計算進度條寬度
        if self._maximum > 0:
            progress_ratio = self._value / self._maximum
        else:
            progress_ratio = 0
            
        bar_width = rect.width() * progress_ratio
        
        # 2. 繪製進度條 (使用 Clip Path 處理圓角)
        path = QtGui.QPainterPath()
        path.addRoundedRect(QtCore.QRectF(rect), 8, 8)
        painter.setClipPath(path)

        if bar_width > 0:
            # 進度條底色
            bar_rect = QtCore.QRectF(0, 0, bar_width, rect.height())
            painter.setBrush(self.bar_color)
            painter.drawRect(bar_rect)

            # 動態斜紋
            painter.setBrush(self.stripe_color)
            stripe_width = 20
            gap = 20
            total_step = stripe_width + gap
            
            start_x = -total_step + self._animation_offset
            end_x = int(bar_width) + total_step
            h = rect.height()
            
            for x in range(int(start_x), end_x, total_step):
                points = [
                    QtCore.QPointF(x, 0),
                    QtCore.QPointF(x + stripe_width, 0),
                    QtCore.QPointF(x + stripe_width - h, h),
                    QtCore.QPointF(x - h, h)
                ]
                painter.drawPolygon(QtGui.QPolygonF(points))

        # 3. 繪製文字 (置中，自動變色)
        painter.setClipping(False) # 先取消裁切
        
        if self._text:
            font = painter.font()
            font.setBold(True)
            font.setPointSize(12) # 字體加大到 12
            font.setFamily("Segoe UI")
            painter.setFont(font)
            
            text_rect = QtCore.QRectF(rect)
            
            # 第一層：畫在背景上的文字 (使用深藍色)
            painter.setPen(self.bar_color)
            painter.drawText(text_rect, QtCore.Qt.AlignmentFlag.AlignCenter, self._text)
            
            # 第二層：畫在進度條上的文字 (使用白色)
            # 設定裁切區域為進度條部分
            if bar_width > 0:
                # 重新設定 Clip Path 為進度條區域與圓角矩形的交集
                progress_rect_path = QtGui.QPainterPath()
                progress_rect_path.addRect(0, 0, bar_width, rect.height())
                final_clip = path.intersected(progress_rect_path)
                
                painter.setClipPath(final_clip)
                
                painter.setPen(self.text_color)
                painter.drawText(text_rect, QtCore.Qt.AlignmentFlag.AlignCenter, self._text)

# ============================================================================
# CL Tighten Calculator Widget
# ============================================================================

class CLTightenWidget(QtWidgets.QWidget):
    """CL Tighten Calculator 頁面"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_app = parent
        self.filepath = resource_path('input/All_Chart_Information.xlsx')
        self.raw_data_directory = resource_path('input/raw_charts/')
        self.calculator = None
        self.results_data = {}
        
        self.init_ui()
    
    def refresh_ui_texts(self):
        """刷新UI文字（當語言切換時）"""
        # 更新按鈕文字
        self.start_button.setText(tr("start_cl_calculation"))
        self.export_button.setText(f"📁 {tr('export_results')}")
        if hasattr(self, 'page_title_label'):
            self.page_title_label.setText(tr("cl_tighten"))
            self.page_subtitle_label.setText(tr("cl_tighten_subtitle"))
        
        # 更新 Range 標籤
        if hasattr(self, 'range_label'):
            self.range_label.setText(tr("calculation_range"))
        
        # 更新GroupBox標題
        if hasattr(self, 'chart_list_panel'):
            self.chart_list_panel.setTitle(tr("chart_list"))
        if hasattr(self, 'chart_detail_panel'):
            self.chart_detail_panel.setTitle(tr("chart_details"))
        
        # 更新搜尋框 placeholder
        if hasattr(self, 'search_input'):
            self.search_input.setPlaceholderText(tr("search_placeholder"))
        
        # 更新統計標籤
        if hasattr(self, 'stats_label'):
            current_text = self.stats_label.text()
            if current_text in ["No data loaded", "未載入資料", "N/A"]:
                self.stats_label.setText(tr("no_data_loaded"))
        
        # 更新預設標籤
        if hasattr(self, 'default_label'):
            self.default_label.setText(tr("select_chart_prompt"))
        
        # 更新圖例文字
        if hasattr(self, 'legend_widgets'):
            for text_label, text_key, color in self.legend_widgets:
                text_label.setText(tr(text_key))
                text_label.setStyleSheet(f"color: {color}; background-color: transparent;")
        
        print("CLTightenWidget UI texts refreshed")
        
    def init_ui(self):
        """初始化 UI - 頂部對齊優化版"""
        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(24, 22, 24, 24)
        main_layout.setSpacing(16)

        self.page_title_label = QtWidgets.QLabel(tr("cl_tighten"))
        self.page_subtitle_label = QtWidgets.QLabel(tr(
            "cl_tighten_subtitle",
            "Compare current control limits and focus on charts that need adjustment."
        ))
        self.page_subtitle_label.setWordWrap(True)
        main_layout.addWidget(self.page_title_label)
        main_layout.addWidget(self.page_subtitle_label)

        # === 1. 頂部工具列容器 (使用控制高度的 Widget 代替 GroupBox) ===
        top_toolbar = QtWidgets.QWidget()
        top_toolbar.setObjectName("TopToolbar")
        # 這裡統一樣式，讓工具列看起來像一個整體
        top_toolbar.setStyleSheet("""
            #TopToolbar {
                background-color: #FFFFFF;
                border: 1px solid #E5E7EB;
                border-radius: 8px;
            }
            QLabel { color: #64748B; font-weight: bold; font-size: 13px; }
        """)
        
        # 工具列內部佈局
        toolbar_layout = QtWidgets.QHBoxLayout(top_toolbar)
        toolbar_layout.setContentsMargins(10, 8, 10, 8) # 統一邊距
        toolbar_layout.setSpacing(10) # 減少整體間距
        toolbar_layout.setAlignment(QtCore.Qt.AlignmentFlag.AlignVCenter) # 垂直置中對齊

        # --- 時間選擇區 (直接放 Label 和 Edit，不加 GroupBox) ---
        self.range_label = QtWidgets.QLabel(tr("calculation_range"))
        self.range_label.setFixedHeight(32) # 統一高度
        self.range_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignVCenter) # 垂直置中
        self.range_label.setSizePolicy(QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Fixed) # 確保只佔用文字所需的最小寬度
        self.range_label.setStyleSheet("background-color: transparent;") # 移除背景色
        toolbar_layout.addWidget(self.range_label)
        toolbar_layout.addSpacing(3) # Range 標籤後添加小間距
        
        self.start_date_edit = QDateTimeEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setDateTime(QDateTime.currentDateTime().addYears(-2))
        self.start_date_edit.setFixedSize(132, 32)
        self.start_date_edit.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        self.end_date_edit = QDateTimeEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_date_edit.setDateTime(QDateTime.currentDateTime())
        self.end_date_edit.setFixedSize(132, 32)
        self.end_date_edit.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        # 保留原字級並預留足夠空間，避免日期末尾被日曆按鈕遮住。
        compact_date_style = """
            QDateTimeEdit {
                padding: 4px 6px;
                padding-right: 24px;
                font-size: 14px;
            }
            QDateTimeEdit::drop-down {
                width: 20px;
            }
        """
        self.start_date_edit.setStyleSheet(compact_date_style)
        self.end_date_edit.setStyleSheet(compact_date_style)

        toolbar_layout.addWidget(self.start_date_edit)
        toolbar_layout.addSpacing(3) # 減少日期選擇器之間的間距
        separator_label = QtWidgets.QLabel("~")
        separator_label.setFixedHeight(32) # 統一高度
        separator_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignVCenter) # 垂直置中
        separator_label.setStyleSheet("background-color: transparent;") # 移除背景色
        toolbar_layout.addWidget(separator_label)
        toolbar_layout.addSpacing(3) # 減少日期選擇器之間的間距
        toolbar_layout.addWidget(self.end_date_edit)

        # --- 分隔線 (垂直小線條，增加質感) ---
        v_line = QtWidgets.QFrame()
        v_line.setFrameShape(QtWidgets.QFrame.Shape.VLine)
        v_line.setFrameShadow(QtWidgets.QFrame.Shadow.Plain)
        v_line.setStyleSheet("color: #E5E7EB; max-height: 20px;")
        toolbar_layout.addWidget(v_line)

        # --- 按鈕區 ---
        self.start_button = QtWidgets.QPushButton(tr("start_cl_calculation"))
        self.start_button.setFixedHeight(32)
        self.start_button.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        # 覆蓋全局樣式，讓按鈕在這一排看起來更精緻
        btn_style = "QPushButton { padding: 0 15px; font-size: 13px; border-radius: 4px; }"
        self.start_button.setStyleSheet(self.start_button.styleSheet() + btn_style)

        self.export_button = QtWidgets.QPushButton(f"📁 {tr('export_results')}")
        self.export_button.setFixedHeight(32)
        self.export_button.setEnabled(False)
        self.export_button.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.export_button.setStyleSheet(self.export_button.styleSheet() + btn_style)

        toolbar_layout.addWidget(self.start_button)
        toolbar_layout.addWidget(self.export_button)
        
        # 連接按鈕事件
        self.start_button.clicked.connect(self.start_calculation)
        self.export_button.clicked.connect(self.export_results)

        # --- 進度條區 (靠左顯示) ---
        self.progress_bar = None

        # === 在所有元件之後加入彈簧，把所有內容推到左邊 ===
        toolbar_layout.addStretch(1)

        main_layout.addWidget(top_toolbar)

        # === 2. 下方主要內容區域 (維持你原本的 QHBoxLayout) ===
        content_layout = QtWidgets.QHBoxLayout()
        
        left_panel = self.create_chart_list_panel()
        left_panel.setFixedWidth(300)
        
        right_panel = self.create_chart_detail_panel() # 記得照前一封信改成 QScrollArea
        
        content_layout.addWidget(left_panel)
        content_layout.addWidget(right_panel, 1)
        main_layout.addLayout(content_layout, 1)

        self.chart_list_panel.setStyleSheet("")
        self.chart_detail_panel.setStyleSheet("")
        from modern_ui import apply_modern_feature_page
        apply_modern_feature_page(
            self,
            primary_buttons=(self.start_button,),
            secondary_buttons=(self.export_button,),
            cards=(top_toolbar, self.chart_list_panel, self.chart_detail_panel),
            title_labels=(self.page_title_label,),
            subtitle_labels=(self.page_subtitle_label,),
        )
    def create_chart_list_panel(self):
            """創建左側圖表清單面板"""
            self.chart_list_panel = QtWidgets.QGroupBox(tr("chart_list"))
            self.chart_list_panel.setFont(get_app_font(12, QtGui.QFont.Weight.Bold))
            self.chart_list_panel.setStyleSheet("""
                QGroupBox {
                    background-color: #FFFFFF;
                    color: #1A1F2C;
                    border: 1px solid #e5e7eb;
                    border-radius: 8px;
                    margin-top: 12px;
                    padding: 15px;
                    font-weight: bold;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 12px;
                    padding: 0 8px;
                    color: #1A1F2C;
                }
            """)
            layout = QtWidgets.QVBoxLayout(self.chart_list_panel)
            
            # 搜尋框 - 整合放大鏡和輸入框
            search_container = QtWidgets.QWidget()
            search_container.setStyleSheet("""
                QWidget {
                    background-color: white;
                    border: 1px solid #E5E7EB;
                    border-radius: 6px;
                }
            """)
            search_layout = QtWidgets.QHBoxLayout(search_container)
            search_layout.setContentsMargins(10, 6, 10, 6)
            search_layout.setSpacing(8)
            
            search_label = QtWidgets.QLabel("🔍")
            search_label.setStyleSheet("border: none; background: transparent;")
            
            self.search_input = QtWidgets.QLineEdit()
            self.search_input.setPlaceholderText(tr("search_placeholder"))
            self.search_input.setStyleSheet("""
                QLineEdit {
                    border: none;
                    background: transparent;
                    padding: 0px;
                }
            """)
            self.search_input.textChanged.connect(self.filter_charts)
            
            search_layout.addWidget(search_label)
            search_layout.addWidget(self.search_input)
            layout.addWidget(search_container)
            
            # 統計信息
            self.stats_label = QtWidgets.QLabel(tr("no_data_loaded"))
            self.stats_label.setFont(get_app_font(9))
            self.stats_label.setStyleSheet("""
                background-color: #F4F6F9;
                color: #1A1F2C;
                padding: 10px;
                border-radius: 6px;
                border: 1px solid #e5e7eb;
            """)
            layout.addWidget(self.stats_label)
            
            # 圖表清單
            self.chart_list = QtWidgets.QListWidget()
            self.chart_list.itemClicked.connect(self.on_chart_selected)
            layout.addWidget(self.chart_list)
            
            # 狀態圖例
            legend_layout = QtWidgets.QVBoxLayout()
            
            self.legend_items_data = [
                ("⚠️", "need_tighten", "#f59e0b"),
                ("✅", "no_tighten_needed", "#10b981"),
                ("📁❌", "no_data_file", "#64748B"),
                ("🧮❌", "calc_error", "#ef4444"),
                ("📖❌", "read_error", "#ef4444")
            ]
            
            self.legend_widgets = []
            for icon, text_key, color in self.legend_items_data:
                item_layout = QtWidgets.QHBoxLayout()
                icon_label = QtWidgets.QLabel(icon)
                icon_label.setStyleSheet("background-color: transparent;")
                text_label = QtWidgets.QLabel(tr(text_key))
                text_label.setStyleSheet(f"color: {color}; background-color: transparent;")
                item_layout.addWidget(icon_label)
                item_layout.addWidget(text_label)
                item_layout.addStretch()
                legend_layout.addLayout(item_layout)
                self.legend_widgets.append((text_label, text_key, color))
                self.legend_widgets.append((text_label, text_key, color))
                
            layout.addLayout(legend_layout)
            return self.chart_list_panel
            
    def create_chart_detail_panel(self):
            """創建右側圖表詳細資訊面板"""
            self.chart_detail_panel = QtWidgets.QGroupBox(tr("chart_details"))
            self.chart_detail_panel.setFont(get_app_font(12, QtGui.QFont.Weight.Bold))
            self.chart_detail_panel.setStyleSheet("""
                QGroupBox {
                    background-color: #FFFFFF;
                    color: #1A1F2C;
                    border: 1px solid #e5e7eb;
                    border-radius: 8px;
                    margin-top: 12px;
                    padding: 15px;
                    font-weight: bold;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 12px;
                    padding: 0 8px;
                    color: #1A1F2C;
                }
            """)
            layout = QtWidgets.QVBoxLayout(self.chart_detail_panel)
            
            # === 添加外層 ScrollArea 來容納所有內容 ===
            main_scroll_area = QtWidgets.QScrollArea()
            main_scroll_area.setWidgetResizable(True)
            main_scroll_area.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            main_scroll_area.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            main_scroll_area.setStyleSheet("""
                QScrollArea {
                    background-color: white;
                    border: none;
                }
                QScrollBar:vertical {
                    border: none;
                    background: #F1F5F9;
                    width: 10px;
                    border-radius: 5px;
                }
                QScrollBar::handle:vertical {
                    background: #CBD5E1;
                    border-radius: 5px;
                    min-height: 20px;
                }
                QScrollBar::handle:vertical:hover {
                    background: #94A3B8;
                }
            """)
            
            # 創建容器 widget 放置所有內容
            content_container = QtWidgets.QWidget()
            content_layout = QtWidgets.QVBoxLayout(content_container)
            content_layout.setContentsMargins(0, 0, 0, 0)
            content_layout.setSpacing(10)
            
            # 圖表顯示區域（可自動擴展高度以適應原始圖片大小）
            self.chart_display_area = QtWidgets.QWidget()
            self.chart_display_area.setMinimumHeight(300)
            self.chart_display_area.setStyleSheet("background-color: white;")
            chart_display_layout = QtWidgets.QVBoxLayout(self.chart_display_area)
            chart_display_layout.setContentsMargins(0, 0, 0, 0)
            
            # 預設顯示
            self.default_label = QtWidgets.QLabel(tr("select_chart_prompt"))
            self.default_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            self.default_label.setStyleSheet("color: #64748B; font-size: 14px;")
            chart_display_layout.addWidget(self.default_label)
            
            content_layout.addWidget(self.chart_display_area)
            
            # 詳細資訊區域
            self.info_widget = QtWidgets.QWidget()
            self.info_layout = QtWidgets.QVBoxLayout(self.info_widget)
            self.info_layout.setContentsMargins(0, 0, 0, 0)
            self.info_layout.setSpacing(6)
            
            content_layout.addWidget(self.info_widget)
            content_layout.addStretch()
            
            # 將容器設置為 ScrollArea 的 widget
            main_scroll_area.setWidget(content_container)
            
            # 將 ScrollArea 添加到主 layout
            layout.addWidget(main_scroll_area)
            
            return self.chart_detail_panel
            
    def start_calculation(self):
        progress_dialog = None
        """開始CL計算"""
        try:
            if hasattr(self.parent_app, 'ensure_charts_selected') and not self.parent_app.ensure_charts_selected():
                return
            # 檢查必要檔案
            if not os.path.exists(self.filepath):
                QtWidgets.QMessageBox.warning(self, "Warning", 
                    f"Chart information file not found: {self.filepath}")
                return
                
            if not os.path.exists(self.raw_data_directory):
                QtWidgets.QMessageBox.warning(self, "Warning", 
                    f"Raw data directory not found: {self.raw_data_directory}")
                return
            
            # 取得日期範圍
            start_date = self.start_date_edit.dateTime().toPyDateTime()
            end_date = self.end_date_edit.dateTime().toPyDateTime()
            
            # 驗證日期範圍
            if start_date >= end_date:
                QtWidgets.QMessageBox.warning(self, "警告", 
                    "起始日期必須早於結束日期！")
                return
            
            # 創建計算器實例，傳入自訂日期範圍
            self.calculator = CLTightenCalculator(
                chart_info_path=self.filepath,
                raw_data_dir=self.raw_data_directory,
                start_date=start_date,
                end_date=end_date,
                selected_customers=self.parent_app.get_selected_customers(),
                selected_chart_keys=self.parent_app.get_selected_chart_keys()
            )
            
            # 顯示並重置進度條
            from modern_ui import ModernProgressDialog
            progress_dialog = ModernProgressDialog(
                tr("cl_progress_title", "Calculating Control Limits"),
                tr("cl_progress_preparing", "Preparing chart data..."),
                0,
                0,
                self,
            )
            self.progress_bar = progress_dialog
            self.start_button.setEnabled(False)
            progress_dialog.show()
            QtWidgets.QApplication.processEvents()
            
            def progress_callback(current, total):
                progress_dialog.setRange(0, max(1, total))
                progress_dialog.setValue(current)
                # 顯示格式：目前/總數
                progress_dialog.setLabelText(f"{current} / {total}")
                QtWidgets.QApplication.processEvents()

            # 執行計算
            results_df = self.calculator.run_calculation(
                'CL_Calculation_Results.xlsx',
                progress_callback=progress_callback
            )
            
            progress_dialog.setValue(progress_dialog.maximum())
            progress_dialog.setLabelText(tr("complete", "Complete"))
            QtWidgets.QApplication.processEvents()
            
            if results_df is not None and not results_df.empty:
                self.results_df = results_df  # 保存結果 DataFrame 以供匯出使用
                self.load_results(results_df)
                self.export_button.setEnabled(True)
            else:
                raise Exception("No results generated")
                
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Calculation failed: {str(e)}")
        finally:
            if progress_dialog is not None:
                progress_dialog.close()
                progress_dialog.deleteLater()
            self.progress_bar = None
            self.start_button.setEnabled(True)
            
    def load_results(self, results_df):
        """載入計算結果"""
        self.results_data = {}
        self.chart_list.clear()
        
        # 創建排序後的項目清單
        chart_items = []
        
        for _, row in results_df.iterrows():
            chart_key = f"{row.get('GroupName', 'N/A')}_{row.get('ChartName', 'N/A')}"
            self.results_data[chart_key] = row.to_dict()
            
            # 根據狀態設定圖示、顏色和排序優先級
            status = row.get('Status', 'Unknown')
            
            # 使用 Excel 中的 TightenNeeded 欄位（基於容差百分比計算）
            tighten_needed_raw = row.get('TightenNeeded', False)
            # 調試：顯示原始值和類型
            if chart_key == f"{row.get('GroupName', 'N/A')}_{row.get('ChartName', 'N/A')}" and status == 'Success':
                print(f"Chart: {chart_key}, TightenNeeded原始值: {tighten_needed_raw}, 類型: {type(tighten_needed_raw)}")
            
            # 處理各種可能的數據類型
            if pd.isna(tighten_needed_raw):
                tighten_needed = False
            elif isinstance(tighten_needed_raw, bool):
                # 已經是布林值，直接使用
                tighten_needed = tighten_needed_raw
            elif isinstance(tighten_needed_raw, str):
                # 字符串轉布林值
                tighten_needed = tighten_needed_raw.lower() in ['true', '1', 'yes', 'y']
            elif isinstance(tighten_needed_raw, (int, float)):
                # 數字轉布林值
                tighten_needed = bool(tighten_needed_raw)
            else:
                # 其他類型轉布林值
                tighten_needed = bool(tighten_needed_raw)
            
            # 調試：顯示轉換後的值
            if chart_key == f"{row.get('GroupName', 'N/A')}_{row.get('ChartName', 'N/A')}" and status == 'Success':
                print(f"  → 轉換後: {tighten_needed}")
            
            # 設定圖示、顏色、標籤和排序優先級
            if status == 'Success':
                if tighten_needed:
                    icon = "⚠️"
                    color = "#f59e0b"
                    status_label = "Need Tighten"
                    sort_priority = 1
                else:
                    icon = "✅"
                    color = "#10b981"
                    status_label = "No Tighten Needed"
                    sort_priority = 2
            elif status == 'No Raw Data':
                icon = "📁❌"
                color = "#64748B"
                status_label = "No Data File"
                sort_priority = 3
            elif status == 'Calculation Error':
                icon = "🧮❌"
                color = "#ef4444"
                status_label = "Calc Error"
                sort_priority = 4
            elif status == 'File Read Error':
                icon = "📖❌"
                color = "#ef4444"
                status_label = "Read Error"
                sort_priority = 5
            else:
                icon = "❓"
                color = "#ef4444"
                status_label = status
                sort_priority = 6
            
            chart_items.append({
                'chart_key': chart_key,
                'icon': icon,
                'color': color,
                'status_label': status_label,
                'sort_priority': sort_priority,
                'group_name': row.get('GroupName', 'N/A'),
                'chart_name': row.get('ChartName', 'N/A')
            })
        
        # 按排序優先級排序：先排 Tighten Needed，再排 No Tighten，最後排各種錯誤
        chart_items.sort(key=lambda x: (x['sort_priority'], x['group_name'], x['chart_name']))
        
        # 添加到清單
        for item_data in chart_items:
            item = QtWidgets.QListWidgetItem()
            
            # 創建更詳細的顯示文字
            display_text = f"{item_data['icon']} {item_data['chart_key']} ({item_data['status_label']})"
            item.setText(display_text)
            item.setData(QtCore.Qt.ItemDataRole.UserRole, item_data['chart_key'])
            
            # 設定顏色
            item.setForeground(QtGui.QColor(item_data['color']))
            
            self.chart_list.addItem(item)
            
        # 更新統計信息
        self.update_stats_display(chart_items)
        
    def update_stats_display(self, chart_items):
        """更新統計信息顯示"""
        if not chart_items:
            self.stats_label.setText("No data loaded")
            return
            
        # 統計各類別數量
        stats = {}
        for item in chart_items:
            status = item['status_label']
            stats[status] = stats.get(status, 0) + 1
            
        total = len(chart_items)
        
        # 構建統計文字
        stats_text = f"Total: {total}\n"
        
        # 按優先級排序顯示
        priority_order = ["Need Tighten", "No Tighten Needed", "No Data File", "Calc Error", "Read Error"]
        for status in priority_order:
            if status in stats:
                count = stats[status]
                percentage = (count / total * 100) if total > 0 else 0
                stats_text += f"• {status}: {count} ({percentage:.1f}%)\n"
        
        # 顯示其他未分類的
        for status, count in stats.items():
            if status not in priority_order:
                percentage = (count / total * 100) if total > 0 else 0
                stats_text += f"• {status}: {count} ({percentage:.1f}%)\n"
                
        self.stats_label.setText(stats_text.strip())
            
    def on_chart_selected(self, item):
        """當選擇圖表時"""
        chart_key = item.data(QtCore.Qt.ItemDataRole.UserRole)
        if chart_key not in self.results_data:
            return
            
        result = self.results_data[chart_key]
        
        # 顯示圖表圖片
        plot_file = result.get('PlotFile')
        if plot_file and isinstance(plot_file, str) and os.path.exists(plot_file):
            self.display_chart_image(plot_file)
        else:
            self.show_no_image()
            
        # 顯示詳細資訊
        self.display_chart_info(result)
        
    def display_chart_image(self, image_path):
        """顯示圖表圖片 (支援 SVG 和 PNG)"""
        try:
            from PyQt6.QtSvgWidgets import QSvgWidget
            import os
            
            # 清空現有內容
            for i in reversed(range(self.chart_display_area.layout().count())):
                child = self.chart_display_area.layout().itemAt(i)
                if child and child.widget():
                    child.widget().deleteLater()
            
            # 檢查檔案格式
            if image_path.lower().endswith('.svg'):
                # 使用 SVG widget 顯示向量圖（原始尺寸的 1.2 倍）
                svg_widget = QSvgWidget(image_path)
                # 獲取 SVG 原始尺寸並縮放 1.2 倍
                original_size = svg_widget.sizeHint()
                scaled_width = int(original_size.width() * 1.2)
                scaled_height = int(original_size.height() * 1.2)
                svg_widget.setFixedSize(scaled_width, scaled_height)
                
                # 創建容器來置中 SVG widget
                container = QtWidgets.QWidget()
                container.setStyleSheet("background-color: transparent;")
                container_layout = QtWidgets.QHBoxLayout(container)
                container_layout.addStretch()
                container_layout.addWidget(svg_widget)
                container_layout.addStretch()
                container_layout.setContentsMargins(0, 0, 0, 0)
                
                self.chart_display_area.layout().addWidget(container)
            else:
                # PNG 格式使用 QPixmap（高解析度自適應縮放）
                pixmap = QtGui.QPixmap(image_path)
                
                #  關鍵：獲取設備像素比率 (Device Pixel Ratio)
                device_pixel_ratio = self.devicePixelRatio()
                logger.debug("Device Pixel Ratio: %s", device_pixel_ratio)
                
                # 獲取chart_display_area的邏輯寬度
                logical_width = self.chart_display_area.width() - 20  # 留10px邊距
                if logical_width <= 0:
                    logical_width = 800  # 預設寬度
                
                # 🎯 UI顯示縮放因子（調整圖片在UI中的顯示大小）
                ui_scale_factor = 0.9
                logical_width = int(logical_width * ui_scale_factor)
                
                # 計算物理寬度（考慮高分屏縮放）
                physical_width = int(logical_width * device_pixel_ratio)
                logger.debug("Logical width: %dpx (0.95x), Physical width: %dpx", logical_width, physical_width)
                
                # 如果原始圖片寬度小於等於物理寬度，使用原始大小
                if pixmap.width() <= physical_width:
                    # 設定高DPI支援
                    pixmap.setDevicePixelRatio(device_pixel_ratio)
                    scaled_pixmap = pixmap
                    logger.debug("使用原始圖片 (%dx%dpx)", pixmap.width(), pixmap.height())
                else:
                    # 縮放到物理像素大小，然後設定正確的DPR
                    scaled_pixmap = pixmap.scaled(physical_width, int(pixmap.height() * physical_width / pixmap.width()), 
                                                QtCore.Qt.AspectRatioMode.KeepAspectRatio,
                                                QtCore.Qt.TransformationMode.SmoothTransformation)
                    # 設定設備像素比率，讓Qt知道這是高分辨率圖片
                    scaled_pixmap.setDevicePixelRatio(device_pixel_ratio)
                    logger.debug("縮放後圖片: %dx%dpx, DPR: %s", scaled_pixmap.width(), scaled_pixmap.height(), device_pixel_ratio)
                
                image_label = QtWidgets.QLabel()
                image_label.setPixmap(scaled_pixmap)
                image_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                # 確保Label也支持高DPI
                image_label.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground, False)
                self.chart_display_area.layout().addWidget(image_label)
            
        except Exception as e:
            self.show_no_image(f"Error loading image: {str(e)}")
            
    def show_no_image(self, message="No chart image available"):
        """顯示無圖片訊息"""
        # 清空現有內容
        for i in reversed(range(self.chart_display_area.layout().count())):
            child = self.chart_display_area.layout().itemAt(i)
            if child and child.widget():
                child.widget().deleteLater()
        
        # 添加無圖片訊息
        label = QtWidgets.QLabel(message)
        label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("color: #64748B; font-size: 14px;")
        self.chart_display_area.layout().addWidget(label)
        
    def _format_tighten_needed(self, result):
        """格式化 TightenNeeded 顯示值 - 使用 Excel 欄位值"""
        # 使用 Excel 中的 TightenNeeded 欄位（基於容差百分比計算）
        tighten_needed = result.get('TightenNeeded', False)
        
        # 處理各種可能的數據類型
        if pd.isna(tighten_needed):
            tighten_needed = False
        elif isinstance(tighten_needed, bool):
            pass  # 已經是布林值
        elif isinstance(tighten_needed, str):
            tighten_needed = tighten_needed.lower() in ['true', '1', 'yes', 'y']
        elif isinstance(tighten_needed, (int, float)):
            tighten_needed = bool(tighten_needed)
        else:
            tighten_needed = bool(tighten_needed)
        
        return "🟡 Yes" if tighten_needed else "🟢 No"
        
    def display_chart_info(self, result):
        """顯示圖表詳細資訊 - 卡片式設計"""
        # 清除舊資訊
        for i in reversed(range(self.info_layout.count())):
            child = self.info_layout.itemAt(i)
            if child:
                widget = child.widget()
                if widget:
                    widget.deleteLater()
                else:
                    self.clear_layout(child.layout())
        
        # === 標題卡片 - 圖表名稱與狀態 ===
        header_card = self.create_header_card(result)
        self.info_layout.addWidget(header_card)
        
        # === 關鍵指標卡片 ===
        metrics_card = self.create_metrics_card(result)
        self.info_layout.addWidget(metrics_card)
        
        # === 管制線比較卡片 ===
        cl_comparison_card = self.create_cl_comparison_card(result)
        self.info_layout.addWidget(cl_comparison_card)
        
        self.info_layout.addStretch()
    
    def create_header_card(self, result):
        """創建標題卡片 - 緊湊版"""
        card = QtWidgets.QWidget()
        card.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #344CB7, stop:1 #577BC1);
                border-radius: 6px;
                padding: 8px 10px;
            }
        """)
        
        layout = QtWidgets.QHBoxLayout(card)
        layout.setSpacing(8)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 圖表名稱
        chart_name = f"{result.get('GroupName', 'N/A')} @ {result.get('ChartName', 'N/A')}"
        name_label = QtWidgets.QLabel(chart_name)
        name_label.setFont(get_app_font(9, QtGui.QFont.Weight.Bold))
        name_label.setStyleSheet("color: #FFFFFF; background: transparent;")
        name_label.setWordWrap(True)
        
        # Tighten 狀態標籤
        tighten_value = self._format_tighten_needed(result)
        
        if "Yes" in tighten_value:
            status_badge = QtWidgets.QLabel("⚠️ TIGHTEN")
            status_badge.setStyleSheet("""
                background-color: #f59e0b;
                color: #FFFFFF;
                padding: 6px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 11px;
            """)
        else:
            status_badge = QtWidgets.QLabel("✅ OK")
            status_badge.setStyleSheet("""
                background-color: #10b981;
                color: #FFFFFF;
                padding: 6px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 11px;
            """)
        
        layout.addWidget(name_label, 1)
        layout.addWidget(status_badge)
        
        return card
    
    def create_metrics_card(self, result):
        """創建關鍵指標卡片 - 緊湊版"""
        card = QtWidgets.QWidget()
        card.setStyleSheet("""
            QWidget {
                background-color: #FFFFFF;
                border: 1px solid #e5e7eb;
                border-radius: 6px;
                padding: 8px 10px;
            }
        """)
        
        layout = QtWidgets.QVBoxLayout(card)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 指標網格 - 直接顯示,不要標題
        metrics_grid = QtWidgets.QGridLayout()
        metrics_grid.setSpacing(4)
        metrics_grid.setHorizontalSpacing(8)
        metrics_grid.setVerticalSpacing(3)
        metrics_grid.setColumnStretch(1, 1)
        metrics_grid.setColumnStretch(3, 1)
        
        # 計算數據
        initial_ooc = result.get('Ori_OOC_Count', 0)  # ✅ 修正：使用正確的欄位名
        final_ooc = result.get('Final_OOC_Count', 'N/A')
        data_count = result.get('DataCountUsed', 0)
        ooc_rate = f"{(initial_ooc / data_count * 100):.1f}%" if data_count > 0 else "N/A"
        
        # 處理 Pattern 文字（確保是字符串類型）
        pattern_text = result.get('Pattern', 'N/A')
        if pd.isna(pattern_text) or not isinstance(pattern_text, str):
            pattern_text = 'N/A'
        
        hard_rule = result.get('HardRule', 'None')
        if pd.notna(hard_rule) and str(hard_rule) != 'None' and str(hard_rule) not in pattern_text:
            pattern_text += f" ({hard_rule})"
        
        # K值
        ori_k_set = result.get('Ori_K_Set', np.nan)
        sug_k_set = result.get('Sug_K_Set', np.nan)
        ori_k_display = f"{ori_k_set:.1f}σ" if not pd.isna(ori_k_set) else "N/A"
        sug_k_display = f"{sug_k_set:.1f}σ" if not pd.isna(sug_k_set) else "N/A"
        
        # 第一行
        self.add_metric_item(metrics_grid, 0, 0, "Pattern", pattern_text)
        self.add_metric_item(metrics_grid, 0, 2, "Data Count", str(data_count))
        
        # 第二行
        self.add_metric_item(metrics_grid, 1, 0, "OOC Count", f"{initial_ooc} → {final_ooc}", highlight=True)
        self.add_metric_item(metrics_grid, 1, 2, "OOC Rate", ooc_rate, highlight=True)
        
        # 第三行
        self.add_metric_item(metrics_grid, 2, 0, "Ori K Set", ori_k_display)
        self.add_metric_item(metrics_grid, 2, 2, "Sug K Set", sug_k_display, color="#344CB7")
        
        layout.addLayout(metrics_grid)
        
        return card
    
    def create_cl_comparison_card(self, result):
        """創建管制線比較卡片 - 緊湊版"""
        card = QtWidgets.QWidget()
        card.setStyleSheet("""
            QWidget {
                background-color: #FFFFFF;
                border: 1px solid #e5e7eb;
                border-radius: 6px;
                padding: 8px 10px;
            }
        """)
        
        layout = QtWidgets.QVBoxLayout(card)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 比較表格 - 直接顯示,不要標題
        comparison_layout = QtWidgets.QGridLayout()
        comparison_layout.setSpacing(4)
        comparison_layout.setHorizontalSpacing(6)
        comparison_layout.setVerticalSpacing(3)
        
        # 表頭
        headers = ["", "UCL", "CL", "LCL"]
        for col, header in enumerate(headers):
            header_label = QtWidgets.QLabel(header)
            header_label.setFont(get_app_font(9, QtGui.QFont.Weight.Bold))
            header_label.setStyleSheet("color: #64748B; background: transparent;")
            header_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            comparison_layout.addWidget(header_label, 0, col)
        
        # 原始值
        ori_ucl = result.get('UCL', np.nan)
        ori_lcl = result.get('LCL', np.nan)
        cl_center = result.get('CL_Center', np.nan)
        
        ori_label = QtWidgets.QLabel("Original")
        ori_label.setFont(get_app_font(9))
        ori_label.setStyleSheet("color: #64748B; background: transparent;")
        comparison_layout.addWidget(ori_label, 1, 0)
        
        # 使用完整精度顯示，與圖表一致
        ori_ucl_text = f"{ori_ucl}" if not pd.isna(ori_ucl) else "N/A"
        ori_cl_text = f"{cl_center}" if not pd.isna(cl_center) else "N/A"
        ori_lcl_text = f"{ori_lcl}" if not pd.isna(ori_lcl) else "N/A"
        
        for col, text in enumerate([ori_ucl_text, ori_cl_text, ori_lcl_text], 1):
            value_label = QtWidgets.QLabel(text)
            value_label.setFont(get_app_font(9))
            value_label.setStyleSheet("color: #1A1F2C; background: transparent;")
            value_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            comparison_layout.addWidget(value_label, 1, col)
        
        # 建議值
        sug_ucl = result.get('Suggest UCL', np.nan)
        sug_lcl = result.get('Suggest LCL', np.nan)
        
        sug_label = QtWidgets.QLabel("Suggested")
        sug_label.setFont(get_app_font(9, QtGui.QFont.Weight.Bold))
        sug_label.setStyleSheet("color: #344CB7; background: transparent;")
        comparison_layout.addWidget(sug_label, 2, 0)
        
        # 使用完整精度顯示，與圖表一致
        sug_ucl_text = f"{sug_ucl}" if not pd.isna(sug_ucl) else "N/A"
        sug_cl_text = ori_cl_text  # CL 不變
        sug_lcl_text = f"{sug_lcl}" if not pd.isna(sug_lcl) else "N/A"
        
        for col, text in enumerate([sug_ucl_text, sug_cl_text, sug_lcl_text], 1):
            value_label = QtWidgets.QLabel(text)
            value_label.setFont(get_app_font(9, QtGui.QFont.Weight.Bold))
            value_label.setStyleSheet("color: #344CB7; background: transparent;")
            value_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            comparison_layout.addWidget(value_label, 2, col)
        
        layout.addLayout(comparison_layout)
        
        return card
    
    def add_metric_item(self, grid_layout, row, col, label, value, highlight=False, color=None):
        """添加指標項目到網格"""
        # 標籤
        label_widget = QtWidgets.QLabel(f"{label}:")
        label_widget.setFont(get_app_font(8))
        label_widget.setStyleSheet("color: #64748B; background: transparent;")
        grid_layout.addWidget(label_widget, row, col)
        
        # 值
        value_widget = QtWidgets.QLabel(str(value))
        font = get_app_font(8)
        if highlight:
            font.setBold(True)
        value_widget.setFont(font)
        
        if color:
            value_widget.setStyleSheet(f"color: {color}; background: transparent;")
        elif highlight:
            value_widget.setStyleSheet("color: #1A1F2C; background: transparent; font-weight: bold;")
        else:
            value_widget.setStyleSheet("color: #1A1F2C; background: transparent;")
        
        value_widget.setWordWrap(True)
        grid_layout.addWidget(value_widget, row, col + 1)
    
    def add_info_row(self, layout, row, label, value, color="#2c3e50", bold_value=False):
        """添加資訊行 - 舊方法保持相容性"""
        self.add_info_row_to_layout(layout, row, label, value, color, bold_value)
    
    def add_info_row_to_layout(self, layout, row, label, value, color="#1A1F2C", bold_value=False):
        """添加資訊行到指定的佈局"""
        label_widget = QtWidgets.QLabel(f"{label}:")
        label_widget.setFont(get_app_font(9))
        label_widget.setStyleSheet("color: #64748B; font-weight: bold;")
        
        value_widget = QtWidgets.QLabel(str(value))
        font = get_app_font(9)
        if bold_value:
            font.setBold(True)
        value_widget.setFont(font)
        value_widget.setStyleSheet(f"color: {color};")
        value_widget.setWordWrap(True)
        
        layout.addWidget(label_widget, row, 0, QtCore.Qt.AlignmentFlag.AlignTop)
        layout.addWidget(value_widget, row, 1, QtCore.Qt.AlignmentFlag.AlignTop)
    
    def clear_layout(self, layout):
        """清除布局中的所有元素"""
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget:
                    widget.deleteLater()
                else:
                    self.clear_layout(item.layout())
    
    def create_info_section(self, title, items):
        """創建資訊區塊"""
        section = QtWidgets.QGroupBox(title)
        section.setFont(get_app_font(10, QtGui.QFont.Weight.Bold))
        section.setStyleSheet("""
            QGroupBox {
                border: 1px solid #ccc;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: #f8f9fa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        
        layout = QtWidgets.QGridLayout(section)
        layout.setSpacing(8)
        
        for row, (label, value) in enumerate(items):
            label_widget = QtWidgets.QLabel(f"{label}:")
            label_widget.setFont(get_app_font(9))
            label_widget.setStyleSheet("color: #6c757d;")
            
            value_widget = QtWidgets.QLabel(str(value))
            value_widget.setFont(get_app_font(9, QtGui.QFont.Weight.Bold))
            value_widget.setStyleSheet("color: #2c3e50;")
            
            layout.addWidget(label_widget, row, 0)
            layout.addWidget(value_widget, row, 1)
        
        return section
    
    def create_cl_info_section(self, result):
        """創建管制線資訊區塊"""
        section = QtWidgets.QGroupBox("管制線資訊")
        section.setFont(get_app_font(10, QtGui.QFont.Weight.Bold))
        section.setStyleSheet("""
            QGroupBox {
                border: 1px solid #ccc;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: #f8f9fa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        
        layout = QtWidgets.QGridLayout(section)
        layout.setSpacing(8)
        
        # UCL
        layout.addWidget(QtWidgets.QLabel("UCL:"), 0, 0)
        orig_ucl = f"{result.get('UCL', 'N/A'):.4f}" if isinstance(result.get('UCL'), (int, float)) else str(result.get('UCL', 'N/A'))
        sug_ucl = f"{result.get('Suggest UCL', 'N/A'):.4f}" if isinstance(result.get('Suggest UCL'), (int, float)) else str(result.get('Suggest UCL', 'N/A'))
        
        orig_ucl_label = QtWidgets.QLabel(f"Original: {orig_ucl}")
        orig_ucl_label.setStyleSheet("color: #6c757d;")
        layout.addWidget(orig_ucl_label, 0, 1)
        
        sug_ucl_label = QtWidgets.QLabel(f"Suggest: {sug_ucl}")
        sug_ucl_label.setStyleSheet("color: #007bff; font-weight: bold;")
        layout.addWidget(sug_ucl_label, 0, 2)
        
        # LCL
        layout.addWidget(QtWidgets.QLabel("LCL:"), 1, 0)
        orig_lcl = f"{result.get('LCL', 'N/A'):.4f}" if isinstance(result.get('LCL'), (int, float)) else str(result.get('LCL', 'N/A'))
        sug_lcl = f"{result.get('Suggest LCL', 'N/A'):.4f}" if isinstance(result.get('Suggest LCL'), (int, float)) else str(result.get('Suggest LCL', 'N/A'))
        
        orig_lcl_label = QtWidgets.QLabel(f"Original: {orig_lcl}")
        orig_lcl_label.setStyleSheet("color: #6c757d;")
        layout.addWidget(orig_lcl_label, 1, 1)
        
        sug_lcl_label = QtWidgets.QLabel(f"Suggest: {sug_lcl}")
        sug_lcl_label.setStyleSheet("color: #007bff; font-weight: bold;")
        layout.addWidget(sug_lcl_label, 1, 2)
        
        return section
    
    def create_tighten_section(self, result):
        """創建 Tighten 判斷區塊"""
        section = QtWidgets.QGroupBox("Tighten 判斷")
        section.setFont(get_app_font(10, QtGui.QFont.Weight.Bold))
        
        tighten_value = self._format_tighten_needed(result)
        if "Yes" in tighten_value:
            section.setStyleSheet("""
                QGroupBox {
                    border: 2px solid #ffc107;
                    border-radius: 5px;
                    margin-top: 10px;
                    padding-top: 10px;
                    background-color: #fff8e1;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0 5px;
                    color: #f57c00;
                }
            """)
        else:
            section.setStyleSheet("""
                QGroupBox {
                    border: 2px solid #28a745;
                    border-radius: 5px;
                    margin-top: 10px;
                    padding-top: 10px;
                    background-color: #f1f8f4;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0 5px;
                    color: #28a745;
                }
            """)
        
        layout = QtWidgets.QGridLayout(section)
        layout.setSpacing(8)
        
        # Tighten Needed
        layout.addWidget(QtWidgets.QLabel("Tighten Needed:"), 0, 0)
        tighten_label = QtWidgets.QLabel(tighten_value)
        tighten_label.setFont(get_app_font(10, QtGui.QFont.Weight.Bold))
        if "Yes" in tighten_value:
            tighten_label.setStyleSheet("color: #f57c00;")
        else:
            tighten_label.setStyleSheet("color: #28a745;")
        layout.addWidget(tighten_label, 0, 1, 1, 2)
        
        # Diff Ratio
        diff_ratio = result.get('Diff_Ratio_%', np.nan)
        if not pd.isna(diff_ratio):
            layout.addWidget(QtWidgets.QLabel("Diff Ratio:"), 1, 0)
            diff_label = QtWidgets.QLabel(f"{diff_ratio:.2f}%")
            diff_label.setStyleSheet("color: #2c3e50; font-weight: bold;")
            layout.addWidget(diff_label, 1, 1)
            
            # Threshold
            threshold = result.get('Tighten_Threshold_%', np.nan)
            if not pd.isna(threshold):
                layout.addWidget(QtWidgets.QLabel("Threshold:"), 1, 2)
                threshold_label = QtWidgets.QLabel(f"{threshold:.0f}%")
                threshold_label.setStyleSheet("color: #6c757d;")
                layout.addWidget(threshold_label, 1, 3)
        
        return section
        
    def filter_charts(self):
        """過濾圖表清單"""
        search_text = self.search_input.text().lower()
        
        for i in range(self.chart_list.count()):
            item = self.chart_list.item(i)
            chart_name = item.text().lower()
            item.setHidden(search_text not in chart_name)
            
    def view_full_chart(self):
        """查看完整圖表"""
        current_item = self.chart_list.currentItem()
        if not current_item:
            return
            
        chart_key = current_item.data(QtCore.Qt.ItemDataRole.UserRole)
        result = self.results_data.get(chart_key, {})
        plot_file = result.get('PlotFile')
        
        if plot_file and isinstance(plot_file, str) and os.path.exists(plot_file):
            self.show_full_image_dialog(plot_file, chart_key)
        else:
            QtWidgets.QMessageBox.information(self, "Info", "Chart image not available")
            
    def show_full_image_dialog(self, image_path, chart_name):
        """顯示完整圖表對話框"""
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle(f"Control Chart - {chart_name}")
        dialog.setModal(True)
        dialog.resize(1400, 800)
        
        layout = QtWidgets.QVBoxLayout(dialog)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 檢查是否為SVG格式
        if image_path.lower().endswith('.svg'):
            from PyQt6.QtSvgWidgets import QSvgWidget
            
            # 使用SVG widget,不使用ScrollArea
            svg_widget = QSvgWidget(image_path)
            
            # 創建容器並設置白色背景
            container = QtWidgets.QWidget()
            container.setStyleSheet("background-color: white;")
            container_layout = QtWidgets.QVBoxLayout(container)
            container_layout.addStretch()
            
            # 內層容器來置中SVG
            svg_container = QtWidgets.QWidget()
            svg_container.setStyleSheet("background-color: transparent;")
            svg_h_layout = QtWidgets.QHBoxLayout(svg_container)
            svg_h_layout.addStretch()
            svg_h_layout.addWidget(svg_widget)
            svg_h_layout.addStretch()
            svg_h_layout.setContentsMargins(0, 0, 0, 0)
            
            container_layout.addWidget(svg_container)
            container_layout.addStretch()
            container_layout.setContentsMargins(0, 0, 0, 0)
            
            layout.addWidget(container)
        else:
            # PNG格式使用ScrollArea
            scroll_area = QtWidgets.QScrollArea()
            scroll_area.setWidgetResizable(True)
            scroll_area.setStyleSheet("background-color: white;")
            
            image_label = QtWidgets.QLabel()
            pixmap = QtGui.QPixmap(image_path)
            image_label.setPixmap(pixmap)
            image_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            scroll_area.setWidget(image_label)
            
            layout.addWidget(scroll_area)
        
        # 關閉按鈕
        close_button = QtWidgets.QPushButton("Close")
        close_button.clicked.connect(dialog.close)
        layout.addWidget(close_button, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)
        
        dialog.exec()
        
    def apply_new_cl(self):
        """應用新的管制線"""
        current_item = self.chart_list.currentItem()
        if not current_item:
            return
            
        chart_key = current_item.data(QtCore.Qt.ItemDataRole.UserRole)
        result = self.results_data.get(chart_key, {})
        
        # 確認對話框
        reply = QtWidgets.QMessageBox.question(
            self, "Confirm Action",
            f"Apply new control limits for {chart_key}?\n\n"
            f"Original UCL: {result.get('UCL', 'N/A')}\n"
            f"Original LCL: {result.get('LCL', 'N/A')}\n\n"
            f"New UCL: {result.get('Suggest UCL', 'N/A')}\n"
            f"New LCL: {result.get('Suggest LCL', 'N/A')}",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No
        )
        
        if reply == QtWidgets.QMessageBox.StandardButton.Yes:
            # 這裡可以實現實際的應用邏輯
            # 例如：更新資料庫、生成新的配置檔等
            QtWidgets.QMessageBox.information(self, "Success", 
                f"New control limits applied for {chart_key}")
                
    def export_results(self):
        """匯出結果"""
        if not self.results_data:
            QtWidgets.QMessageBox.warning(self, "Warning", "No results to export")
            return
            
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Export Results", "CL_Tighten_Results.xlsx", 
            "Excel Files (*.xlsx);;All Files (*)"
        )
        
        if file_path:
            from modern_ui import ModernProgressDialog
            progress = ModernProgressDialog(
                tr("export_progress", "Export Progress"),
                tr("exporting_results", "Saving results to Excel..."),
                0,
                0,
                self,
            )
            self.export_button.setEnabled(False)
            progress.show()
            QtWidgets.QApplication.processEvents()
            try:
                # 如果有 calculator 和結果 DataFrame，使用完整的匯出功能（包含圖片）
                if hasattr(self, 'calculator') and hasattr(self, 'results_df') and self.results_df is not None:
                    success = self.calculator.export_results(self.results_df, file_path)
                    if success:
                        QtWidgets.QMessageBox.information(self, "Success", 
                            f"Results exported to {file_path}")
                    else:
                        QtWidgets.QMessageBox.critical(self, "Error", 
                            "Export failed. Please check console for details.")
                else:
                    # 備用方案：使用簡單的 DataFrame 匯出
                    import pandas as pd
                    df = pd.DataFrame(list(self.results_data.values()))
                    df.to_excel(file_path, index=False)
                    
                    QtWidgets.QMessageBox.information(self, "Success", 
                        f"Results exported to {file_path}")
                        
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", 
                    f"Export failed: {str(e)}")
            finally:
                progress.close()
                progress.deleteLater()
                self.export_button.setEnabled(True)

if __name__ == "__main__":
    # ========== 編碼與環境設定（防止 Big5 環境閃退）==========
    import locale

    # 0. 設定 logging（WARNING 以上才輸出，避免 DEBUG/INFO 淹沒 console）
    logging.basicConfig(
        level=logging.WARNING,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # 1. 設定環境變數，強制使用 UTF-8
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    os.environ['PYTHONUTF8'] = '1'  # Python 3.7+ 啟用 UTF-8 模式
    
    # 2. 設定 stdout/stderr 編碼（錯誤處理更寬容）
    try:
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        if hasattr(sys.stderr, 'reconfigure'):
            sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception as e:
        print(f"Warning: Failed to reconfigure stdout/stderr encoding: {e}")
    
    # 3. 設定系統 locale（處理中文路徑）
    try:
        if sys.platform == 'win32':
            # Windows 環境嘗試設定繁體中文
            try:
                locale.setlocale(locale.LC_ALL, 'Chinese (Traditional)_Taiwan.950')
            except locale.Error:
                try:
                    locale.setlocale(locale.LC_ALL, 'zh_TW.UTF-8')
                except locale.Error:
                    locale.setlocale(locale.LC_ALL, '')  # 使用系統預設
        else:
            locale.setlocale(locale.LC_ALL, 'zh_TW.UTF-8')
    except Exception as e:
        print(f"Warning: Failed to set locale: {e}")
        try:
            locale.setlocale(locale.LC_ALL, '')
        except Exception:
            pass
    
    # 4. 確保工作目錄正確（特別是打包後）
    if getattr(sys, 'frozen', False):
        # 打包環境：切換到 exe 所在目錄
        os.chdir(os.path.dirname(sys.executable))
    # ======================================
    
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("windowsvista")
    
    # =======================================================
    #  [新增] 全域設定：將所有視窗預設背景改為 #F8FAFC
    # =======================================================
    palette = app.palette()
    # ColorRole.Window 控制一般視窗與 Widget 的背景色
    palette.setColor(QtGui.QPalette.ColorRole.Window, QtGui.QColor("#F8FAFC"))
    # ColorRole.Base 控制輸入框(TextEdit)等的背景色，通常維持白色，這裡不改它以保持卡片白色
    app.setPalette(palette)
    # =======================================================

    spc_app = SPCApp()
    spc_app.show()
    sys.exit(app.exec())
