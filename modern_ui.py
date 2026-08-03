"""Shared visual language for the application's feature pages."""

from PyQt6 import QtCore, QtWidgets


FEATURE_PAGE_STYLE = r"""
QWidget#ModernFeaturePage {
    background-color: #F6F8FC;
    color: #172033;
    font-family: "Segoe UI", "Microsoft JhengHei", "Yu Gothic UI", "Malgun Gothic", "Meiryo UI", "PingFang TC", "Noto Sans CJK JP", "Noto Sans CJK KR", sans-serif;
    font-size: 13px;
}
QWidget#ModernFeaturePage QLabel {
    background: transparent;
}
QWidget#ModernFeaturePage QLabel[uiRole="pageTitle"] {
    color: #172033;
    font-size: 24px;
    font-weight: 750;
    background: transparent;
}
QWidget#ModernFeaturePage QLabel[uiRole="pageSubtitle"] {
    color: #667085;
    font-size: 13px;
    background: transparent;
}
QWidget#ModernFeaturePage QLabel[uiRole="status"] {
    color: #475467;
    background: #EEF2FF;
    border: 1px solid #DDE3FF;
    border-radius: 9px;
    padding: 7px 11px;
}
QWidget#ModernFeaturePage QLabel#plainLabel,
QWidget#ModernFeaturePage QLabel#sectionTitle {
    color: #27324A;
    background: transparent;
    font-size: 14px;
    font-weight: 700;
}
QWidget#ModernFeaturePage QFrame#metricCard QLabel#metricTitle {
    color: #667085;
    background: transparent;
    font-size: 12px;
    font-weight: 650;
}
QWidget#ModernFeaturePage QFrame#metricCard QLabel#metricValue {
    color: #172033;
    background: transparent;
    font-size: 24px;
    font-weight: 750;
}
QWidget#ModernFeaturePage QFrame[uiCard="true"],
QWidget#ModernFeaturePage QWidget[uiCard="true"] {
    background: #FFFFFF;
    border: 1px solid #E1E7F0;
    border-radius: 14px;
}
QWidget#ModernFeaturePage QGroupBox[uiCard="true"] {
    background: #FFFFFF;
    color: #27324A;
    border: 1px solid #E1E7F0;
    border-radius: 14px;
    margin-top: 14px;
    padding: 18px 14px 12px 14px;
    font-size: 14px;
    font-weight: 700;
}
QWidget#ModernFeaturePage QGroupBox[uiCard="true"]::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 0 7px;
    color: #344054;
    background: #FFFFFF;
}
QWidget#ModernFeaturePage QPushButton {
    min-height: 38px;
    padding: 0 16px;
    border-radius: 9px;
    border: 1px solid #D0D8E5;
    background: #FFFFFF;
    color: #344054;
    font-weight: 650;
}
QWidget#ModernFeaturePage QPushButton:hover {
    background: #F7F9FC;
    border-color: #98A5B8;
}
QWidget#ModernFeaturePage QPushButton:pressed {
    background: #EEF2F7;
}
QWidget#ModernFeaturePage QPushButton[uiRole="primary"] {
    background: #3D4FC3;
    color: #FFFFFF;
    border-color: #3D4FC3;
}
QWidget#ModernFeaturePage QPushButton[uiRole="primary"]:hover {
    background: #3343AC;
    border-color: #3343AC;
}
QWidget#ModernFeaturePage QPushButton[uiRole="success"] {
    background: #087A55;
    color: #FFFFFF;
    border-color: #087A55;
}
QWidget#ModernFeaturePage QPushButton[uiRole="success"]:hover {
    background: #066747;
    border-color: #066747;
}
QWidget#ModernFeaturePage QPushButton:disabled {
    background: #E9EDF4;
    color: #98A2B3;
    border-color: #E1E6EE;
}
QWidget#ModernFeaturePage QLineEdit,
QWidget#ModernFeaturePage QComboBox,
QWidget#ModernFeaturePage QDateEdit,
QWidget#ModernFeaturePage QDateTimeEdit,
QWidget#ModernFeaturePage QDoubleSpinBox,
QWidget#ModernFeaturePage QSpinBox {
    min-height: 36px;
    padding: 0 11px;
    color: #27324A;
    background: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 9px;
    selection-background-color: #C7D2FE;
}
QWidget#ModernFeaturePage QLineEdit:focus,
QWidget#ModernFeaturePage QComboBox:focus,
QWidget#ModernFeaturePage QDateEdit:focus,
QWidget#ModernFeaturePage QDateTimeEdit:focus {
    border: 2px solid #6574D9;
}
QWidget#ModernFeaturePage QTableWidget,
QWidget#ModernFeaturePage QListWidget {
    background: #FFFFFF;
    alternate-background-color: #F8FAFD;
    border: 1px solid #E1E7F0;
    border-radius: 12px;
    gridline-color: #EDF1F6;
    color: #27324A;
    outline: none;
}
QWidget#ModernFeaturePage QTableWidget::item,
QWidget#ModernFeaturePage QListWidget::item {
    padding: 7px;
    border-bottom: 1px solid #F0F3F7;
}
QWidget#ModernFeaturePage QTableWidget::item:selected,
QWidget#ModernFeaturePage QListWidget::item:selected {
    color: #283593;
    background: #EEF2FF;
}
QWidget#ModernFeaturePage QHeaderView::section {
    background: #F1F4F9;
    color: #475467;
    border: none;
    border-bottom: 1px solid #DCE3ED;
    padding: 10px 8px;
    font-weight: 700;
}
QWidget#ModernFeaturePage QProgressBar {
    min-height: 9px;
    max-height: 9px;
    border: none;
    border-radius: 4px;
    background: #E1E7F0;
    color: transparent;
    text-align: center;
}
QWidget#ModernFeaturePage QProgressBar::chunk {
    border-radius: 4px;
    background: #5366D6;
}
QWidget#ModernFeaturePage QCheckBox {
    spacing: 8px;
    color: #475467;
    font-weight: 600;
}
QWidget#ModernFeaturePage QCheckBox::indicator {
    width: 18px;
    height: 18px;
}
QWidget#ModernFeaturePage QScrollBar:vertical {
    width: 10px;
    margin: 2px;
    border: none;
    background: transparent;
}
QWidget#ModernFeaturePage QScrollBar::handle:vertical {
    min-height: 28px;
    border-radius: 4px;
    background: #C4CDDA;
}
QWidget#ModernFeaturePage QScrollBar::add-line:vertical,
QWidget#ModernFeaturePage QScrollBar::sub-line:vertical {
    height: 0;
}
"""


class ModernProgressDialog(QtWidgets.QDialog):
    """Non-blocking progress popup shared by long-running feature actions."""

    canceled = QtCore.pyqtSignal()

    def __init__(
        self,
        title,
        label_text="",
        minimum=0,
        maximum=100,
        parent=None,
        *,
        cancelable=False,
        cancel_text="Cancel",
        cancelling_text="Cancelling...",
    ):
        super().__init__(parent)
        self._was_canceled = False
        self._cancelling_text = str(cancelling_text)
        self.setWindowTitle(str(title))
        self.setWindowModality(QtCore.Qt.WindowModality.WindowModal)
        self.setMinimumWidth(540)
        self.setWindowFlag(QtCore.Qt.WindowType.WindowContextHelpButtonHint, False)
        if not cancelable:
            self.setWindowFlag(QtCore.Qt.WindowType.WindowCloseButtonHint, False)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(14)

        self.title_label = QtWidgets.QLabel(str(title))
        self.title_label.setObjectName("ProgressTitle")
        self.status_label = QtWidgets.QLabel(str(label_text))
        self.status_label.setObjectName("ProgressStatus")
        self.status_label.setWordWrap(True)
        self.status_label.setMinimumHeight(38)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setRange(minimum, maximum)
        self.progress_bar.setValue(minimum)
        self.progress_bar.setMinimumHeight(24)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")

        layout.addWidget(self.title_label)
        layout.addWidget(self.status_label)
        layout.addWidget(self.progress_bar)

        self.cancel_button = None
        if cancelable:
            actions = QtWidgets.QHBoxLayout()
            actions.addStretch()
            self.cancel_button = QtWidgets.QPushButton(str(cancel_text))
            self.cancel_button.setObjectName("ProgressCancel")
            self.cancel_button.clicked.connect(self.cancel)
            actions.addWidget(self.cancel_button)
            layout.addLayout(actions)

        self.setStyleSheet("""
            QDialog {
                background: #F8FAFD;
                color: #172033;
                font-family: "Segoe UI", "Microsoft JhengHei", "Yu Gothic UI", "Malgun Gothic", "Meiryo UI", "PingFang TC", "Noto Sans CJK JP", "Noto Sans CJK KR", sans-serif;
            }
            QLabel#ProgressTitle {
                color: #172033;
                font-size: 18px;
                font-weight: 750;
                background: transparent;
            }
            QLabel#ProgressStatus {
                color: #526078;
                font-size: 13px;
                background: transparent;
            }
            QProgressBar {
                color: white;
                background: #DDE4EE;
                border: none;
                border-radius: 12px;
                text-align: center;
                font-weight: 750;
            }
            QProgressBar::chunk {
                background: #4F46E5;
                border-radius: 12px;
            }
            QPushButton#ProgressCancel {
                min-width: 90px;
                min-height: 36px;
                padding: 0 14px;
                color: #344054;
                background: white;
                border: 1px solid #CBD5E1;
                border-radius: 9px;
                font-weight: 650;
            }
            QPushButton#ProgressCancel:hover { background: #F1F5F9; }
            QPushButton#ProgressCancel:disabled { color: #98A2B3; background: #EEF2F6; }
        """)

    def cancel(self):
        if self._was_canceled:
            return
        self._was_canceled = True
        if self.cancel_button is not None:
            self.cancel_button.setEnabled(False)
        self.setLabelText(self._cancelling_text)
        self.canceled.emit()

    def wasCanceled(self):
        return self._was_canceled

    def setLabelText(self, text):
        self.status_label.setText(str(text))

    def setFormat(self, text):
        self.setLabelText(text)

    def setRange(self, minimum, maximum):
        self.progress_bar.setRange(minimum, maximum)

    def setMinimum(self, minimum):
        self.progress_bar.setMinimum(minimum)

    def setMaximum(self, maximum):
        self.progress_bar.setMaximum(maximum)

    def maximum(self):
        return self.progress_bar.maximum()

    def setValue(self, value):
        self.progress_bar.setValue(value)

    def value(self):
        return self.progress_bar.value()

    def setCancelButton(self, button):
        if button is None and self.cancel_button is not None:
            self.cancel_button.hide()

    def setMinimumDuration(self, _duration):
        pass

    def setAutoClose(self, _enabled):
        pass

    def setAutoReset(self, _enabled):
        pass


def _set_role(widget, property_name, value, clear_local_style=False):
    if widget is None:
        return
    widget.setProperty(property_name, value)
    if clear_local_style:
        widget.setStyleSheet("")
    widget.style().unpolish(widget)
    widget.style().polish(widget)


def apply_modern_feature_page(
    page,
    *,
    primary_buttons=(),
    secondary_buttons=(),
    success_buttons=(),
    cards=(),
    title_labels=(),
    subtitle_labels=(),
    status_labels=(),
    clear_table_styles=(),
):
    """Apply the shared theme and semantic roles to a feature page."""
    page.setObjectName("ModernFeaturePage")
    page.setAttribute(QtCore.Qt.WidgetAttribute.WA_StyledBackground, True)

    for button in primary_buttons:
        _set_role(button, "uiRole", "primary", True)
        button.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
    for button in secondary_buttons:
        _set_role(button, "uiRole", "secondary", True)
        button.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
    for button in success_buttons:
        _set_role(button, "uiRole", "success", True)
        button.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
    for card in cards:
        _set_role(card, "uiCard", True)
    for label in title_labels:
        _set_role(label, "uiRole", "pageTitle", True)
    for label in subtitle_labels:
        _set_role(label, "uiRole", "pageSubtitle", True)
    for label in status_labels:
        _set_role(label, "uiRole", "status", True)
    for table in clear_table_styles:
        if table is not None:
            table.setStyleSheet("")

    page.setStyleSheet(FEATURE_PAGE_STYLE)
    page.style().unpolish(page)
    page.style().polish(page)
