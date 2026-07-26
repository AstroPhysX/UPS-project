"""
PySide6 GUI wrapper for the UPS bid-processing script.

Put this file in the same folder as your existing project files, for example:
    Trips_Extractor.py
    Lines_Extractor.py
    master_lines_creation.py
    master_to_pandas.py
    export_to_excel.py
    Processing_fucntions.py

Run with:
    python bid_gui_pyside6.py

Install PySide6 first if needed:
    pip install PySide6

Notes:
    - This is a PySide6 conversion of the Tkinter UPS Bid Analyzer GUI.
    - It keeps the same PDF loading, cached extraction, sorting settings,
      bid-string generation, Excel export, and export-complete dialog behavior.
    - For the Visualizer button, this file expects a PySide6 visualizer module
      such as GUI_spreadsheet_pyside6.py or excel_killer_pyside6.py.
"""

from __future__ import annotations

import ctypes
import json
import math
import os
import platform
import queue
import re
import subprocess
import sys
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

import pandas as pd

from PySide6.QtCore import QDate, QMimeData, QPoint, QSize, QTimer, Qt, Signal
from PySide6.QtGui import QColor, QDrag, QIcon, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCalendarWidget,
    QCheckBox,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from pdf_extractors import extract_trips_from_pdf, parse_line_report_pdf, matching_bid_period
from master_lines_creation import creating_master_line
from master_to_pandas import (
    get_sort_percent_contributions,
    master_lines_to_dataframe,
    drop_empty_sort_columns,
    sort_dataframe_by_conditions,
)
from export_to_excel import export_master_lines_to_excel_table
import Processing_fucntions as pf


CONFIG_PATH = Path("bid_config.json")

DEFAULT_SORTING_SETTINGS = {
    "default_mode": "weighted",
    "weighting_style": "soft",
    "soft_max_weight": 3.0,
    "soft_min_weight": 1.0,
    "keep_score_columns": True,
}

DEFAULT_NUMBER_OF_LINES_TO_BID = 20
DEFAULT_HOURLY_RATE = 50.33

LINE_TYPE_CODES = [
    "TRIPS",
    "VTO",
    "RA",
    "RB",
    "SA",
    "SB",
    "SBA",
    "SBG",
    "VOR",
]

DEFAULT_LINE_TYPE_PREFERENCE_ORDER = list(LINE_TYPE_CODES)

MIN_SORT_CRITERIA_ROWS = 3

SORT_DIRECTION_LABEL_TO_VALUE = {
    "High to Low": "high_to_low",
    "Low to High": "low_to_high",
}
SORT_DIRECTION_VALUE_TO_LABEL = {
    value: label for label, value in SORT_DIRECTION_LABEL_TO_VALUE.items()
}

SORT_MODE_LABEL_TO_VALUE = {
    "High Priority": "strict",
    "Weighted": "weighted",
    "Equal to Previous": "equal",
}
SORT_MODE_VALUE_TO_LABEL = {
    value: label for label, value in SORT_MODE_LABEL_TO_VALUE.items()
}

DEFAULT_MODE_DESCRIPTIONS = {
    "strict": "Normal priority / tie-breaker sort. The first selected column dominates, then the next column breaks ties, and so on.",
    "weighted": "Consecutive weighted conditions are blended into a combined percentile-rank score.",
}

WEIGHTING_STYLE_DESCRIPTIONS = {
    "equal": "Every weighted item in a group gets the same weight: 1.",
    "hard": "Position-based weights. Earlier selected columns matter much more, for example 4, 3, 2, 1.",
    "soft": "Softer position-based weights. With defaults, four weighted items become about 3.0, 2.33, 1.67, 1.0.",
}

UPS_BROWN = "#351C15"
UPS_BROWN_2 = "#4B2618"
UPS_GOLD = "#FFB500"
UPS_BLUE = "#1F6FEB"
UPS_BLUE_ACTIVE = "#1557B0"
UPS_GREEN = "#2E7D32"
UPS_GREEN_ACTIVE = "#1B5E20"
UPS_TEXT = "#FFF7E6"
UPS_PANEL = "#432116"
UPS_FIELD_BG = "#FFF8DC"


# ---------------------------------------------------------------------------
# PyInstaller / auto-py-to-exe resource helpers
# ---------------------------------------------------------------------------

def resource_path(relative_path: str) -> Path:
    """Return a file path that works as .py and inside PyInstaller builds."""
    try:
        base_path = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    except AttributeError:
        base_path = Path(__file__).resolve().parent

    return base_path / relative_path


def set_windows_app_id(app_id: str = "UPS.BidAnalyzer.App") -> None:
    """Help Windows use the app icon in the taskbar instead of the default icon."""
    if platform.system() != "Windows":
        return

    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def apply_window_icon(window: QWidget) -> None:
    """Apply app_icon.ico to a Qt window if the icon file is available."""
    icon_path = resource_path("app_icon.ico")
    if icon_path.exists():
        window.setWindowIcon(QIcon(str(icon_path)))


# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------

def load_saved_config() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        return {}

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(config: dict[str, Any]) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)


def get_os_name() -> str:
    return platform.system()


def clean_filename(filename: str) -> str:
    filename = filename.strip().strip('"').strip("'")
    filename = Path(filename).stem
    filename = re.sub(r'[<>:"/\\|?*]', "_", filename)
    filename = filename.strip().rstrip(".")
    return filename or "Bid_Results"


def validate_positive_int(value: str, field_name: str) -> int:
    text = str(value).strip()
    if not text:
        raise ValueError(f"{field_name} is required.")

    try:
        number = int(text)
    except ValueError as exc:
        raise ValueError(f"{field_name} must be a whole number, such as 20.") from exc

    if number <= 0:
        raise ValueError(f"{field_name} must be greater than zero.")

    return number


def validate_positive_float(value: str, field_name: str) -> float:
    text = str(value).strip()
    if not text:
        raise ValueError(f"{field_name} is required.")

    try:
        number = float(text)
    except ValueError as exc:
        raise ValueError(f"{field_name} must be a number, such as 50.33.") from exc

    if number <= 0:
        raise ValueError(f"{field_name} must be greater than zero.")

    return number


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

def validate_date_or_blank(value: str, field_name: str) -> str | None:
    value = value.strip()
    if not value:
        return None

    try:
        datetime.strptime(value, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError(f"{field_name} must be YYYY-MM-DD or blank.") from exc

    return value


def validate_required_date(value: str, field_name: str) -> str:
    value = value.strip()
    if not value:
        raise ValueError(f"{field_name} is required.")

    try:
        datetime.strptime(value, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError(f"{field_name} must be YYYY-MM-DD.") from exc

    return value


def iso_to_date(value: str | None) -> date:
    if value:
        try:
            return datetime.strptime(value, "%Y-%m-%d").date()
        except ValueError:
            pass
    return date.today()


def date_to_qdate(value: date) -> QDate:
    return QDate(value.year, value.month, value.day)


def make_calendar_icon(size: int = 22) -> QIcon:
    """Create a small calendar icon instead of relying on emoji font support."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)

    border = QColor(255, 255, 255)
    header = QColor(255, 181, 0)
    body = QColor(255, 248, 220)
    dark = QColor(53, 28, 21)

    painter.setPen(QPen(border, 1.6))
    painter.setBrush(body)
    painter.drawRoundedRect(3, 4, size - 6, size - 7, 3, 3)

    painter.setPen(Qt.NoPen)
    painter.setBrush(header)
    painter.drawRoundedRect(3, 4, size - 6, 6, 3, 3)
    painter.drawRect(3, 8, size - 6, 3)

    painter.setPen(QPen(dark, 1.4))
    painter.drawLine(8, 2, 8, 7)
    painter.drawLine(size - 8, 2, size - 8, 7)

    painter.setPen(QPen(dark, 1.1))
    for y in (13, 17):
        painter.drawLine(7, y, size - 7, y)
    for x in (11, 15):
        painter.drawLine(x, 11, x, size - 6)

    painter.end()
    return QIcon(pixmap)


# ---------------------------------------------------------------------------
# Small Qt widgets/dialogs
# ---------------------------------------------------------------------------

class CalendarPopup(QDialog):
    """Small calendar popup that returns an ISO date string."""

    def __init__(
        self,
        parent: QWidget,
        initial_value: str | None,
        callback: Callable[[str], None],
    ) -> None:
        super().__init__(parent)
        self.callback = callback
        self.setWindowTitle("Choose date")
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        initial_date = iso_to_date(initial_value)

        self.calendar = QCalendarWidget(self)
        self.calendar.setSelectedDate(date_to_qdate(initial_date))
        self.calendar.activated.connect(self._choose_selected_date)

        today_button = QPushButton("Today")
        clear_button = QPushButton("Clear")
        cancel_button = QPushButton("Cancel")
        choose_button = QPushButton("Choose")

        today_button.clicked.connect(self._choose_today)
        clear_button.clicked.connect(self._clear_date)
        cancel_button.clicked.connect(self.reject)
        choose_button.clicked.connect(self._choose_selected_date)

        buttons = QHBoxLayout()
        buttons.addWidget(today_button)
        buttons.addWidget(clear_button)
        buttons.addStretch(1)
        buttons.addWidget(cancel_button)
        buttons.addWidget(choose_button)

        layout = QVBoxLayout(self)
        layout.addWidget(self.calendar)
        layout.addLayout(buttons)

        self.resize(360, 300)
        self._position_near_parent(parent)

    def _position_near_parent(self, parent: QWidget) -> None:
        try:
            point = parent.mapToGlobal(QPoint(0, parent.height()))
            self.move(point)
        except Exception:
            pass

    def _qdate_to_iso(self, value: QDate) -> str:
        return value.toString("yyyy-MM-dd")

    def _choose_selected_date(self, _selected: QDate | None = None) -> None:
        self.callback(self._qdate_to_iso(self.calendar.selectedDate()))
        self.accept()

    def _choose_today(self) -> None:
        self.callback(date.today().isoformat())
        self.accept()

    def _clear_date(self) -> None:
        self.callback("")
        self.accept()


class DateEntry(QWidget):
    """Line edit + calendar button. Stores dates as YYYY-MM-DD."""

    def __init__(self, parent: QWidget | None = None, width: int = 14) -> None:
        super().__init__(parent)
        self.line_edit = QLineEdit(self)
        self.line_edit.setPlaceholderText("YYYY-MM-DD")
        self.line_edit.setMaximumWidth(width * 10)

        self.button = QPushButton(self)
        self.button.setIcon(make_calendar_icon())
        self.button.setIconSize(QSize(22, 22))
        self.button.setText("")
        self.button.setToolTip("Open calendar")
        self.button.setFixedWidth(38)
        self.button.clicked.connect(self._open_calendar)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.line_edit)
        layout.addWidget(self.button)

    def text(self) -> str:
        return self.line_edit.text().strip()

    def setText(self, value: str | None) -> None:
        self.line_edit.setText(value or "")

    def _open_calendar(self) -> None:
        CalendarPopup(self, self.text(), self.setText).exec()


class VacationRangeDialog(QDialog):
    """Dialog for adding or editing a vacation range and its OCV setting."""

    def __init__(
        self,
        parent: QWidget,
        title: str,
        initial_start: str = "",
        initial_end: str = "",
        initial_pp_drop: bool = True,
    ) -> None:
        super().__init__(parent)
        self.result: dict[str, Any] | None = None
        self.setWindowTitle(title)
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self.start_entry = DateEntry(self)
        self.start_entry.setText(initial_start)
        self.end_entry = DateEntry(self)
        self.end_entry.setText(initial_end)
        self.ocv_check = QCheckBox("Enable OCV / pay-period drop")
        self.ocv_check.setChecked(bool(initial_pp_drop))
        self.ocv_check.setToolTip(
            "Checked stores pp_drop=True for this vacation range. "
            "Unchecked stores pp_drop=False."
        )

        main = QGridLayout(self)
        main.addWidget(QLabel("Start date:"), 0, 0)
        main.addWidget(self.start_entry, 0, 1)
        main.addWidget(QLabel("End date:"), 1, 0)
        main.addWidget(self.end_entry, 1, 1)
        main.addWidget(self.ocv_check, 2, 0, 1, 2)

        save_button = QPushButton("Save")
        cancel_button = QPushButton("Cancel")
        save_button.clicked.connect(self._save)
        cancel_button.clicked.connect(self.reject)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        buttons.addWidget(cancel_button)
        buttons.addWidget(save_button)
        main.addLayout(buttons, 3, 0, 1, 2)

        self.resize(390, 165)

    def _save(self) -> None:
        try:
            start = validate_required_date(self.start_entry.text(), "Vacation start")
            end = validate_required_date(self.end_entry.text(), "Vacation end")
            if end < start:
                raise ValueError("Vacation end date is before vacation start date.")
            self.result = {
                "start": start,
                "end": end,
                "pp_drop": self.ocv_check.isChecked(),
            }
            self.accept()
        except Exception as exc:
            QMessageBox.critical(self, "Vacation date error", str(exc))


class RequestedDateRangeDialog(QDialog):
    """Dialog for a requested single day off or a requested date range."""

    def __init__(
        self,
        parent: QWidget,
        title: str,
        initial_start: str = "",
        initial_end: str = "",
        initial_note: str = "",
    ) -> None:
        super().__init__(parent)
        self.result: dict[str, str] | None = None
        self.setWindowTitle(title)
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self.note_edit = QLineEdit(self)
        self.note_edit.setText(initial_note)
        self.note_edit.setPlaceholderText("Optional note")
        self.note_edit.setToolTip(
            "This note is saved only in your local bid_config.json and is not "
            "passed into the requested-days scoring function."
        )

        self.start_entry = DateEntry(self)
        self.start_entry.setText(initial_start)
        self.end_entry = DateEntry(self)
        self.end_entry.setText(initial_end)

        help_label = QLabel(
            "For one requested day, enter only the start date. "
            "For a range, enter both dates."
        )
        help_label.setWordWrap(True)

        privacy_label = QLabel(
            "Notes are saved for your reference and do not affect scoring."
        )
        privacy_label.setWordWrap(True)
        privacy_label.setStyleSheet("color: #E6D7C5;")

        main = QGridLayout(self)
        main.setColumnStretch(1, 1)
        main.addWidget(help_label, 0, 0, 1, 2)
        main.addWidget(QLabel("Notes:"), 1, 0)
        main.addWidget(self.note_edit, 1, 1)
        main.addWidget(QLabel("Date / start:"), 2, 0)
        main.addWidget(self.start_entry, 2, 1)
        main.addWidget(QLabel("Optional end:"), 3, 0)
        main.addWidget(self.end_entry, 3, 1)
        main.addWidget(privacy_label, 4, 0, 1, 2)

        save_button = QPushButton("Save")
        cancel_button = QPushButton("Cancel")
        save_button.clicked.connect(self._save)
        cancel_button.clicked.connect(self.reject)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        buttons.addWidget(cancel_button)
        buttons.addWidget(save_button)
        main.addLayout(buttons, 5, 0, 1, 2)

        self.resize(540, 240)

    def _save(self) -> None:
        try:
            start = validate_required_date(self.start_entry.text(), "Requested date")
            end = validate_date_or_blank(self.end_entry.text(), "Requested end date") or ""
            if end and end < start:
                raise ValueError("Requested end date is before the start date.")
            self.result = {
                "note": self.note_edit.text().strip(),
                "start": start,
                "end": end,
            }
            self.accept()
        except Exception as exc:
            QMessageBox.critical(self, "Requested date error", str(exc))


class ExportCompleteDialog(QDialog):
    """Export-complete dialog with buttons to open the Excel file or folder."""

    def __init__(
        self,
        parent: QWidget,
        output_path: Path,
        open_file_callback: Callable[[Path], None],
        open_folder_callback: Callable[[Path], None],
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Export Complete")
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        title = QLabel("Excel file created successfully.")
        title.setObjectName("ExportTitle")

        path_entry = QLineEdit(str(output_path))
        path_entry.setReadOnly(True)
        path_entry.setMinimumWidth(700)

        open_file_button = QPushButton("Open Excel file")
        open_file_button.setObjectName("GreenButton")
        open_folder_button = QPushButton("Open folder")
        close_button = QPushButton("Close")

        open_file_button.clicked.connect(lambda: open_file_callback(output_path))
        open_folder_button.clicked.connect(lambda: open_folder_callback(output_path))
        close_button.clicked.connect(self.accept)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        buttons.addWidget(open_file_button)
        buttons.addWidget(open_folder_button)
        buttons.addWidget(close_button)

        layout = QVBoxLayout(self)
        layout.addWidget(title)
        layout.addWidget(path_entry)
        layout.addLayout(buttons)


# ---------------------------------------------------------------------------
# Sorting helpers copied from the Tk GUI logic
# ---------------------------------------------------------------------------

def is_calendar_date_column(col: Any) -> bool:
    if isinstance(col, (date, datetime, pd.Timestamp)):
        return True

    if not isinstance(col, str):
        return False

    text = col.strip()

    formats_with_year = [
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
    ]

    formats_without_year = [
        "%a, %b %d",
        "%A, %b %d",
        "%a, %B %d",
        "%A, %B %d",
    ]

    for fmt in formats_with_year:
        try:
            datetime.strptime(text, fmt)
            return True
        except ValueError:
            pass

    for fmt in formats_without_year:
        try:
            datetime.strptime(f"{text} 2000", f"{fmt} %Y")
            return True
        except ValueError:
            pass

    return False


def get_sortable_columns_from_df(df: pd.DataFrame, include_text_columns: bool = False) -> list[str]:
    sortable_columns: list[str] = []

    for col in df.columns:
        if is_calendar_date_column(col):
            continue

        if include_text_columns:
            sortable_columns.append(str(col))
            continue

        cleaned = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.strip()
            .replace({"": None, "None": None, "nan": None, "NaN": None})
        )

        numeric_values = pd.to_numeric(cleaned, errors="coerce")
        if numeric_values.notna().any():
            sortable_columns.append(str(col))

    return sortable_columns


# ---------------------------------------------------------------------------
# Drag-and-drop sorting-row helpers
# ---------------------------------------------------------------------------

class WheelSafeComboBox(QComboBox):
    """Combo box that ignores mouse-wheel changes unless its popup is open."""

    def wheelEvent(self, event: Any) -> None:
        # Page scrolling should never accidentally change a sorting selection.
        event.ignore()


class NoInternalScrollListWidget(QListWidget):
    """Compact list that shows every item and lets the page handle scrolling."""

    def wheelEvent(self, event: Any) -> None:
        # Do not scroll the line-type list internally. Let the outer page receive
        # the wheel event instead.
        event.ignore()

    def resize_to_all_items(self) -> None:
        count = self.count()
        if count <= 0:
            self.setFixedHeight(32)
            return

        row_height = max(
            22,
            max(self.sizeHintForRow(row) for row in range(count)),
        )
        spacing = max(0, self.spacing())
        content_height = row_height * count + spacing * max(0, count - 1)
        frame_height = self.frameWidth() * 2
        self.setFixedHeight(content_height + frame_height + 6)


SORT_ROW_MIME_TYPE = "application/x-ups-sort-criterion-row"


class SortCriteriaListWidget(QWidget):
    """
    Expanding sorting-row container with live, widget-safe drag reordering.

    The rows are regular widgets in a QVBoxLayout rather than item widgets
    embedded in QListWidget. During a drag, the source row is hidden and a
    visible placeholder moves through the layout. This makes the surrounding
    rows shift immediately while avoiding Qt deleting or detaching controls.
    """

    orderPreviewed = Signal()
    orderCommitted = Signal()
    orderCancelled = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.rows_layout = QVBoxLayout(self)
        self.rows_layout.setContentsMargins(0, 0, 0, 0)
        self.rows_layout.setSpacing(5)

        self._items: list[QListWidgetItem] = []
        self._widgets: dict[int, QWidget] = {}

        self._drag_source_item: QListWidgetItem | None = None
        self._drag_original_items: list[QListWidgetItem] = []
        self._drag_placeholder: QFrame | None = None
        self._drag_committed = False

    @staticmethod
    def _item_key(item: QListWidgetItem) -> int:
        """Return a stable, hashable key for a QListWidgetItem wrapper."""
        return id(item)

    # QListWidget-like compatibility helpers used by the surrounding GUI.
    def addItem(self, item: QListWidgetItem) -> None:
        self._items.append(item)
        self._rebuild_layout()

    def setItemWidget(self, item: QListWidgetItem, widget: QWidget) -> None:
        self._widgets[self._item_key(item)] = widget
        self._rebuild_layout()

    def itemWidget(self, item: QListWidgetItem) -> QWidget | None:
        return self._widgets.get(self._item_key(item))

    def count(self) -> int:
        return len(self._items)

    def item(self, index: int) -> QListWidgetItem | None:
        if 0 <= index < len(self._items):
            return self._items[index]
        return None

    def row(self, item: QListWidgetItem) -> int:
        try:
            return self._items.index(item)
        except ValueError:
            return -1

    def takeItem(self, index: int) -> QListWidgetItem | None:
        if not 0 <= index < len(self._items):
            return None
        item = self._items.pop(index)
        widget = self._widgets.pop(self._item_key(item), None)
        if widget is not None:
            self.rows_layout.removeWidget(widget)
        self._rebuild_layout()
        return item

    def clear(self) -> None:
        if self._drag_source_item is not None:
            self._finish_drag(commit=False)

        while self.rows_layout.count():
            layout_item = self.rows_layout.takeAt(0)
            widget = layout_item.widget()
            if widget is not None:
                widget.setParent(None)

        for widget in self._widgets.values():
            widget.deleteLater()

        self._items.clear()
        self._widgets.clear()
        self.updateGeometry()

    def setSpacing(self, spacing: int) -> None:
        self.rows_layout.setSpacing(spacing)
        self.updateGeometry()

    def sizeHint(self) -> QSize:
        width = 900
        height = self.rows_layout.contentsMargins().top() + self.rows_layout.contentsMargins().bottom()
        visible_count = 0
        for item in self._items:
            widget = self._widgets.get(self._item_key(item))
            if widget is None:
                continue
            height += max(widget.minimumHeight(), widget.sizeHint().height())
            visible_count += 1
        if visible_count > 1:
            height += self.rows_layout.spacing() * (visible_count - 1)
        return QSize(width, max(1, height))

    def minimumSizeHint(self) -> QSize:
        return self.sizeHint()

    @staticmethod
    def _refresh_widget_style(widget: QWidget | None) -> None:
        if widget is None:
            return
        widget.style().unpolish(widget)
        widget.style().polish(widget)
        widget.update()

    def _create_placeholder(self, source_widget: QWidget) -> QFrame:
        placeholder = QFrame(self)
        placeholder.setObjectName("SortDropPlaceholder")
        placeholder.setFixedHeight(max(46, source_widget.height(), source_widget.sizeHint().height()))
        placeholder_layout = QHBoxLayout(placeholder)
        placeholder_layout.setContentsMargins(12, 0, 12, 0)
        placeholder_label = QLabel("Release to place sorting criterion")
        placeholder_label.setAlignment(Qt.AlignCenter)
        placeholder_label.setAttribute(Qt.WA_TransparentForMouseEvents)
        placeholder_layout.addWidget(placeholder_label)
        return placeholder

    def _clear_layout_only(self) -> None:
        while self.rows_layout.count():
            self.rows_layout.takeAt(0)

    def _rebuild_layout(self) -> None:
        self._clear_layout_only()

        for item in self._items:
            if item is self._drag_source_item and self._drag_placeholder is not None:
                self.rows_layout.addWidget(self._drag_placeholder)
                continue

            widget = self._widgets.get(self._item_key(item))
            if widget is not None:
                widget.show()
                self.rows_layout.addWidget(widget)

        self.rows_layout.invalidate()
        self.rows_layout.activate()
        self.updateGeometry()
        self.adjustSize()
        self.update()

    def _mime_row_id(self, mime_data: QMimeData) -> int | None:
        if not mime_data.hasFormat(SORT_ROW_MIME_TYPE):
            return None
        try:
            return int(bytes(mime_data.data(SORT_ROW_MIME_TYPE)).decode("utf-8"))
        except (TypeError, ValueError, UnicodeDecodeError):
            return None

    def _source_row_id(self) -> int | None:
        if self._drag_source_item is None:
            return None
        try:
            return int(self._drag_source_item.data(Qt.UserRole))
        except (TypeError, ValueError):
            return None

    def _target_index_for_y(self, y_position: int) -> int:
        if self._drag_source_item is None:
            return 0

        remaining = [item for item in self._items if item is not self._drag_source_item]
        insertion_index = 0

        for item in remaining:
            widget = self._widgets.get(self._item_key(item))
            if widget is None or not widget.isVisible():
                continue
            if y_position < widget.geometry().center().y():
                return insertion_index
            insertion_index += 1

        return len(remaining)

    def _preview_source_at(self, insertion_index: int) -> None:
        source_item = self._drag_source_item
        if source_item is None:
            return

        remaining = [item for item in self._items if item is not source_item]
        insertion_index = max(0, min(int(insertion_index), len(remaining)))
        new_items = list(remaining)
        new_items.insert(insertion_index, source_item)

        if new_items == self._items:
            return

        self._items = new_items
        self._rebuild_layout()
        self.orderPreviewed.emit()

    def begin_drag_for_item(self, item: QListWidgetItem) -> None:
        if item not in self._items or self._drag_source_item is not None:
            return

        source_widget = self._widgets.get(self._item_key(item))
        if source_widget is None:
            return

        try:
            row_id = int(item.data(Qt.UserRole))
        except (TypeError, ValueError):
            return

        source_pixmap = source_widget.grab()
        self._drag_source_item = item
        self._drag_original_items = list(self._items)
        self._drag_placeholder = self._create_placeholder(source_widget)
        self._drag_committed = False

        source_widget.setProperty("dragging", True)
        self._refresh_widget_style(source_widget)
        source_widget.hide()
        self._rebuild_layout()

        mime_data = QMimeData()
        mime_data.setData(SORT_ROW_MIME_TYPE, str(row_id).encode("utf-8"))

        drag = QDrag(self)
        drag.setMimeData(mime_data)

        padding = 10
        ghost = QPixmap(
            source_pixmap.width() + padding * 2,
            source_pixmap.height() + padding * 2,
        )
        ghost.fill(Qt.transparent)

        painter = QPainter(ghost)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(0, 0, 0, 110))
        painter.drawRoundedRect(
            padding,
            padding + 4,
            source_pixmap.width(),
            source_pixmap.height(),
            8,
            8,
        )
        painter.setOpacity(0.96)
        painter.drawPixmap(padding, padding, source_pixmap)
        painter.end()

        drag.setPixmap(ghost)
        drag.setHotSpot(QPoint(min(40, ghost.width() // 3), ghost.height() // 2))

        result = drag.exec(Qt.MoveAction)
        if not self._drag_committed or result != Qt.MoveAction:
            self._finish_drag(commit=False)

    def _finish_drag(self, *, commit: bool) -> None:
        source_item = self._drag_source_item
        if source_item is None:
            return

        source_widget = self._widgets.get(self._item_key(source_item))

        if not commit:
            self._items = list(self._drag_original_items)

        placeholder = self._drag_placeholder
        self._drag_placeholder = None
        self._drag_source_item = None

        if placeholder is not None:
            self.rows_layout.removeWidget(placeholder)
            placeholder.deleteLater()

        if source_widget is not None:
            source_widget.setProperty("dragging", False)
            self._refresh_widget_style(source_widget)
            source_widget.show()

        self._rebuild_layout()

        self._drag_original_items = []
        if commit:
            self._drag_committed = True
            self.orderCommitted.emit()
        else:
            self._drag_committed = False
            self.orderCancelled.emit()

    def dragEnterEvent(self, event: Any) -> None:
        row_id = self._mime_row_id(event.mimeData())
        if row_id is None or row_id != self._source_row_id():
            event.ignore()
            return
        event.setDropAction(Qt.MoveAction)
        event.accept()

    def dragMoveEvent(self, event: Any) -> None:
        row_id = self._mime_row_id(event.mimeData())
        if row_id is None or row_id != self._source_row_id():
            event.ignore()
            return

        insertion_index = self._target_index_for_y(event.position().toPoint().y())
        self._preview_source_at(insertion_index)

        event.setDropAction(Qt.MoveAction)
        event.accept()

    def dropEvent(self, event: Any) -> None:
        row_id = self._mime_row_id(event.mimeData())
        if row_id is None or row_id != self._source_row_id():
            event.ignore()
            return

        self._finish_drag(commit=True)
        event.setDropAction(Qt.MoveAction)
        event.accept()


class SortCriteriaDragHandle(QWidget):
    """Six-dot handle that starts a polished live sorting-row drag."""

    def __init__(
        self,
        list_widget: SortCriteriaListWidget,
        item: QListWidgetItem,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.list_widget = list_widget
        self.item = item
        self.drag_start_position: QPoint | None = None
        self._hovered = False
        self._pressed = False

        self.setFixedSize(28, 38)
        self.setToolTip("Drag to reorder this sorting criterion.")
        self.setCursor(Qt.OpenHandCursor)

    def paintEvent(self, event: Any) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        if self._pressed:
            background = QColor(255, 181, 0, 105)
        elif self._hovered:
            background = QColor(255, 181, 0, 58)
        else:
            background = QColor(255, 181, 0, 25)

        painter.setPen(Qt.NoPen)
        painter.setBrush(background)
        painter.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), 5, 5)

        painter.setBrush(QColor(255, 181, 0))
        for x in (10, 18):
            for y in (11, 19, 27):
                painter.drawEllipse(QPoint(x, y), 2, 2)
        painter.end()

    def enterEvent(self, event: Any) -> None:
        self._hovered = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event: Any) -> None:
        self._hovered = False
        self._pressed = False
        self.setCursor(Qt.OpenHandCursor)
        self.update()
        super().leaveEvent(event)

    def mousePressEvent(self, event: Any) -> None:
        if event.button() == Qt.LeftButton:
            self.drag_start_position = event.position().toPoint()
            self._pressed = True
            self.setCursor(Qt.ClosedHandCursor)
            self.update()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: Any) -> None:
        if not (event.buttons() & Qt.LeftButton):
            super().mouseMoveEvent(event)
            return

        if self.drag_start_position is None:
            self.drag_start_position = event.position().toPoint()

        distance = (
            event.position().toPoint() - self.drag_start_position
        ).manhattanLength()
        if distance >= QApplication.startDragDistance():
            self._pressed = False
            self.setCursor(Qt.OpenHandCursor)
            self.update()
            self.list_widget.begin_drag_for_item(self.item)
            self.drag_start_position = None

        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: Any) -> None:
        self.drag_start_position = None
        self._pressed = False
        self.setCursor(Qt.OpenHandCursor)
        self.update()
        super().mouseReleaseEvent(event)


# ---------------------------------------------------------------------------
# Main GUI
# ---------------------------------------------------------------------------

class BidGUI(QMainWindow):
    def __init__(self) -> None:
        set_windows_app_id()
        super().__init__()

        self.setWindowTitle("UPS Bid Analyzer")
        self.resize(1080, 900)
        self.setMinimumSize(960, 640)
        apply_window_icon(self)

        self.config_data = load_saved_config()
        self.message_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self.preview_df: pd.DataFrame | None = None
        self.cached_lines: dict[str, Any] | None = None
        self.cached_trips: dict[str, Any] | None = None
        self.cached_pdf_key: tuple[str, str] | None = None
        self.cached_bid_period_key: tuple[str, str] | None = None
        self.cached_bid_period: str | None = None
        self.cached_airport_lookup_key: tuple[str, str] | None = None
        self.cached_airport_lookup: dict[str, Any] | None = None
        self.cached_unmatched_airports: Any = None
        self.cached_matched_airports_df: pd.DataFrame | None = None

        self.sort_order: list[list[str]] = []
        self.latest_bid_string = ""
        self.visualizer_windows: list[QWidget] = []

        self._loading_saved_values = True
        self.preference_refresh_pending = False

        self._setup_style()
        self._build_ui()

        self.preference_refresh_timer = QTimer(self)
        self.preference_refresh_timer.setSingleShot(True)
        self.preference_refresh_timer.timeout.connect(self._refresh_after_preference_change)

        self._load_saved_values()
        self._loading_saved_values = False

        self.queue_timer = QTimer(self)
        self.queue_timer.timeout.connect(self._poll_queue)
        self.queue_timer.start(100)

    # -------------------------- UI construction --------------------------

    def _setup_style(self) -> None:
        self.setStyleSheet(f"""
            QMainWindow, QWidget {{
                background: {UPS_BROWN};
                color: {UPS_TEXT};
                font-family: Segoe UI, Arial, sans-serif;
                font-size: 10pt;
            }}
            QScrollArea {{
                border: none;
                background: {UPS_BROWN};
            }}
            QScrollArea#MainScrollArea QScrollBar:vertical {{
                background: white;
                width: 18px;
                margin: 0;
                border: none;
            }}
            QScrollArea#MainScrollArea QScrollBar::handle:vertical {{
                background: #9A9A9A;
                min-height: 30px;
                margin: 2px;
                border-radius: 5px;
            }}
            QScrollArea#MainScrollArea QScrollBar::handle:vertical:hover {{
                background: #777777;
            }}
            QScrollArea#MainScrollArea QScrollBar::add-line:vertical,
            QScrollArea#MainScrollArea QScrollBar::sub-line:vertical {{
                height: 0;
                border: none;
                background: white;
            }}
            QScrollArea#MainScrollArea QScrollBar::add-page:vertical,
            QScrollArea#MainScrollArea QScrollBar::sub-page:vertical {{
                background: white;
            }}
            QGroupBox {{
                border: 2px solid {UPS_GOLD};
                border-radius: 6px;
                margin-top: 14px;
                padding: 10px;
                color: {UPS_TEXT};
                font-weight: bold;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
                color: {UPS_GOLD};
                background: {UPS_BROWN};
            }}
            QLabel#TitleLabel {{
                color: {UPS_GOLD};
                font-size: 20pt;
                font-weight: bold;
            }}
            QLabel#SubtitleLabel {{
                color: {UPS_TEXT};
            }}
            QLabel#ExportTitle {{
                font-size: 11pt;
                font-weight: bold;
            }}
            QLineEdit, QDoubleSpinBox, QTextEdit, QListWidget, QTableWidget, QComboBox {{
                background: white;
                color: black;
                selection-background-color: {UPS_BLUE};
                selection-color: white;
            }}
            QComboBox QAbstractItemView {{
                background-color: white;
                color: black;
                selection-background-color: {UPS_BLUE};
                selection-color: white;
                border: 1px solid #888888;
                outline: 0;
            }}
            QListWidget#LineTypePreferenceList {{
                background: white;
                color: black;
                border: 1px solid #8A8A8A;
                border-radius: 4px;
                padding: 2px;
            }}
            QListWidget#LineTypePreferenceList::item {{
                min-height: 18px;
                padding: 1px 7px;
                border-bottom: 1px solid #E5E5E5;
            }}
            QListWidget#LineTypePreferenceList::item:selected {{
                background: {UPS_BLUE};
                color: white;
            }}
            QFrame#PreferenceDateCard {{
                background-color: {UPS_BROWN_2};
                border: 1px solid #71402E;
                border-radius: 8px;
            }}
            QLabel#PreferenceCardTitle {{
                color: {UPS_GOLD};
                font-size: 11pt;
                font-weight: bold;
            }}
            QLabel#PreferenceCardSubtitle {{
                color: #E9DDD2;
            }}
            QLabel#PreferenceCardCount {{
                color: {UPS_BROWN};
                background: {UPS_GOLD};
                border-radius: 9px;
                padding: 2px 8px;
                font-weight: bold;
            }}
            QTableWidget#PreferenceDateTable {{
                background: white;
                alternate-background-color: #F4F4F4;
                color: black;
                border: 1px solid #C7B9AE;
                border-radius: 5px;
                gridline-color: #E5E5E5;
            }}
            QTableWidget#PreferenceDateTable::item {{
                padding: 4px;
            }}
            QWidget#SortCriteriaList {{
                background: transparent;
                border: none;
            }}
            QWidget#SortCriterionCard {{
                background-color: {UPS_BROWN_2};
                border: 1px solid #6A3A29;
                border-radius: 7px;
            }}
            QWidget#SortCriterionCard:hover {{
                background-color: #552C1D;
                border: 1px solid {UPS_GOLD};
            }}
            QWidget#SortCriterionCard[dragging="true"] {{
                background-color: #5D382A;
                border: 2px dashed {UPS_GOLD};
            }}
            QFrame#SortDropPlaceholder {{
                background-color: rgba(255, 181, 0, 32);
                border: 2px dashed {UPS_GOLD};
                border-radius: 7px;
                color: {UPS_GOLD};
                font-weight: bold;
            }}
            QTextEdit#LogText {{
                background: {UPS_FIELD_BG};
            }}
            QHeaderView::section {{
                background: {UPS_GOLD};
                color: {UPS_BROWN};
                font-weight: bold;
                padding: 4px;
                border: 1px solid {UPS_BROWN_2};
            }}
            QPushButton {{
                background: {UPS_BLUE};
                color: white;
                border: 1px solid {UPS_BLUE};
                border-radius: 4px;
                padding: 6px 10px;
            }}
            QPushButton:hover {{
                background: {UPS_BLUE_ACTIVE};
            }}
            QPushButton:disabled {{
                background: #777777;
                color: #dddddd;
                border-color: #777777;
            }}
            QPushButton#GreenButton {{
                background: {UPS_GREEN};
                border-color: {UPS_GREEN};
                font-weight: bold;
            }}
            QPushButton#GreenButton:hover {{
                background: {UPS_GREEN_ACTIVE};
            }}
            QProgressBar {{
                border: 1px solid {UPS_GOLD};
                border-radius: 4px;
                background: {UPS_BROWN_2};
                color: {UPS_TEXT};
                text-align: center;
            }}
            QProgressBar::chunk {{
                background: {UPS_GOLD};
            }}
        """)

        # Combo-box popup lists are separate Qt windows on some platforms.
        # Applying the same stylesheet at application level keeps every popup
        # white and readable, including dropdowns opened from dialogs.
        app = QApplication.instance()
        if app is not None:
            app.setStyleSheet(self.styleSheet())

    def _build_ui(self) -> None:
        scroll_area = QScrollArea(self)
        scroll_area.setObjectName("MainScrollArea")
        scroll_area.setWidgetResizable(True)
        scroll_area.verticalScrollBar().setStyleSheet("""
            QScrollBar:vertical {
                background: white;
                width: 18px;
                margin: 0;
                border: none;
            }
            QScrollBar::handle:vertical {
                background: #A0A0A0;
                min-height: 34px;
                margin: 2px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical:hover {
                background: #777777;
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0;
                border: none;
                background: white;
            }
            QScrollBar::add-page:vertical,
            QScrollBar::sub-page:vertical {
                background: white;
            }
        """)
        self.setCentralWidget(scroll_area)

        self.scrollable_frame = QWidget()
        scroll_area.setWidget(self.scrollable_frame)

        container = QVBoxLayout(self.scrollable_frame)
        container.setContentsMargins(14, 14, 14, 14)
        container.setSpacing(10)

        header_frame = QWidget()
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(0, 0, 0, 0)
        title = QLabel("UPS Bid Analyzer")
        title.setObjectName("TitleLabel")
        subtitle = QLabel("Load bid-package PDFs, apply your scoring preferences, and export the ranked Excel file.")
        subtitle.setObjectName("SubtitleLabel")
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        container.addWidget(header_frame)

        self._build_pdf_section(container)
        self._build_preferences_section(container)
        self._build_sorting_section(container)
        self._build_bid_string_section(container)
        self._build_output_section(container)
        self._build_action_section(container)
        self._build_status_section(container)

    def _build_pdf_section(self, container: QVBoxLayout) -> None:
        file_frame = QGroupBox("PDF Files")
        grid = QGridLayout(file_frame)
        grid.setColumnStretch(1, 1)

        self.trips_path_edit = QLineEdit()
        self.trips_path_edit.textChanged.connect(lambda _text: self._mark_pdf_paths_changed())
        trips_browse = QPushButton("Browse")
        trips_browse.clicked.connect(self._browse_trips)

        self.lines_path_edit = QLineEdit()
        self.lines_path_edit.textChanged.connect(lambda _text: self._mark_pdf_paths_changed())
        lines_browse = QPushButton("Browse")
        lines_browse.clicked.connect(self._browse_lines)

        self.load_button = QPushButton("Load PDFs into UPS Bid Analyzer")
        self.load_button.clicked.connect(self.load_pdfs)
        self.load_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)

        grid.addWidget(QLabel("TRIPS PDF:"), 0, 0)
        grid.addWidget(self.trips_path_edit, 0, 1)
        grid.addWidget(trips_browse, 0, 2)
        grid.addWidget(self.load_button, 0, 3, 2, 1)

        grid.addWidget(QLabel("LINES PDF:"), 1, 0)
        grid.addWidget(self.lines_path_edit, 1, 1)
        grid.addWidget(lines_browse, 1, 2)

        self.pdf_status_label = QLabel("PDFs not loaded yet.")
        grid.addWidget(self.pdf_status_label, 2, 0, 1, 4)

        self.trip_progress = QProgressBar()
        self.trip_progress.setRange(0, 100)
        self.trip_progress.setValue(0)
        grid.addWidget(self.trip_progress, 3, 0, 1, 4)

        self.trip_progress_text_label = QLabel("Trips extraction progress: not started.")
        grid.addWidget(self.trip_progress_text_label, 4, 0, 1, 4)

        container.addWidget(file_frame)

    def _build_preferences_section(self, container: QVBoxLayout) -> None:
        prefs_frame = QGroupBox("Preferences")
        main_layout = QVBoxLayout(prefs_frame)
        main_layout.setSpacing(12)

        # Vacation ranges -------------------------------------------------
        vacation_card = QFrame()
        vacation_card.setObjectName("PreferenceDateCard")
        vacation_card_layout = QVBoxLayout(vacation_card)
        vacation_card_layout.setContentsMargins(10, 9, 10, 10)
        vacation_card_layout.setSpacing(8)

        vacation_header = QHBoxLayout()
        vacation_title_area = QVBoxLayout()
        vacation_title_area.setSpacing(1)
        vacation_title = QLabel("Vacation ranges")
        vacation_title.setObjectName("PreferenceCardTitle")
        vacation_subtitle = QLabel(
            "Add vacation periods and choose whether OCV / pay-period drop applies to each range."
        )
        vacation_subtitle.setObjectName("PreferenceCardSubtitle")
        vacation_subtitle.setWordWrap(True)
        vacation_title_area.addWidget(vacation_title)
        vacation_title_area.addWidget(vacation_subtitle)

        add_vacation = QPushButton("Add range")
        edit_vacation = QPushButton("Edit")
        remove_vacation = QPushButton("Remove")
        clear_vacation = QPushButton("Clear")
        add_vacation.clicked.connect(self._add_vacation_range)
        edit_vacation.clicked.connect(self._edit_vacation_range)
        remove_vacation.clicked.connect(self._remove_vacation_range)
        clear_vacation.clicked.connect(self._clear_vacation_ranges)

        vacation_header.addLayout(vacation_title_area, 1)
        for button in (add_vacation, edit_vacation, remove_vacation, clear_vacation):
            vacation_header.addWidget(button)

        self.vacation_table = QTableWidget(0, 3)
        self.vacation_table.setObjectName("PreferenceDateTable")
        self.vacation_table.setHorizontalHeaderLabels(["Start", "End", "OCV"])
        self.vacation_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.vacation_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.vacation_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.vacation_table.verticalHeader().setVisible(False)
        self.vacation_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.vacation_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.vacation_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.vacation_table.setAlternatingRowColors(True)
        self.vacation_table.setShowGrid(False)
        self.vacation_table.setMinimumHeight(125)
        self.vacation_table.itemChanged.connect(self._on_vacation_table_item_changed)

        vacation_card_layout.addLayout(vacation_header)
        vacation_card_layout.addWidget(self.vacation_table)
        main_layout.addWidget(vacation_card)

        # Requested days --------------------------------------------------
        requested_card = QFrame()
        requested_card.setObjectName("PreferenceDateCard")
        requested_card_layout = QVBoxLayout(requested_card)
        requested_card_layout.setContentsMargins(10, 9, 10, 10)
        requested_card_layout.setSpacing(8)

        requested_header = QHBoxLayout()
        requested_title_area = QVBoxLayout()
        requested_title_area.setSpacing(1)
        requested_title = QLabel("Requested days off")
        requested_title.setObjectName("PreferenceCardTitle")
        requested_subtitle = QLabel(
            "Add a single date or range. Notes are saved locally and never affect scoring."
        )
        requested_subtitle.setObjectName("PreferenceCardSubtitle")
        requested_subtitle.setWordWrap(True)
        requested_title_area.addWidget(requested_title)
        requested_title_area.addWidget(requested_subtitle)

        add_requested = QPushButton("Add day or range")
        edit_requested = QPushButton("Edit")
        remove_requested = QPushButton("Remove")
        clear_requested = QPushButton("Clear")
        add_requested.clicked.connect(self._add_requested_date_range)
        edit_requested.clicked.connect(self._edit_requested_date_range)
        remove_requested.clicked.connect(self._remove_requested_date_range)
        clear_requested.clicked.connect(self._clear_requested_date_ranges)

        requested_header.addLayout(requested_title_area, 1)
        for button in (add_requested, edit_requested, remove_requested, clear_requested):
            requested_header.addWidget(button)

        self.requested_dates_table = QTableWidget(0, 3)
        self.requested_dates_table.setObjectName("PreferenceDateTable")
        self.requested_dates_table.setHorizontalHeaderLabels(
            ["Notes", "Date / Start", "End (optional)"]
        )
        self.requested_dates_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.requested_dates_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.requested_dates_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.requested_dates_table.verticalHeader().setVisible(False)
        self.requested_dates_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.requested_dates_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.requested_dates_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.requested_dates_table.setAlternatingRowColors(True)
        self.requested_dates_table.setShowGrid(False)
        self.requested_dates_table.setMinimumHeight(125)

        requested_card_layout.addLayout(requested_header)
        requested_card_layout.addWidget(self.requested_dates_table)
        main_layout.addWidget(requested_card)

        # Lower preferences: normal inputs on the left, line types on right.
        lower_area = QWidget()
        lower_layout = QHBoxLayout(lower_area)
        lower_layout.setContentsMargins(0, 2, 0, 0)
        lower_layout.setSpacing(28)

        left_preferences = QWidget()
        left_grid = QGridLayout(left_preferences)
        left_grid.setContentsMargins(0, 0, 0, 0)
        left_grid.setHorizontalSpacing(10)
        left_grid.setVerticalSpacing(10)
        left_grid.setColumnStretch(1, 1)

        self.training_start_entry = DateEntry(self)
        self.training_end_entry = DateEntry(self)
        self.training_start_entry.line_edit.textChanged.connect(
            lambda: self._on_date_preferences_changed(
                "Training dates changed. Sorting columns will refresh automatically."
            )
        )
        self.training_end_entry.line_edit.textChanged.connect(
            lambda: self._on_date_preferences_changed(
                "Training dates changed. Sorting columns will refresh automatically."
            )
        )

        training_row = QWidget()
        training_layout = QHBoxLayout(training_row)
        training_layout.setContentsMargins(0, 0, 0, 0)
        training_layout.setSpacing(8)
        training_layout.addWidget(QLabel("Start:"))
        training_layout.addWidget(self.training_start_entry)
        training_layout.addSpacing(10)
        training_layout.addWidget(QLabel("End:"))
        training_layout.addWidget(self.training_end_entry)
        training_layout.addStretch(1)

        self.bid_edge_combo = QComboBox()
        self.bid_edge_combo.addItems(["none", "start", "end", "both"])
        self.bid_edge_combo.currentTextChanged.connect(lambda _text: self._on_bid_edge_changed())

        self.hourly_rate_edit = QDoubleSpinBox()
        self.hourly_rate_edit.setPrefix("$")
        self.hourly_rate_edit.setDecimals(2)
        self.hourly_rate_edit.setRange(0.01, 10000.00)
        self.hourly_rate_edit.setSingleStep(1.00)
        self.hourly_rate_edit.setValue(DEFAULT_HOURLY_RATE)
        self.hourly_rate_edit.setMaximumWidth(125)
        self.hourly_rate_edit.editingFinished.connect(self._on_hourly_rate_changed)

        left_grid.addWidget(QLabel("Training dates:"), 0, 0)
        left_grid.addWidget(training_row, 0, 1)
        left_grid.addWidget(QLabel("Bid edge days off:"), 1, 0)
        left_grid.addWidget(self.bid_edge_combo, 1, 1, alignment=Qt.AlignLeft)
        left_grid.addWidget(QLabel("Hourly pay rate:"), 2, 0)
        left_grid.addWidget(self.hourly_rate_edit, 2, 1, alignment=Qt.AlignLeft)
        left_grid.setRowStretch(3, 1)

        right_preferences = QWidget()
        right_grid = QGridLayout(right_preferences)
        right_grid.setContentsMargins(0, 0, 0, 0)
        right_grid.setHorizontalSpacing(10)

        self.line_type_preference_list = NoInternalScrollListWidget()
        self.line_type_preference_list.setObjectName("LineTypePreferenceList")
        self.line_type_preference_list.setSpacing(1)
        self.line_type_preference_list.setMinimumWidth(145)
        self.line_type_preference_list.setMaximumWidth(180)
        self.line_type_preference_list.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.line_type_preference_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.line_type_preference_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.line_type_preference_list.setDragEnabled(True)
        self.line_type_preference_list.viewport().setAcceptDrops(True)
        self.line_type_preference_list.setDropIndicatorShown(True)
        self.line_type_preference_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.line_type_preference_list.setDefaultDropAction(Qt.MoveAction)
        self.line_type_preference_list.setDragDropOverwriteMode(False)
        self.line_type_preference_list.model().rowsMoved.connect(
            lambda *_args: self._line_type_preference_changed()
        )

        line_type_controls = QVBoxLayout()
        reset_type_order = QPushButton("Reset order")
        reset_type_order.clicked.connect(self._reset_line_type_preference_order)
        line_type_controls.addWidget(reset_type_order)
        line_type_controls.addStretch(1)

        line_type_help = QLabel("Drag items to reorder.\nTop = most preferred.")
        line_type_help.setWordWrap(True)
        line_type_help.setMaximumWidth(170)

        right_grid.addWidget(QLabel("Line-type preference order:"), 0, 0, alignment=Qt.AlignTop)
        right_grid.addWidget(self.line_type_preference_list, 0, 1, 2, 1)
        right_grid.addLayout(line_type_controls, 0, 2, 2, 1)
        right_grid.addWidget(line_type_help, 1, 0, alignment=Qt.AlignTop)

        lower_layout.addWidget(left_preferences, 1)
        lower_layout.addWidget(right_preferences, 1)
        main_layout.addWidget(lower_area)

        container.addWidget(prefs_frame)

    def _build_sorting_section(self, container: QVBoxLayout) -> None:
        sort_frame = QGroupBox("Sorting")
        main_layout = QVBoxLayout(sort_frame)
        main_layout.setSpacing(8)

        # These settings are used immediately by the live contribution labels.
        self.default_mode = str(DEFAULT_SORTING_SETTINGS["default_mode"])
        self.weighting_style = str(DEFAULT_SORTING_SETTINGS["weighting_style"])
        self.soft_max_weight = float(DEFAULT_SORTING_SETTINGS["soft_max_weight"])
        self.soft_min_weight = float(DEFAULT_SORTING_SETTINGS["soft_min_weight"])
        self.keep_score_columns = bool(DEFAULT_SORTING_SETTINGS["keep_score_columns"])

        self.sortable_columns: list[str] = []
        self.sort_criteria_rows: list[dict[str, Any]] = []
        self._sort_row_by_id: dict[int, dict[str, Any]] = {}
        self._next_sort_row_id = 1
        self._rebuilding_sort_rows = False

        # Intuitive, immediately visible control for the top-to-bottom
        # contribution curve used by soft weighting.
        emphasis_widget = QWidget()
        emphasis_layout = QHBoxLayout(emphasis_widget)
        emphasis_layout.setContentsMargins(0, 0, 0, 0)
        emphasis_layout.setSpacing(8)

        emphasis_label = QLabel("Priority emphasis:")
        self.priority_emphasis_spin = QDoubleSpinBox()
        self.priority_emphasis_spin.setDecimals(2)
        self.priority_emphasis_spin.setRange(self.soft_min_weight, 100.0)
        self.priority_emphasis_spin.setSingleStep(0.25)
        self.priority_emphasis_spin.setValue(self.soft_max_weight)
        self.priority_emphasis_spin.setMaximumWidth(95)
        self.priority_emphasis_spin.setToolTip(
            "Controls how strongly earlier Weighted criteria outrank later ones."
        )
        self.priority_emphasis_spin.valueChanged.connect(
            self._on_priority_emphasis_changed
        )

        emphasis_help = QLabel(
            "Bigger numbers give more contribution to top priorities.\n"
            "Smaller numbers give lower-priority criteria more influence."
        )
        emphasis_help.setWordWrap(True)

        emphasis_layout.addWidget(emphasis_label)
        emphasis_layout.addWidget(self.priority_emphasis_spin)
        emphasis_layout.addWidget(emphasis_help, 1)

        header_widget = QWidget()
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(8)

        drag_header = QLabel("")
        drag_header.setFixedWidth(24)

        number_header = QLabel("#")
        number_header.setFixedWidth(28)
        number_header.setAlignment(Qt.AlignCenter)

        column_header = QLabel("Sort by")
        direction_header = QLabel("Direction")
        mode_header = QLabel("Priority method")
        contribution_header = QLabel("Contribution")
        contribution_header.setAlignment(Qt.AlignCenter)

        column_header.setFixedWidth(245)
        direction_header.setFixedWidth(130)
        mode_header.setFixedWidth(155)
        contribution_header.setFixedWidth(100)

        header_layout.addWidget(drag_header)
        header_layout.addWidget(number_header)
        header_layout.addWidget(column_header)
        header_layout.addWidget(direction_header)
        remove_header = QLabel("")
        remove_header.setFixedWidth(78)

        header_layout.addWidget(mode_header)
        header_layout.addWidget(contribution_header)
        header_layout.addSpacing(4)
        header_layout.addWidget(remove_header)
        header_layout.addStretch(1)

        self.sort_rows_list = SortCriteriaListWidget()
        self.sort_rows_list.setObjectName("SortCriteriaList")
        self.sort_rows_list.setSpacing(5)
        self.sort_rows_list.orderPreviewed.connect(self._on_sort_rows_previewed)
        self.sort_rows_list.orderCommitted.connect(self._on_sort_rows_reordered)
        self.sort_rows_list.orderCancelled.connect(self._on_sort_rows_previewed)

        controls = QHBoxLayout()
        add_criterion_button = QPushButton("Add sorting criterion")
        clear_criteria_button = QPushButton("Clear sorting")
        advanced_button = QPushButton("Advanced settings...")

        add_criterion_button.clicked.connect(
            lambda _checked=False: self._add_sort_criteria_row()
        )
        clear_criteria_button.clicked.connect(self._clear_sort_order)
        advanced_button.clicked.connect(self._open_sorting_settings_dialog)

        controls.addWidget(add_criterion_button)
        controls.addWidget(clear_criteria_button)
        controls.addStretch(1)
        controls.addWidget(advanced_button)

        help_label = QLabel(
            "Drag the gold handle to pick up a criterion; the other rows move out of the way as you hover.\n"
            "High Priority is a strict tie-breaker. Weighted starts a new weight level. "
            "Equal to Previous shares both the previous weight and its number."
        )
        help_label.setWordWrap(True)

        main_layout.addWidget(emphasis_widget)
        main_layout.addWidget(header_widget)
        main_layout.addWidget(self.sort_rows_list)
        main_layout.addLayout(controls)
        main_layout.addWidget(help_label)

        for _ in range(MIN_SORT_CRITERIA_ROWS):
            self._add_sort_criteria_row(notify=False)

        container.addWidget(sort_frame)

    def _build_output_section(self, container: QVBoxLayout) -> None:
        output_frame = QGroupBox("Output")
        grid = QGridLayout(output_frame)
        grid.setColumnStretch(1, 1)

        self.output_folder_edit = QLineEdit()
        output_browse = QPushButton("Browse")
        output_browse.clicked.connect(self._browse_output_folder)

        self.output_filename_edit = QLineEdit("Bid_Results")
        extension_label = QLabel(".xlsx")

        filename_row = QWidget()
        filename_layout = QHBoxLayout(filename_row)
        filename_layout.setContentsMargins(0, 0, 0, 0)
        filename_layout.addWidget(self.output_filename_edit)
        filename_layout.addWidget(extension_label)
        filename_layout.addStretch(1)

        grid.addWidget(QLabel("Output folder:"), 0, 0)
        grid.addWidget(self.output_folder_edit, 0, 1)
        grid.addWidget(output_browse, 0, 2)
        grid.addWidget(QLabel("File name:"), 1, 0)
        grid.addWidget(filename_row, 1, 1, 1, 2)

        container.addWidget(output_frame)

    def _build_bid_string_section(self, container: QVBoxLayout) -> None:
        bid_frame = QGroupBox("Copy Line Numbers")
        layout = QVBoxLayout(bid_frame)

        number_row = QHBoxLayout()
        number_row.addWidget(QLabel("Number of lines you would like to bid:"))
        self.number_of_lines_edit = QLineEdit(str(DEFAULT_NUMBER_OF_LINES_TO_BID))
        self.number_of_lines_edit.setMaximumWidth(90)
        self.number_of_lines_edit.textChanged.connect(lambda _text: self._mark_bid_string_stale())
        number_row.addWidget(self.number_of_lines_edit)
        number_row.addStretch(1)
        layout.addLayout(number_row)

        bid_string_row = QHBoxLayout()
        bid_string_row.addWidget(QLabel("Bid string:"), alignment=Qt.AlignTop)
        self.bid_string_text = QTextEdit()
        self.bid_string_text.setReadOnly(True)
        self.bid_string_text.setMinimumHeight(100)
        bid_string_row.addWidget(self.bid_string_text, 1)
        layout.addLayout(bid_string_row)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.generate_bid_string_button = QPushButton("Generate bid string")
        self.copy_bid_string_button = QPushButton("Copy bid")
        self.generate_bid_string_button.clicked.connect(self.generate_bid_string)
        self.copy_bid_string_button.clicked.connect(self.copy_bid_string)
        buttons.addWidget(self.generate_bid_string_button)
        buttons.addWidget(self.copy_bid_string_button)
        layout.addLayout(buttons)

        self.bid_string_status_label = QLabel("Bid string not generated yet.")
        layout.addWidget(self.bid_string_status_label)

        container.addWidget(bid_frame)

    def _build_action_section(self, container: QVBoxLayout) -> None:
        action_widget = QWidget()
        action_layout = QHBoxLayout(action_widget)
        action_layout.setContentsMargins(0, 0, 0, 0)

        self.export_button = QPushButton("Export Excel")
        self.export_button.setObjectName("GreenButton")
        self.export_button.clicked.connect(self.export_excel)

        self.visualizer_button = QPushButton("Open Visualizer")
        self.visualizer_button.clicked.connect(self.open_visualizer)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)

        action_layout.addWidget(self.export_button)
        action_layout.addWidget(self.visualizer_button)
        action_layout.addWidget(self.progress, 1)

        container.addWidget(action_widget)

    def _build_status_section(self, container: QVBoxLayout) -> None:
        log_frame = QGroupBox("Status")
        layout = QVBoxLayout(log_frame)
        self.log_text = QTextEdit()
        self.log_text.setObjectName("LogText")
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(180)
        layout.addWidget(self.log_text)
        container.addWidget(log_frame, 1)

    def _load_saved_values(self) -> None:
        self._set_vacation_ranges(self.config_data.get("vacation_ranges", []))
        self._set_requested_date_entries(self.config_data.get("requested_dates", []))

        self.training_start_entry.setText(self.config_data.get("training_start") or "")
        self.training_end_entry.setText(self.config_data.get("training_end") or "")

        saved_hourly_rate = self.config_data.get("hourly_rate", DEFAULT_HOURLY_RATE)
        try:
            saved_hourly_rate = validate_positive_float(str(saved_hourly_rate), "Hourly pay rate")
        except ValueError:
            saved_hourly_rate = DEFAULT_HOURLY_RATE
        self.hourly_rate_edit.setValue(saved_hourly_rate)

        saved_preference_order = self.config_data.get(
            "line_type_preference_order",
            DEFAULT_LINE_TYPE_PREFERENCE_ORDER,
        )
        self._set_line_type_preference_order(saved_preference_order)

        bid_edge = self.config_data.get("bid_edge") or "none"
        index = self.bid_edge_combo.findText(bid_edge)
        self.bid_edge_combo.blockSignals(True)
        self.bid_edge_combo.setCurrentIndex(index if index >= 0 else 0)
        self.bid_edge_combo.blockSignals(False)

        output_paths = self.config_data.get("output_paths", {})
        saved_output_folder = output_paths.get(get_os_name(), "")
        self.output_folder_edit.setText(saved_output_folder or str(Path.cwd()))

        saved_number_of_lines = self.config_data.get(
            "number_of_lines_to_bid",
            DEFAULT_NUMBER_OF_LINES_TO_BID,
        )
        try:
            saved_number_of_lines = int(saved_number_of_lines)
            if saved_number_of_lines <= 0:
                saved_number_of_lines = DEFAULT_NUMBER_OF_LINES_TO_BID
        except (TypeError, ValueError):
            saved_number_of_lines = DEFAULT_NUMBER_OF_LINES_TO_BID
        self.number_of_lines_edit.setText(str(saved_number_of_lines))

        # Load weighting settings before restoring rows so percentages are
        # calculated with the user's saved hard/soft/equal configuration.
        saved_sorting_settings = self.config_data.get("sorting_settings", {})
        if not isinstance(saved_sorting_settings, dict):
            saved_sorting_settings = {}
        merged_settings = {**DEFAULT_SORTING_SETTINGS, **saved_sorting_settings}
        self._apply_sorting_settings_to_ui(merged_settings)

        saved_sort_order = self.config_data.get("sort_order", [])
        self._set_sort_order(saved_sort_order)

    # -------------------------- Browse buttons --------------------------

    def _browse_trips(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select TRIPS PDF", "", "PDF files (*.pdf);;All files (*.*)")
        if path:
            self.trips_path_edit.setText(path)

    def _browse_lines(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select LINES PDF", "", "PDF files (*.pdf);;All files (*.*)")
        if path:
            self.lines_path_edit.setText(path)

    def _browse_output_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select output folder", self.output_folder_edit.text() or str(Path.cwd()))
        if path:
            self.output_folder_edit.setText(path)

    def _mark_pdf_paths_changed(self) -> None:
        current_key = (self.trips_path_edit.text().strip(), self.lines_path_edit.text().strip())
        if self.cached_pdf_key and current_key != self.cached_pdf_key:
            self.pdf_status_label.setText("PDF paths changed. Click Load PDFs again.")
            self.preview_df = None
            self.cached_bid_period_key = None
            self.cached_bid_period = None
            self.cached_airport_lookup_key = None
            self.cached_airport_lookup = None
            self.cached_unmatched_airports = None
            self.cached_matched_airports_df = None
            self._refresh_available_columns_list([])
            self._clear_bid_string("PDF paths changed. Generate the bid string again after loading/sorting.")

    # -------------------------- Vacation / requested-date actions --------------------------

    def _make_ocv_item(self, checked: bool) -> QTableWidgetItem:
        item = QTableWidgetItem()
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
        item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
        item.setTextAlignment(Qt.AlignCenter)
        return item

    def _update_date_section_counts(self) -> None:
        """Compatibility hook retained after removing the saved-count badges."""
        return

    def _append_vacation_row(self, start: str, end: str, pp_drop: bool = True) -> None:
        row = self.vacation_table.rowCount()
        self.vacation_table.insertRow(row)
        self.vacation_table.setItem(row, 0, QTableWidgetItem(start))
        self.vacation_table.setItem(row, 1, QTableWidgetItem(end))
        self.vacation_table.setItem(row, 2, self._make_ocv_item(pp_drop))

    def _set_vacation_ranges(self, vacation_ranges: list[dict[str, Any]] | None) -> None:
        self.vacation_table.blockSignals(True)
        try:
            self.vacation_table.setRowCount(0)
            for vacation in vacation_ranges or []:
                if not isinstance(vacation, dict):
                    continue
                start = str(vacation.get("start", "") or "")
                end = str(vacation.get("end", "") or "")
                pp_drop = bool(vacation.get("pp_drop", True))
                if start and end:
                    self._append_vacation_row(start, end, pp_drop)
        finally:
            self.vacation_table.blockSignals(False)
        self._update_date_section_counts()

    def _get_vacation_ranges(self) -> list[dict[str, Any]]:
        ranges: list[dict[str, Any]] = []
        for row in range(self.vacation_table.rowCount()):
            start_item = self.vacation_table.item(row, 0)
            end_item = self.vacation_table.item(row, 1)
            ocv_item = self.vacation_table.item(row, 2)
            start = validate_required_date(start_item.text() if start_item else "", "Vacation start")
            end = validate_required_date(end_item.text() if end_item else "", "Vacation end")
            if end < start:
                raise ValueError(f"Vacation range {start} to {end}: end date is before start date.")
            ranges.append({
                "start": start,
                "end": end,
                "pp_drop": bool(ocv_item and ocv_item.checkState() == Qt.Checked),
            })
        return ranges

    def _selected_vacation_row(self) -> int | None:
        rows = self.vacation_table.selectionModel().selectedRows()
        if not rows:
            return None
        return rows[0].row()

    def _add_vacation_range(self) -> None:
        dialog = VacationRangeDialog(self, "Add vacation range")
        if dialog.exec() == QDialog.Accepted and dialog.result:
            self.vacation_table.blockSignals(True)
            try:
                self._append_vacation_row(
                    dialog.result["start"],
                    dialog.result["end"],
                    bool(dialog.result.get("pp_drop", True)),
                )
            finally:
                self.vacation_table.blockSignals(False)
            self._update_date_section_counts()
            self._save_vacation_ranges_to_config()
            self._schedule_preference_refresh(
                "Vacation ranges changed. Sorting columns will refresh automatically."
            )

    def _edit_vacation_range(self) -> None:
        row = self._selected_vacation_row()
        if row is None:
            QMessageBox.information(self, "Edit vacation range", "Select a vacation range first.")
            return

        start = self.vacation_table.item(row, 0).text() if self.vacation_table.item(row, 0) else ""
        end = self.vacation_table.item(row, 1).text() if self.vacation_table.item(row, 1) else ""
        ocv_item = self.vacation_table.item(row, 2)
        pp_drop = bool(ocv_item and ocv_item.checkState() == Qt.Checked)
        dialog = VacationRangeDialog(self, "Edit vacation range", start, end, pp_drop)
        if dialog.exec() == QDialog.Accepted and dialog.result:
            self.vacation_table.blockSignals(True)
            try:
                self.vacation_table.setItem(row, 0, QTableWidgetItem(dialog.result["start"]))
                self.vacation_table.setItem(row, 1, QTableWidgetItem(dialog.result["end"]))
                self.vacation_table.setItem(
                    row,
                    2,
                    self._make_ocv_item(bool(dialog.result.get("pp_drop", True))),
                )
            finally:
                self.vacation_table.blockSignals(False)
            self._update_date_section_counts()
            self._save_vacation_ranges_to_config()
            self._schedule_preference_refresh(
                "Vacation ranges changed. Sorting columns will refresh automatically."
            )

    def _remove_vacation_range(self) -> None:
        row = self._selected_vacation_row()
        if row is not None:
            self.vacation_table.removeRow(row)
            self._update_date_section_counts()
            self._save_vacation_ranges_to_config()
            self._schedule_preference_refresh(
                "Vacation ranges changed. Sorting columns will refresh automatically."
            )

    def _clear_vacation_ranges(self) -> None:
        self.vacation_table.setRowCount(0)
        self._update_date_section_counts()
        self._save_vacation_ranges_to_config()
        self._schedule_preference_refresh(
            "Vacation ranges changed. Sorting columns will refresh automatically."
        )

    def _on_vacation_table_item_changed(self, _item: QTableWidgetItem) -> None:
        if self._loading_saved_values:
            return
        self._update_date_section_counts()
        self._save_vacation_ranges_to_config()
        self._schedule_preference_refresh(
            "Vacation OCV setting changed. Sorting columns will refresh automatically."
        )

    def _save_vacation_ranges_to_config(self) -> None:
        try:
            self.config_data["vacation_ranges"] = self._get_vacation_ranges() or None
            save_config(self.config_data)
        except Exception:
            pass

    # Requested single dates and ranges ---------------------------------

    def _append_requested_date_row(
        self,
        start: str,
        end: str = "",
        note: str = "",
    ) -> None:
        row = self.requested_dates_table.rowCount()
        self.requested_dates_table.insertRow(row)
        note_item = QTableWidgetItem(note)
        note_item.setToolTip(note)
        self.requested_dates_table.setItem(row, 0, note_item)
        self.requested_dates_table.setItem(row, 1, QTableWidgetItem(start))
        self.requested_dates_table.setItem(row, 2, QTableWidgetItem(end))
        self._update_date_section_counts()

    def _normalize_saved_requested_date_entries(self, entries: Any) -> list[dict[str, str]]:
        normalized: list[dict[str, str]] = []
        if entries is None:
            return normalized

        if isinstance(entries, (str, tuple)):
            entries = [entries]

        if not isinstance(entries, list):
            return normalized

        for entry in entries:
            note = ""
            start = ""
            end = ""
            if isinstance(entry, str):
                start = entry
            elif isinstance(entry, dict):
                note = str(entry.get("note", "") or "")
                start = str(entry.get("start", "") or "")
                end = str(entry.get("end", "") or "")
            elif isinstance(entry, (list, tuple)) and len(entry) == 2:
                # Backward compatibility with the old (start, end) format.
                start = str(entry[0] or "")
                end = str(entry[1] or "")
            elif isinstance(entry, (list, tuple)) and len(entry) >= 3:
                note = str(entry[0] or "")
                start = str(entry[1] or "")
                end = str(entry[2] or "")

            if start:
                normalized.append({"note": note, "start": start, "end": end})

        return normalized

    def _set_requested_date_entries(self, entries: Any) -> None:
        self.requested_dates_table.setRowCount(0)
        for entry in self._normalize_saved_requested_date_entries(entries):
            self._append_requested_date_row(
                entry["start"],
                entry["end"],
                entry.get("note", ""),
            )
        self._update_date_section_counts()

    def _get_requested_date_entries(self) -> list[dict[str, str]]:
        entries: list[dict[str, str]] = []
        for row in range(self.requested_dates_table.rowCount()):
            note_item = self.requested_dates_table.item(row, 0)
            start_item = self.requested_dates_table.item(row, 1)
            end_item = self.requested_dates_table.item(row, 2)
            note = note_item.text().strip() if note_item else ""
            start = validate_required_date(start_item.text() if start_item else "", "Requested date")
            end = validate_date_or_blank(end_item.text() if end_item else "", "Requested end date") or ""
            if end and end < start:
                raise ValueError(f"Requested range {start} to {end}: end date is before start date.")
            entries.append({"note": note, "start": start, "end": end})
        return entries

    def _get_requested_dates_for_scoring(self) -> list[Any]:
        requested_dates: list[Any] = []
        for entry in self._get_requested_date_entries():
            start = entry["start"]
            end = entry["end"]
            if end and end != start:
                requested_dates.append((start, end))
            else:
                requested_dates.append(start)
        return requested_dates

    def _selected_requested_date_row(self) -> int | None:
        rows = self.requested_dates_table.selectionModel().selectedRows()
        if not rows:
            return None
        return rows[0].row()

    def _add_requested_date_range(self) -> None:
        dialog = RequestedDateRangeDialog(self, "Add requested day or range")
        if dialog.exec() == QDialog.Accepted and dialog.result:
            self._append_requested_date_row(
                dialog.result["start"],
                dialog.result["end"],
                dialog.result.get("note", ""),
            )
            self._save_requested_dates_to_config()
            self._schedule_preference_refresh(
                "Requested days changed. Sorting columns will refresh automatically."
            )

    def _edit_requested_date_range(self) -> None:
        row = self._selected_requested_date_row()
        if row is None:
            QMessageBox.information(self, "Edit requested day", "Select a requested day or range first.")
            return

        note = self.requested_dates_table.item(row, 0).text() if self.requested_dates_table.item(row, 0) else ""
        start = self.requested_dates_table.item(row, 1).text() if self.requested_dates_table.item(row, 1) else ""
        end = self.requested_dates_table.item(row, 2).text() if self.requested_dates_table.item(row, 2) else ""
        dialog = RequestedDateRangeDialog(
            self,
            "Edit requested day or range",
            start,
            end,
            note,
        )
        if dialog.exec() == QDialog.Accepted and dialog.result:
            note_item = QTableWidgetItem(dialog.result.get("note", ""))
            note_item.setToolTip(dialog.result.get("note", ""))
            self.requested_dates_table.setItem(row, 0, note_item)
            self.requested_dates_table.setItem(row, 1, QTableWidgetItem(dialog.result["start"]))
            self.requested_dates_table.setItem(row, 2, QTableWidgetItem(dialog.result["end"]))
            self._save_requested_dates_to_config()
            self._schedule_preference_refresh(
                "Requested days changed. Sorting columns will refresh automatically."
            )

    def _remove_requested_date_range(self) -> None:
        row = self._selected_requested_date_row()
        if row is not None:
            self.requested_dates_table.removeRow(row)
            self._update_date_section_counts()
            self._save_requested_dates_to_config()
            self._schedule_preference_refresh(
                "Requested days changed. Sorting columns will refresh automatically."
            )

    def _clear_requested_date_ranges(self) -> None:
        self.requested_dates_table.setRowCount(0)
        self._update_date_section_counts()
        self._save_requested_dates_to_config()
        self._schedule_preference_refresh(
            "Requested days changed. Sorting columns will refresh automatically."
        )

    def _save_requested_dates_to_config(self) -> None:
        try:
            self.config_data["requested_dates"] = self._get_requested_date_entries()
            save_config(self.config_data)
        except Exception:
            pass

    # Line-type preference order ----------------------------------------

    def _set_line_type_preference_order(self, order: Any) -> None:
        cleaned: list[str] = []
        if isinstance(order, (list, tuple)):
            for value in order:
                code = str(value).strip().upper()
                if code in LINE_TYPE_CODES and code not in cleaned:
                    cleaned.append(code)
        for code in LINE_TYPE_CODES:
            if code not in cleaned:
                cleaned.append(code)

        self.line_type_preference_list.clear()
        self.line_type_preference_list.addItems(cleaned)
        self.line_type_preference_list.resize_to_all_items()

    def _get_line_type_preference_order(self) -> list[str]:
        order = [
            self.line_type_preference_list.item(row).text().strip().upper()
            for row in range(self.line_type_preference_list.count())
        ]
        if len(order) != len(LINE_TYPE_CODES) or set(order) != set(LINE_TYPE_CODES):
            raise ValueError(
                "Line-type preference order must contain TRIPS, VTO, RA, RB, "
                "SA, SB, SBA, SBG, and VOR exactly once."
            )
        return order

    def _move_line_type_preference_up(self) -> None:
        row = self.line_type_preference_list.currentRow()
        if row <= 0:
            return
        item = self.line_type_preference_list.takeItem(row)
        self.line_type_preference_list.insertItem(row - 1, item)
        self.line_type_preference_list.setCurrentRow(row - 1)
        self._line_type_preference_changed()

    def _move_line_type_preference_down(self) -> None:
        row = self.line_type_preference_list.currentRow()
        if row < 0 or row >= self.line_type_preference_list.count() - 1:
            return
        item = self.line_type_preference_list.takeItem(row)
        self.line_type_preference_list.insertItem(row + 1, item)
        self.line_type_preference_list.setCurrentRow(row + 1)
        self._line_type_preference_changed()

    def _reset_line_type_preference_order(self) -> None:
        self._set_line_type_preference_order(DEFAULT_LINE_TYPE_PREFERENCE_ORDER)
        self.line_type_preference_list.setCurrentRow(0)
        self._line_type_preference_changed()

    def _line_type_preference_changed(self) -> None:
        self.config_data["line_type_preference_order"] = self._get_line_type_preference_order()
        save_config(self.config_data)
        self._schedule_preference_refresh(
            "Line-type preference order changed. Analyzer values will refresh automatically."
        )

    # Preference persistence and automatic column refresh ---------------

    def _on_hourly_rate_changed(self) -> None:
        if self._loading_saved_values:
            return
        hourly_rate = round(float(self.hourly_rate_edit.value()), 2)
        self.config_data["hourly_rate"] = hourly_rate
        save_config(self.config_data)

    def _on_date_preferences_changed(self, status_message: str) -> None:
        if self._loading_saved_values:
            return
        self._schedule_preference_refresh(status_message)

    def _schedule_preference_refresh(self, status_message: str) -> None:
        if self._loading_saved_values:
            return
        self._clear_bid_string(status_message)
        self.preference_refresh_pending = True
        self.preference_refresh_timer.start(500)

    def _refresh_after_preference_change(self) -> None:
        if not self.preference_refresh_pending:
            return

        if self.worker_thread and self.worker_thread.is_alive():
            self.preference_refresh_timer.start(200)
            return

        if self.cached_trips is None or self.cached_lines is None:
            self.preference_refresh_pending = False
            self.pdf_status_label.setText(
                "Preferences saved. Load the PDFs to build the updated sorting columns."
            )
            return

        try:
            inputs = self._collect_inputs()
            self._save_inputs_to_config(inputs)
        except Exception as exc:
            self.preference_refresh_pending = False
            self.pdf_status_label.setText(
                "Preference refresh is waiting for complete, valid inputs."
            )
            self._write_log(f"Automatic preference refresh skipped: {exc}")
            return

        self.preference_refresh_pending = False
        self.pdf_status_label.setText("Preferences changed. Refreshing analyzer columns...")
        self._start_worker(self._load_worker, inputs, False)

    # -------------------------- Sorting criteria rows --------------------------

    @staticmethod
    def _normalize_sort_direction(value: Any) -> str:
        text = str(value or "").strip().lower()
        aliases = {
            "desc": "high_to_low",
            "descending": "high_to_low",
            "high-to-low": "high_to_low",
            "high to low": "high_to_low",
            "high_to_low": "high_to_low",
            "asc": "low_to_high",
            "ascending": "low_to_high",
            "low-to-high": "low_to_high",
            "low to high": "low_to_high",
            "low_to_high": "low_to_high",
        }
        return aliases.get(text, "high_to_low")

    @staticmethod
    def _normalize_sort_mode(value: Any, *, fallback: str = "weighted") -> str:
        text = str(value or "").strip().lower()
        aliases = {
            "strict": "strict",
            "high priority": "strict",
            "high_priority": "strict",
            "weighted": "weighted",
            "weight": "weighted",
            "equal": "equal",
            "equal to previous": "equal",
            "equal_to_previous": "equal",
        }
        normalized = aliases.get(text, fallback)
        if normalized not in {"strict", "weighted", "equal"}:
            return fallback
        return normalized

    def _normalize_saved_sort_order(self, sort_order: Any) -> list[list[str]]:
        normalized: list[list[str]] = []
        if not isinstance(sort_order, (list, tuple)):
            return normalized

        fallback_mode = self._normalize_sort_mode(
            getattr(self, "default_mode", "weighted"),
            fallback="weighted",
        )

        for item in sort_order:
            if not isinstance(item, (list, tuple)) or len(item) < 2:
                continue

            column = str(item[0] or "").strip()
            if not column:
                continue

            direction = self._normalize_sort_direction(item[1])
            mode = (
                self._normalize_sort_mode(item[2], fallback=fallback_mode)
                if len(item) >= 3
                else fallback_mode
            )

            # "Equal to Previous" cannot be the first actual criterion.
            if not normalized and mode == "equal":
                mode = "weighted"

            normalized.append([column, direction, mode])

        return normalized

    def _add_sort_criteria_row(
        self,
        criterion: list[str] | tuple[str, ...] | None = None,
        *,
        notify: bool = True,
    ) -> None:
        list_item = QListWidgetItem()
        list_item.setFlags(
            list_item.flags()
            | Qt.ItemIsEnabled
            | Qt.ItemIsSelectable
            | Qt.ItemIsDragEnabled
            | Qt.ItemIsDropEnabled
        )

        row_id = self._next_sort_row_id
        self._next_sort_row_id += 1
        list_item.setData(Qt.UserRole, row_id)

        row_widget = QWidget()
        row_widget.setObjectName("SortCriterionCard")
        row_widget.setMinimumHeight(46)
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(6, 3, 6, 3)
        row_layout.setSpacing(8)

        drag_handle = SortCriteriaDragHandle(
            self.sort_rows_list,
            list_item,
            row_widget,
        )

        number_label = QLabel()
        number_label.setFixedWidth(28)
        number_label.setAlignment(Qt.AlignCenter)

        column_combo = WheelSafeComboBox()
        column_combo.setFixedWidth(245)
        column_combo.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        column_combo.addItem("Select column...", "")
        for column in self.sortable_columns:
            column_combo.addItem(column, column)

        direction_combo = WheelSafeComboBox()
        direction_combo.setFixedWidth(130)
        direction_combo.addItems(SORT_DIRECTION_LABEL_TO_VALUE.keys())

        mode_combo = WheelSafeComboBox()
        mode_combo.setFixedWidth(155)

        contribution_label = QLabel("—")
        contribution_label.setFixedWidth(100)
        contribution_label.setAlignment(Qt.AlignCenter)
        contribution_label.setToolTip(
            "Calculated by get_sort_percent_contributions() using the current "
            "weighting style and priority-emphasis setting."
        )

        remove_button = QPushButton("Remove")
        remove_button.setFixedWidth(78)

        row_layout.addWidget(drag_handle)
        row_layout.addWidget(number_label)
        row_layout.addWidget(column_combo)
        row_layout.addWidget(direction_combo)
        row_layout.addWidget(mode_combo)
        row_layout.addWidget(contribution_label)
        row_layout.addSpacing(4)
        row_layout.addWidget(remove_button)
        row_layout.addStretch(1)

        row_data: dict[str, Any] = {
            "row_id": row_id,
            "item": list_item,
            "widget": row_widget,
            "drag_handle": drag_handle,
            "number_label": number_label,
            "column_combo": column_combo,
            "direction_combo": direction_combo,
            "mode_combo": mode_combo,
            "contribution_label": contribution_label,
            "remove_button": remove_button,
        }

        self._sort_row_by_id[row_id] = row_data
        self.sort_criteria_rows.append(row_data)
        self.sort_rows_list.addItem(list_item)
        list_item.setSizeHint(QSize(0, 52))
        self.sort_rows_list.setItemWidget(list_item, row_widget)

        column_combo.currentIndexChanged.connect(self._on_sort_criteria_changed)
        direction_combo.currentIndexChanged.connect(self._on_sort_criteria_changed)
        mode_combo.currentIndexChanged.connect(self._on_sort_criteria_changed)
        remove_button.clicked.connect(
            lambda _checked=False, row=row_data: self._remove_sort_criteria_row(row)
        )

        self._update_sort_row_positions()

        if criterion:
            self._set_sort_row_values(row_data, criterion)
        else:
            self._set_sort_row_values(row_data, None)

        self._update_sort_percentages()

        if notify:
            self._on_sort_criteria_changed()

    def _sync_sort_criteria_rows_from_list(self) -> None:
        ordered_rows: list[dict[str, Any]] = []
        for index in range(self.sort_rows_list.count()):
            item = self.sort_rows_list.item(index)
            try:
                row_id = int(item.data(Qt.UserRole))
            except (TypeError, ValueError):
                continue
            row_data = self._sort_row_by_id.get(row_id)
            if row_data is not None:
                ordered_rows.append(row_data)
        self.sort_criteria_rows = ordered_rows

    def _remove_sort_criteria_row(self, row_data: dict[str, Any]) -> None:
        self._sync_sort_criteria_rows_from_list()

        if len(self.sort_criteria_rows) <= MIN_SORT_CRITERIA_ROWS:
            self._set_sort_row_values(row_data, None)
            self._on_sort_criteria_changed()
            return

        if row_data not in self.sort_criteria_rows:
            return

        list_item: QListWidgetItem = row_data["item"]
        row_index = self.sort_rows_list.row(list_item)
        if row_index >= 0:
            self.sort_rows_list.takeItem(row_index)

        self._sort_row_by_id.pop(row_data["row_id"], None)
        row_data["widget"].deleteLater()
        del list_item

        self._sync_sort_criteria_rows_from_list()
        self._update_sort_row_positions()
        self._on_sort_criteria_changed()

    def _on_sort_rows_previewed(self, *_args: Any) -> None:
        """Refresh numbering and percentages while rows visibly shift."""
        if self._rebuilding_sort_rows:
            return

        self._sync_sort_criteria_rows_from_list()
        self._update_sort_row_positions()
        self._update_sort_percentages()

    def _on_sort_rows_reordered(self, *_args: Any) -> None:
        """Commit and save the order after the criterion is dropped."""
        if self._rebuilding_sort_rows:
            return

        self._sync_sort_criteria_rows_from_list()
        self._update_sort_row_positions()
        self._on_sort_criteria_changed()

    def _update_sort_row_positions(self) -> None:
        self._sync_sort_criteria_rows_from_list()

        displayed_number = 0
        for index, row in enumerate(self.sort_criteria_rows):
            row["remove_button"].setEnabled(
                len(self.sort_criteria_rows) > MIN_SORT_CRITERIA_ROWS
                or bool(row["column_combo"].currentData())
            )

            mode_combo: QComboBox = row["mode_combo"]
            current_mode = self._normalize_sort_mode(
                mode_combo.currentData() or mode_combo.currentText(),
                fallback="weighted",
            )

            mode_combo.blockSignals(True)
            mode_combo.clear()
            mode_combo.addItem("High Priority", "strict")
            mode_combo.addItem("Normal Priority", "weighted")
            if index > 0:
                mode_combo.addItem("Equal to Previous Priority", "equal")

            if index == 0 and current_mode == "equal":
                current_mode = "weighted"

            mode_index = mode_combo.findData(current_mode)
            mode_combo.setCurrentIndex(mode_index if mode_index >= 0 else 1)
            mode_combo.blockSignals(False)

            # Equal criteria visibly belong to the same numbered priority level.
            if index == 0 or current_mode != "equal":
                displayed_number += 1
            row["number_label"].setText(f"{displayed_number}.")

    def _set_sort_row_values(
        self,
        row_data: dict[str, Any],
        criterion: list[str] | tuple[str, ...] | None,
    ) -> None:
        column_combo: QComboBox = row_data["column_combo"]
        direction_combo: QComboBox = row_data["direction_combo"]
        mode_combo: QComboBox = row_data["mode_combo"]

        for combo in (column_combo, direction_combo, mode_combo):
            combo.blockSignals(True)

        try:
            if not criterion:
                column_combo.setCurrentIndex(0)
                direction_combo.setCurrentText("High to Low")
                default_mode = "strict" if self.sort_criteria_rows.index(row_data) == 0 else "weighted"
                mode_index = mode_combo.findData(default_mode)
                mode_combo.setCurrentIndex(mode_index if mode_index >= 0 else 0)
                return

            if len(criterion) < 2:
                return

            column = str(criterion[0] or "").strip()
            if not column:
                return

            direction = self._normalize_sort_direction(criterion[1])
            mode = (
                self._normalize_sort_mode(
                    criterion[2],
                    fallback=self._normalize_sort_mode(
                        getattr(self, "default_mode", "weighted"),
                        fallback="weighted",
                    ),
                )
                if len(criterion) >= 3
                else self._normalize_sort_mode(
                    getattr(self, "default_mode", "weighted"),
                    fallback="weighted",
                )
            )

            column_index = column_combo.findData(column)
            if column_index < 0:
                # Keep a saved selection visible until the DataFrame is loaded.
                column_combo.addItem(column, column)
                column_index = column_combo.findData(column)
            column_combo.setCurrentIndex(column_index)

            direction_label = SORT_DIRECTION_VALUE_TO_LABEL.get(direction, "High to Low")
            direction_combo.setCurrentText(direction_label)

            row_index = self.sort_criteria_rows.index(row_data)
            if row_index == 0 and mode == "equal":
                mode = "weighted"
            mode_index = mode_combo.findData(mode)
            mode_combo.setCurrentIndex(mode_index if mode_index >= 0 else 1)
        finally:
            for combo in (column_combo, direction_combo, mode_combo):
                combo.blockSignals(False)

    def _set_sort_order(self, sort_order: Any) -> None:
        normalized = self._normalize_saved_sort_order(sort_order)

        self._rebuilding_sort_rows = True
        try:
            self.sort_rows_list.clear()
            self.sort_criteria_rows = []
            self._sort_row_by_id.clear()

            total_rows = max(MIN_SORT_CRITERIA_ROWS, len(normalized))
            for index in range(total_rows):
                criterion = normalized[index] if index < len(normalized) else None
                self._add_sort_criteria_row(criterion, notify=False)
        finally:
            self._rebuilding_sort_rows = False

        self.sort_order = normalized
        self._update_sort_row_positions()
        self._update_sort_percentages()

    def _clear_sort_order(self) -> None:
        self._set_sort_order([])
        self._on_sort_criteria_changed()

    def _get_sort_order_from_rows(self, *, validate: bool) -> list[list[str]]:
        self._sync_sort_criteria_rows_from_list()
        sort_order: list[list[str]] = []
        seen_columns: set[str] = set()
        found_blank = False

        for row_number, row in enumerate(self.sort_criteria_rows, start=1):
            column = str(row["column_combo"].currentData() or "").strip()

            if not column:
                found_blank = True
                continue

            if validate and found_blank:
                raise ValueError(
                    "Sorting criteria must be filled in order. "
                    f"Choose a column in row {row_number - 1} before using row {row_number}."
                )

            if validate and column in seen_columns:
                raise ValueError(
                    f'The sorting column "{column}" is selected more than once.'
                )

            direction = SORT_DIRECTION_LABEL_TO_VALUE.get(
                row["direction_combo"].currentText(),
                "high_to_low",
            )
            mode = self._normalize_sort_mode(
                row["mode_combo"].currentData() or row["mode_combo"].currentText(),
                fallback="weighted",
            )

            if not sort_order and mode == "equal":
                if validate:
                    raise ValueError(
                        '"Equal to Previous" cannot be used for the first sorting criterion.'
                    )
                mode = "weighted"

            sort_order.append([column, direction, mode])
            seen_columns.add(column)

        return sort_order

    def _on_sort_criteria_changed(self, *_args: Any) -> None:
        self.sort_order = self._get_sort_order_from_rows(validate=False)
        self._update_sort_row_positions()
        self._update_sort_percentages()

        if self._loading_saved_values:
            return

        self.config_data["sort_order"] = self.sort_order
        save_config(self.config_data)
        self._clear_bid_string(
            "Sorting criteria changed. Generate the bid string again."
        )

    def _update_sort_percentages(self) -> None:
        active_rows: list[dict[str, Any]] = []
        sort_order: list[list[str]] = []

        contribution_tooltip = (
            "Calculated by get_sort_percent_contributions() using the current "
            "weighting style and priority-emphasis setting."
        )

        for row in self.sort_criteria_rows:
            row["contribution_label"].setText("—")
            row["contribution_label"].setToolTip(contribution_tooltip)
            column = str(row["column_combo"].currentData() or "").strip()
            if not column:
                continue

            direction = SORT_DIRECTION_LABEL_TO_VALUE.get(
                row["direction_combo"].currentText(),
                "high_to_low",
            )
            mode = self._normalize_sort_mode(
                row["mode_combo"].currentData() or row["mode_combo"].currentText(),
                fallback="weighted",
            )
            if not sort_order and mode == "equal":
                mode = "weighted"

            active_rows.append(row)
            sort_order.append([column, direction, mode])

        if not sort_order:
            return

        try:
            contributions = get_sort_percent_contributions(
                sort_order,
                weighting_style=self.weighting_style,
                soft_max_weight=self.soft_max_weight,
                soft_min_weight=self.soft_min_weight,
                round_digits=2,
            )
        except Exception as exc:
            for row in active_rows:
                row["contribution_label"].setText("Error")
                row["contribution_label"].setToolTip(
                    f"Could not calculate sorting contribution: {exc}"
                )
            return

        for row, contribution in zip(active_rows, contributions):
            if contribution is None:
                row["contribution_label"].setText("—")
                continue

            try:
                numeric = float(contribution)
            except (TypeError, ValueError):
                row["contribution_label"].setText("—")
                continue

            if math.isnan(numeric):
                row["contribution_label"].setText("—")
            else:
                row["contribution_label"].setText(f"{numeric:.2f}%")

    def _refresh_available_columns_list(self, columns: list[str]) -> None:
        self.sortable_columns = [str(column) for column in columns]

        for row in self.sort_criteria_rows:
            combo: QComboBox = row["column_combo"]
            selected = str(combo.currentData() or "").strip()

            combo.blockSignals(True)
            combo.clear()
            combo.addItem("Select column...", "")
            for column in self.sortable_columns:
                combo.addItem(column, column)

            if selected:
                index = combo.findData(selected)
                if index < 0:
                    combo.addItem(selected, selected)
                    index = combo.findData(selected)
                combo.setCurrentIndex(index)
            combo.blockSignals(False)

        self._update_sort_percentages()

    def _clean_sort_order_for_columns(self, columns: list[str]) -> None:
        valid_columns = {str(column) for column in columns}
        current_order = self._get_sort_order_from_rows(validate=False)

        valid_order = [rule for rule in current_order if rule[0] in valid_columns]
        skipped = [rule[0] for rule in current_order if rule[0] not in valid_columns]

        if skipped:
            self._write_log(
                "Skipped saved sorting columns not found in this DataFrame: "
                + ", ".join(skipped)
            )

        self._set_sort_order(valid_order)
        self.sort_order = valid_order

    # -------------------------- Sorting settings --------------------------

    def _on_priority_emphasis_changed(self, value: float) -> None:
        self.soft_max_weight = max(float(value), float(self.soft_min_weight))
        self._update_sort_percentages()

        if self._loading_saved_values:
            return

        settings = self._get_sorting_settings()
        self.config_data["sorting_settings"] = settings
        save_config(self.config_data)
        self._clear_bid_string(
            "Priority emphasis changed. Generate the bid string again."
        )

    def _validate_sorting_settings_values(
        self,
        default_mode: str,
        weighting_style: str,
        soft_max_weight_text: str,
        soft_min_weight_text: str,
        keep_score_columns: bool,
    ) -> dict[str, Any]:
        default_mode = default_mode.strip().lower() or DEFAULT_SORTING_SETTINGS["default_mode"]
        weighting_style = weighting_style.strip().lower() or DEFAULT_SORTING_SETTINGS["weighting_style"]

        if default_mode not in DEFAULT_MODE_DESCRIPTIONS:
            raise ValueError("Default mode must be either strict or weighted.")
        if weighting_style not in WEIGHTING_STYLE_DESCRIPTIONS:
            raise ValueError("Weighting style must be equal, hard, or soft.")

        try:
            soft_max_weight = float(str(soft_max_weight_text).strip())
        except ValueError as exc:
            raise ValueError("Soft max weight must be a number, such as 3.0.") from exc

        try:
            soft_min_weight = float(str(soft_min_weight_text).strip())
        except ValueError as exc:
            raise ValueError("Soft min weight must be a number, such as 1.0.") from exc

        if soft_max_weight <= 0 or soft_min_weight <= 0:
            raise ValueError("Soft max weight and soft min weight must both be greater than zero.")
        if soft_max_weight < soft_min_weight:
            raise ValueError("Soft max weight should be greater than or equal to soft min weight.")

        return {
            "default_mode": default_mode,
            "weighting_style": weighting_style,
            "soft_max_weight": soft_max_weight,
            "soft_min_weight": soft_min_weight,
            "keep_score_columns": bool(keep_score_columns),
        }

    def _get_sorting_settings(self) -> dict[str, Any]:
        return self._validate_sorting_settings_values(
            self.default_mode,
            self.weighting_style,
            str(self.soft_max_weight),
            str(self.soft_min_weight),
            self.keep_score_columns,
        )

    def _apply_sorting_settings_to_ui(self, sorting_settings: dict[str, Any]) -> None:
        self.default_mode = str(sorting_settings["default_mode"])
        self.weighting_style = str(sorting_settings["weighting_style"])
        self.soft_max_weight = float(sorting_settings["soft_max_weight"])
        self.soft_min_weight = float(sorting_settings["soft_min_weight"])
        self.keep_score_columns = bool(sorting_settings["keep_score_columns"])

        if hasattr(self, "priority_emphasis_spin"):
            self.priority_emphasis_spin.blockSignals(True)
            self.priority_emphasis_spin.setMinimum(max(0.01, self.soft_min_weight))
            self.priority_emphasis_spin.setValue(
                max(self.soft_max_weight, self.soft_min_weight)
            )
            self.priority_emphasis_spin.blockSignals(False)

    def _save_sorting_settings(self, sorting_settings: dict[str, Any], *, show_message: bool = True) -> None:
        self._apply_sorting_settings_to_ui(sorting_settings)
        self.config_data["sorting_settings"] = sorting_settings
        save_config(self.config_data)
        self._write_log(
            "Saved advanced sorting settings: "
            f"{sorting_settings['default_mode']}, "
            f"{sorting_settings['weighting_style']}, "
            f"soft weights {sorting_settings['soft_min_weight']}–{sorting_settings['soft_max_weight']}."
        )
        self._update_sort_percentages()
        self._clear_bid_string("Advanced sorting settings changed. Generate the bid string again.")
        if show_message:
            QMessageBox.information(self, "Saved", "Advanced sorting settings saved.")

    def _open_sorting_settings_dialog(self) -> None:
        current = self._get_sorting_settings()

        dialog = QDialog(self)
        dialog.setWindowTitle("Advanced Sorting Settings")
        dialog.setModal(True)
        dialog.setWindowFlags(dialog.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QGridLayout(dialog)
        layout.setColumnStretch(1, 1)

        intro = QLabel(
            "These settings affect how selected sorting columns are combined. "
            "Most users should leave them at the saved defaults."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro, 0, 0, 1, 3)

        default_mode_combo = QComboBox()
        default_mode_combo.addItems(list(DEFAULT_MODE_DESCRIPTIONS))
        default_mode_combo.setCurrentText(str(current["default_mode"]))
        default_tip = "strict: " + DEFAULT_MODE_DESCRIPTIONS["strict"] + "\n\nweighted: " + DEFAULT_MODE_DESCRIPTIONS["weighted"]
        default_mode_combo.setToolTip(default_tip)

        weighting_style_combo = QComboBox()
        weighting_style_combo.addItems(list(WEIGHTING_STYLE_DESCRIPTIONS))
        weighting_style_combo.setCurrentText(str(current["weighting_style"]))
        weighting_tip = (
            "equal: " + WEIGHTING_STYLE_DESCRIPTIONS["equal"] + "\n\n"
            "hard: " + WEIGHTING_STYLE_DESCRIPTIONS["hard"] + "\n\n"
            "soft: " + WEIGHTING_STYLE_DESCRIPTIONS["soft"]
        )
        weighting_style_combo.setToolTip(weighting_tip)

        soft_max_entry = QLineEdit(str(current["soft_max_weight"]))
        soft_max_entry.setToolTip("Priority emphasis for soft weighting. Bigger values favor earlier criteria more strongly. Default: 3.0.")
        soft_min_entry = QLineEdit(str(current["soft_min_weight"]))
        soft_min_entry.setToolTip("Baseline contribution weight for the lowest-priority criterion. Default: 1.0.")
        keep_score_columns_check = QCheckBox("Keep score columns in Excel")
        keep_score_columns_check.setChecked(bool(current["keep_score_columns"]))
        keep_score_columns_check.setToolTip("When enabled, any extra score/helper columns created by weighted sorting remain in the exported Excel file.")

        def add_row(row: int, label_text: str, widget: QWidget, tip: str | None = None) -> None:
            label = QLabel(label_text)
            if tip:
                label.setToolTip(tip)
            layout.addWidget(label, row, 0)
            layout.addWidget(widget, row, 1)
            if tip:
                help_label = QLabel("?")
                help_label.setToolTip(tip)
                help_label.setStyleSheet(f"color: {UPS_GOLD}; font-weight: bold;")
                layout.addWidget(help_label, row, 2)

        add_row(1, "Default mode:", default_mode_combo, default_tip)
        add_row(2, "Weighting style:", weighting_style_combo, weighting_tip)
        add_row(3, "Priority emphasis:", soft_max_entry, soft_max_entry.toolTip())
        add_row(4, "Lower-priority baseline:", soft_min_entry, soft_min_entry.toolTip())
        layout.addWidget(keep_score_columns_check, 5, 0, 1, 2)

        restore_defaults_button = QPushButton("Restore defaults")
        cancel_button = QPushButton("Cancel")
        save_button = QPushButton("Save")

        buttons = QHBoxLayout()
        buttons.addWidget(restore_defaults_button)
        buttons.addStretch(1)
        buttons.addWidget(cancel_button)
        buttons.addWidget(save_button)
        layout.addLayout(buttons, 6, 0, 1, 3)

        def restore_defaults() -> None:
            default_mode_combo.setCurrentText(str(DEFAULT_SORTING_SETTINGS["default_mode"]))
            weighting_style_combo.setCurrentText(str(DEFAULT_SORTING_SETTINGS["weighting_style"]))
            soft_max_entry.setText(str(DEFAULT_SORTING_SETTINGS["soft_max_weight"]))
            soft_min_entry.setText(str(DEFAULT_SORTING_SETTINGS["soft_min_weight"]))
            keep_score_columns_check.setChecked(bool(DEFAULT_SORTING_SETTINGS["keep_score_columns"]))

        def save_and_close() -> None:
            try:
                settings = self._validate_sorting_settings_values(
                    default_mode_combo.currentText(),
                    weighting_style_combo.currentText(),
                    soft_max_entry.text(),
                    soft_min_entry.text(),
                    keep_score_columns_check.isChecked(),
                )
            except Exception as exc:
                QMessageBox.critical(dialog, "Sorting settings error", str(exc))
                return

            self._save_sorting_settings(settings, show_message=False)
            dialog.accept()

        restore_defaults_button.clicked.connect(restore_defaults)
        cancel_button.clicked.connect(dialog.reject)
        save_button.clicked.connect(save_and_close)

        dialog.resize(620, 260)
        dialog.exec()

    # -------------------------- Validation / config --------------------------

    def _collect_inputs(self) -> dict[str, Any]:
        trips_pdf_path = self.trips_path_edit.text().strip().strip('"').strip("'")
        lines_pdf_path = self.lines_path_edit.text().strip().strip('"').strip("'")

        if not trips_pdf_path:
            raise ValueError("Please choose the TRIPS PDF.")
        if not lines_pdf_path:
            raise ValueError("Please choose the LINES PDF.")
        if not Path(trips_pdf_path).exists():
            raise ValueError("The TRIPS PDF path does not exist.")
        if not Path(lines_pdf_path).exists():
            raise ValueError("The LINES PDF path does not exist.")

        vacation_ranges = self._get_vacation_ranges()
        if not vacation_ranges:
            vacation_ranges = None

        requested_date_entries = self._get_requested_date_entries()
        requested_dates = self._get_requested_dates_for_scoring()
        if not requested_dates:
            requested_dates = None

        training_start = validate_date_or_blank(self.training_start_entry.text(), "Training start")
        training_end = validate_date_or_blank(self.training_end_entry.text(), "Training end")

        if bool(training_start) != bool(training_end):
            raise ValueError("Enter both training start and training end, or leave both blank.")
        if training_start and training_end and training_end < training_start:
            raise ValueError("Training end date is before training start date.")

        hourly_rate = round(float(self.hourly_rate_edit.value()), 2)
        line_type_preference_order = self._get_line_type_preference_order()

        bid_edge = self.bid_edge_combo.currentText().strip().lower() or "none"
        if bid_edge not in {"none", "start", "end", "both"}:
            raise ValueError("Bid edge preference must be none, start, end, or both.")

        output_folder = Path(self.output_folder_edit.text().strip().strip('"').strip("'")).expanduser()
        if not output_folder.exists():
            output_folder.mkdir(parents=True, exist_ok=True)
        if not output_folder.is_dir():
            raise ValueError("Output folder is not a folder.")

        output_filename = clean_filename(self.output_filename_edit.text())
        output_path = output_folder / f"{output_filename}.xlsx"

        number_of_lines_to_bid = validate_positive_int(
            self.number_of_lines_edit.text(),
            "Number of lines to bid",
        )

        return {
            "trips_pdf_path": trips_pdf_path,
            "lines_pdf_path": lines_pdf_path,
            "vacation_ranges": vacation_ranges,
            "requested_date_entries": requested_date_entries,
            "requested_dates": requested_dates,
            "training_start": training_start,
            "training_end": training_end,
            "hourly_rate": hourly_rate,
            "line_type_preference_order": line_type_preference_order,
            "bid_edge": bid_edge,
            "output_folder": output_folder,
            "output_path": output_path,
            "number_of_lines_to_bid": number_of_lines_to_bid,
            "sort_order": self._get_sort_order_from_rows(validate=True),
            "sorting_settings": self._get_sorting_settings(),
        }


    def _check_matching_bid_period_or_warn(self, inputs: dict[str, Any]) -> bool:
        """
        Verifies that the TRIPS and LINES PDFs are from the same bid period.

        If they match:
            - stores the bid period string
            - sets the Excel filename to that bid period
            - updates inputs["output_path"] so export uses the new filename

        If they do not match:
            - shows a popup
            - returns False
            - prevents PDF extraction from starting
        """
        pdf_key = (inputs["trips_pdf_path"], inputs["lines_pdf_path"])

        if self.cached_bid_period_key == pdf_key and self.cached_bid_period:
            bid_period = self.cached_bid_period
        else:
            self.pdf_status_label.setText("Checking bid period match...")
            QApplication.processEvents()

            try:
                bid_period = matching_bid_period(
                    inputs["trips_pdf_path"],
                    inputs["lines_pdf_path"],
                )
            except Exception as exc:
                self.pdf_status_label.setText("Bid period check failed.")
                QMessageBox.critical(
                    self,
                    "Bid period check error",
                    f"Could not verify that the LINES and TRIPS PDFs match.\n\n{exc}",
                )
                return False

            if bid_period is None:
                self.pdf_status_label.setText("Lines and trips packages do not match.")
                self._write_log("PDF load stopped: Lines and trips package provided do not match.")

                QMessageBox.critical(
                    self,
                    "PDF package mismatch",
                    "Lines and trips package provided do not match.\n\n"
                    "Please choose LINES and TRIPS PDFs from the same bid period.",
                )
                return False

            bid_period = str(bid_period).strip()
            self.cached_bid_period_key = pdf_key
            self.cached_bid_period = bid_period

        output_filename = clean_filename(bid_period)
        self.output_filename_edit.setText(output_filename)

        inputs["bid_period"] = bid_period
        inputs["output_path"] = Path(inputs["output_folder"]) / f"{output_filename}.xlsx"

        self.pdf_status_label.setText(f"Bid period verified: {bid_period}")
        self._write_log(f"Bid period verified: {bid_period}. Excel filename set to {output_filename}.xlsx")

        return True


    def _save_inputs_to_config(self, inputs: dict[str, Any]) -> None:
        self.config_data["vacation_ranges"] = inputs["vacation_ranges"]
        self.config_data["requested_dates"] = inputs["requested_date_entries"]
        self.config_data["training_start"] = inputs["training_start"]
        self.config_data["training_end"] = inputs["training_end"]
        self.config_data["hourly_rate"] = inputs["hourly_rate"]
        self.config_data["line_type_preference_order"] = inputs["line_type_preference_order"]
        self.config_data["bid_edge"] = inputs["bid_edge"]
        self.config_data["sort_order"] = inputs["sort_order"]
        self.config_data["sorting_settings"] = inputs["sorting_settings"]
        self.config_data["number_of_lines_to_bid"] = inputs["number_of_lines_to_bid"]
        output_paths = self.config_data.setdefault("output_paths", {})
        output_paths[get_os_name()] = str(inputs["output_folder"])
        save_config(self.config_data)

    # -------------------------- Worker control --------------------------

    def _reset_trip_progress(self, message: str = "Trips extraction progress: starting...") -> None:
        self.trip_progress.setRange(0, 100)
        self.trip_progress.setValue(0)
        self.trip_progress_text_label.setText(message)

    def _queue_trip_progress(self, progress_data: dict[str, Any]) -> None:
        self.message_queue.put(("trip_progress", progress_data))

    def _handle_trip_progress(self, progress_data: dict[str, Any]) -> None:
        try:
            current = int(progress_data.get("current") or 0)
            total = int(progress_data.get("total") or 0)
        except (TypeError, ValueError):
            current = 0
            total = 0

        status = str(progress_data.get("status") or "running").lower()
        message = str(progress_data.get("message") or "").strip()
        total_trips = progress_data.get("total_trips")

        percent = max(0.0, min(100.0, (current / total) * 100.0)) if total > 0 else 0.0
        if status in {"done", "cached"}:
            percent = 100.0

        self.trip_progress.setRange(0, 100)
        self.trip_progress.setValue(int(percent))

        label = f"Trips extraction: {percent:.0f}%"
        if message:
            label += f" — {message}"
        if total_trips is not None:
            label += f" | Trips found: {total_trips}"
        self.trip_progress_text_label.setText(label)

    def _set_busy(self, busy: bool) -> None:
        for button in (
            self.load_button,
            self.export_button,
            self.visualizer_button,
            self.generate_bid_string_button,
            self.copy_bid_string_button,
        ):
            button.setEnabled(not busy)

        if busy:
            self.progress.setRange(0, 0)
        else:
            self.progress.setRange(0, 100)
            self.progress.setValue(0)
            if self.preference_refresh_pending:
                QTimer.singleShot(150, self._refresh_after_preference_change)

    def _log(self, message: str) -> None:
        self.message_queue.put(("log", message))

    def _write_log(self, message: str) -> None:
        self.log_text.append(str(message))

    def _poll_queue(self) -> None:
        try:
            while True:
                event, payload = self.message_queue.get_nowait()

                if event == "log":
                    self._write_log(str(payload))
                elif event == "trip_progress":
                    if isinstance(payload, dict):
                        self._handle_trip_progress(payload)
                elif event == "error":
                    self._set_busy(False)
                    self.pdf_status_label.setText("Error. See status box below.")
                    self._write_log(f"ERROR: {payload}")
                    QMessageBox.critical(self, "Error", str(payload))
                elif event == "loaded":
                    self._handle_loaded_payload(payload)
                elif event == "bid_string_ready":
                    self._handle_bid_string_ready_payload(payload)
                elif event == "visualizer_ready":
                    self._set_busy(False)
                    if isinstance(payload, dict):
                        self.preview_df = payload.get("df")
                    try:
                        self._open_visualizer_from_payload(payload)
                    except Exception as exc:
                        self.pdf_status_label.setText("Visualizer error. See status box below.")
                        self._write_log(f"ERROR opening visualizer: {exc}")
                        QMessageBox.critical(self, "Visualizer error", str(exc))
                elif event == "exported":
                    self._handle_exported_payload(payload)
        except queue.Empty:
            pass

    def _handle_loaded_payload(self, payload: Any) -> None:
        self._set_busy(False)
        if isinstance(payload, dict):
            df = payload["df"]
            show_ready_message = payload.get("show_ready_message", True)
            status_text = payload.get("status_text") or "PDFs loaded. Sorting columns are ready."
            log_text = payload.get("log_text") or "Loaded PDFs"
        else:
            df = payload
            show_ready_message = True
            status_text = "PDFs loaded. Sorting columns are ready."
            log_text = "Loaded PDFs"

        self.preview_df = df
        columns = get_sortable_columns_from_df(df)
        self._refresh_available_columns_list(columns)
        self._clean_sort_order_for_columns(columns)
        self.pdf_status_label.setText(status_text)
        self._write_log(f"{log_text} and prepared {len(columns)} sortable columns.")

        if show_ready_message:
            QMessageBox.information(self, "Ready", "PDFs are loaded. Choose your sorting priority, then export.")

    def _handle_bid_string_ready_payload(self, payload: Any) -> None:
        self._set_busy(False)
        if isinstance(payload, dict):
            bid_string = str(payload.get("bid_string") or "")
            number_of_lines = payload.get("number_of_lines")
            copy_after = bool(payload.get("copy_after"))
        else:
            bid_string = str(payload)
            number_of_lines = None
            copy_after = False

        self._set_bid_string(
            bid_string,
            f"Bid string ready for {number_of_lines} lines." if number_of_lines else "Bid string ready.",
        )
        self.pdf_status_label.setText("Bid string ready.")
        self._write_log("Generated bid string: " + bid_string)

        if copy_after:
            self._copy_bid_string_to_clipboard()

    def _handle_exported_payload(self, payload: Any) -> None:
        self._set_busy(False)
        if isinstance(payload, dict):
            output_path = payload.get("output_path")
            bid_string = str(payload.get("bid_string") or "")
            number_of_lines = payload.get("number_of_lines")
        else:
            output_path = payload
            bid_string = ""
            number_of_lines = None

        if bid_string:
            self._set_bid_string(
                bid_string,
                f"Bid string ready for {number_of_lines} lines." if number_of_lines else "Bid string ready.",
            )

        self.pdf_status_label.setText("Export complete.")
        self._write_log(f"Finished export: {output_path}")
        if bid_string:
            self._write_log("Generated bid string: " + bid_string)
        self._show_export_complete_dialog(output_path)

    def _start_worker(self, target: Callable[..., None], *args: Any) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            QMessageBox.warning(self, "Busy", "A job is already running.")
            return

        self._set_busy(True)
        self.worker_thread = threading.Thread(target=target, args=args, daemon=True)
        self.worker_thread.start()

    def _open_file_with_default_app(self, file_path: str | Path) -> None:
        path = Path(file_path)
        if not path.exists():
            QMessageBox.critical(self, "Open file", f"File not found:\n{path}")
            return

        try:
            system_name = platform.system()
            if system_name == "Windows":
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif system_name == "Darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            QMessageBox.critical(self, "Open file", f"Could not open file:\n{path}\n\n{exc}")

    def _open_containing_folder(self, file_path: str | Path) -> None:
        path = Path(file_path)
        folder = path.parent
        if not folder.exists():
            QMessageBox.critical(self, "Open folder", f"Folder not found:\n{folder}")
            return

        try:
            system_name = platform.system()
            if system_name == "Windows":
                if path.exists():
                    subprocess.Popen(["explorer", f"/select,{path}"])
                else:
                    os.startfile(str(folder))  # type: ignore[attr-defined]
            elif system_name == "Darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as exc:
            QMessageBox.critical(self, "Open folder", f"Could not open folder:\n{folder}\n\n{exc}")

    def _show_export_complete_dialog(self, output_path: str | Path | None) -> None:
        if output_path is None:
            QMessageBox.information(self, "Finished", "Excel file created.")
            return

        path = Path(output_path)
        dialog = ExportCompleteDialog(
            self,
            path,
            self._open_file_with_default_app,
            self._open_containing_folder,
        )
        dialog.exec()

    # -------------------------- Bid string actions --------------------------

    def _mark_bid_string_stale(self, status_message: str | None = None) -> None:
        self._clear_bid_string(status_message or "Inputs changed. Generate the bid string again after sorting.")

    def _display_bid_string(self, bid_string: str) -> None:
        self.bid_string_text.setPlainText(bid_string or "")

    def _get_displayed_bid_string(self) -> str:
        return self.bid_string_text.toPlainText().strip()

    def _clear_bid_string(self, status_message: str | None = None) -> None:
        self.latest_bid_string = ""
        self._display_bid_string("")
        if status_message:
            self.bid_string_status_label.setText(status_message)

    def _set_bid_string(self, bid_string: str, status_message: str = "Bid string ready.") -> None:
        self.latest_bid_string = bid_string
        self._display_bid_string(bid_string)
        self.bid_string_status_label.setText(status_message)

    def _copy_bid_string_to_clipboard(self) -> None:
        bid_string = self.latest_bid_string.strip() or self._get_displayed_bid_string()
        if not bid_string:
            QMessageBox.information(self, "Bid string", "Generate the bid string first.")
            return

        QApplication.clipboard().setText(bid_string)
        self.bid_string_status_label.setText("Bid string copied to clipboard.")

    def generate_bid_string(self, *, copy_after: bool = False) -> None:
        try:
            inputs = self._collect_inputs()

            if not self._check_matching_bid_period_or_warn(inputs):
                return

            self._save_inputs_to_config(inputs)

        except Exception as exc:
            QMessageBox.critical(self, "Input error", str(exc))
            return

        current_key = (inputs["trips_pdf_path"], inputs["lines_pdf_path"])
        if self.cached_pdf_key != current_key:
            self.pdf_status_label.setText("PDFs not loaded for these paths. Loading first, then generating bid string...")
            self._reset_trip_progress("Trips extraction progress: waiting to start...")

        self.bid_string_status_label.setText("Generating bid string...")
        self._start_worker(self._bid_string_worker, inputs, copy_after)

    def copy_bid_string(self) -> None:
        if self.latest_bid_string.strip() or self._get_displayed_bid_string():
            self._copy_bid_string_to_clipboard()
            return

        self.generate_bid_string(copy_after=True)

    # -------------------------- Processing logic --------------------------

    def _on_bid_edge_changed(self) -> None:
        if self._loading_saved_values:
            return
        self.config_data["bid_edge"] = self.bid_edge_combo.currentText().strip().lower() or "none"
        save_config(self.config_data)
        self._schedule_preference_refresh(
            "Bid edge preference changed. Sorting columns will refresh automatically."
        )

    def load_pdfs(self) -> None:
        try:
            inputs = self._collect_inputs()

            if not self._check_matching_bid_period_or_warn(inputs):
                return

            self._save_inputs_to_config(inputs)

        except Exception as exc:
            QMessageBox.critical(self, "Input error", str(exc))
            return

        self.pdf_status_label.setText("Loading PDFs...")
        self._clear_bid_string("PDFs are loading. Generate the bid string after sorting.")
        self._reset_trip_progress("Trips extraction progress: waiting to start...")
        self._start_worker(self._load_worker, inputs, True)

    def export_excel(self) -> None:
        try:
            inputs = self._collect_inputs()

            if not self._check_matching_bid_period_or_warn(inputs):
                return

            self._save_inputs_to_config(inputs)

        except Exception as exc:
            QMessageBox.critical(self, "Input error", str(exc))
            return

        current_key = (inputs["trips_pdf_path"], inputs["lines_pdf_path"])
        if self.cached_pdf_key != current_key:
            self.pdf_status_label.setText("PDFs not loaded for these paths. Loading first, then exporting...")
            self._reset_trip_progress("Trips extraction progress: waiting to start...")

        self._start_worker(self._export_worker, inputs)

    def open_visualizer(self) -> None:
        try:
            inputs = self._collect_inputs()

            if not self._check_matching_bid_period_or_warn(inputs):
                return

            self._save_inputs_to_config(inputs)

        except Exception as exc:
            QMessageBox.critical(self, "Input error", str(exc))
            return

        current_key = (inputs["trips_pdf_path"], inputs["lines_pdf_path"])
        if self.cached_pdf_key != current_key:
            self.pdf_status_label.setText("PDFs not loaded for these paths. Loading first, then opening visualizer...")
            self._reset_trip_progress("Trips extraction progress: waiting to start...")
        else:
            self.pdf_status_label.setText("Opening visualizer...")

        self._start_worker(self._visualizer_worker, inputs)

    def _extract_trips_with_progress(self, trips_pdf_path: str) -> dict[str, Any]:
        try:
            return extract_trips_from_pdf(
                trips_pdf_path,
                first_page=2,
                progress_callback=self._queue_trip_progress,
            )
        except TypeError as exc:
            if "progress_callback" not in str(exc):
                raise

            self._log("Trip extractor does not support progress updates; loading trips without page progress.")
            self._queue_trip_progress({
                "current": 0,
                "total": 0,
                "status": "running",
                "message": "Trip extractor does not support progress updates.",
            })
            trips = extract_trips_from_pdf(trips_pdf_path, first_page=2)
            self._queue_trip_progress({
                "current": 1,
                "total": 1,
                "status": "done",
                "message": f"Finished extracting {len(trips)} trips.",
                "total_trips": len(trips),
            })
            return trips

    def _extract_pdfs(self, inputs: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        pdf_key = (inputs["trips_pdf_path"], inputs["lines_pdf_path"])

        if self.cached_pdf_key == pdf_key and self.cached_trips is not None and self.cached_lines is not None:
            self._log("Using already-loaded PDF data.")
            self._queue_trip_progress({
                "current": 1,
                "total": 1,
                "status": "cached",
                "message": "Using already-loaded trip data.",
                "total_trips": len(self.cached_trips),
            })
            return self.cached_trips, self.cached_lines

        self._log("Extracting PDFs...")
        self._queue_trip_progress({
            "current": 0,
            "total": 1,
            "status": "starting",
            "message": "Starting trip extraction...",
            "total_trips": 0,
        })

        with ThreadPoolExecutor(max_workers=2) as executor:
            trips_future = executor.submit(self._extract_trips_with_progress, inputs["trips_pdf_path"])
            lines_future = executor.submit(parse_line_report_pdf, inputs["lines_pdf_path"], first_calendar_page=3)

            trips = trips_future.result()
            lines = lines_future.result()

        self.cached_pdf_key = pdf_key
        self.cached_trips = trips
        self.cached_lines = lines

        self._log("PDF extraction complete.")
        return trips, lines

    @staticmethod
    def _pay_hours(value: Any) -> float:
        """Convert common UPS time formats to decimal hours."""
        if value is None:
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)

        text = str(value).strip()
        if not text or text in {"-", "None", "nan", "NaN"}:
            return 0.0

        hhmm_match = re.fullmatch(r"(\d+):(\d{1,2})", text)
        if hhmm_match:
            return int(hhmm_match.group(1)) + int(hhmm_match.group(2)) / 60.0

        trip_time_match = re.fullmatch(r"(\d+)h(\d{1,2})(?:[A-Za-z])?", text)
        if trip_time_match:
            return int(trip_time_match.group(1)) + int(trip_time_match.group(2)) / 60.0

        try:
            return float(text.replace(",", "").replace("$", ""))
        except ValueError:
            return 0.0

    @staticmethod
    def _pay_number(value: Any) -> float:
        if value is None:
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        text = str(value).strip().replace(",", "").replace("$", "")
        if not text or text == "-":
            return 0.0
        try:
            return float(text)
        except ValueError:
            return 0.0

    def _add_pay_fallback(
        self,
        master_lines: dict[Any, dict[str, Any]],
        hourly_rate: float,
        default_pp_guarantee_hours: float = 75.0,
    ) -> None:
        """
        Apply the documented pay formula when the current processing function
        encounters its local-variable ``pp`` bug.

        The normal Processing_fucntions.add_pay_to_master_lines() function is
        always attempted first. This fallback only runs for that specific bug.
        """
        hourly_rate = float(hourly_rate)

        for line_data in master_lines.values():
            pay_period_details: list[dict[str, Any]] = []
            total_extracted_credit = 0.0
            total_paid_credit = 0.0
            total_base_pay = 0.0
            total_premium = 0.0
            total_per_diem = 0.0

            for pp_index, pay_period in enumerate(line_data.get("PPs") or [], start=1):
                extracted_credit = self._pay_hours(pay_period.get("CT"))
                guarantee_hours = float(default_pp_guarantee_hours)
                paid_credit = max(extracted_credit, guarantee_hours)
                guarantee_added = max(0.0, paid_credit - extracted_credit)
                base_pay = paid_credit * hourly_rate

                total_extracted_credit += extracted_credit
                total_paid_credit += paid_credit
                total_base_pay += base_pay

                for assignment in pay_period.get("assignments") or []:
                    total_premium += self._pay_number(assignment.get("premium"))
                    total_per_diem += self._pay_number(assignment.get("per_diem"))

                pay_period_details.append({
                    "pp": pay_period.get("pp") or f"PP{pp_index}",
                    "extracted_credit_hours": round(extracted_credit, 2),
                    "guarantee_hours": round(guarantee_hours, 2),
                    "paid_credit_hours": round(paid_credit, 2),
                    "guarantee_credit_added": round(guarantee_added, 2),
                    "base_pay": round(base_pay, 2),
                })

            guarantee_credit_added = max(0.0, total_paid_credit - total_extracted_credit)
            taxable_pay = total_base_pay + total_premium
            cash_pay = taxable_pay + total_per_diem

            line_data["pay"] = {
                "hourly_rate": round(hourly_rate, 2),
                "pay_periods": pay_period_details,
                "extracted_CT": round(total_extracted_credit, 2),
                "paid_CT": round(total_paid_credit, 2),
                "guarantee_credit_added": round(guarantee_credit_added, 2),
                "total_base_pay": round(total_base_pay, 2),
                "total_premium": round(total_premium, 2),
                "total_per_diem": round(total_per_diem, 2),
                "taxable_pay_estimate": round(taxable_pay, 2),
                "cash_pay_estimate": round(cash_pay, 2),
            }

            # Flat numeric fields used by master_lines_to_dataframe and sorting.
            line_data["pay_cash_estimate"] = round(cash_pay, 2)
            line_data["pay_taxable_estimate"] = round(taxable_pay, 2)
            line_data["pay_base"] = round(total_base_pay, 2)
            line_data["pay_premium"] = round(total_premium, 2)
            line_data["pay_per_diem"] = round(total_per_diem, 2)
            line_data["paid_CT"] = round(total_paid_credit, 2)
            line_data["guarantee_credit_added"] = round(guarantee_credit_added, 2)

    def _build_dataframe(
        self,
        inputs: dict[str, Any],
        *,
        apply_sort: bool,
    ) -> tuple[pd.DataFrame, list[dict[str, Any]] | None]:
        trips, lines = self._extract_pdfs(inputs)

        bid_period_info = {
            key: lines[key]
            for key in ("bid_period_date_range", "pay_period_date_ranges")
        }

        # Build the bid-period airport lookup before creating_master_line.
        pdf_key = (inputs["trips_pdf_path"], inputs["lines_pdf_path"])
        if self.cached_airport_lookup_key == pdf_key and self.cached_airport_lookup is not None:
            self._log("Using cached bid-period airport lookup...")
            airport_lookup = self.cached_airport_lookup
        else:
            airports_csv_path = resource_path("airports.csv")
            if not airports_csv_path.exists():
                raise FileNotFoundError(
                    "airports.csv was not found next to the program. "
                    "Add it to the project folder and to the PyInstaller data files."
                )

            self._log("Collecting bid-period destination airports...")
            destination_codes = pf.collect_unique_arrival_destinations_from_trips(
                trips,
                ignore_sba_sbg=True,
            )

            self._log("Building bid-period airport lookup...")
            airport_lookup, unmatched_airports, matched_airports_df = (
                pf.build_bid_period_airport_lookup(
                    str(airports_csv_path),
                    destination_codes,
                    allowed_types=("large_airport", "medium_airport"),
                )
            )

            self.cached_airport_lookup_key = pdf_key
            self.cached_airport_lookup = airport_lookup
            self.cached_unmatched_airports = unmatched_airports
            self.cached_matched_airports_df = matched_airports_df

            try:
                unmatched_count = len(unmatched_airports)
            except TypeError:
                unmatched_count = 0

            if unmatched_count:
                self._log(
                    f"Airport lookup completed with {unmatched_count} "
                    "unmatched destination code(s)."
                )
            else:
                self._log("Airport lookup completed with all destination codes matched.")

        self._log("Creating master lines...")
        master_lines = creating_master_line(trips, lines)

        # Keep the processing calls explicit here so this method is the single,
        # easy-to-read list of everything applied to master_lines.
        self._log("Adding blockiness scores...")
        pf.add_blockiness_scores(master_lines, bid_period_info)

        self._log("Adding company-ticket percentages...")
        pf.add_company_ticket_percentages(master_lines)

        self._log("Adding line-type preference scores...")
        pf.add_line_type_preference_scores(
            master_lines,
            inputs["line_type_preference_order"],
            power_law_coeff=3
        )

        self._log("Adding estimated pay...")
        try:
            pf.add_pay(
                master_lines,
                inputs["hourly_rate"],
            )
        except UnboundLocalError as exc:
            if "local variable 'pp'" not in str(exc) and 'local variable "pp"' not in str(exc):
                raise
            self._log(
                "The current add_pay function hit its local-variable "
                "'pp' bug. Applying the same documented pay formula with the GUI fallback."
            )
            self._add_pay_fallback(master_lines, inputs["hourly_rate"])

        self._log("Adding complete-weekends-off percentages...")
        pf.add_weekends_off_percentage(master_lines)

        self._log("Adding international-destination scores...")
        pf.add_international_destination_scores(
            master_lines,
            airport_lookup,
        )

        if inputs["requested_dates"] is not None:
            self._log("Adding requested-days-off scores...")
            pf.add_requested_days_off_scores(
                master_lines,
                bid_period_info,
                inputs["requested_dates"],
            )

        if inputs["vacation_ranges"] is not None:
            self._log("Adding vacation scores...")
            new_vacation_ranges = pf.add_vacation_days_off_score(
                master_lines,
                inputs["vacation_ranges"],
                bid_period_info,
                save_details=False,
            )
        else:
            new_vacation_ranges = None

        if inputs["training_start"] is not None:
            self._log("Adding training-fit scores...")
            pf.add_training_fit_score(
                master_lines,
                inputs["training_start"],
                inputs["training_end"],
                bid_period_info,
            )

        self._log("Adding average legs per work day...")
        pf.add_avg_legs_per_work_day(master_lines)

        if inputs["bid_edge"] != "none":
            self._log("Adding bid-edge days-off scores...")
            pf.add_bid_edge_days_off(
                master_lines,
                bid_period_info,
                edge=inputs["bid_edge"],
            )

        self._log("Creating DataFrame...")
        df = master_lines_to_dataframe(master_lines, bid_period_info)
        

        if apply_sort and inputs["sort_order"]:
            sorting_settings = inputs.get("sorting_settings") or DEFAULT_SORTING_SETTINGS
            self._log(
                "Sorting DataFrame "
                f"({sorting_settings['default_mode']}, "
                f"{sorting_settings['weighting_style']}, "
                f"soft weights {sorting_settings['soft_min_weight']}–{sorting_settings['soft_max_weight']})."
            )
            df = drop_empty_sort_columns(df,check_all_columns=True)
            
            df = sort_dataframe_by_conditions(
                df,
                inputs["sort_order"],
                default_mode=sorting_settings["default_mode"],
                weighting_style=sorting_settings["weighting_style"],
                soft_max_weight=sorting_settings["soft_max_weight"],
                soft_min_weight=sorting_settings["soft_min_weight"],
            )

        return df, new_vacation_ranges

    def _load_worker(self, inputs: dict[str, Any], show_ready_message: bool = True) -> None:
        try:
            df, _ = self._build_dataframe(inputs, apply_sort=False)
            self.message_queue.put((
                "loaded",
                {
                    "df": df,
                    "show_ready_message": show_ready_message,
                    "status_text": "PDFs loaded. Sorting columns are ready."
                    if show_ready_message
                    else "Analyzer refreshed. Sorting columns are ready.",
                    "log_text": "Loaded PDFs" if show_ready_message else "Refreshed analyzer",
                },
            ))
        except Exception as exc:
            self.message_queue.put(("error", exc))

    def _get_bid_spreadsheet_viewer_class(self) -> Any:
        """Return a PySide6 BidSpreadsheetWindow/BidSpreadsheetViewer class, or raise a clear import error."""
        errors: list[str] = []

        try:
            from GUI_spreadsheet_pyside6 import BidSpreadsheetWindow
            return BidSpreadsheetWindow
        except Exception as exc:
            errors.append(f"GUI_spreadsheet_pyside6.BidSpreadsheetWindow: {exc}")

        try:
            from GUI_spreadsheet_pyside6 import BidSpreadsheetViewer
            return BidSpreadsheetViewer
        except Exception as exc:
            errors.append(f"GUI_spreadsheet_pyside6.BidSpreadsheetViewer: {exc}")

        try:
            from excel_killer_pyside6 import BidSpreadsheetWindow
            return BidSpreadsheetWindow
        except Exception as exc:
            errors.append(f"excel_killer_pyside6.BidSpreadsheetWindow: {exc}")

        try:
            from excel_killer_pyside6 import BidSpreadsheetViewer
            return BidSpreadsheetViewer
        except Exception as exc:
            errors.append(f"excel_killer_pyside6.BidSpreadsheetViewer: {exc}")

        raise RuntimeError(
            "Could not import a PySide6 visualizer. Put GUI_spreadsheet_pyside6.py "
            "or excel_killer_pyside6.py in the same folder as this GUI.\n\n"
            + "\n".join(errors)
        )

    def _open_visualizer_from_payload(self, payload: Any) -> None:
        if not isinstance(payload, dict):
            raise RuntimeError("Visualizer payload was not valid.")

        sorted_df = payload.get("df")
        if sorted_df is None or getattr(sorted_df, "empty", False):
            raise RuntimeError("There is no sorted DataFrame to display yet.")

        training_start = payload.get("training_start")
        training_end = payload.get("training_end")
        vacation_ranges = payload.get("vacation_ranges") or []
        requested_days_off_dates = payload.get("requested_days_off_dates") or []

        viewer_class = self._get_bid_spreadsheet_viewer_class()

        try:
            # Works with BidSpreadsheetWindow from GUI_spreadsheet_pyside6.py.
            viewer_window = viewer_class(
                sorted_df,
                training_start=training_start,
                training_end=training_end,
                vacation_ranges=vacation_ranges,
                requested_days_off_dates= requested_days_off_dates,
            )
        except TypeError:
            # Fallback for a QWidget-based viewer that expects parent=None and a later load_dataframe call.
            viewer_window = QMainWindow(self)
            viewer_window.setWindowTitle("Bid Table Viewer")
            viewer_window.resize(1400, 800)
            viewer = viewer_class(parent=viewer_window)
            viewer_window.setCentralWidget(viewer)
            viewer.load_dataframe(
                sorted_df,
                training_start=training_start,
                training_end=training_end,
                vacation_ranges=vacation_ranges,
                requested_days_off_dates= requested_days_off_dates,
            )

        if isinstance(viewer_window, QWidget):
            viewer_window.setWindowTitle("Bid Table Viewer")
            viewer_window.resize(1400, 800)
            viewer_window.setMinimumSize(900, 500)
            apply_window_icon(viewer_window)
            viewer_window.show()
            self.visualizer_windows.append(viewer_window)
        else:
            raise RuntimeError("The imported visualizer does not appear to be a PySide6 QWidget/QMainWindow.")

        self.pdf_status_label.setText("Visualizer opened.")
        self._write_log("Opened Bid Table Viewer with sorted DataFrame.")

    def _visualizer_worker(self, inputs: dict[str, Any]) -> None:
        try:
            sorted_df, vacation_ranges = self._build_dataframe(inputs, apply_sort=True)
            self.message_queue.put((
                "visualizer_ready",
                {
                    "df": sorted_df,
                    "training_start": inputs["training_start"],
                    "training_end": inputs["training_end"],
                    "vacation_ranges": vacation_ranges,
                    "requested_days_off_dates": inputs["requested_dates"],
                },
            ))
        except Exception as exc:
            self.message_queue.put(("error", exc))

    def _export_worker(self, inputs: dict[str, Any]) -> None:
        try:
            df, new_vacation_ranges = self._build_dataframe(inputs, apply_sort=True)

            self._log("Generating bid string...")
            bid_string = pf.line_numbers_to_bid_string(df, inputs["number_of_lines_to_bid"])

            self._log("Exporting Excel file...")
            export_master_lines_to_excel_table(
                df,
                str(inputs["output_path"]),
                training_start=inputs["training_start"],
                training_end=inputs["training_end"],
                vacation_ranges=new_vacation_ranges,
            )

            self.message_queue.put((
                "exported",
                {
                    "output_path": inputs["output_path"],
                    "bid_string": bid_string,
                    "number_of_lines": inputs["number_of_lines_to_bid"],
                },
            ))
        except Exception as exc:
            self.message_queue.put(("error", exc))

    def _bid_string_worker(self, inputs: dict[str, Any], copy_after: bool = False) -> None:
        try:
            df, _ = self._build_dataframe(inputs, apply_sort=True)
            self._log("Generating bid string...")
            bid_string = pf.line_numbers_to_bid_string(df, inputs["number_of_lines_to_bid"])

            self.message_queue.put((
                "bid_string_ready",
                {
                    "bid_string": bid_string,
                    "number_of_lines": inputs["number_of_lines_to_bid"],
                    "copy_after": copy_after,
                },
            ))
        except Exception as exc:
            self.message_queue.put(("error", exc))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BidGUI()
    window.show()
    sys.exit(app.exec())
