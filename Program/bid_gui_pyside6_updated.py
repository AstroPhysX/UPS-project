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
      such as bid_spreadsheet_viewer_pyside6.py or excel_killer_pyside6.py.
"""

from __future__ import annotations

import ctypes
import json
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

from PySide6.QtCore import QDate, QPoint, QSize, QTimer, Qt
from PySide6.QtGui import QColor, QIcon, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCalendarWidget,
    QCheckBox,
    QComboBox,
    QDialog,
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

from Trips_Extractor import extract_trips_from_pdf
from Lines_Extractor import parse_line_report_pdf
from master_lines_creation import creating_master_line
from master_to_pandas import master_lines_to_dataframe, sort_dataframe_by_conditions
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
    """Dialog for adding or editing a vacation range."""

    def __init__(
        self,
        parent: QWidget,
        title: str,
        initial_start: str = "",
        initial_end: str = "",
    ) -> None:
        super().__init__(parent)
        self.result: dict[str, str] | None = None
        self.setWindowTitle(title)
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self.start_entry = DateEntry(self)
        self.start_entry.setText(initial_start)
        self.end_entry = DateEntry(self)
        self.end_entry.setText(initial_end)

        main = QGridLayout(self)
        main.addWidget(QLabel("Start date:"), 0, 0)
        main.addWidget(self.start_entry, 0, 1)
        main.addWidget(QLabel("End date:"), 1, 0)
        main.addWidget(self.end_entry, 1, 1)

        save_button = QPushButton("Save")
        cancel_button = QPushButton("Cancel")
        save_button.clicked.connect(self._save)
        cancel_button.clicked.connect(self.reject)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        buttons.addWidget(cancel_button)
        buttons.addWidget(save_button)
        main.addLayout(buttons, 2, 0, 1, 2)

        self.resize(360, 130)

    def _save(self) -> None:
        try:
            start = validate_required_date(self.start_entry.text(), "Vacation start")
            end = validate_required_date(self.end_entry.text(), "Vacation end")
            if end < start:
                raise ValueError("Vacation end date is before vacation start date.")
            self.result = {"start": start, "end": end}
            self.accept()
        except Exception as exc:
            QMessageBox.critical(self, "Vacation date error", str(exc))


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
# Main GUI
# ---------------------------------------------------------------------------

class BidGUI(QMainWindow):
    def __init__(self) -> None:
        set_windows_app_id()
        super().__init__()

        self.setWindowTitle("UPS Bid Analyzer")
        self.resize(1080, 780)
        self.setMinimumSize(960, 640)
        apply_window_icon(self)

        self.config_data = load_saved_config()
        self.message_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self.preview_df: pd.DataFrame | None = None
        self.cached_lines: dict[str, Any] | None = None
        self.cached_trips: dict[str, Any] | None = None
        self.cached_pdf_key: tuple[str, str] | None = None

        self.sort_order: list[list[str]] = []
        self.latest_bid_string = ""
        self.visualizer_windows: list[QWidget] = []

        self._setup_style()
        self._build_ui()
        self._load_saved_values()

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
            QLineEdit, QTextEdit, QListWidget, QTableWidget, QComboBox {{
                background: white;
                color: black;
                selection-background-color: {UPS_BLUE};
                selection-color: white;
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

    def _build_ui(self) -> None:
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
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
        grid = QGridLayout(prefs_frame)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)

        grid.addWidget(QLabel("Vacation ranges:"), 0, 0, alignment=Qt.AlignTop)

        vacation_area = QWidget()
        vacation_layout = QHBoxLayout(vacation_area)
        vacation_layout.setContentsMargins(0, 0, 0, 0)

        self.vacation_table = QTableWidget(0, 2)
        self.vacation_table.setHorizontalHeaderLabels(["Start", "End"])
        self.vacation_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.vacation_table.verticalHeader().setVisible(False)
        self.vacation_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.vacation_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.vacation_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.vacation_table.setMinimumHeight(120)
        vacation_layout.addWidget(self.vacation_table, 1)

        vacation_buttons = QVBoxLayout()
        add_vacation = QPushButton("Add vacation range")
        edit_vacation = QPushButton("Edit selected")
        remove_vacation = QPushButton("Remove selected")
        clear_vacation = QPushButton("Clear all")
        add_vacation.clicked.connect(self._add_vacation_range)
        edit_vacation.clicked.connect(self._edit_vacation_range)
        remove_vacation.clicked.connect(self._remove_vacation_range)
        clear_vacation.clicked.connect(self._clear_vacation_ranges)
        for button in (add_vacation, edit_vacation, remove_vacation, clear_vacation):
            vacation_buttons.addWidget(button)
        vacation_buttons.addStretch(1)
        vacation_layout.addLayout(vacation_buttons)

        grid.addWidget(vacation_area, 0, 1, 1, 3)

        self.training_start_entry = DateEntry(self)
        self.training_end_entry = DateEntry(self)
        self.training_start_entry.line_edit.textChanged.connect(
            lambda: self._mark_bid_string_stale("Training dates changed. Generate the bid string again after sorting.")
        )
        self.training_end_entry.line_edit.textChanged.connect(
            lambda: self._mark_bid_string_stale("Training dates changed. Generate the bid string again after sorting.")
        )

        grid.addWidget(QLabel("Training start:"), 1, 0)
        grid.addWidget(self.training_start_entry, 1, 1, alignment=Qt.AlignLeft)
        grid.addWidget(QLabel("Training end:"), 1, 2)
        grid.addWidget(self.training_end_entry, 1, 3, alignment=Qt.AlignLeft)

        self.bid_edge_combo = QComboBox()
        self.bid_edge_combo.addItems(["none", "start", "end", "both"])
        self.bid_edge_combo.currentTextChanged.connect(lambda _text: self._on_bid_edge_changed())

        grid.addWidget(QLabel("Bid edge days off:"), 2, 0)
        grid.addWidget(self.bid_edge_combo, 2, 1, alignment=Qt.AlignLeft)

        container.addWidget(prefs_frame)

    def _build_sorting_section(self, container: QVBoxLayout) -> None:
        sort_frame = QGroupBox("Sorting")
        grid = QGridLayout(sort_frame)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(2, 1)

        grid.addWidget(QLabel("Available columns"), 0, 0)
        grid.addWidget(QLabel("Selected sorting priority"), 0, 2)

        self.available_columns_list = QListWidget()
        self.available_columns_list.setMinimumHeight(180)
        self.selected_sort_list = QListWidget()
        self.selected_sort_list.setMinimumHeight(180)

        grid.addWidget(self.available_columns_list, 1, 0)
        grid.addWidget(self.selected_sort_list, 1, 2)

        sort_buttons_widget = QWidget()
        sort_buttons = QVBoxLayout(sort_buttons_widget)
        sort_buttons.setContentsMargins(0, 0, 0, 0)

        add_high = QPushButton("Add high →")
        add_low = QPushButton("Add low →")
        move_up = QPushButton("Move up")
        move_down = QPushButton("Move down")
        remove = QPushButton("Remove")
        clear = QPushButton("Clear")
        advanced = QPushButton("Advanced settings...")

        add_high.clicked.connect(lambda: self._add_sort_column("desc"))
        add_low.clicked.connect(lambda: self._add_sort_column("asc"))
        move_up.clicked.connect(self._move_sort_up)
        move_down.clicked.connect(self._move_sort_down)
        remove.clicked.connect(self._remove_sort_column)
        clear.clicked.connect(self._clear_sort_order)
        advanced.clicked.connect(self._open_sorting_settings_dialog)

        for button in (add_high, add_low):
            sort_buttons.addWidget(button)
        line1 = QFrame()
        line1.setFrameShape(QFrame.HLine)
        sort_buttons.addWidget(line1)
        for button in (move_up, move_down, remove, clear):
            sort_buttons.addWidget(button)
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        sort_buttons.addWidget(line2)
        sort_buttons.addWidget(advanced)
        sort_buttons.addStretch(1)

        grid.addWidget(sort_buttons_widget, 1, 1)

        self.default_mode = str(DEFAULT_SORTING_SETTINGS["default_mode"])
        self.weighting_style = str(DEFAULT_SORTING_SETTINGS["weighting_style"])
        self.soft_max_weight = float(DEFAULT_SORTING_SETTINGS["soft_max_weight"])
        self.soft_min_weight = float(DEFAULT_SORTING_SETTINGS["soft_min_weight"])
        self.keep_score_columns = bool(DEFAULT_SORTING_SETTINGS["keep_score_columns"])

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

        self.training_start_entry.setText(self.config_data.get("training_start") or "")
        self.training_end_entry.setText(self.config_data.get("training_end") or "")

        bid_edge = self.config_data.get("bid_edge") or "none"
        index = self.bid_edge_combo.findText(bid_edge)
        self.bid_edge_combo.blockSignals(True)
        self.bid_edge_combo.setCurrentIndex(index if index >= 0 else 0)
        self.bid_edge_combo.blockSignals(False)

        output_paths = self.config_data.get("output_paths", {})
        saved_output_folder = output_paths.get(get_os_name(), "")
        self.output_folder_edit.setText(saved_output_folder or str(Path.cwd()))

        saved_number_of_lines = self.config_data.get("number_of_lines_to_bid", DEFAULT_NUMBER_OF_LINES_TO_BID)
        try:
            saved_number_of_lines = int(saved_number_of_lines)
            if saved_number_of_lines <= 0:
                saved_number_of_lines = DEFAULT_NUMBER_OF_LINES_TO_BID
        except (TypeError, ValueError):
            saved_number_of_lines = DEFAULT_NUMBER_OF_LINES_TO_BID
        self.number_of_lines_edit.setText(str(saved_number_of_lines))

        saved_sort_order = self.config_data.get("sort_order", [])
        if isinstance(saved_sort_order, list):
            self.sort_order = [list(item) for item in saved_sort_order if isinstance(item, list) and len(item) == 2]
            self._refresh_selected_sort_list()

        saved_sorting_settings = self.config_data.get("sorting_settings", {})
        if not isinstance(saved_sorting_settings, dict):
            saved_sorting_settings = {}
        merged_settings = {**DEFAULT_SORTING_SETTINGS, **saved_sorting_settings}
        self._apply_sorting_settings_to_ui(merged_settings)

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
            self._refresh_available_columns_list([])
            self._clear_bid_string("PDF paths changed. Generate the bid string again after loading/sorting.")

    # -------------------------- Vacation range actions --------------------------

    def _set_vacation_ranges(self, vacation_ranges: list[dict[str, str]] | None) -> None:
        self.vacation_table.setRowCount(0)
        for vacation in vacation_ranges or []:
            start = vacation.get("start", "")
            end = vacation.get("end", "")
            if start and end:
                row = self.vacation_table.rowCount()
                self.vacation_table.insertRow(row)
                self.vacation_table.setItem(row, 0, QTableWidgetItem(start))
                self.vacation_table.setItem(row, 1, QTableWidgetItem(end))

    def _get_vacation_ranges(self) -> list[dict[str, str]]:
        ranges: list[dict[str, str]] = []
        for row in range(self.vacation_table.rowCount()):
            start_item = self.vacation_table.item(row, 0)
            end_item = self.vacation_table.item(row, 1)
            start = validate_required_date(start_item.text() if start_item else "", "Vacation start")
            end = validate_required_date(end_item.text() if end_item else "", "Vacation end")
            if end < start:
                raise ValueError(f"Vacation range {start} to {end}: end date is before start date.")
            ranges.append({"start": start, "end": end})
        return ranges

    def _selected_vacation_row(self) -> int | None:
        rows = self.vacation_table.selectionModel().selectedRows()
        if not rows:
            return None
        return rows[0].row()

    def _add_vacation_range(self) -> None:
        dialog = VacationRangeDialog(self, "Add vacation range")
        if dialog.exec() == QDialog.Accepted and dialog.result:
            row = self.vacation_table.rowCount()
            self.vacation_table.insertRow(row)
            self.vacation_table.setItem(row, 0, QTableWidgetItem(dialog.result["start"]))
            self.vacation_table.setItem(row, 1, QTableWidgetItem(dialog.result["end"]))
            self._clear_bid_string("Vacation ranges changed. Generate the bid string again after sorting.")

    def _edit_vacation_range(self) -> None:
        row = self._selected_vacation_row()
        if row is None:
            QMessageBox.information(self, "Edit vacation range", "Select a vacation range first.")
            return

        start = self.vacation_table.item(row, 0).text() if self.vacation_table.item(row, 0) else ""
        end = self.vacation_table.item(row, 1).text() if self.vacation_table.item(row, 1) else ""
        dialog = VacationRangeDialog(self, "Edit vacation range", start, end)
        if dialog.exec() == QDialog.Accepted and dialog.result:
            self.vacation_table.setItem(row, 0, QTableWidgetItem(dialog.result["start"]))
            self.vacation_table.setItem(row, 1, QTableWidgetItem(dialog.result["end"]))
            self._clear_bid_string("Vacation ranges changed. Generate the bid string again after sorting.")

    def _remove_vacation_range(self) -> None:
        row = self._selected_vacation_row()
        if row is not None:
            self.vacation_table.removeRow(row)
            self._clear_bid_string("Vacation ranges changed. Generate the bid string again after sorting.")

    def _clear_vacation_ranges(self) -> None:
        self.vacation_table.setRowCount(0)
        self._clear_bid_string("Vacation ranges changed. Generate the bid string again after sorting.")

    # -------------------------- Sorting list actions --------------------------

    def _add_sort_column(self, direction: str) -> None:
        item = self.available_columns_list.currentItem()
        if item is None:
            return

        col = item.text()
        self.sort_order = [rule for rule in self.sort_order if rule[0] != col]
        self.sort_order.append([col, direction])
        self._refresh_selected_sort_list()
        self._clear_bid_string("Sorting priority changed. Generate the bid string again.")

    def _remove_sort_column(self) -> None:
        row = self.selected_sort_list.currentRow()
        if row < 0:
            return
        del self.sort_order[row]
        self._refresh_selected_sort_list()
        self._clear_bid_string("Sorting priority changed. Generate the bid string again.")

    def _clear_sort_order(self) -> None:
        self.sort_order = []
        self._refresh_selected_sort_list()
        self._clear_bid_string("Sorting priority changed. Generate the bid string again.")

    def _move_sort_up(self) -> None:
        row = self.selected_sort_list.currentRow()
        if row <= 0:
            return
        self.sort_order[row - 1], self.sort_order[row] = self.sort_order[row], self.sort_order[row - 1]
        self._refresh_selected_sort_list(select_index=row - 1)
        self._clear_bid_string("Sorting priority changed. Generate the bid string again.")

    def _move_sort_down(self) -> None:
        row = self.selected_sort_list.currentRow()
        if row < 0 or row >= len(self.sort_order) - 1:
            return
        self.sort_order[row + 1], self.sort_order[row] = self.sort_order[row], self.sort_order[row + 1]
        self._refresh_selected_sort_list(select_index=row + 1)
        self._clear_bid_string("Sorting priority changed. Generate the bid string again.")

    def _refresh_available_columns_list(self, columns: list[str]) -> None:
        self.available_columns_list.clear()
        for col in columns:
            self.available_columns_list.addItem(str(col))

    def _refresh_selected_sort_list(self, select_index: int | None = None) -> None:
        self.selected_sort_list.clear()
        for col, direction in self.sort_order:
            label = "high-to-low" if direction == "desc" else "low-to-high"
            self.selected_sort_list.addItem(f"{col} ({label})")

        if select_index is not None and 0 <= select_index < len(self.sort_order):
            self.selected_sort_list.setCurrentRow(select_index)

    def _clean_sort_order_for_columns(self, columns: list[str]) -> None:
        valid_columns = set(columns)
        old_sort_order = list(self.sort_order)
        self.sort_order = [rule for rule in self.sort_order if len(rule) == 2 and rule[0] in valid_columns]
        self._refresh_selected_sort_list()

        skipped = [rule[0] for rule in old_sort_order if len(rule) == 2 and rule[0] not in valid_columns]
        if skipped:
            self._write_log("Skipped saved sorting columns not found in this DataFrame: " + ", ".join(skipped))

    # -------------------------- Sorting settings --------------------------

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
        soft_max_entry.setToolTip("Used by the soft weighting style. This is the weight given to the first item in each weighted group. Default: 3.0.")
        soft_min_entry = QLineEdit(str(current["soft_min_weight"]))
        soft_min_entry.setToolTip("Used by the soft weighting style. This is the weight given to the last item in each weighted group. Default: 1.0.")
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
        add_row(3, "Soft max weight:", soft_max_entry, soft_max_entry.toolTip())
        add_row(4, "Soft min weight:", soft_min_entry, soft_min_entry.toolTip())
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

        training_start = validate_date_or_blank(self.training_start_entry.text(), "Training start")
        training_end = validate_date_or_blank(self.training_end_entry.text(), "Training end")

        if bool(training_start) != bool(training_end):
            raise ValueError("Enter both training start and training end, or leave both blank.")
        if training_start and training_end and training_end < training_start:
            raise ValueError("Training end date is before training start date.")

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
            "training_start": training_start,
            "training_end": training_end,
            "bid_edge": bid_edge,
            "output_folder": output_folder,
            "output_path": output_path,
            "number_of_lines_to_bid": number_of_lines_to_bid,
            "sort_order": self.sort_order,
            "sorting_settings": self._get_sorting_settings(),
        }

    def _save_inputs_to_config(self, inputs: dict[str, Any]) -> None:
        self.config_data["vacation_ranges"] = inputs["vacation_ranges"]
        self.config_data["training_start"] = inputs["training_start"]
        self.config_data["training_end"] = inputs["training_end"]
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
        self.config_data["bid_edge"] = self.bid_edge_combo.currentText().strip().lower() or "none"
        save_config(self.config_data)
        self._clear_bid_string("Bid edge changed. Generate the bid string again after sorting.")

        if not self.trips_path_edit.text().strip() or not self.lines_path_edit.text().strip():
            self.pdf_status_label.setText("Bid edge preference saved. Choose PDFs when ready.")
            return

        if self.worker_thread and self.worker_thread.is_alive():
            self.pdf_status_label.setText("Bid edge changed. Refresh after the current job finishes.")
            return

        try:
            inputs = self._collect_inputs()
            self._save_inputs_to_config(inputs)
        except Exception as exc:
            self.pdf_status_label.setText("Bid edge changed. Analyzer refresh skipped.")
            self._write_log(f"Bid edge changed, but analyzer could not refresh yet: {exc}")
            return

        if self.cached_trips is None or self.cached_lines is None:
            self.pdf_status_label.setText("Bid edge changed. Click Load PDFs into UPS Bid Analyzer.")
            return

        self.pdf_status_label.setText("Bid edge changed. Refreshing analyzer...")
        self._start_worker(self._load_worker, inputs, False)

    def load_pdfs(self) -> None:
        try:
            inputs = self._collect_inputs()
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

    def _build_dataframe(self, inputs: dict[str, Any], *, apply_sort: bool) -> tuple[pd.DataFrame, list[dict[str, str]] | None]:
        trips, lines = self._extract_pdfs(inputs)

        bid_period_info = {x: lines[x] for x in ("bid_period_date_range", "pay_period_date_ranges")}

        self._log("Creating master lines...")
        master_lines = creating_master_line(trips, lines)

        self._log("Adding scores...")
        pf.add_blockiness_scores(master_lines, bid_period_info)
        pf.add_company_ticket_percentages(master_lines)

        if inputs["vacation_ranges"] is not None:
            new_vacation_ranges = pf.add_vacation_days_off_score(
                master_lines,
                inputs["vacation_ranges"],
                bid_period_info,
                save_details=False,
            )
        else:
            new_vacation_ranges = None

        if inputs["training_start"] is not None or inputs["training_end"] is not None:
            pf.add_training_fit_score(
                master_lines,
                inputs["training_start"],
                inputs["training_end"],
                bid_period_info,
            )

        pf.add_avg_legs_per_work_day(master_lines)

        if inputs["bid_edge"] != "none":
            pf.add_bid_edge_days_off(master_lines, bid_period_info, edge=inputs["bid_edge"])

        self._log("Creating DataFrame...")
        df = master_lines_to_dataframe(master_lines, bid_period_info)

        if apply_sort and inputs["sort_order"]:
            sorting_settings = inputs.get("sorting_settings") or DEFAULT_SORTING_SETTINGS
            self._log(
                "Sorting DataFrame "
                f"({sorting_settings['default_mode']}, "
                f"{sorting_settings['weighting_style']}, "
                f"soft weights {sorting_settings['soft_min_weight']}–{sorting_settings['soft_max_weight']})..."
            )
            df = sort_dataframe_by_conditions(
                df,
                inputs["sort_order"],
                default_mode=sorting_settings["default_mode"],
                weighting_style=sorting_settings["weighting_style"],
                soft_max_weight=sorting_settings["soft_max_weight"],
                soft_min_weight=sorting_settings["soft_min_weight"],
                keep_score_columns=sorting_settings["keep_score_columns"],
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
            from bid_spreadsheet_viewer_pyside6 import BidSpreadsheetWindow
            return BidSpreadsheetWindow
        except Exception as exc:
            errors.append(f"bid_spreadsheet_viewer_pyside6.BidSpreadsheetWindow: {exc}")

        try:
            from bid_spreadsheet_viewer_pyside6 import BidSpreadsheetViewer
            return BidSpreadsheetViewer
        except Exception as exc:
            errors.append(f"bid_spreadsheet_viewer_pyside6.BidSpreadsheetViewer: {exc}")

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
            "Could not import a PySide6 visualizer. Put bid_spreadsheet_viewer_pyside6.py "
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

        viewer_class = self._get_bid_spreadsheet_viewer_class()

        try:
            # Works with BidSpreadsheetWindow from bid_spreadsheet_viewer_pyside6.py.
            viewer_window = viewer_class(
                sorted_df,
                training_start=training_start,
                training_end=training_end,
                vacation_ranges=vacation_ranges,
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
