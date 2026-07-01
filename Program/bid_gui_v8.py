"""
Tkinter GUI wrapper for the UPS bid-processing script.

Put this file in the same folder as your existing project files, for example:
    Trips_Extractor.py
    Lines_Extractor.py
    master_lines_creation.py
    master_to_pandas.py
    export_to_excel.py
    Processing_fucntions.py

Run with:
    python bid_gui_v8.py

This version adds:
    - Calendar popups for training dates and vacation ranges
    - Add/Edit/Remove buttons for vacation ranges instead of a large text box
    - A "Load PDFs into UPS Bid Analyzer" button next to the PDF selectors
    - Reuse of already-extracted PDF data during export, unless PDF paths change
    - Automatic analyzer refresh when the bid-edge days-off option changes
    - UPS-inspired brown/gold theme with blue buttons and green export button
    - Main-window scrollbar for smaller screens
    - Saved advanced sorting settings, opened from a popup instead of the main workflow
    - Hover tooltips explaining sorting modes and weighting styles
"""

from __future__ import annotations

import calendar
import json
import platform
import queue
import re
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

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

DEFAULT_MODE_DESCRIPTIONS = {
    "strict": "Normal priority / tie-breaker sort. The first selected column dominates, then the next column breaks ties, and so on.",
    "weighted": "Consecutive weighted conditions are blended into a combined percentile-rank score.",
}

WEIGHTING_STYLE_DESCRIPTIONS = {
    "equal": "Every weighted item in a group gets the same weight: 1.",
    "hard": "Position-based weights. Earlier selected columns matter much more, for example 4, 3, 2, 1.",
    "soft": "Softer position-based weights. With defaults, four weighted items become about 3.0, 2.33, 1.67, 1.0.",
}

SORT_TOOLTIP_DELAY_MS = 450

UPS_BROWN = "#351C15"
UPS_BROWN_2 = "#4B2618"
UPS_GOLD = "#FFB500"
UPS_GOLD_DARK = "#C99700"
UPS_BLUE = "#1F6FEB"
UPS_BLUE_ACTIVE = "#1557B0"
UPS_GREEN = "#2E7D32"
UPS_GREEN_ACTIVE = "#1B5E20"
UPS_TEXT = "#FFF7E6"
UPS_PANEL = "#432116"
UPS_FIELD_BG = "#FFF8DC"

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


# ---------------------------------------------------------------------------
# Calendar popup and date entry widgets
# ---------------------------------------------------------------------------

class CalendarPopup(tk.Toplevel):
    """Small dependency-free calendar popup that returns an ISO date string."""

    def __init__(
        self,
        parent: tk.Widget,
        initial_value: str | None,
        callback: Callable[[str], None],
    ) -> None:
        super().__init__(parent)

        self.callback = callback
        initial_date = iso_to_date(initial_value)
        self.year = initial_date.year
        self.month = initial_date.month

        self.title("Choose date")
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        self.grab_set()

        self.header_var = tk.StringVar()

        main = ttk.Frame(self, padding=8)
        main.pack(fill="both", expand=True)

        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 6))

        ttk.Button(header, text="‹", width=3, command=self._previous_month).pack(side="left")
        ttk.Label(header, textvariable=self.header_var, width=22, anchor="center").pack(side="left", expand=True)
        ttk.Button(header, text="›", width=3, command=self._next_month).pack(side="right")

        self.days_frame = ttk.Frame(main)
        self.days_frame.pack(fill="both", expand=True)

        footer = ttk.Frame(main)
        footer.pack(fill="x", pady=(8, 0))
        ttk.Button(footer, text="Today", command=self._choose_today).pack(side="left")
        ttk.Button(footer, text="Clear", command=self._clear_date).pack(side="left", padx=(6, 0))
        ttk.Button(footer, text="Cancel", command=self.destroy).pack(side="right")

        self._draw_calendar()
        self._position_near_parent(parent)

    def _position_near_parent(self, parent: tk.Widget) -> None:
        self.update_idletasks()
        try:
            x = parent.winfo_rootx()
            y = parent.winfo_rooty() + parent.winfo_height()
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _draw_calendar(self) -> None:
        for child in self.days_frame.winfo_children():
            child.destroy()

        self.header_var.set(f"{calendar.month_name[self.month]} {self.year}")

        for col, day_name in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]):
            ttk.Label(self.days_frame, text=day_name, anchor="center", width=5).grid(
                row=0,
                column=col,
                padx=1,
                pady=1,
            )

        month_calendar = calendar.monthcalendar(self.year, self.month)

        for row_index, week in enumerate(month_calendar, start=1):
            for col_index, day_number in enumerate(week):
                if day_number == 0:
                    ttk.Label(self.days_frame, text="", width=5).grid(
                        row=row_index,
                        column=col_index,
                        padx=1,
                        pady=1,
                    )
                    continue

                chosen = date(self.year, self.month, day_number)
                ttk.Button(
                    self.days_frame,
                    text=str(day_number),
                    width=5,
                    command=lambda d=chosen: self._choose_date(d),
                ).grid(row=row_index, column=col_index, padx=1, pady=1)

    def _previous_month(self) -> None:
        if self.month == 1:
            self.month = 12
            self.year -= 1
        else:
            self.month -= 1
        self._draw_calendar()

    def _next_month(self) -> None:
        if self.month == 12:
            self.month = 1
            self.year += 1
        else:
            self.month += 1
        self._draw_calendar()

    def _choose_today(self) -> None:
        self._choose_date(date.today())

    def _clear_date(self) -> None:
        self.callback("")
        self.destroy()

    def _choose_date(self, chosen: date) -> None:
        self.callback(chosen.isoformat())
        self.destroy()


class DateEntry(ttk.Frame):
    """Entry + calendar button. Stores dates as YYYY-MM-DD."""

    def __init__(self, parent: tk.Widget, textvariable: tk.StringVar | None = None, width: int = 14) -> None:
        super().__init__(parent)
        self.variable = textvariable or tk.StringVar()

        self.entry = ttk.Entry(self, textvariable=self.variable, width=width)
        self.entry.pack(side="left", fill="x", expand=True)

        self.button = ttk.Button(self, text="📅", width=3, command=self._open_calendar)
        self.button.pack(side="left", padx=(4, 0))

    def get(self) -> str:
        return self.variable.get().strip()

    def set(self, value: str | None) -> None:
        self.variable.set(value or "")

    def _open_calendar(self) -> None:
        CalendarPopup(self, self.get(), self.set)


class ToolTip:
    """Small hover tooltip for explaining advanced settings."""

    def __init__(self, widget: tk.Widget, text: str, *, wraplength: int = 340) -> None:
        self.widget = widget
        self.text = text
        self.wraplength = wraplength
        self.tip_window: tk.Toplevel | None = None
        self.after_id: str | None = None

        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event: tk.Event | None = None) -> None:
        self._cancel_schedule()
        self.after_id = self.widget.after(SORT_TOOLTIP_DELAY_MS, self._show)

    def _cancel_schedule(self) -> None:
        if self.after_id is not None:
            try:
                self.widget.after_cancel(self.after_id)
            except Exception:
                pass
            self.after_id = None

    def _show(self) -> None:
        self.after_id = None
        if self.tip_window is not None or not self.text:
            return

        try:
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        except Exception:
            x, y = 100, 100

        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")

        label = tk.Label(
            self.tip_window,
            text=self.text,
            justify="left",
            background="#FFF8DC",
            foreground="black",
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=6,
            wraplength=self.wraplength,
        )
        label.pack()

    def _hide(self, _event: tk.Event | None = None) -> None:
        self._cancel_schedule()
        if self.tip_window is not None:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None


class VacationRangeDialog(tk.Toplevel):
    """Dialog for adding or editing a vacation range."""

    def __init__(
        self,
        parent: tk.Widget,
        title: str,
        initial_start: str = "",
        initial_end: str = "",
    ) -> None:
        super().__init__(parent)

        self.result: dict[str, str] | None = None

        self.title(title)
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        self.grab_set()

        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)

        ttk.Label(main, text="Start date:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
        self.start_var = tk.StringVar(value=initial_start)
        self.start_entry = DateEntry(main, self.start_var)
        self.start_entry.grid(row=0, column=1, sticky="ew", pady=6)

        ttk.Label(main, text="End date:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        self.end_var = tk.StringVar(value=initial_end)
        self.end_entry = DateEntry(main, self.end_var)
        self.end_entry.grid(row=1, column=1, sticky="ew", pady=6)

        buttons = ttk.Frame(main)
        buttons.grid(row=2, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Save", command=self._save).pack(side="right", padx=(0, 8))

        self.bind("<Return>", lambda _event: self._save())
        self.bind("<Escape>", lambda _event: self.destroy())

        self.update_idletasks()
        self._center_on_parent(parent)
        self.start_entry.entry.focus_set()

    def _center_on_parent(self, parent: tk.Widget) -> None:
        try:
            parent_root = parent.winfo_toplevel()
            x = parent_root.winfo_rootx() + 80
            y = parent_root.winfo_rooty() + 80
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _save(self) -> None:
        try:
            start = validate_required_date(self.start_var.get(), "Vacation start")
            end = validate_required_date(self.end_var.get(), "Vacation end")

            if end < start:
                raise ValueError("Vacation end date is before vacation start date.")

            self.result = {"start": start, "end": end}
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Vacation date error", str(exc), parent=self)


# ---------------------------------------------------------------------------
# Sorting helpers copied from your CLI logic, but used by the GUI.
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

class BidGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title("UPS Bid Analyzer")
        self.geometry("1080x780")
        self.minsize(960, 640)

        self.config_data = load_saved_config()
        self._setup_style()
        self.message_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self.preview_df: pd.DataFrame | None = None
        self.cached_lines: dict[str, Any] | None = None
        self.cached_trips: dict[str, Any] | None = None
        self.cached_pdf_key: tuple[str, str] | None = None

        self.sort_order: list[list[str]] = []

        self._build_ui()
        self._load_saved_values()
        self.after(100, self._poll_queue)

    # -------------------------- UI construction --------------------------

    def _setup_style(self) -> None:
        """Apply a UPS-inspired color theme.

        The clam theme is used because it respects custom button and frame
        colors more consistently than the default platform themes.
        """
        self.configure(bg=UPS_BROWN)
        self.style = ttk.Style(self)

        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        default_font = ("Segoe UI", 10)
        heading_font = ("Segoe UI", 11, "bold")
        title_font = ("Segoe UI", 20, "bold")

        self.option_add("*Font", default_font)

        self.style.configure(".", background=UPS_BROWN, foreground=UPS_TEXT, font=default_font)
        self.style.configure("TFrame", background=UPS_BROWN)
        self.style.configure("UPS.TFrame", background=UPS_BROWN)
        self.style.configure("Panel.TFrame", background=UPS_PANEL)

        self.style.configure("TLabel", background=UPS_BROWN, foreground=UPS_TEXT)
        self.style.configure("Title.TLabel", background=UPS_BROWN, foreground=UPS_GOLD, font=title_font)
        self.style.configure("Subtitle.TLabel", background=UPS_BROWN, foreground=UPS_TEXT, font=("Segoe UI", 10))

        self.style.configure(
            "TLabelframe",
            background=UPS_BROWN,
            foreground=UPS_TEXT,
            bordercolor=UPS_GOLD,
            lightcolor=UPS_GOLD,
            darkcolor=UPS_GOLD,
            relief="solid",
            borderwidth=2,
        )
        self.style.configure(
            "TLabelframe.Label",
            background=UPS_BROWN,
            foreground=UPS_GOLD,
            font=heading_font,
        )

        self.style.configure(
            "TButton",
            background=UPS_BLUE,
            foreground="white",
            bordercolor=UPS_BLUE,
            lightcolor=UPS_BLUE,
            darkcolor=UPS_BLUE_ACTIVE,
            padding=(10, 6),
        )
        self.style.map(
            "TButton",
            background=[("active", UPS_BLUE_ACTIVE), ("disabled", "#7A7A7A")],
            foreground=[("disabled", "#DDDDDD")],
        )

        self.style.configure(
            "Green.TButton",
            background=UPS_GREEN,
            foreground="white",
            bordercolor=UPS_GREEN,
            lightcolor=UPS_GREEN,
            darkcolor=UPS_GREEN_ACTIVE,
            padding=(12, 7),
            font=("Segoe UI", 10, "bold"),
        )
        self.style.map(
            "Green.TButton",
            background=[("active", UPS_GREEN_ACTIVE), ("disabled", "#7A7A7A")],
            foreground=[("disabled", "#DDDDDD")],
        )

        self.style.configure("TEntry", fieldbackground="white", foreground="black", insertcolor="black")
        self.style.configure("TCombobox", fieldbackground="white", foreground="black", arrowcolor=UPS_BROWN)
        self.style.configure("TCheckbutton", background=UPS_BROWN, foreground=UPS_TEXT)
        self.style.map("TCheckbutton", background=[("active", UPS_BROWN)], foreground=[("active", UPS_TEXT)])

        self.style.configure(
            "Treeview",
            background="white",
            fieldbackground="white",
            foreground="black",
            rowheight=24,
            bordercolor=UPS_GOLD,
        )
        self.style.configure(
            "Treeview.Heading",
            background=UPS_GOLD,
            foreground=UPS_BROWN,
            font=("Segoe UI", 10, "bold"),
        )
        self.style.map("Treeview", background=[("selected", UPS_BLUE)], foreground=[("selected", "white")])

        self.style.configure("TProgressbar", background=UPS_GOLD, troughcolor=UPS_BROWN_2)
        self.style.configure("Vertical.TScrollbar", background=UPS_GOLD, troughcolor=UPS_BROWN_2, arrowcolor=UPS_BROWN)

    def _bind_mousewheel(self, widget: tk.Widget) -> None:
        widget.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        widget.bind_all("<Button-4>", self._on_mousewheel, add="+")
        widget.bind_all("<Button-5>", self._on_mousewheel, add="+")

    def _on_mousewheel(self, event: tk.Event) -> None:
        if not hasattr(self, "main_canvas"):
            return

        if getattr(event, "num", None) == 4:
            self.main_canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            self.main_canvas.yview_scroll(1, "units")
        else:
            delta = int(-1 * (event.delta / 120))
            self.main_canvas.yview_scroll(delta, "units")

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, style="UPS.TFrame")
        outer.pack(fill="both", expand=True)
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        self.main_canvas = tk.Canvas(outer, bg=UPS_BROWN, highlightthickness=0)
        self.main_canvas.grid(row=0, column=0, sticky="nsew")

        main_scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.main_canvas.yview)
        main_scrollbar.grid(row=0, column=1, sticky="ns")
        self.main_canvas.configure(yscrollcommand=main_scrollbar.set)

        self.scrollable_frame = ttk.Frame(self.main_canvas, style="UPS.TFrame")
        self.scroll_window = self.main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda _event: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")),
        )
        self.main_canvas.bind(
            "<Configure>",
            lambda event: self.main_canvas.itemconfigure(self.scroll_window, width=event.width),
        )
        self._bind_mousewheel(self.main_canvas)

        container = ttk.Frame(self.scrollable_frame, padding=14, style="UPS.TFrame")
        container.pack(fill="both", expand=True)

        container.columnconfigure(0, weight=1)

        header_frame = ttk.Frame(container, style="UPS.TFrame")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header_frame.columnconfigure(0, weight=1)

        ttk.Label(header_frame, text="UPS Bid Analyzer", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header_frame,
            text="Load bid-package PDFs, apply your scoring preferences, and export the ranked Excel file.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w")

        file_frame = ttk.LabelFrame(container, text="PDF Files", padding=10)
        file_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="TRIPS PDF:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.trips_path_var = tk.StringVar()
        self.trips_path_var.trace_add("write", lambda *_: self._mark_pdf_paths_changed())
        ttk.Entry(file_frame, textvariable=self.trips_path_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(file_frame, text="Browse", command=self._browse_trips).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(file_frame, text="LINES PDF:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.lines_path_var = tk.StringVar()
        self.lines_path_var.trace_add("write", lambda *_: self._mark_pdf_paths_changed())
        ttk.Entry(file_frame, textvariable=self.lines_path_var).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(file_frame, text="Browse", command=self._browse_lines).grid(row=1, column=2, padx=(8, 0), pady=4)

        self.load_button = ttk.Button(
            file_frame,
            text="Load PDFs into UPS Bid Analyzer",
            command=self.load_pdfs,
        )
        self.load_button.grid(row=0, column=3, rowspan=2, sticky="ns", padx=(14, 0), pady=4)

        self.pdf_status_var = tk.StringVar(value="PDFs not loaded yet.")
        ttk.Label(file_frame, textvariable=self.pdf_status_var).grid(
            row=2,
            column=0,
            columnspan=4,
            sticky="w",
            pady=(8, 0),
        )

        prefs_frame = ttk.LabelFrame(container, text="Preferences", padding=10)
        prefs_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        prefs_frame.columnconfigure(1, weight=1)
        prefs_frame.columnconfigure(3, weight=1)

        ttk.Label(prefs_frame, text="Vacation ranges:").grid(row=0, column=0, sticky="nw", padx=(0, 8), pady=4)

        vacation_area = ttk.Frame(prefs_frame)
        vacation_area.grid(row=0, column=1, columnspan=3, sticky="nsew", pady=4)
        vacation_area.columnconfigure(0, weight=1)

        self.vacation_tree = ttk.Treeview(
            vacation_area,
            columns=("start", "end"),
            show="headings",
            height=4,
            selectmode="browse",
        )
        self.vacation_tree.heading("start", text="Start")
        self.vacation_tree.heading("end", text="End")
        self.vacation_tree.column("start", width=140, anchor="center")
        self.vacation_tree.column("end", width=140, anchor="center")
        self.vacation_tree.grid(row=0, column=0, sticky="nsew")

        vacation_scrollbar = ttk.Scrollbar(vacation_area, orient="vertical", command=self.vacation_tree.yview)
        vacation_scrollbar.grid(row=0, column=1, sticky="ns")
        self.vacation_tree.configure(yscrollcommand=vacation_scrollbar.set)

        vacation_buttons = ttk.Frame(vacation_area)
        vacation_buttons.grid(row=0, column=2, sticky="ns", padx=(10, 0))
        ttk.Button(vacation_buttons, text="Add vacation range", command=self._add_vacation_range).pack(fill="x", pady=(0, 4))
        ttk.Button(vacation_buttons, text="Edit selected", command=self._edit_vacation_range).pack(fill="x", pady=4)
        ttk.Button(vacation_buttons, text="Remove selected", command=self._remove_vacation_range).pack(fill="x", pady=4)
        ttk.Button(vacation_buttons, text="Clear all", command=self._clear_vacation_ranges).pack(fill="x", pady=4)

        ttk.Label(prefs_frame, text="Training start:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=8)
        self.training_start_var = tk.StringVar()
        self.training_start_entry = DateEntry(prefs_frame, self.training_start_var)
        self.training_start_entry.grid(row=1, column=1, sticky="w", pady=8)

        ttk.Label(prefs_frame, text="Training end:").grid(row=1, column=2, sticky="w", padx=(16, 8), pady=8)
        self.training_end_var = tk.StringVar()
        self.training_end_entry = DateEntry(prefs_frame, self.training_end_var)
        self.training_end_entry.grid(row=1, column=3, sticky="w", pady=8)

        ttk.Label(prefs_frame, text="Bid edge days off:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self.bid_edge_var = tk.StringVar(value="none")
        self.bid_edge_combo = ttk.Combobox(
            prefs_frame,
            textvariable=self.bid_edge_var,
            values=["none", "start", "end", "both"],
            state="readonly",
            width=16,
        )
        self.bid_edge_combo.grid(row=2, column=1, sticky="w", pady=4)
        self.bid_edge_combo.bind("<<ComboboxSelected>>", self._on_bid_edge_changed)

        sort_frame = ttk.LabelFrame(container, text="Sorting", padding=10)
        sort_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        sort_frame.columnconfigure(0, weight=1)
        sort_frame.columnconfigure(2, weight=1)

        ttk.Label(sort_frame, text="Available columns").grid(row=0, column=0, sticky="w")
        ttk.Label(sort_frame, text="Selected sorting priority").grid(row=0, column=2, sticky="w")

        self.available_columns_list = tk.Listbox(
            sort_frame,
            height=8,
            exportselection=False,
            bg="white",
            fg="black",
            selectbackground=UPS_BLUE,
            selectforeground="white",
            highlightbackground=UPS_GOLD,
            highlightcolor=UPS_GOLD,
        )
        self.available_columns_list.grid(row=1, column=0, sticky="nsew", pady=4)

        sort_buttons = ttk.Frame(sort_frame)
        sort_buttons.grid(row=1, column=1, padx=10, sticky="ns")
        ttk.Button(sort_buttons, text="Add high →", command=lambda: self._add_sort_column("desc")).pack(fill="x", pady=3)
        ttk.Button(sort_buttons, text="Add low →", command=lambda: self._add_sort_column("asc")).pack(fill="x", pady=3)
        ttk.Separator(sort_buttons).pack(fill="x", pady=8)
        ttk.Button(sort_buttons, text="Move up", command=self._move_sort_up).pack(fill="x", pady=3)
        ttk.Button(sort_buttons, text="Move down", command=self._move_sort_down).pack(fill="x", pady=3)
        ttk.Button(sort_buttons, text="Remove", command=self._remove_sort_column).pack(fill="x", pady=3)
        ttk.Button(sort_buttons, text="Clear", command=self._clear_sort_order).pack(fill="x", pady=3)

        self.selected_sort_list = tk.Listbox(
            sort_frame,
            height=8,
            exportselection=False,
            bg="white",
            fg="black",
            selectbackground=UPS_BLUE,
            selectforeground="white",
            highlightbackground=UPS_GOLD,
            highlightcolor=UPS_GOLD,
        )
        self.selected_sort_list.grid(row=1, column=2, sticky="nsew", pady=4)

        # Advanced sorting settings are intentionally not shown in the main
        # workflow. They are stored here and edited through a popup to reduce
        # accidental changes by novice users.
        self.default_mode_var = tk.StringVar(value=DEFAULT_SORTING_SETTINGS["default_mode"])
        self.weighting_style_var = tk.StringVar(value=DEFAULT_SORTING_SETTINGS["weighting_style"])
        self.soft_max_weight_var = tk.StringVar(value=str(DEFAULT_SORTING_SETTINGS["soft_max_weight"]))
        self.soft_min_weight_var = tk.StringVar(value=str(DEFAULT_SORTING_SETTINGS["soft_min_weight"]))
        self.keep_score_columns_var = tk.BooleanVar(value=DEFAULT_SORTING_SETTINGS["keep_score_columns"])

        ttk.Separator(sort_buttons).pack(fill="x", pady=8)
        ttk.Button(
            sort_buttons,
            text="Advanced settings...",
            command=self._open_sorting_settings_dialog,
        ).pack(fill="x", pady=3)

        output_frame = ttk.LabelFrame(container, text="Output", padding=10)
        output_frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)

        ttk.Label(output_frame, text="Output folder:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.output_folder_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_folder_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(output_frame, text="Browse", command=self._browse_output_folder).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(output_frame, text="File name:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.output_filename_var = tk.StringVar(value="Bid_Results")
        filename_row = ttk.Frame(output_frame)
        filename_row.grid(row=1, column=1, sticky="w", pady=4)
        ttk.Entry(filename_row, textvariable=self.output_filename_var, width=36).pack(side="left")
        ttk.Label(filename_row, text=".xlsx").pack(side="left", padx=(4, 0))

        action_frame = ttk.Frame(container)
        action_frame.grid(row=5, column=0, sticky="ew", pady=(0, 10))
        action_frame.columnconfigure(1, weight=1)

        self.export_button = ttk.Button(
            action_frame,
            text="Export to Excel",
            command=self.export_excel,
            style="Green.TButton",
        )
        self.export_button.grid(row=0, column=0, padx=(0, 8))

        self.progress = ttk.Progressbar(action_frame, mode="indeterminate")
        self.progress.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        log_frame = ttk.LabelFrame(container, text="Status", padding=10)
        log_frame.grid(row=6, column=0, sticky="nsew")
        container.rowconfigure(6, weight=1)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame,
            height=10,
            wrap="word",
            state="disabled",
            bg=UPS_FIELD_BG,
            fg="black",
            insertbackground="black",
            highlightbackground=UPS_GOLD,
            highlightcolor=UPS_GOLD,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _load_saved_values(self) -> None:
        self._set_vacation_ranges(self.config_data.get("vacation_ranges", []))

        self.training_start_var.set(self.config_data.get("training_start") or "")
        self.training_end_var.set(self.config_data.get("training_end") or "")
        self.bid_edge_var.set(self.config_data.get("bid_edge") or "none")

        output_paths = self.config_data.get("output_paths", {})
        saved_output_folder = output_paths.get(get_os_name(), "")
        self.output_folder_var.set(saved_output_folder or str(Path.cwd()))

        saved_sort_order = self.config_data.get("sort_order", [])
        if isinstance(saved_sort_order, list):
            self.sort_order = [list(item) for item in saved_sort_order if isinstance(item, list) and len(item) == 2]
            self._refresh_selected_sort_list()

        saved_sorting_settings = self.config_data.get("sorting_settings", {})
        if not isinstance(saved_sorting_settings, dict):
            saved_sorting_settings = {}
        merged_settings = {**DEFAULT_SORTING_SETTINGS, **saved_sorting_settings}

        self.default_mode_var.set(str(merged_settings["default_mode"]))
        self.weighting_style_var.set(str(merged_settings["weighting_style"]))
        self.soft_max_weight_var.set(str(merged_settings["soft_max_weight"]))
        self.soft_min_weight_var.set(str(merged_settings["soft_min_weight"]))
        self.keep_score_columns_var.set(bool(merged_settings["keep_score_columns"]))

    # -------------------------- Browse buttons --------------------------

    def _browse_trips(self) -> None:
        path = filedialog.askopenfilename(title="Select TRIPS PDF", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if path:
            self.trips_path_var.set(path)

    def _browse_lines(self) -> None:
        path = filedialog.askopenfilename(title="Select LINES PDF", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if path:
            self.lines_path_var.set(path)

    def _browse_output_folder(self) -> None:
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_folder_var.set(path)

    def _mark_pdf_paths_changed(self) -> None:
        current_key = (self.trips_path_var.get().strip(), self.lines_path_var.get().strip())
        if self.cached_pdf_key and current_key != self.cached_pdf_key:
            self.pdf_status_var.set("PDF paths changed. Click Load PDFs again.")
            self.preview_df = None
            self._refresh_available_columns_list([])

    # -------------------------- Vacation range actions --------------------------

    def _set_vacation_ranges(self, vacation_ranges: list[dict[str, str]]) -> None:
        for item_id in self.vacation_tree.get_children():
            self.vacation_tree.delete(item_id)

        for vacation in vacation_ranges or []:
            start = vacation.get("start", "")
            end = vacation.get("end", "")
            if start and end:
                self.vacation_tree.insert("", tk.END, values=(start, end))

    def _get_vacation_ranges(self) -> list[dict[str, str]]:
        ranges: list[dict[str, str]] = []
        for item_id in self.vacation_tree.get_children():
            start, end = self.vacation_tree.item(item_id, "values")
            start = validate_required_date(str(start), "Vacation start")
            end = validate_required_date(str(end), "Vacation end")
            if end < start:
                raise ValueError(f"Vacation range {start} to {end}: end date is before start date.")
            ranges.append({"start": start, "end": end})
        return ranges

    def _add_vacation_range(self) -> None:
        dialog = VacationRangeDialog(self, "Add vacation range")
        self.wait_window(dialog)
        if dialog.result:
            self.vacation_tree.insert("", tk.END, values=(dialog.result["start"], dialog.result["end"]))

    def _edit_vacation_range(self) -> None:
        selection = self.vacation_tree.selection()
        if not selection:
            messagebox.showinfo("Edit vacation range", "Select a vacation range first.")
            return

        item_id = selection[0]
        start, end = self.vacation_tree.item(item_id, "values")
        dialog = VacationRangeDialog(self, "Edit vacation range", str(start), str(end))
        self.wait_window(dialog)
        if dialog.result:
            self.vacation_tree.item(item_id, values=(dialog.result["start"], dialog.result["end"]))

    def _remove_vacation_range(self) -> None:
        selection = self.vacation_tree.selection()
        if not selection:
            return
        self.vacation_tree.delete(selection[0])

    def _clear_vacation_ranges(self) -> None:
        for item_id in self.vacation_tree.get_children():
            self.vacation_tree.delete(item_id)

    # -------------------------- Sorting list actions --------------------------

    def _add_sort_column(self, direction: str) -> None:
        selection = self.available_columns_list.curselection()
        if not selection:
            return

        col = self.available_columns_list.get(selection[0])
        self.sort_order = [rule for rule in self.sort_order if rule[0] != col]
        self.sort_order.append([col, direction])
        self._refresh_selected_sort_list()

    def _remove_sort_column(self) -> None:
        selection = self.selected_sort_list.curselection()
        if not selection:
            return

        index = selection[0]
        del self.sort_order[index]
        self._refresh_selected_sort_list()

    def _clear_sort_order(self) -> None:
        self.sort_order = []
        self._refresh_selected_sort_list()

    def _move_sort_up(self) -> None:
        selection = self.selected_sort_list.curselection()
        if not selection:
            return
        index = selection[0]
        if index == 0:
            return
        self.sort_order[index - 1], self.sort_order[index] = self.sort_order[index], self.sort_order[index - 1]
        self._refresh_selected_sort_list(select_index=index - 1)

    def _move_sort_down(self) -> None:
        selection = self.selected_sort_list.curselection()
        if not selection:
            return
        index = selection[0]
        if index >= len(self.sort_order) - 1:
            return
        self.sort_order[index + 1], self.sort_order[index] = self.sort_order[index], self.sort_order[index + 1]
        self._refresh_selected_sort_list(select_index=index + 1)

    def _refresh_available_columns_list(self, columns: list[str]) -> None:
        self.available_columns_list.delete(0, tk.END)
        for col in columns:
            self.available_columns_list.insert(tk.END, col)

    def _refresh_selected_sort_list(self, select_index: int | None = None) -> None:
        self.selected_sort_list.delete(0, tk.END)
        for col, direction in self.sort_order:
            label = "high-to-low" if direction == "desc" else "low-to-high"
            self.selected_sort_list.insert(tk.END, f"{col} ({label})")

        if select_index is not None and 0 <= select_index < len(self.sort_order):
            self.selected_sort_list.selection_set(select_index)
            self.selected_sort_list.activate(select_index)

    def _clean_sort_order_for_columns(self, columns: list[str]) -> None:
        valid_columns = set(columns)
        old_sort_order = list(self.sort_order)
        self.sort_order = [rule for rule in self.sort_order if len(rule) == 2 and rule[0] in valid_columns]
        self._refresh_selected_sort_list()

        skipped = [rule[0] for rule in old_sort_order if len(rule) == 2 and rule[0] not in valid_columns]
        if skipped:
            self._write_log("Skipped saved sorting columns not found in this DataFrame: " + ", ".join(skipped))

    # -------------------------- Validation / config --------------------------

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
            self.default_mode_var.get(),
            self.weighting_style_var.get(),
            self.soft_max_weight_var.get(),
            self.soft_min_weight_var.get(),
            self.keep_score_columns_var.get(),
        )

    def _apply_sorting_settings_to_ui(self, sorting_settings: dict[str, Any]) -> None:
        self.default_mode_var.set(str(sorting_settings["default_mode"]))
        self.weighting_style_var.set(str(sorting_settings["weighting_style"]))
        self.soft_max_weight_var.set(str(sorting_settings["soft_max_weight"]))
        self.soft_min_weight_var.set(str(sorting_settings["soft_min_weight"]))
        self.keep_score_columns_var.set(bool(sorting_settings["keep_score_columns"]))

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
        if show_message:
            messagebox.showinfo("Saved", "Advanced sorting settings saved.")

    def _save_sorting_settings_from_ui(self) -> None:
        try:
            sorting_settings = self._get_sorting_settings()
        except Exception as exc:
            messagebox.showerror("Sorting settings error", str(exc))
            return

        self._save_sorting_settings(sorting_settings)

    def _open_sorting_settings_dialog(self) -> None:
        """Open a protected popup for less-common sorting parameters."""
        current = self._get_sorting_settings()

        dialog = tk.Toplevel(self)
        dialog.title("Advanced Sorting Settings")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        main = ttk.Frame(dialog, padding=14)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)

        default_mode_var = tk.StringVar(value=current["default_mode"])
        weighting_style_var = tk.StringVar(value=current["weighting_style"])
        soft_max_weight_var = tk.StringVar(value=str(current["soft_max_weight"]))
        soft_min_weight_var = tk.StringVar(value=str(current["soft_min_weight"]))
        keep_score_columns_var = tk.BooleanVar(value=current["keep_score_columns"])

        intro = ttk.Label(
            main,
            text=(
                "These settings affect how selected sorting columns are combined. "
                "Most users should leave them at the saved defaults."
            ),
            wraplength=520,
        )
        intro.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))

        default_label = ttk.Label(main, text="Default mode:")
        default_label.grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        default_mode_combo = ttk.Combobox(
            main,
            textvariable=default_mode_var,
            values=list(DEFAULT_MODE_DESCRIPTIONS),
            state="readonly",
            width=18,
        )
        default_mode_combo.grid(row=1, column=1, sticky="w", pady=6)
        default_help = ttk.Label(main, text="?", foreground=UPS_GOLD)
        default_help.grid(row=1, column=2, sticky="w", padx=(8, 0), pady=6)
        ToolTip(default_label, "strict: " + DEFAULT_MODE_DESCRIPTIONS["strict"] + "\n\nweighted: " + DEFAULT_MODE_DESCRIPTIONS["weighted"])
        ToolTip(default_mode_combo, "strict: " + DEFAULT_MODE_DESCRIPTIONS["strict"] + "\n\nweighted: " + DEFAULT_MODE_DESCRIPTIONS["weighted"])
        ToolTip(default_help, "strict: " + DEFAULT_MODE_DESCRIPTIONS["strict"] + "\n\nweighted: " + DEFAULT_MODE_DESCRIPTIONS["weighted"])

        weighting_label = ttk.Label(main, text="Weighting style:")
        weighting_label.grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
        weighting_style_combo = ttk.Combobox(
            main,
            textvariable=weighting_style_var,
            values=list(WEIGHTING_STYLE_DESCRIPTIONS),
            state="readonly",
            width=18,
        )
        weighting_style_combo.grid(row=2, column=1, sticky="w", pady=6)
        weighting_help = ttk.Label(main, text="?", foreground=UPS_GOLD)
        weighting_help.grid(row=2, column=2, sticky="w", padx=(8, 0), pady=6)
        weighting_tip = (
            "equal: " + WEIGHTING_STYLE_DESCRIPTIONS["equal"] + "\n\n"
            "hard: " + WEIGHTING_STYLE_DESCRIPTIONS["hard"] + "\n\n"
            "soft: " + WEIGHTING_STYLE_DESCRIPTIONS["soft"]
        )
        ToolTip(weighting_label, weighting_tip)
        ToolTip(weighting_style_combo, weighting_tip)
        ToolTip(weighting_help, weighting_tip)

        soft_max_label = ttk.Label(main, text="Soft max weight:")
        soft_max_label.grid(row=3, column=0, sticky="w", padx=(0, 8), pady=6)
        soft_max_entry = ttk.Entry(main, textvariable=soft_max_weight_var, width=12)
        soft_max_entry.grid(row=3, column=1, sticky="w", pady=6)
        soft_max_help = ttk.Label(main, text="?", foreground=UPS_GOLD)
        soft_max_help.grid(row=3, column=2, sticky="w", padx=(8, 0), pady=6)
        ToolTip(soft_max_label, "Used by the soft weighting style. This is the weight given to the first item in each weighted group. Default: 3.0.")
        ToolTip(soft_max_entry, "Used by the soft weighting style. This is the weight given to the first item in each weighted group. Default: 3.0.")
        ToolTip(soft_max_help, "Used by the soft weighting style. This is the weight given to the first item in each weighted group. Default: 3.0.")

        soft_min_label = ttk.Label(main, text="Soft min weight:")
        soft_min_label.grid(row=4, column=0, sticky="w", padx=(0, 8), pady=6)
        soft_min_entry = ttk.Entry(main, textvariable=soft_min_weight_var, width=12)
        soft_min_entry.grid(row=4, column=1, sticky="w", pady=6)
        soft_min_help = ttk.Label(main, text="?", foreground=UPS_GOLD)
        soft_min_help.grid(row=4, column=2, sticky="w", padx=(8, 0), pady=6)
        ToolTip(soft_min_label, "Used by the soft weighting style. This is the weight given to the last item in each weighted group. Default: 1.0.")
        ToolTip(soft_min_entry, "Used by the soft weighting style. This is the weight given to the last item in each weighted group. Default: 1.0.")
        ToolTip(soft_min_help, "Used by the soft weighting style. This is the weight given to the last item in each weighted group. Default: 1.0.")

        keep_check = ttk.Checkbutton(
            main,
            text="Keep score columns in Excel",
            variable=keep_score_columns_var,
        )
        keep_check.grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        keep_help = ttk.Label(main, text="?", foreground=UPS_GOLD)
        keep_help.grid(row=5, column=2, sticky="w", padx=(8, 0), pady=(8, 0))
        ToolTip(keep_check, "When enabled, any extra score/helper columns created by weighted sorting remain in the exported Excel file.")
        ToolTip(keep_help, "When enabled, any extra score/helper columns created by weighted sorting remain in the exported Excel file.")

        buttons = ttk.Frame(main)
        buttons.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(16, 0))
        buttons.columnconfigure(0, weight=1)

        def restore_defaults() -> None:
            default_mode_var.set(str(DEFAULT_SORTING_SETTINGS["default_mode"]))
            weighting_style_var.set(str(DEFAULT_SORTING_SETTINGS["weighting_style"]))
            soft_max_weight_var.set(str(DEFAULT_SORTING_SETTINGS["soft_max_weight"]))
            soft_min_weight_var.set(str(DEFAULT_SORTING_SETTINGS["soft_min_weight"]))
            keep_score_columns_var.set(bool(DEFAULT_SORTING_SETTINGS["keep_score_columns"]))

        def save_and_close() -> None:
            try:
                settings = self._validate_sorting_settings_values(
                    default_mode_var.get(),
                    weighting_style_var.get(),
                    soft_max_weight_var.get(),
                    soft_min_weight_var.get(),
                    keep_score_columns_var.get(),
                )
            except Exception as exc:
                messagebox.showerror("Sorting settings error", str(exc), parent=dialog)
                return

            self._save_sorting_settings(settings, show_message=False)
            dialog.destroy()

        ttk.Button(buttons, text="Restore defaults", command=restore_defaults).grid(row=0, column=0, sticky="w")
        ttk.Button(buttons, text="Cancel", command=dialog.destroy).grid(row=0, column=1, sticky="e", padx=(8, 0))
        ttk.Button(buttons, text="Save", command=save_and_close).grid(row=0, column=2, sticky="e", padx=(8, 0))

        dialog.bind("<Return>", lambda _event: save_and_close())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

        dialog.update_idletasks()
        try:
            x = self.winfo_rootx() + 120
            y = self.winfo_rooty() + 120
            dialog.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _collect_inputs(self) -> dict[str, Any]:
        trips_pdf_path = self.trips_path_var.get().strip().strip('"').strip("'")
        lines_pdf_path = self.lines_path_var.get().strip().strip('"').strip("'")

        if not trips_pdf_path:
            raise ValueError("Please choose the TRIPS PDF.")
        if not lines_pdf_path:
            raise ValueError("Please choose the LINES PDF.")
        if not Path(trips_pdf_path).exists():
            raise ValueError("The TRIPS PDF path does not exist.")
        if not Path(lines_pdf_path).exists():
            raise ValueError("The LINES PDF path does not exist.")

        vacation_ranges = self._get_vacation_ranges()
        training_start = validate_date_or_blank(self.training_start_var.get(), "Training start")
        training_end = validate_date_or_blank(self.training_end_var.get(), "Training end")

        if bool(training_start) != bool(training_end):
            raise ValueError("Enter both training start and training end, or leave both blank.")

        if training_start and training_end and training_end < training_start:
            raise ValueError("Training end date is before training start date.")

        bid_edge = self.bid_edge_var.get().strip().lower() or "none"
        if bid_edge not in {"none", "start", "end", "both"}:
            raise ValueError("Bid edge preference must be none, start, end, or both.")

        output_folder = Path(self.output_folder_var.get().strip().strip('"').strip("'")).expanduser()
        if not output_folder.exists():
            output_folder.mkdir(parents=True, exist_ok=True)
        if not output_folder.is_dir():
            raise ValueError("Output folder is not a folder.")

        output_filename = clean_filename(self.output_filename_var.get())
        output_path = output_folder / f"{output_filename}.xlsx"

        return {
            "trips_pdf_path": trips_pdf_path,
            "lines_pdf_path": lines_pdf_path,
            "vacation_ranges": vacation_ranges,
            "training_start": training_start,
            "training_end": training_end,
            "bid_edge": bid_edge,
            "output_folder": output_folder,
            "output_path": output_path,
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
        output_paths = self.config_data.setdefault("output_paths", {})
        output_paths[get_os_name()] = str(inputs["output_folder"])

        save_config(self.config_data)

    # -------------------------- Worker control --------------------------

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.load_button.configure(state=state)
        self.export_button.configure(state=state)

        if busy:
            self.progress.start(10)
        else:
            self.progress.stop()

    def _log(self, message: str) -> None:
        self.message_queue.put(("log", message))

    def _write_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _poll_queue(self) -> None:
        try:
            while True:
                event, payload = self.message_queue.get_nowait()

                if event == "log":
                    self._write_log(str(payload))
                elif event == "error":
                    self._set_busy(False)
                    self.pdf_status_var.set("Error. See status box below.")
                    self._write_log(f"ERROR: {payload}")
                    messagebox.showerror("Error", str(payload))
                elif event == "loaded":
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
                    self.pdf_status_var.set(status_text)
                    self._write_log(f"{log_text} and prepared {len(columns)} sortable columns.")

                    if show_ready_message:
                        messagebox.showinfo("Ready", "PDFs are loaded. Choose your sorting priority, then export.")
                elif event == "exported":
                    self._set_busy(False)
                    output_path = payload
                    self.pdf_status_var.set("Export complete.")
                    self._write_log(f"Finished export: {output_path}")
                    messagebox.showinfo("Finished", f"Excel file created:\n{output_path}")
        except queue.Empty:
            pass

        self.after(100, self._poll_queue)

    def _start_worker(self, target: Callable[..., None], *args: Any) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Busy", "A job is already running.")
            return

        self._set_busy(True)
        self.worker_thread = threading.Thread(target=target, args=args, daemon=True)
        self.worker_thread.start()

    # -------------------------- Processing logic --------------------------

    def _on_bid_edge_changed(self, _event: tk.Event | None = None) -> None:
        """Rebuild the analyzer when Start/End/Both/None changes.

        This refreshes the DataFrame and available sorting columns immediately.
        If the same PDFs are already loaded, it reuses the cached extraction. If
        the PDF paths changed, it will extract the new PDF pair.
        """
        self.config_data["bid_edge"] = self.bid_edge_var.get().strip().lower() or "none"
        save_config(self.config_data)

        if not self.trips_path_var.get().strip() or not self.lines_path_var.get().strip():
            self.pdf_status_var.set("Bid edge preference saved. Choose PDFs when ready.")
            return

        if self.worker_thread and self.worker_thread.is_alive():
            self.pdf_status_var.set("Bid edge changed. Refresh after the current job finishes.")
            return

        try:
            inputs = self._collect_inputs()
            self._save_inputs_to_config(inputs)
        except Exception as exc:
            self.pdf_status_var.set("Bid edge changed. Analyzer refresh skipped.")
            self._write_log(f"Bid edge changed, but analyzer could not refresh yet: {exc}")
            return

        if self.cached_trips is None or self.cached_lines is None:
            self.pdf_status_var.set("Bid edge changed. Click Load PDFs into UPS Bid Analyzer.")
            return

        self.pdf_status_var.set("Bid edge changed. Refreshing analyzer...")
        self._start_worker(self._load_worker, inputs, False)

    def load_pdfs(self) -> None:
        try:
            inputs = self._collect_inputs()
            self._save_inputs_to_config(inputs)
        except Exception as exc:
            messagebox.showerror("Input error", str(exc))
            return

        self.pdf_status_var.set("Loading PDFs...")
        self._start_worker(self._load_worker, inputs, True)

    def export_excel(self) -> None:
        try:
            inputs = self._collect_inputs()
            self._save_inputs_to_config(inputs)
        except Exception as exc:
            messagebox.showerror("Input error", str(exc))
            return

        current_key = (inputs["trips_pdf_path"], inputs["lines_pdf_path"])
        if self.cached_pdf_key != current_key:
            self.pdf_status_var.set("PDFs not loaded for these paths. Loading first, then exporting...")

        self._start_worker(self._export_worker, inputs)

    def _extract_pdfs(self, inputs: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        pdf_key = (inputs["trips_pdf_path"], inputs["lines_pdf_path"])

        if self.cached_pdf_key == pdf_key and self.cached_trips is not None and self.cached_lines is not None:
            self._log("Using already-loaded PDF data.")
            return self.cached_trips, self.cached_lines

        self._log("Extracting PDFs...")

        with ThreadPoolExecutor(max_workers=2) as executor:
            trips_future = executor.submit(extract_trips_from_pdf, inputs["trips_pdf_path"], first_page=2)
            lines_future = executor.submit(parse_line_report_pdf, inputs["lines_pdf_path"], first_calendar_page=3)

            trips = trips_future.result()
            lines = lines_future.result()

        self.cached_pdf_key = pdf_key
        self.cached_trips = trips
        self.cached_lines = lines

        self._log("PDF extraction complete.")
        return trips, lines

    def _build_dataframe(self, inputs: dict[str, Any], *, apply_sort: bool) -> tuple[pd.DataFrame, list[dict[str, str]]]:
        trips, lines = self._extract_pdfs(inputs)

        bid_period_info = {x: lines[x] for x in ("bid_period_date_range", "pay_period_date_ranges")}

        self._log("Creating master lines...")
        master_lines = creating_master_line(trips, lines)

        self._log("Adding scores...")
        pf.add_blockiness_scores(master_lines, bid_period_info)
        pf.add_company_ticket_percentages(master_lines)
        new_vacation_ranges = pf.add_vacation_days_off_score(
            master_lines,
            inputs["vacation_ranges"],
            bid_period_info,
            save_details=False,
        )
        if inputs["training_start"] != None and inputs["training_end"] != None:
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

    def _export_worker(self, inputs: dict[str, Any]) -> None:
        try:
            df, new_vacation_ranges = self._build_dataframe(inputs, apply_sort=True)

            self._log("Exporting Excel file...")
            export_master_lines_to_excel_table(
                df,
                str(inputs["output_path"]),
                training_start=inputs["training_start"],
                training_end=inputs["training_end"],
                vacation_ranges=new_vacation_ranges,
            )

            self.message_queue.put(("exported", inputs["output_path"]))
        except Exception as exc:
            self.message_queue.put(("error", exc))


if __name__ == "__main__":
    app = BidGUI()
    app.mainloop()
