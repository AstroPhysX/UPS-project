# bid_table_viewer.py

import re
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date, datetime, timedelta

import pandas as pd

try:
    from tksheet import Sheet
except ImportError:
    raise SystemExit(
        "Missing dependency: tksheet\n"
        "Install it with:\n\n"
        "    pip install tksheet\n"
    )


class BidSpreadsheetViewer(ttk.Frame):
    """
    Simple Excel-like DataFrame viewer for the UPS bid analyzer.

    Features:
        - Displays a pandas DataFrame in a spreadsheet-style table
        - Horizontal and vertical scrolling
        - Resizable rows/columns
        - Copy/select behavior
        - Search highlighting
        - Calendar columns formatted like: Wed, May 27
        - Non-empty calendar cells highlighted green
        - Vacation headers highlighted purple
        - Training headers highlighted orange
        - Narrow colored marker columns for vacation/training start/end dates

    Important:
        This is a viewer. Pandas remains the source of truth.
    """

    WORK_FILL = "#C6EFCE"

    HEADER_FILL = "#4F4F4F"
    HEADER_TEXT = "#FFFFFF"

    VACATION_HEADER_FILL = "#800080"
    TRAINING_HEADER_FILL = "#FFA500"

    SEARCH_FILL = "#FFF2CC"

    def __init__(
        self,
        parent,
        *,
        calendar_col_width=125,
        calendar_row_height=45,
        header_row_height=36,
        non_calendar_default_width=110,
        non_calendar_max_width=220,
        line_number_col_width=90,
        marker_col_width=6,
        show_boundary_markers=True,
    ):
        super().__init__(parent)

        self.original_df = None
        self.source_df = None
        self.display_df = None

        self.display_headers = []
        self.display_to_source_col = {}
        self.display_date_by_col_idx = {}
        self.marker_cols = {}

        self.calendar_cols = None
        self.training_start = None
        self.training_end = None
        self.vacation_ranges = []

        self.calendar_col_width = calendar_col_width
        self.calendar_row_height = calendar_row_height
        self.header_row_height = header_row_height
        self.non_calendar_default_width = non_calendar_default_width
        self.non_calendar_max_width = non_calendar_max_width
        self.line_number_col_width = line_number_col_width
        self.marker_col_width = marker_col_width
        self.show_boundary_markers = show_boundary_markers

        self.search_var = tk.StringVar()
        self.status_var = tk.StringVar(value="No data loaded.")

        self._build_gui()

    # ------------------------------------------------------------
    # GUI layout
    # ------------------------------------------------------------

    def _build_gui(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", padx=6, pady=(6, 3))

        ttk.Label(toolbar, text="Search:").pack(side="left")

        search_entry = ttk.Entry(
            toolbar,
            textvariable=self.search_var,
            width=24,
        )
        search_entry.pack(side="left", padx=(4, 4))
        search_entry.bind("<Return>", lambda event: self.apply_search())

        ttk.Button(
            toolbar,
            text="Find",
            command=self.apply_search,
        ).pack(side="left", padx=(0, 4))

        ttk.Button(
            toolbar,
            text="Clear",
            command=self.clear_search,
        ).pack(side="left", padx=(0, 12))

        self.sheet = Sheet(
            self,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
            show_header=True,
            show_row_index=True,
            width=1200,
            height=650,
        )
        self.sheet.pack(fill="both", expand=True, padx=6, pady=3)

        self._enable_sheet_bindings()

        status_bar = ttk.Label(
            self,
            textvariable=self.status_var,
            anchor="w",
        )
        status_bar.pack(fill="x", padx=6, pady=(3, 6))

    def _enable_sheet_bindings(self):
        """
        Enable Excel-like interactions but not editing.
        Sorting is handled by the main GUI for now.
        """

        self.sheet.enable_bindings(
            "single_select",
            "drag_select",
            "select_all",
            "column_select",
            "row_select",
            "arrowkeys",
            "right_click_popup_menu",
            "rc_select",
            "copy",
            "find",
            "column_width_resize",
            "row_height_resize",
            "double_click_column_resize",
        )

    # ------------------------------------------------------------
    # Public loading method
    # ------------------------------------------------------------

    def load_dataframe(
        self,
        df,
        *,
        calendar_cols=None,
        training_start=None,
        training_end=None,
        vacation_ranges=None,
        reset_original=True,
        calendar_col_width=None,
        calendar_row_height=None,
        header_row_height=None,
        non_calendar_max_width=None,
    ):
        """
        Load/reload a DataFrame into the spreadsheet viewer.

        Optional sizing overrides work similarly to your Excel export function.

        Note:
            tksheet column widths are pixels, not Excel character-width units.
        """

        if df is None or df.empty:
            messagebox.showwarning("No Data", "There is no DataFrame to display.")
            return

        if calendar_col_width is not None:
            self.calendar_col_width = calendar_col_width

        if calendar_row_height is not None:
            self.calendar_row_height = calendar_row_height

        if header_row_height is not None:
            self.header_row_height = header_row_height

        if non_calendar_max_width is not None:
            self.non_calendar_max_width = non_calendar_max_width

        df = df.copy()

        if reset_original:
            self.original_df = df.copy()

        self.source_df = df.copy()
        self.calendar_cols = calendar_cols
        self.training_start = self._normalize_date(training_start)
        self.training_end = self._normalize_date(training_end)
        self.vacation_ranges = self._normalize_date_ranges(vacation_ranges)

        self.display_df = self._make_display_dataframe(
            self.source_df,
            calendar_cols=self.calendar_cols,
        )

        sheet_data = [
            [self._display_value(value) for value in row]
            for row in self.display_df.itertuples(index=False, name=None)
        ]

        self.sheet.set_sheet_data(
            sheet_data,
            reset_col_positions=True,
            reset_row_positions=True,
            redraw=False,
            reset_highlights=True,
            keep_formatting=False,
            delete_options=True,
        )

        self.sheet.headers(self.display_headers, redraw=False)

        self.sheet.row_index(
            [str(i + 1) for i in range(len(self.display_df))],
            redraw=False,
        )

        self._size_columns_and_rows()
        self._apply_base_formatting()

        self.sheet.refresh()

        self.status_var.set(
            f"Loaded {len(self.source_df):,} rows and {len(self.source_df.columns):,} source columns."
        )

    # ------------------------------------------------------------
    # Display preparation
    # ------------------------------------------------------------

    def _make_display_dataframe(self, df, calendar_cols=None):
        """
        Creates a display copy of df with calendar headers renamed.

        Also optionally inserts very narrow marker columns:
            - before training/vacation start dates
            - after training/vacation end dates

        These marker columns visually approximate Excel-style vertical borders.
        """

        self.display_headers = []
        self.display_to_source_col = {}
        self.display_date_by_col_idx = {}
        self.marker_cols = {}

        calendar_date_set = set()

        if calendar_cols is not None:
            for col in calendar_cols:
                d = self._normalize_date(col)
                if d is not None:
                    calendar_date_set.add(d)

        before_markers = {}
        after_markers = {}

        # Vacation markers
        for vacation_start, vacation_end in self.vacation_ranges:
            before_markers[vacation_start] = {
                "kind": "start",
                "color": self.VACATION_HEADER_FILL,
                "name": "Vacation start",
            }
            after_markers[vacation_end] = {
                "kind": "end",
                "color": self.VACATION_HEADER_FILL,
                "name": "Vacation end",
            }

        # Training markers overwrite vacation if they overlap
        if self.training_start is not None:
            before_markers[self.training_start] = {
                "kind": "start",
                "color": self.TRAINING_HEADER_FILL,
                "name": "Training start",
            }

        if self.training_end is not None:
            after_markers[self.training_end] = {
                "kind": "end",
                "color": self.TRAINING_HEADER_FILL,
                "name": "Training end",
            }

        used_headers = set()
        display_data = {}
        marker_counter = 1

        def make_unique_header(header):
            base = str(header).strip() if header is not None else "Column"

            if not base:
                base = "Column"

            candidate = base
            counter = 2

            while candidate in used_headers:
                candidate = f"{base}_{counter}"
                counter += 1

            used_headers.add(candidate)
            return candidate

        def add_marker(marker_info):
            nonlocal marker_counter

            internal_header = f"__marker_{marker_counter}"
            marker_counter += 1

            display_col_idx = len(display_data)

            display_data[internal_header] = [""] * len(df)
            self.display_headers.append("")

            self.marker_cols[display_col_idx] = marker_info

        def add_real_column(source_col, display_header, real_date=None):
            internal_header = make_unique_header(display_header)
            display_col_idx = len(display_data)

            display_data[internal_header] = df[source_col].tolist()
            self.display_headers.append(display_header)

            self.display_to_source_col[internal_header] = source_col

            if real_date is not None:
                self.display_date_by_col_idx[display_col_idx] = real_date

        for source_col in df.columns:
            col_date = self._normalize_date(source_col)

            if calendar_cols is None:
                is_calendar_col = col_date is not None
            else:
                is_calendar_col = col_date in calendar_date_set

            if is_calendar_col and col_date is not None:
                if self.show_boundary_markers and col_date in before_markers:
                    add_marker(before_markers[col_date])

                display_header = self._format_calendar_header(col_date)
                add_real_column(source_col, display_header, real_date=col_date)

                if self.show_boundary_markers and col_date in after_markers:
                    add_marker(after_markers[col_date])

            else:
                add_real_column(source_col, str(source_col), real_date=None)

        return pd.DataFrame(display_data)

    @staticmethod
    def _display_value(value):
        """
        Convert values to user-friendly display values.
        """

        if value is None:
            return ""

        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass

        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d %H:%M")

        if isinstance(value, date):
            return value.strftime("%Y-%m-%d")

        return str(value)

    # ------------------------------------------------------------
    # Formatting
    # ------------------------------------------------------------

    def _size_columns_and_rows(self):
        """
        Rough equivalent of the Excel column width / row height formatting.
        """

        self.sheet.column_width(
            "all",
            width=self.non_calendar_default_width,
            redraw=False,
        )

        self.sheet.row_height(
            "all",
            height=self.calendar_row_height,
            redraw=False,
        )

        self._set_header_height_safely()

        internal_headers = list(self.display_df.columns)

        for col_idx, internal_header in enumerate(internal_headers):
            if col_idx in self.marker_cols:
                self.sheet.column_width(
                    col_idx,
                    width=self.marker_col_width,
                    redraw=False,
                )
                continue

            source_col = self.display_to_source_col.get(internal_header, internal_header)

            if str(source_col).lower() in {"line number", "line"}:
                self.sheet.column_width(
                    col_idx,
                    width=self.line_number_col_width,
                    redraw=False,
                )

            elif col_idx in self.display_date_by_col_idx:
                self.sheet.column_width(
                    col_idx,
                    width=self.calendar_col_width,
                    redraw=False,
                )

            else:
                text_len = max(
                    len(str(source_col)),
                    min(
                        28,
                        max(
                            len(str(v))
                            for v in self.display_df[internal_header].head(50).fillna("")
                        ),
                    ),
                )

                width = max(
                    self.non_calendar_default_width,
                    min(self.non_calendar_max_width, text_len * 8),
                )

                self.sheet.column_width(
                    col_idx,
                    width=width,
                    redraw=False,
                )

    def _set_header_height_safely(self):
        """
        tksheet header-height support varies by version.
        These attempts are harmless if unsupported.
        """

        try:
            self.sheet.header_height(self.header_row_height)
            return
        except Exception:
            pass

        try:
            self.sheet.set_options(header_height=self.header_row_height)
            return
        except Exception:
            pass

        try:
            self.sheet.CH.set_height(self.header_row_height)
            return
        except Exception:
            pass

    def _apply_base_formatting(self):
        """
        Applies:
            - dark header background
            - green fill to non-empty calendar cells
            - purple headers for vacation dates
            - orange headers for training dates
            - narrow marker columns for start/end boundaries
        """

        self.sheet.dehighlight_all(redraw=False)

        total_cols = len(self.display_df.columns)

        # Base header color
        for col_idx in range(total_cols):
            self.sheet.highlight_cells(
                row=0,
                column=col_idx,
                canvas="header",
                bg=self.HEADER_FILL,
                fg=self.HEADER_TEXT,
                redraw=False,
                overwrite=True,
            )

        # Marker columns
        for col_idx, marker_info in self.marker_cols.items():
            color = marker_info["color"]
            kind = marker_info["kind"]

            # Header marker
            self.sheet.highlight_cells(
                row=0,
                column=col_idx,
                canvas="header",
                bg=color,
                fg=self.HEADER_TEXT,
                redraw=False,
                overwrite=True,
            )

            for row_idx in range(len(self.display_df)):
                if kind == "start":
                    # Solid thick-looking marker
                    should_color = True
                else:
                    # Dashed-looking marker
                    should_color = row_idx % 2 == 0

                if should_color:
                    self.sheet.highlight_cells(
                        row=row_idx,
                        column=col_idx,
                        bg=color,
                        fg=color,
                        redraw=False,
                        overwrite=True,
                    )

        # Calendar cells: green if non-empty
        for col_idx, real_date in self.display_date_by_col_idx.items():
            internal_header = self.display_df.columns[col_idx]

            for row_idx, value in enumerate(self.display_df[internal_header]):
                if str(value).strip():
                    self.sheet.highlight_cells(
                        row=row_idx,
                        column=col_idx,
                        bg=self.WORK_FILL,
                        redraw=False,
                        overwrite=False,
                    )

        # Vacation headers
        for col_idx, real_date in self.display_date_by_col_idx.items():
            if self._date_in_any_range(real_date, self.vacation_ranges):
                self.sheet.highlight_cells(
                    row=0,
                    column=col_idx,
                    canvas="header",
                    bg=self.VACATION_HEADER_FILL,
                    fg=self.HEADER_TEXT,
                    redraw=False,
                    overwrite=True,
                )

        # Training headers second, so training wins if there is overlap
        if self.training_start is not None and self.training_end is not None:
            start = min(self.training_start, self.training_end)
            end = max(self.training_start, self.training_end)

            for col_idx, real_date in self.display_date_by_col_idx.items():
                if start <= real_date <= end:
                    self.sheet.highlight_cells(
                        row=0,
                        column=col_idx,
                        canvas="header",
                        bg=self.TRAINING_HEADER_FILL,
                        fg=self.HEADER_TEXT,
                        redraw=False,
                        overwrite=True,
                    )

    # ------------------------------------------------------------
    # Search
    # ------------------------------------------------------------

    def apply_search(self):
        if self.display_df is None:
            return

        search_text = self.search_var.get().strip().lower()

        self._apply_base_formatting()

        if not search_text:
            self.sheet.refresh()
            self.status_var.set("Search cleared.")
            return

        match_count = 0

        for row_idx, row in enumerate(self.display_df.itertuples(index=False, name=None)):
            for col_idx, value in enumerate(row):
                if search_text in str(value).lower():
                    self.sheet.highlight_cells(
                        row=row_idx,
                        column=col_idx,
                        bg=self.SEARCH_FILL,
                        redraw=False,
                        overwrite=True,
                    )
                    match_count += 1

        self.sheet.refresh()

        self.status_var.set(
            f"Found {match_count:,} matching cell(s) for: {self.search_var.get()}"
        )

    def clear_search(self):
        self.search_var.set("")
        self._apply_base_formatting()
        self.sheet.refresh()

        if self.display_df is not None:
            self.status_var.set(
                f"Loaded {len(self.source_df):,} rows and {len(self.source_df.columns):,} source columns."
            )

    # ------------------------------------------------------------
    # Date helpers
    # ------------------------------------------------------------

    @staticmethod
    def _normalize_date(value):
        if value is None or value == "":
            return None

        if isinstance(value, datetime):
            return value.date()

        if isinstance(value, date):
            return value

        text = str(value).strip()

        known_formats = [
            "%Y-%m-%d",
            "%m/%d/%Y",
            "%m/%d/%y",
        ]

        for fmt in known_formats:
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                pass

        try:
            return pd.to_datetime(text, errors="raise").date()
        except Exception:
            return None

    @staticmethod
    def _format_calendar_header(d):
        return f"{d:%a}, {d:%b} {d.day}"

    def _normalize_date_ranges(self, ranges):
        if not ranges:
            return []

        normalized = []

        for item in ranges:
            if isinstance(item, dict):
                start = item.get("start")
                end = item.get("end")
            else:
                try:
                    start, end = item
                except Exception:
                    continue

            start = self._normalize_date(start)
            end = self._normalize_date(end)

            if start is None or end is None:
                continue

            if end < start:
                start, end = end, start

            normalized.append((start, end))

        return normalized

    @staticmethod
    def _date_in_any_range(d, ranges):
        return any(start <= d <= end for start, end in ranges)


def main():
    root = tk.Tk()
    root.title("UPS Bid Analyzer - Table Viewer")
    root.geometry("1400x800")

    viewer = BidSpreadsheetViewer(root)
    viewer.pack(fill="both", expand=True)

    sample_df = build_sample_dataframe()

    viewer.load_dataframe(
        sample_df,
        training_start="2026-07-20",
        training_end="2026-07-23",
        vacation_ranges=[
            {"start": "2026-07-26", "end": "2026-08-01"},
        ],
    )

    root.mainloop()


if __name__ == "__main__":
    main()







