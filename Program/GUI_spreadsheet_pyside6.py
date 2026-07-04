from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Iterable, Optional
from Processing_fucntions import line_numbers_to_bid_string

import pandas as pd

from PySide6.QtCore import (
    QAbstractTableModel,
    QItemSelection,
    QItemSelectionModel,
    QMimeData,
    QModelIndex,
    QPoint,
    QRect,
    Qt,
    QTimer,
)
from PySide6.QtGui import (
    QAction,
    QBrush,
    QColor,
    QDrag,
    QFont,
    QFontMetrics,
    QKeySequence,
    QPainter,
    QPalette,
    QPen,
    QPixmap,
    QShortcut,
)
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QHBoxLayout,
    QComboBox,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QStyledItemDelegate,
    QStyleOptionHeader,
    QTableView,
    QVBoxLayout,
    QWidget,
)


# -----------------------------
# Date / formatting helpers
# -----------------------------

def normalize_date(value) -> Optional[date]:
    """Convert a column name or user input into a date, when possible."""
    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    try:
        converted = pd.to_datetime(value, errors="coerce")
    except Exception:
        return None

    if pd.isna(converted):
        return None

    return converted.date()


def normalize_date_ranges(ranges) -> list[tuple[date, date]]:
    """
    Accepts either:
        [{"start": "2026-08-02", "end": "2026-08-08"}]
    or:
        [("2026-08-02", "2026-08-08")]
    """
    if not ranges:
        return []

    result: list[tuple[date, date]] = []

    for item in ranges:
        if isinstance(item, dict):
            start = item.get("start")
            end = item.get("end")
        else:
            try:
                start, end = item
            except Exception:
                continue

        start_date = normalize_date(start)
        end_date = normalize_date(end)

        if start_date is None or end_date is None:
            continue

        if end_date < start_date:
            start_date, end_date = end_date, start_date

        result.append((start_date, end_date))

    return result


def format_calendar_header(d: date) -> str:
    """Cross-platform date header like: Wed, May 27."""
    return f"{d:%a}, {d:%b} {d.day}"


def date_in_any_range(d: date, ranges: Iterable[tuple[date, date]]) -> bool:
    return any(start <= d <= end for start, end in ranges)


def is_blank(value) -> bool:
    if value is None or value == "":
        return True

    try:
        return bool(pd.isna(value))
    except Exception:
        return False


def excel_width_to_pixels(width: int | float) -> int:
    """
    Rough conversion from Excel-style character width to Qt pixels.
    This keeps your existing Excel numbers useful in the GUI.
    """
    return int((float(width) * 7) + 12)


def contiguous_ranges(rows: Iterable[int]) -> list[tuple[int, int]]:
    rows = sorted(set(rows))
    if not rows:
        return []

    ranges: list[tuple[int, int]] = []
    start = previous = rows[0]

    for row in rows[1:]:
        if row == previous + 1:
            previous = row
        else:
            ranges.append((start, previous))
            start = previous = row

    ranges.append((start, previous))
    return ranges

# -----------------------------
# Config helpers
# -----------------------------

def load_bid_config(config_path: str | Path) -> dict:
    path = Path(config_path)
    if not path.exists():
        return {}

    try:
        with path.open("r", encoding="utf-8") as file:
            data = json.load(file)
    except Exception:
        return {}

    if isinstance(data, dict):
        return data

    return {}


def save_bid_config_value(config_path: str | Path, key: str, value) -> None:
    path = Path(config_path)
    config = load_bid_config(path)
    config[key] = value

    try:
        with path.open("w", encoding="utf-8") as file:
            json.dump(config, file, indent=4)
    except Exception:
        # Do not crash the GUI just because config saving failed.
        pass


# -----------------------------
# Table model
# -----------------------------



@dataclass(frozen=True)
class TableTheme:
    background: QColor
    alternate_background: QColor
    text: QColor
    grid: QColor
    selection_background: QColor
    selection_text: QColor
    header_background: QColor
    header_text: QColor
    calendar_occupied_fill: QColor


TABLE_THEMES = {
    "light": TableTheme(
        background=QColor("#FFFFFF"),
        alternate_background=QColor("#F7F7F7"),
        text=QColor("#000000"),
        grid=QColor("#D0D0D0"),
        selection_background=QColor("#D7E9FF"),
        selection_text=QColor("#000000"),
        header_background=QColor("#007FFF"),
        header_text=QColor("#FFFFFF"),
        calendar_occupied_fill=QColor("#C6EFCE"),
    ),
    "dark": TableTheme(
        background=QColor("#1E1E1E"),
        alternate_background=QColor("#2A2A2A"),
        text=QColor("#F2F2F2"),
        grid=QColor("#555555"),
        selection_background=QColor("#375A7F"),
        selection_text=QColor("#FFFFFF"),
        header_background=QColor("#6395EE"),
        header_text=QColor("#FFFFFF"),
        calendar_occupied_fill=QColor("#2F6B3B"),
    ),
}


def normalize_theme_name(value: str | None) -> str:
    name = str(value or "light").strip().lower()
    if name not in TABLE_THEMES:
        return "light"
    return name


def normalize_column_key(value) -> str:
    return "".join(ch for ch in str(value).lower() if ch.isalnum())


@dataclass(frozen=True)
class BorderSpec:
    color: QColor
    width: int
    style: Qt.PenStyle


class DataFrameTableModel(QAbstractTableModel):
    """
    Pandas-backed table model.

    Important:
        Moving or dragging rows changes only self._row_order.
        It does not reorder the original DataFrame passed into the viewer.
    """

    MIME_TYPE = "application/x-bid-spreadsheet-rows"

    green_fill = QColor("#C6EFCE")
    vacation_fill = QColor("#800080")
    training_fill = QColor("#FFA500")
    white = QColor("#FFFFFF")
    black = QColor("#000000")

    def __init__(
        self,
        df: pd.DataFrame,
        *,
        calendar_cols=None,
        training_start=None,
        training_end=None,
        vacation_ranges=None,
        editable: bool = False,
        copy_data: bool = True,
        theme: str = "light",
        body_font_point_size: int = 12,
        parent=None,
    ):
        super().__init__(parent)

        self._theme_name = normalize_theme_name(theme)
        self._theme = TABLE_THEMES[self._theme_name]
        self._body_font_point_size = int(body_font_point_size)

        self._source_df = df
        self._df = df.copy(deep=copy_data)
        self._row_order = list(range(len(self._df)))
        self._last_moved_view_rows: list[int] = []
        self._editable = editable

        self.training_start = normalize_date(training_start)
        self.training_end = normalize_date(training_end)
        self.vacation_ranges = normalize_date_ranges(vacation_ranges)

        self._columns = list(self._df.columns)
        self._calendar_dates_by_col = self._build_calendar_column_map(calendar_cols)

    def _build_calendar_column_map(self, calendar_cols) -> dict[int, date]:
        if calendar_cols is None:
            calendar_date_set = {
                normalize_date(col)
                for col in self._columns
                if normalize_date(col) is not None
            }
        else:
            calendar_date_set = {
                normalize_date(col)
                for col in calendar_cols
                if normalize_date(col) is not None
            }

        calendar_date_set.discard(None)

        result: dict[int, date] = {}

        for col_idx, column_name in enumerate(self._columns):
            column_date = normalize_date(column_name)
            if column_date in calendar_date_set:
                result[col_idx] = column_date

        return result

    def rowCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._row_order)

    def columnCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._columns)

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        view_row = index.row()
        col = index.column()
        source_row = self._row_order[view_row]
        value = self._df.iat[source_row, col]
        column_date = self._calendar_dates_by_col.get(col)

        if role in (Qt.DisplayRole, Qt.EditRole):
            if is_blank(value):
                return ""
            return str(value)

        if role == Qt.BackgroundRole:
            # Match the Excel export: non-empty calendar cells get a green fill.
            # Dark mode uses a darker green so the text stays readable.
            if column_date is not None and not is_blank(value):
                return QBrush(self._theme.calendar_occupied_fill)
            return None

        if role == Qt.ForegroundRole:
            return QBrush(self._theme.text)

        if role == Qt.FontRole:
            font = QFont()
            font.setPointSize(self._body_font_point_size)

            # Keep your Line Number / Line Numbers cells bold.
            if self.is_line_number_column(col):
                font.setBold(True)

            return font

        if role == Qt.TextAlignmentRole:
            if column_date is not None:
                return Qt.AlignCenter
            return Qt.AlignVCenter | Qt.AlignLeft

        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role=Qt.DisplayRole):
        if orientation == Qt.Vertical:
            if role == Qt.DisplayRole:
                return str(section + 1)

            if role == Qt.TextAlignmentRole:
                return Qt.AlignCenter

            if role == Qt.ForegroundRole:
                return QBrush(QColor("#FFFFFF"))  # row number color only

            if role == Qt.FontRole:
                font = QFont()
                font.setPointSize(13)  # row number font size
                font.setBold(False)
                return font

            return None

        if orientation != Qt.Horizontal:
            return None

        column_date = self._calendar_dates_by_col.get(section)
        column_name = self._columns[section]

        if role == Qt.DisplayRole:
            if column_date is not None:
                return format_calendar_header(column_date)
            return str(column_name)

        if role == Qt.TextAlignmentRole:
            return Qt.AlignCenter

        if role == Qt.BackgroundRole:
            return QBrush(self.header_background_for_column(section))

        if role == Qt.ForegroundRole:
            return QBrush(self.header_foreground_for_column(section))

        if role == Qt.FontRole:
            font = QFont()
            font.setBold(False)
            font.setPointSize(14)
            return font

        return None

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.ItemIsDropEnabled

        flags = (
            Qt.ItemIsEnabled
            | Qt.ItemIsSelectable
            | Qt.ItemIsDragEnabled
            | Qt.ItemIsDropEnabled
        )

        if self._editable:
            flags |= Qt.ItemIsEditable

        return flags

    def setData(self, index: QModelIndex, value, role=Qt.EditRole) -> bool:
        if not self._editable or role != Qt.EditRole or not index.isValid():
            return False

        source_row = self._row_order[index.row()]
        self._df.iat[source_row, index.column()] = value
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        return True

    def supportedDropActions(self):
        return Qt.MoveAction

    def supportedDragActions(self):
        return Qt.MoveAction

    def mimeTypes(self) -> list[str]:
        return [self.MIME_TYPE]

    def mimeData(self, indexes) -> QMimeData:
        mime_data = QMimeData()
        rows = sorted({index.row() for index in indexes if index.isValid()})
        payload = json.dumps(rows).encode("utf-8")
        mime_data.setData(self.MIME_TYPE, payload)
        return mime_data

    def canDropMimeData(self, data, action, row, column, parent) -> bool:
        return action in (Qt.MoveAction, Qt.IgnoreAction) and data.hasFormat(self.MIME_TYPE)

    def dropMimeData(self, data, action, row, column, parent) -> bool:
        if action == Qt.IgnoreAction:
            return True

        if action != Qt.MoveAction or not data.hasFormat(self.MIME_TYPE):
            return False

        try:
            rows = json.loads(bytes(data.data(self.MIME_TYPE)).decode("utf-8"))
        except Exception:
            return False

        if row == -1 and parent.isValid():
            destination_row = parent.row()
        elif row == -1:
            destination_row = self.rowCount()
        else:
            destination_row = row

        self._last_moved_view_rows = self.move_rows_to(rows, destination_row)
        return True

    def take_last_moved_view_rows(self) -> list[int]:
        """Return and clear the most recent rows produced by a drag/drop move."""
        rows = list(self._last_moved_view_rows)
        self._last_moved_view_rows = []
        return rows

    def is_calendar_column(self, col: int) -> bool:
        return col in self._calendar_dates_by_col

    def column_key(self, col: int) -> str:
        if 0 <= col < len(self._columns):
            return normalize_column_key(self._columns[col])
        return ""

    def is_line_number_column(self, col: int) -> bool:
        return self.column_key(col) in {"linenumber", "linenumbers"}

    def is_extra_vacation_column(self, col: int) -> bool:
        return self.column_key(col) in {"extravacationdays", "extravacationday"}

    def is_training_column(self, col: int) -> bool:
        return self.column_key(col) == "training"

    def set_theme(self, theme: str):
        self._theme_name = normalize_theme_name(theme)
        self._theme = TABLE_THEMES[self._theme_name]

        if self.rowCount() and self.columnCount():
            top_left = self.index(0, 0)
            bottom_right = self.index(self.rowCount() - 1, self.columnCount() - 1)
            self.dataChanged.emit(
                top_left,
                bottom_right,
                [Qt.BackgroundRole, Qt.ForegroundRole, Qt.FontRole],
            )

        self.headerDataChanged.emit(Qt.Horizontal, 0, max(0, self.columnCount() - 1))
        self.headerDataChanged.emit(Qt.Vertical, 0, max(0, self.rowCount() - 1))

    def set_body_font_point_size(self, point_size: int):
        """
        Change only the body cell font size.

        This intentionally does not affect the horizontal column headers or the
        vertical row-number headers, because those are controlled separately in
        headerData().
        """
        point_size = max(6, int(point_size))

        if point_size == self._body_font_point_size:
            return

        self._body_font_point_size = point_size

        if self.rowCount() and self.columnCount():
            top_left = self.index(0, 0)
            bottom_right = self.index(self.rowCount() - 1, self.columnCount() - 1)
            self.dataChanged.emit(top_left, bottom_right, [Qt.FontRole])

    def column_uses_training_header_color(self, col: int) -> bool:
        if self.is_training_column(col):
            return True

        column_date = self._calendar_dates_by_col.get(col)
        if column_date is None:
            return False

        # Training wins over vacation for calendar header color, matching the Excel export.
        if self.training_start is not None and self.training_end is not None:
            start = min(self.training_start, self.training_end)
            end = max(self.training_start, self.training_end)
            return start <= column_date <= end

        return False

    def column_uses_vacation_header_color(self, col: int) -> bool:
        if self.is_extra_vacation_column(col):
            return True

        column_date = self._calendar_dates_by_col.get(col)
        if column_date is None:
            return False

        return date_in_any_range(column_date, self.vacation_ranges)

    def header_background_for_column(self, col: int) -> QColor:
        if self.column_uses_training_header_color(col):
            return self.training_fill

        if self.column_uses_vacation_header_color(col):
            return self.vacation_fill

        return self._theme.header_background

    def header_foreground_for_column(self, col: int) -> QColor:
        # Black text is more readable on the orange/yellow training header.
        if self.column_uses_training_header_color(col):
            return self.black

        if self.column_uses_vacation_header_color(col):
            return self.white

        return self._theme.header_text

    def border_specs_for_column(self, col: int) -> tuple[Optional[BorderSpec], Optional[BorderSpec]]:
        """Return left/right border specs for a calendar column."""
        column_date = self._calendar_dates_by_col.get(col)
        if column_date is None:
            return None, None

        left = None
        right = None

        # Training markers.
        if self.training_start is not None and column_date == self.training_start:
            left = BorderSpec(self.training_fill, 3, Qt.SolidLine)

        if self.training_end is not None and column_date == self.training_end:
            right = BorderSpec(self.training_fill, 2, Qt.DashLine)

        # Vacation markers. Do not overwrite training on the same side.
        for vacation_start, vacation_end in self.vacation_ranges:
            if column_date == vacation_start and left is None:
                left = BorderSpec(self.vacation_fill, 3, Qt.SolidLine)
            if column_date == vacation_end and right is None:
                right = BorderSpec(self.vacation_fill, 2, Qt.DashLine)

        return left, right

    def move_rows_to(self, rows: Iterable[int], destination_row: int) -> list[int]:
        """
        Move view rows to a new display position.

        Only self._row_order changes. The original DataFrame and self._df row order
        are not physically reordered.
        """
        selected_rows = sorted({r for r in rows if 0 <= r < self.rowCount()})
        if not selected_rows:
            return []

        destination_row = max(0, min(destination_row, self.rowCount()))

        # If the drop lands inside the same selected block, there is nothing useful to do.
        if destination_row in selected_rows or destination_row == selected_rows[-1] + 1:
            return selected_rows

        self.layoutAboutToBeChanged.emit()

        selected_set = set(selected_rows)
        dragged_source_rows = [self._row_order[r] for r in selected_rows]
        remaining_order = [
            source_row
            for view_row, source_row in enumerate(self._row_order)
            if view_row not in selected_set
        ]

        removed_before_destination = sum(r < destination_row for r in selected_rows)
        insert_at = destination_row - removed_before_destination
        insert_at = max(0, min(insert_at, len(remaining_order)))

        self._row_order = (
            remaining_order[:insert_at]
            + dragged_source_rows
            + remaining_order[insert_at:]
        )

        self.layoutChanged.emit()

        return list(range(insert_at, insert_at + len(dragged_source_rows)))

    def move_rows_up(self, rows: Iterable[int]) -> list[int]:
        """Move selected view rows up by one display position."""
        valid_rows = [r for r in rows if 0 <= r < self.rowCount()]
        ranges = contiguous_ranges(valid_rows)
        if not ranges:
            return []

        new_selected_rows: list[int] = []

        self.layoutAboutToBeChanged.emit()

        for start, end in ranges:
            if start == 0:
                new_selected_rows.extend(range(start, end + 1))
                continue

            before = self._row_order[start - 1]
            selected_block = self._row_order[start : end + 1]
            self._row_order[start - 1 : end + 1] = selected_block + [before]
            new_selected_rows.extend(range(start - 1, end))

        self.layoutChanged.emit()
        return new_selected_rows

    def move_rows_down(self, rows: Iterable[int]) -> list[int]:
        """Move selected view rows down by one display position."""
        valid_rows = [r for r in rows if 0 <= r < self.rowCount()]
        ranges = contiguous_ranges(valid_rows)
        if not ranges:
            return []

        new_selected_rows: list[int] = []

        self.layoutAboutToBeChanged.emit()

        for start, end in reversed(ranges):
            if end >= self.rowCount() - 1:
                new_selected_rows.extend(range(start, end + 1))
                continue

            after = self._row_order[end + 1]
            selected_block = self._row_order[start : end + 1]
            self._row_order[start : end + 2] = [after] + selected_block
            new_selected_rows.extend(range(start + 1, end + 2))

        self.layoutChanged.emit()
        return sorted(new_selected_rows)

    def reset_view_order(self):
        self.layoutAboutToBeChanged.emit()
        self._row_order = list(range(len(self._df)))
        self.layoutChanged.emit()

    def get_view_dataframe(self, *, reset_index: bool = False) -> pd.DataFrame:
        """
        Return a DataFrame in the current visual row order.
        This is a copy and does not mutate the original DataFrame.
        """
        result = self._df.iloc[self._row_order].copy()
        if reset_index:
            result = result.reset_index(drop=True)
        return result

    def get_original_dataframe(self) -> pd.DataFrame:
        """Return the original DataFrame object that was passed into the model."""
        return self._source_df


# -----------------------------
# Delegate for calendar borders
# -----------------------------

class CalendarBorderDelegate(QStyledItemDelegate):
    """Draws thick/dashed vertical date markers on top of normal table cells."""

    def paint(self, painter: QPainter, option, index: QModelIndex):
        super().paint(painter, option, index)

        model = index.model()
        if not hasattr(model, "border_specs_for_column"):
            return

        left_spec, right_spec = model.border_specs_for_column(index.column())
        if left_spec is None and right_spec is None:
            return

        rect = option.rect.adjusted(0, 0, -1, -1)

        painter.save()

        if left_spec is not None:
            pen = QPen(left_spec.color, left_spec.width)
            pen.setStyle(left_spec.style)
            painter.setPen(pen)
            painter.drawLine(rect.left(), rect.top(), rect.left(), rect.bottom())

        if right_spec is not None:
            pen = QPen(right_spec.color, right_spec.width)
            pen.setStyle(right_spec.style)
            painter.setPen(pen)
            painter.drawLine(rect.right(), rect.top(), rect.right(), rect.bottom())

        painter.restore()


class ColorHeaderView(QHeaderView):
    """Header view that reliably paints per-column header colors on all OS themes."""

    def paintSection(self, painter: QPainter, rect, logicalIndex: int):
        if not rect.isValid():
            return

        model = self.model()
        if model is None:
            super().paintSection(painter, rect, logicalIndex)
            return

        orientation = self.orientation()
        text = model.headerData(logicalIndex, orientation, Qt.DisplayRole) or ""
        alignment = model.headerData(logicalIndex, orientation, Qt.TextAlignmentRole)
        if alignment is None:
            alignment = Qt.AlignCenter

        background = model.headerData(logicalIndex, orientation, Qt.BackgroundRole)
        foreground = model.headerData(logicalIndex, orientation, Qt.ForegroundRole)
        font = model.headerData(logicalIndex, orientation, Qt.FontRole)

        if isinstance(background, QBrush):
            background_color = background.color()
        else:
            background_color = self.palette().color(QPalette.Button)

        if isinstance(foreground, QBrush):
            foreground_color = foreground.color()
        else:
            foreground_color = self.palette().color(QPalette.ButtonText)

        painter.save()
        painter.fillRect(rect, background_color)

        grid_color = QColor("#888888")
        if hasattr(model, "_theme"):
            grid_color = model._theme.grid

        painter.setPen(QPen(grid_color, 1))
        painter.drawRect(rect.adjusted(0, 0, -1, -1))

        if isinstance(font, QFont):
            painter.setFont(font)

        painter.setPen(foreground_color)
        painter.drawText(
            rect.adjusted(5, 2, -5, -2),
            alignment | Qt.TextWordWrap,
            str(text),
        )
        painter.restore()



class ZoomableTableView(QTableView):
    """
    QTableView with:
        - Excel-like Ctrl+mouse-wheel zoom support
        - cleaner internal row drag/drop feedback
        - full-width drop marker
        - custom drag pixmap
        - gentle auto-scroll while dragging
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        self._drop_indicator_row: Optional[int] = None
        self._last_drag_pos: Optional[QPoint] = None
        self._drop_indicator_color = QColor("#007FFF")
        self._drop_indicator_fill = QColor("#007FFF")
        self._drop_indicator_fill.setAlpha(35)

        self._auto_scroll_timer = QTimer(self)
        self._auto_scroll_timer.setInterval(30)
        self._auto_scroll_timer.timeout.connect(self._auto_scroll_during_drag)

        # Rows currently being dragged. This is used only for visual feedback
        # such as the QDrag ghost text and the "To Position" label.
        self._drag_rows: list[int] = []

        self.setAutoScroll(True)
        self.setAutoScrollMargin(45)

    def set_drop_indicator_color(self, color: QColor):
        self._drop_indicator_color = QColor(color)
        self._drop_indicator_fill = QColor(color)
        self._drop_indicator_fill.setAlpha(35)
        self.viewport().update()

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            viewer = self.parent()

            if hasattr(viewer, "zoom_in") and hasattr(viewer, "zoom_out"):
                delta = event.angleDelta().y()

                if delta > 0:
                    viewer.zoom_in()
                elif delta < 0:
                    viewer.zoom_out()

                event.accept()
                return

        super().wheelEvent(event)

    def startDrag(self, supported_actions):
        """
        Use a cleaner drag image instead of the default tiny/unclear table drag.
        The actual row movement is still handled by DataFrameTableModel.dropMimeData().
        """
        model = self.model()
        selection_model = self.selectionModel()

        if model is None or selection_model is None:
            super().startDrag(supported_actions)
            return

        indexes = self.selectedIndexes()
        if not indexes:
            return

        self._drag_rows = self._selected_drag_rows()
        if not self._drag_rows:
            return

        drag = QDrag(self)
        drag.setMimeData(model.mimeData(indexes))

        pixmap = self._make_drag_pixmap()
        if not pixmap.isNull():
            drag.setPixmap(pixmap)

            # Offset the drag label so the mouse cursor does not sit on top of
            # the text. Negative hotspot values place the pixmap to the right
            # and slightly below the cursor.
            drag.setHotSpot(QPoint(-28, -10))

        try:
            drag.exec(Qt.MoveAction)
        finally:
            self._clear_drag_feedback()

    def dragEnterEvent(self, event):
        if self._is_own_row_drag(event):
            event.setDropAction(Qt.MoveAction)
            event.accept()
            return

        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if self._is_own_row_drag(event):
            pos = self._event_position(event)
            self._last_drag_pos = pos
            self._update_drop_indicator(pos)

            if not self._auto_scroll_timer.isActive():
                self._auto_scroll_timer.start()

            event.setDropAction(Qt.MoveAction)
            event.accept()

            viewer = self.parent()
            if hasattr(viewer, "status_label"):
                to_position = self._preview_final_insert_position(self._drop_indicator_row)
                viewer.status_label.setText(
                    f"Drop target: Position {to_position}. The blue line shows where they will be inserted."
                )
            return

        super().dragMoveEvent(event)

    def dragLeaveEvent(self, event):
        self._clear_drag_feedback()
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        if self._is_own_row_drag(event):
            model = self.model()
            pos = self._event_position(event)
            destination_row = self._drop_indicator_row

            if destination_row is None:
                destination_row = self._drop_row_from_position(pos)

            ok = model.dropMimeData(
                event.mimeData(),
                Qt.MoveAction,
                destination_row,
                0,
                QModelIndex(),
            )

            self._clear_drag_feedback()

            if ok:
                event.setDropAction(Qt.MoveAction)
                event.accept()
                QTimer.singleShot(0, self._select_rows_moved_by_drop)

                viewer = self.parent()
                if hasattr(viewer, "update_bid_string_preview"):
                    viewer.update_bid_string_preview(copy_to_clipboard=False)
                if hasattr(viewer, "status_label"):
                    viewer.status_label.setText(
                        "Rows moved. Original DataFrame was not reordered."
                    )
                return

        self._clear_drag_feedback()
        super().dropEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)

        if self._drop_indicator_row is None or self.model() is None:
            return

        painter = QPainter(self.viewport())
        painter.setRenderHint(QPainter.Antialiasing, True)

        line_y, highlight_rect = self._drop_indicator_geometry(self._drop_indicator_row)

        if highlight_rect.isValid():
            painter.fillRect(highlight_rect, self._drop_indicator_fill)

        pen = QPen(self._drop_indicator_color, 4, Qt.SolidLine)
        painter.setPen(pen)
        painter.drawLine(0, line_y, self.viewport().width(), line_y)

        # Small end caps make the insertion line easier to see on wide tables.
        cap_height = 12
        painter.fillRect(QRect(0, line_y - cap_height // 2, 8, cap_height), self._drop_indicator_color)
        painter.fillRect(
            QRect(self.viewport().width() - 8, line_y - cap_height // 2, 8, cap_height),
            self._drop_indicator_color,
        )

        self._paint_to_position_label(painter, line_y)

    def _is_own_row_drag(self, event) -> bool:
        model = self.model()
        return (
            model is not None
            and hasattr(model, "MIME_TYPE")
            and event.mimeData().hasFormat(model.MIME_TYPE)
        )

    def _event_position(self, event) -> QPoint:
        # PySide6 uses position(); older Qt APIs used pos().
        if hasattr(event, "position"):
            return event.position().toPoint()
        return event.pos()

    def _update_drop_indicator(self, pos: QPoint):
        new_row = self._drop_row_from_position(pos)

        if new_row == self._drop_indicator_row:
            return

        self._drop_indicator_row = new_row
        self.viewport().update()

    def _drop_row_from_position(self, pos: QPoint) -> int:
        model = self.model()
        if model is None or model.rowCount() == 0:
            return 0

        index = self.indexAt(pos)

        if not index.isValid():
            # Below all visible rows means append to the end.
            return model.rowCount()

        rect = self.visualRect(index)

        if pos.y() < rect.center().y():
            return index.row()

        return index.row() + 1

    def _drop_indicator_geometry(self, drop_row: int) -> tuple[int, QRect]:
        model = self.model()
        width = self.viewport().width()

        if model is None or model.rowCount() == 0:
            return 0, QRect()

        if drop_row <= 0:
            first_index = model.index(0, 0)
            first_rect = self.visualRect(first_index)
            line_y = max(0, first_rect.top())
            highlight_rect = QRect(0, line_y, width, max(1, first_rect.height()))
            return line_y, highlight_rect

        if drop_row >= model.rowCount():
            last_index = model.index(model.rowCount() - 1, 0)
            last_rect = self.visualRect(last_index)
            line_y = min(self.viewport().height() - 1, last_rect.bottom() + 1)
            highlight_rect = QRect(0, max(0, line_y - last_rect.height()), width, max(1, last_rect.height()))
            return line_y, highlight_rect

        target_index = model.index(drop_row, 0)
        target_rect = self.visualRect(target_index)
        line_y = target_rect.top()
        highlight_rect = QRect(0, target_rect.top(), width, max(1, target_rect.height()))
        return line_y, highlight_rect

    def _clear_drag_feedback(self):
        self._drop_indicator_row = None
        self._last_drag_pos = None
        self._drag_rows = []
        self._auto_scroll_timer.stop()
        self.viewport().update()

    def _auto_scroll_during_drag(self):
        if self._last_drag_pos is None:
            self._auto_scroll_timer.stop()
            return

        margin = self.autoScrollMargin()
        y = self._last_drag_pos.y()
        height = self.viewport().height()
        scroll_bar = self.verticalScrollBar()

        if y < margin:
            scroll_bar.setValue(scroll_bar.value() - 1)
        elif y > height - margin:
            scroll_bar.setValue(scroll_bar.value() + 1)

        # After auto-scrolling, the same mouse position may point at a different
        # row, so refresh the insertion line and the position label.
        if self._drag_rows:
            self._update_drop_indicator(self._last_drag_pos)

            viewer = self.parent()
            if hasattr(viewer, "status_label"):
                to_position = self._preview_final_insert_position(self._drop_indicator_row)
                viewer.status_label.setText(
                    f"Drop target: Position {to_position}. The blue line shows where they will be inserted."
                )

    def _select_rows_moved_by_drop(self):
        model = self.model()
        if model is None or not hasattr(model, "take_last_moved_view_rows"):
            return

        moved_rows = model.take_last_moved_view_rows()
        viewer = self.parent()

        if moved_rows and hasattr(viewer, "select_rows"):
            viewer.select_rows(moved_rows)

    def _selected_drag_rows(self) -> list[int]:
        selection_model = self.selectionModel()
        if selection_model is None:
            return []

        rows = sorted(index.row() for index in selection_model.selectedRows())
        if not rows:
            rows = sorted({index.row() for index in selection_model.selectedIndexes()})

        return rows

    def _line_numbers_text_for_rows(self, rows: Iterable[int]) -> str:
        model = self.model()
        if model is None:
            return ""

        line_number_col = None

        for col in range(model.columnCount()):
            if hasattr(model, "is_line_number_column") and model.is_line_number_column(col):
                line_number_col = col
                break

        values: list[str] = []

        if line_number_col is not None:
            for row in rows:
                value = model.index(row, line_number_col).data(Qt.DisplayRole)
                if value not in (None, ""):
                    values.append(str(value))

        if values:
            return ", ".join(values)

        # Fallback if no Line Number column exists.
        return ", ".join(str(row + 1) for row in rows)

    def _preview_final_insert_position(self, destination_row: Optional[int]) -> int:
        """
        Return the 1-based final display position of the first dragged row.

        destination_row is the raw insertion row reported by the view before the
        dragged rows are removed. DataFrameTableModel.move_rows_to() removes the
        dragged rows first, so moving rows downward changes the final insert index.
        This mirrors that logic so the label shows the real final position.
        """
        model = self.model()
        if model is None:
            return 1

        selected_rows = sorted({r for r in self._drag_rows if 0 <= r < model.rowCount()})
        if not selected_rows:
            return 1

        if destination_row is None:
            destination_row = selected_rows[0]

        destination_row = max(0, min(int(destination_row), model.rowCount()))

        # This mirrors the no-op check inside DataFrameTableModel.move_rows_to().
        if destination_row in selected_rows or destination_row == selected_rows[-1] + 1:
            return selected_rows[0] + 1

        removed_before_destination = sum(r < destination_row for r in selected_rows)
        remaining_count = model.rowCount() - len(selected_rows)
        insert_at = destination_row - removed_before_destination
        insert_at = max(0, min(insert_at, remaining_count))

        return insert_at + 1

    def _make_drag_pixmap(self) -> QPixmap:
        if not self._drag_rows:
            return QPixmap()

        width = min(max(360, self.viewport().width() // 2), 720)
        padding_x = 16
        padding_y = 10

        text = f"Moving line numbers: {self._line_numbers_text_for_rows(self._drag_rows)}"

        font = QFont(self.font())
        font.setBold(True)
        font.setPointSize(max(10, font.pointSize() + 1))

        metrics = QFontMetrics(font)
        text_rect = metrics.boundingRect(
            QRect(0, 0, width - (padding_x * 2), 1000),
            Qt.AlignLeft | Qt.TextWordWrap,
            text,
        )

        height = max(48, text_rect.height() + (padding_y * 2))

        pixmap = QPixmap(width, height)
        pixmap.fill(Qt.transparent)

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing, True)

        background = QColor(self.palette().color(QPalette.Highlight))
        background.setAlpha(225)
        border = QColor(self._drop_indicator_color)
        text_color = QColor(self.palette().color(QPalette.HighlightedText))

        painter.setBrush(QBrush(background))
        painter.setPen(QPen(border, 2))
        painter.drawRoundedRect(QRect(1, 1, width - 2, height - 2), 10, 10)

        painter.setFont(font)
        painter.setPen(text_color)
        painter.drawText(
            QRect(padding_x, padding_y, width - (padding_x * 2), height - (padding_y * 2)),
            Qt.AlignLeft | Qt.AlignVCenter | Qt.TextWordWrap,
            text,
        )

        painter.end()
        return pixmap

    def _paint_to_position_label(self, painter: QPainter, line_y: int):
        """Draw the live destination position beside the blue insertion line."""
        if not self._drag_rows or self._drop_indicator_row is None:
            return

        to_position = self._preview_final_insert_position(self._drop_indicator_row)
        text = f"To Position: {to_position}"

        font = QFont(self.font())
        font.setBold(True)
        font.setPointSize(max(10, font.pointSize()))
        painter.setFont(font)

        metrics = QFontMetrics(font)
        padding_x = 10
        padding_y = 5
        text_width = metrics.horizontalAdvance(text)
        text_height = metrics.height()
        label_width = text_width + (padding_x * 2)
        label_height = text_height + (padding_y * 2)

        x = 75
        y = line_y + 8

        # If the insertion line is near the bottom, put the label above it.
        if y + label_height > self.viewport().height():
            y = line_y - label_height - 8

        # Keep it inside the viewport.
        x = max(4, min(x, self.viewport().width() - label_width - 4))
        y = max(4, min(y, self.viewport().height() - label_height - 4))

        rect = QRect(x, y, label_width, label_height)

        background = QColor(self.palette().color(QPalette.Highlight))
        background.setAlpha(235)
        text_color = QColor(self.palette().color(QPalette.HighlightedText))
        border = QColor(self._drop_indicator_color)

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setBrush(QBrush(background))
        painter.setPen(QPen(border, 2))
        painter.drawRoundedRect(rect, 8, 8)
        painter.setPen(text_color)
        painter.drawText(rect, Qt.AlignCenter, text)
        painter.restore()

# -----------------------------
# Viewer widget
# -----------------------------

class BidSpreadsheetViewer(QWidget):
    def __init__(
        self,
        df: pd.DataFrame,
        *,
        calendar_cols=None,
        training_start=None,
        training_end=None,
        vacation_ranges=None,
        calendar_col_width=32,
        calendar_row_height=40,
        header_row_height=45,
        non_calendar_max_width=32,
        editable: bool = False,
        config_path: str | Path = "bid_config.json",
        bid_string_function: Optional[Callable[[pd.DataFrame, int, str], str]] = None,
        bid_line_column: str = "Line Number",
        theme: str = "light",
        body_font_point_size: int = 12,
        zoom_percent: int = 100,
        parent=None,
    ):
        super().__init__(parent)

        self.calendar_col_width = calendar_col_width
        self.calendar_row_height = calendar_row_height
        self.header_row_height = header_row_height
        self.non_calendar_max_width = non_calendar_max_width
        self.config_path = Path(config_path)
        self.bid_string_function = bid_string_function or line_numbers_to_bid_string
        self.bid_line_column = bid_line_column
        self.theme_name = normalize_theme_name(theme)
        self.base_body_font_point_size = int(body_font_point_size)
        self.zoom_percent = max(60, min(200, int(zoom_percent)))

        self.table = ZoomableTableView(self)
        self.table.setHorizontalHeader(ColorHeaderView(Qt.Horizontal, self.table))
        self.table.setVerticalHeader(ColorHeaderView(Qt.Vertical, self.table))
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)  # important: manual row order stays under your control
        self.table.setItemDelegate(CalendarBorderDelegate(self.table))
        self.table.verticalHeader().setDefaultSectionSize(self.calendar_row_height)
        self.table.horizontalHeader().setFixedHeight(self.header_row_height)
        self.table.horizontalHeader().setSectionsClickable(True)

        # Drag-and-drop row reordering.
        self.table.setDragEnabled(True)
        self.table.setAcceptDrops(True)
        self.table.setDropIndicatorShown(False)  # custom full-width drop indicator is drawn by ZoomableTableView
        self.table.setDragDropMode(QAbstractItemView.InternalMove)
        self.table.setDefaultDropAction(Qt.MoveAction)
        self.table.setDragDropOverwriteMode(False)

        self.model = DataFrameTableModel(
            df,
            calendar_cols=calendar_cols,
            training_start=training_start,
            training_end=training_end,
            vacation_ranges=vacation_ranges,
            editable=editable,
            theme=self.theme_name,
            body_font_point_size=self.scaled_body_font_point_size(),
            parent=self,
        )
        self.table.setModel(self.model)

        self.status_label = QLabel(
            "Drag selected row(s) to reorder. The original DataFrame is not reordered."
        )

        self.copy_bid_button = QPushButton("Copy Bid String")
        self.copy_bid_button.clicked.connect(self.copy_bid_string)

        self.number_of_lines_spinbox = QSpinBox()
        self.number_of_lines_spinbox.setMinimum(1)
        self.number_of_lines_spinbox.setMaximum(9999)
        self.number_of_lines_spinbox.setValue(self.load_number_of_lines_to_bid(default=20))
        self.number_of_lines_spinbox.valueChanged.connect(self.save_number_of_lines_to_bid)

        self.bid_string_preview = QLineEdit()
        self.bid_string_preview.setReadOnly(True)
        self.bid_string_preview.setPlaceholderText("Copied bid string will appear here")
        self.bid_string_preview.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.theme_combo = QComboBox()
        self.theme_combo.addItem("Light", "light")
        self.theme_combo.addItem("Dark", "dark")
        self.theme_combo.setCurrentIndex(0 if self.theme_name == "light" else 1)
        self.theme_combo.currentIndexChanged.connect(self.on_theme_changed)

        self.zoom_out_button = QPushButton("−")
        self.zoom_out_button.setToolTip("Zoom out. You can also use Ctrl+mouse wheel down.")
        self.zoom_out_button.clicked.connect(self.zoom_out)

        self.zoom_reset_button = QPushButton(f"{self.zoom_percent}%")
        self.zoom_reset_button.setToolTip("Reset zoom to 100%")
        self.zoom_reset_button.clicked.connect(self.reset_zoom)

        self.zoom_in_button = QPushButton("+")
        self.zoom_in_button.setToolTip("Zoom in. You can also use Ctrl+mouse wheel up.")
        self.zoom_in_button.clicked.connect(self.zoom_in)

        self.move_up_button = QPushButton("Move Row Up")
        self.move_down_button = QPushButton("Move Row Down")
        self.reset_button = QPushButton("Reset Display Order")

        self.move_up_button.clicked.connect(self.move_selected_rows_up)
        self.move_down_button.clicked.connect(self.move_selected_rows_down)
        self.reset_button.clicked.connect(self.reset_display_order)

        bid_layout = QHBoxLayout()
        bid_layout.addWidget(QLabel("Lines to bid:"))
        bid_layout.addWidget(self.number_of_lines_spinbox)
        bid_layout.addWidget(self.copy_bid_button)
        bid_layout.addWidget(QLabel("Bid string:"))
        bid_layout.addWidget(self.bid_string_preview, stretch=1)
        bid_layout.addWidget(QLabel("Theme:"))
        bid_layout.addWidget(self.theme_combo)
        bid_layout.addWidget(QLabel("Zoom:"))
        bid_layout.addWidget(self.zoom_out_button)
        bid_layout.addWidget(self.zoom_reset_button)
        bid_layout.addWidget(self.zoom_in_button)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.move_up_button)
        button_layout.addWidget(self.move_down_button)
        button_layout.addWidget(self.reset_button)
        button_layout.addStretch(1)
        button_layout.addWidget(self.status_label)

        layout = QVBoxLayout(self)
        layout.addLayout(bid_layout)
        layout.addLayout(button_layout)
        layout.addWidget(self.table)

        QShortcut(QKeySequence("Alt+Up"), self, activated=self.move_selected_rows_up)
        QShortcut(QKeySequence("Alt+Down"), self, activated=self.move_selected_rows_down)
        QShortcut(QKeySequence("Ctrl+B"), self, activated=self.copy_bid_string)
        QShortcut(QKeySequence("Ctrl++"), self, activated=self.zoom_in)
        QShortcut(QKeySequence("Ctrl+="), self, activated=self.zoom_in)
        QShortcut(QKeySequence("Ctrl+-"), self, activated=self.zoom_out)
        QShortcut(QKeySequence("Ctrl+0"), self, activated=self.reset_zoom)

        self.model.layoutChanged.connect(
            lambda: self.update_bid_string_preview(copy_to_clipboard=False)
        )

        self.apply_table_theme(self.theme_name)
        self.apply_table_sizing()
        self.update_bid_string_preview(copy_to_clipboard=False)

    def zoom_factor(self) -> float:
        return self.zoom_percent / 100.0

    def scaled_body_font_point_size(self) -> int:
        return max(6, round(self.base_body_font_point_size * self.zoom_factor()))

    def set_zoom_percent(self, value: int):
        """
        Apply an Excel-like zoom to the table body.

        This scales body cells, row heights, and column widths. It intentionally
        does not change your horizontal column-header font size or vertical
        row-number-header font size.
        """
        value = max(60, min(200, int(value)))

        if value == self.zoom_percent:
            return

        self.zoom_percent = value
        self.zoom_reset_button.setText(f"{self.zoom_percent}%")

        self.model.set_body_font_point_size(self.scaled_body_font_point_size())
        self.apply_table_sizing()
        self.table.viewport().update()

        self.status_label.setText(
            f"Zoom set to {self.zoom_percent}%. Column and row headers kept their own font sizes."
        )

    def zoom_in(self):
        self.set_zoom_percent(self.zoom_percent + 10)

    def zoom_out(self):
        self.set_zoom_percent(self.zoom_percent - 10)

    def reset_zoom(self):
        self.set_zoom_percent(100)

    def apply_light_table_palette(self):
        """Backward-compatible helper. Prefer apply_table_theme("light")."""
        self.apply_table_theme("light")

    def on_theme_changed(self):
        theme = self.theme_combo.currentData()
        self.apply_table_theme(theme)

    def apply_table_theme(self, theme: str):
        """Apply a light or dark palette to the spreadsheet area with readable text."""
        self.theme_name = normalize_theme_name(theme)
        theme_data = TABLE_THEMES[self.theme_name]

        self.model.set_theme(self.theme_name)

        palette = self.table.palette()
        palette.setColor(QPalette.Base, theme_data.background)
        palette.setColor(QPalette.AlternateBase, theme_data.alternate_background)
        palette.setColor(QPalette.Text, theme_data.text)
        palette.setColor(QPalette.WindowText, theme_data.text)
        palette.setColor(QPalette.Highlight, theme_data.selection_background)
        palette.setColor(QPalette.HighlightedText, theme_data.selection_text)
        palette.setColor(QPalette.Button, theme_data.header_background)
        palette.setColor(QPalette.ButtonText, theme_data.header_text)

        self.table.setPalette(palette)
        self.table.horizontalHeader().setPalette(palette)
        self.table.verticalHeader().setPalette(palette)

        if hasattr(self.table, "set_drop_indicator_color"):
            self.table.set_drop_indicator_color(theme_data.header_background)

        # Avoid a dark viewport background from a global app stylesheet.
        self.table.viewport().setAutoFillBackground(True)
        self.table.viewport().setPalette(palette)

        self.table.setStyleSheet(
            f"""
            QTableView {{
                background-color: {theme_data.background.name()};
                alternate-background-color: {theme_data.alternate_background.name()};
                color: {theme_data.text.name()};
                gridline-color: {theme_data.grid.name()};
            }}
            QTableView::item:selected {{
                background-color: {theme_data.selection_background.name()};
                color: {theme_data.selection_text.name()};
            }}
            QHeaderView::section {{
                padding: 4px;
            }}
            """
        )

        self.table.horizontalHeader().update()
        self.table.verticalHeader().update()
        self.table.viewport().update()

    def apply_table_sizing(self):
        zoom = self.zoom_factor()

        calendar_width_px = round(excel_width_to_pixels(self.calendar_col_width) * zoom)
        non_calendar_max_px = round(excel_width_to_pixels(self.non_calendar_max_width) * zoom)
        non_calendar_min_px = round(80 * zoom)
        row_height_px = max(18, round(self.calendar_row_height * zoom))

        self.model.set_body_font_point_size(self.scaled_body_font_point_size())

        # This changes body row height only. The column header height remains
        # controlled by self.header_row_height and your existing header font settings.
        self.table.verticalHeader().setDefaultSectionSize(row_height_px)
        self.table.horizontalHeader().setFixedHeight(self.header_row_height)

        for col in range(self.model.columnCount()):
            if self.model.is_calendar_column(col):
                self.table.setColumnWidth(col, calendar_width_px)
            else:
                self.table.resizeColumnToContents(col)
                current = self.table.columnWidth(col)
                self.table.setColumnWidth(
                    col,
                    min(max(current, non_calendar_min_px), non_calendar_max_px),
                )

    def load_number_of_lines_to_bid(self, *, default: int = 20) -> int:
        config = load_bid_config(self.config_path)
        value = config.get("number_of_lines_to_bid", default)

        try:
            value = int(value)
        except Exception:
            value = default

        return max(1, value)

    def save_number_of_lines_to_bid(self, value: int):
        save_bid_config_value(self.config_path, "number_of_lines_to_bid", int(value))
        self.update_bid_string_preview(copy_to_clipboard=False)

    def selected_view_rows(self) -> list[int]:
        selected_rows = self.table.selectionModel().selectedRows()
        if selected_rows:
            return sorted(index.row() for index in selected_rows)

        # Fallback in case selection behavior changes later.
        selected_indexes = self.table.selectionModel().selectedIndexes()
        return sorted({index.row() for index in selected_indexes})

    def select_rows(self, rows: Iterable[int]):
        rows = sorted(set(rows))
        if not rows:
            return

        selection = QItemSelection()
        last_col = self.model.columnCount() - 1

        for row in rows:
            if 0 <= row < self.model.rowCount():
                top_left = self.model.index(row, 0)
                bottom_right = self.model.index(row, last_col)
                selection.select(top_left, bottom_right)

        selection_model = self.table.selectionModel()
        selection_model.clearSelection()
        selection_model.select(selection, QItemSelectionModel.Select | QItemSelectionModel.Rows)
        self.table.scrollTo(self.model.index(rows[0], 0))

    def move_selected_rows_up(self):
        rows = self.selected_view_rows()
        new_rows = self.model.move_rows_up(rows)
        self.select_rows(new_rows)
        self.status_label.setText("Moved selected row(s) up. Original DataFrame was not reordered.")
        self.update_bid_string_preview(copy_to_clipboard=False)

    def move_selected_rows_down(self):
        rows = self.selected_view_rows()
        new_rows = self.model.move_rows_down(rows)
        self.select_rows(new_rows)
        self.status_label.setText("Moved selected row(s) down. Original DataFrame was not reordered.")
        self.update_bid_string_preview(copy_to_clipboard=False)

    def reset_display_order(self):
        self.model.reset_view_order()
        self.status_label.setText("Display order reset. Original DataFrame was not changed.")
        self.update_bid_string_preview(copy_to_clipboard=False)

    def update_bid_string_preview(self, *, copy_to_clipboard: bool) -> str:
        ordered_df = self.get_view_dataframe(reset_index=True)
        number_of_lines = self.number_of_lines_spinbox.value()

        bid_string = self.bid_string_function(
            ordered_df,
            number_of_lines,
            self.bid_line_column,
        )

        self.bid_string_preview.setText(bid_string)

        if copy_to_clipboard:
            QApplication.clipboard().setText(bid_string)

        return bid_string

    def copy_bid_string(self):
        try:
            bid_string = self.update_bid_string_preview(copy_to_clipboard=True)
        except Exception as exc:
            QMessageBox.warning(self, "Could not copy bid string", str(exc))
            return

        self.save_number_of_lines_to_bid(self.number_of_lines_spinbox.value())
        self.status_label.setText(f"Copied bid string: {bid_string}")

    def get_view_dataframe(self, *, reset_index: bool = False) -> pd.DataFrame:
        """Use this when you want the DataFrame in the manually arranged display order."""
        return self.model.get_view_dataframe(reset_index=reset_index)

    def get_original_dataframe(self) -> pd.DataFrame:
        """Returns the exact original DataFrame object passed into the viewer."""
        return self.model.get_original_dataframe()


class BidSpreadsheetWindow(QMainWindow):
    def __init__(self, df: pd.DataFrame, **viewer_kwargs):
        super().__init__()
        self.setWindowTitle("Bid Spreadsheet Viewer")
        self.resize(1400, 800)

        self.viewer = BidSpreadsheetViewer(df, **viewer_kwargs)
        self.setCentralWidget(self.viewer)

        self._build_menu()

    def _build_menu(self):
        file_menu = self.menuBar().addMenu("File")

        close_action = QAction("Close", self)
        close_action.triggered.connect(self.close)
        file_menu.addAction(close_action)

    def get_view_dataframe(self, *, reset_index: bool = False) -> pd.DataFrame:
        return self.viewer.get_view_dataframe(reset_index=reset_index)


# -----------------------------
# Small test/demo
# -----------------------------

if __name__ == "__main__":
    import sys

    bid_dates = pd.date_range("2026-07-12", "2026-07-20", freq="D")

    sample = pd.DataFrame(
        {
            "Line Number": [3, 5, 1, 7, 8, 9, 10, 19],
            "Extra Vacation Days": [0, 1, 0, 2, 0, 0, 0, 1],
            **{
                d.date(): ["TRIP 101", "", "VTO", "TRIP 205", "", "TRIP 301", "", "TRIP 401"]
                for d in bid_dates
            },
            "Training": [80, 20, 0, 100, 50, 40, 60, 10],
            "Blockiness": [91.2, 73.5, 88.0, 95.0, 82.0, 77.0, 93.0, 70.0],
            "Premium": [1200, 900, 0, 1500, 600, 800, 1000, 500],
        }
    )

    app = QApplication(sys.argv)
    window = BidSpreadsheetWindow(
        sample,
        training_start="2026-07-15",
        training_end="2026-07-17",
        vacation_ranges=[{"start": "2026-07-19", "end": "2026-07-20"}],
        editable=False,
        config_path="bid_config.json",
    )
    window.show()
    sys.exit(app.exec())
