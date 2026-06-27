import pandas as pd
from datetime import date, datetime, timedelta
from collections import defaultdict

def build_line_calendar_values(line_data, bid_dates, off_value=""):
    """
    Creates the day-by-day calendar contents for one line.

    Returns:
        {
            date(2023, 6, 1): "{400SDF-ATL-[MDT27.8]",
            date(2023, 6, 2): "MDT-SDF-[MDT15.2]",
            ...
        }
    """

    RESERVE_CODES = {"VTO", "RB", "RA", "SA", "SB", "VOR"}

    def to_date(value):
        if value is None or value == "":
            return None
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        return datetime.strptime(str(value), "%Y-%m-%d").date()

    def date_range_strict(start, end):
        start = to_date(start)
        end = to_date(end)

        if start is None or end is None:
            return []

        if start > end:
            return []

        return [start + timedelta(days=i) for i in range((end - start).days + 1)]

    def format_rest(rest):
        if rest is None:
            return ""

        if isinstance(rest, (int, float)):
            return str(int(rest)) if float(rest).is_integer() else str(round(float(rest), 1))

        text = str(rest).strip()

        if "h" in text:
            hours, minutes = text.split("h", 1)
        elif ":" in text:
            hours, minutes = text.split(":", 1)
        else:
            return text

        hours = int(hours)
        minutes = int(minutes or 0)
        value = hours + minutes / 60

        return str(int(value)) if value.is_integer() else str(round(value, 1))

    def format_route_flags(route_flags):
        if not route_flags:
            return ""

        if isinstance(route_flags, str):
            flags = [route_flags]
        else:
            flags = list(route_flags)

        flags = [str(flag).strip() for flag in flags if str(flag).strip()]

        if not flags:
            return ""

        return f"({','.join(flags)})"

    def arrival_text(arrival, rest=None, close_trip=False):
        if rest is not None:
            text = f"[{arrival}{format_rest(rest)}] "
        else:
            text = str(arrival)

        if close_trip:
            text = text.rstrip() + "}"

        return text

    def append_piece(parts_by_date, d, piece):
        if not piece:
            return

        piece = str(piece)

        if piece == "*":
            if not parts_by_date[d]:
                parts_by_date[d] = "*"
            return

        if parts_by_date[d] == "*":
            parts_by_date[d] = piece
            return

        if not parts_by_date[d]:
            parts_by_date[d] = piece
            return

        if parts_by_date[d].endswith(" "):
            parts_by_date[d] += piece
        else:
            parts_by_date[d] += piece

    def render_trip(assignment):
        trip_id = assignment.get("trip_id")
        flights = assignment.get("flights") or []
        parts_by_date = defaultdict(str)

        previous_arrival = None
        previous_end_date = None
        previous_rest = None

        for index, flight in enumerate(flights):
            dep = flight.get("departure")
            arr = flight.get("arrival")

            start_date = to_date(flight.get("start_date"))
            end_date = to_date(flight.get("end_date"))

            if start_date is None and end_date is None:
                continue

            if start_date is None:
                start_date = end_date

            if end_date is None:
                end_date = start_date

            route_flags = format_route_flags(flight.get("route_flags"))
            rest = flight.get("rest")

            is_first_flight = index == 0
            is_last_flight = index == len(flights) - 1

            trip_open = f"{{{trip_id}" if is_first_flight else ""

            arrival = arrival_text(
                arr,
                rest=rest,
                close_trip=is_last_flight,
            )

            can_compress_departure = (
                not is_first_flight
                and previous_arrival == dep
                and previous_end_date == start_date
                and previous_rest is None
                and parts_by_date[start_date]
            )

            if start_date == end_date:
                if can_compress_departure:
                    piece = f"-{route_flags}{arrival}"
                else:
                    piece = f"{trip_open}{dep}-{route_flags}{arrival}"

                append_piece(parts_by_date, start_date, piece)

            else:
                if can_compress_departure:
                    departure_piece = "-"
                else:
                    departure_piece = f"{trip_open}{dep}-"

                arrival_piece = f"{route_flags}{arrival}"

                append_piece(parts_by_date, start_date, departure_piece)
                append_piece(parts_by_date, end_date, arrival_piece)

            if index < len(flights) - 1:
                next_flight = flights[index + 1]
                next_start = to_date(next_flight.get("start_date"))

                if next_start is not None:
                    gap_start = end_date + timedelta(days=1)
                    gap_end = next_start - timedelta(days=1)

                    for gap_day in date_range_strict(gap_start, gap_end):
                        append_piece(parts_by_date, gap_day, "*")

            previous_arrival = arr
            previous_end_date = end_date
            previous_rest = rest

        return parts_by_date

    def merge_pieces(pieces):
        cleaned = [str(p).strip() for p in pieces if p and str(p).strip()]

        if not cleaned:
            return off_value

        real_pieces = [p for p in cleaned if p != "*"]

        if real_pieces:
            return " ".join(dict.fromkeys(real_pieces))

        return chr(8212)

    text_by_date = defaultdict(list)

    for pp in line_data.get("PPs", line_data.get("pay_periods", [])):
        for assignment in pp.get("assignments", []):

            if assignment.get("flights"):
                rendered = render_trip(assignment)

            elif assignment.get("code") in RESERVE_CODES:
                assignment_date = to_date(assignment.get("date"))
                rendered = {assignment_date: assignment["code"]} if assignment_date else {}

            else:
                rendered = {}

            for d, text in rendered.items():
                if bid_dates[0] <= d <= bid_dates[-1]:
                    text_by_date[d].append(text)

    return {
        d: merge_pieces(text_by_date.get(d, []))
        for d in bid_dates
    }


def _date_range_inclusive(start, end):
    start = _to_date(start)
    end = _to_date(end)

    if start is None or end is None:
        return []

    if end < start:
        start, end = end, start

    return [
        start + timedelta(days=i)
        for i in range((end - start).days + 1)
    ]

def _to_date(value):
    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    return datetime.strptime(str(value), "%Y-%m-%d").date()

def master_lines_to_dataframe(
    master_lines,
    bid_start,
    bid_end,
    *,
    off_value="",
    date_col_format="%Y-%m-%d",
):
    bid_dates = _date_range_inclusive(bid_start, bid_end)
    rows = []

    for line_number, line_data in master_lines.items():

        date_values = build_line_calendar_values(
            line_data,
            bid_dates,
            off_value=off_value,
        )
        
        row = {
            "Line Number": line_number,
            "Extra Vacation Days": line_data.get("extra_vacation_days", 0),
        }

        # Calendar date columns go here
        for d in bid_dates:
            row[d.strftime(date_col_format)] = date_values[d]

        # Score / sorting columns go after the calendar
        row.update({
            "Training": line_data.get("training_fit_score", 0),
            "Blockiness": line_data.get("blockiness_score", 0),
            "Total DO": line_data.get("tot_DO", 0),
            "% tickets paid": line_data.get("company_ticket_pct", 0),
            "Premium": line_data.get("tot_Premium", line_data.get("tot_premium", 0)),
        })
        """
        row = {
            "Line Number": line_number,
            "Extra Vacation Days": line_data.get("extra_vacation_days", 0),
            "Training": line_data.get("training_fit_score", 0),
            "Blockiness": line_data.get("blockiness_score", 0),
            "Total DO": line_data.get("tot_DO", 0),
            "% tickets paid": line_data.get("company_ticket_pct", 0),
            "Premium": line_data.get("tot_premium")
        }

        for d in bid_dates:
            row[d.strftime(date_col_format)] = date_values[d]
        """
        rows.append(row)

    return pd.DataFrame(rows).sort_values("Line Number").reset_index(drop=True)



def sort_dataframe_by_conditions(
    df,
    sort_order,
    drop_all_zero_or_none_cols=True,
    reset_index=True,
    missing_col_action="raise",
):
    """
    Sorts a DataFrame using user-provided sorting conditions.

    Parameters
    ----------
    df:
        Pandas DataFrame.

    sort_order:
        List of sorting rules in priority order.

        Each rule should be either:

            ("Column Name", "desc")

        or:

            ("Column Name", "asc")

        You can also use:

            ("Column Name", "high_to_low")
            ("Column Name", "low_to_high")
            ("Column Name", False)   # False = descending
            ("Column Name", True)    # True = ascending

    drop_all_zero_or_none_cols:
        If True, any sorting column where all values are 0, None, NaN,
        or blank will be dropped from the DataFrame and removed from sorting.

    reset_index:
        If True, resets the index after sorting.

    missing_col_action:
        "raise" -> raise an error if a sorting column is missing.
        "ignore" -> skip missing sorting columns.
    """

    df = df.copy()

    def normalize_sort_direction(direction):
        """
        Returns True for ascending, False for descending.
        """
        if isinstance(direction, bool):
            return direction

        direction = str(direction).lower().strip()

        if direction in ["asc", "ascending", "low_to_high", "small_to_large", "smallest_to_largest"]:
            return True

        if direction in ["desc", "descending", "high_to_low", "large_to_small", "largest_to_smallest"]:
            return False

        raise ValueError(
            f"Invalid sort direction: {direction}. "
            "Use 'asc', 'desc', True, or False."
        )

    def numeric_col(col):
        """
        Converts column to numeric for sorting and zero-checking.

        Handles:
            10
            "10"
            "10%"
            ""
            None
            NaN
        """

        cleaned = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.strip()
            .replace({
                "": None,
                "None": None,
                "nan": None,
                "NaN": None,
            })
        )

        return pd.to_numeric(cleaned, errors="coerce").fillna(0)

    active_sort_cols = []
    ascending_values = []

    for col, direction in sort_order:

        if col not in df.columns:
            if missing_col_action == "raise":
                raise KeyError(f"Missing sorting column: {col}")
            elif missing_col_action == "ignore":
                continue
            else:
                raise ValueError("missing_col_action must be 'raise' or 'ignore'.")

        col_numeric = numeric_col(col)

        all_zero_or_none = col_numeric.eq(0).all()

        if drop_all_zero_or_none_cols and all_zero_or_none:
            df = df.drop(columns=[col])
            continue

        active_sort_cols.append(col)
        ascending_values.append(normalize_sort_direction(direction))

    if not active_sort_cols:
        if reset_index:
            return df.reset_index(drop=True)
        return df

    temp_sort_cols = []

    for col in active_sort_cols:
        temp_col = f"__sort_{col}"

        while temp_col in df.columns:
            temp_col += "_"

        df[temp_col] = numeric_col(col)
        temp_sort_cols.append(temp_col)

    df = df.sort_values(
        by=temp_sort_cols,
        ascending=ascending_values,
        kind="mergesort",
    )

    df = df.drop(columns=temp_sort_cols)

    if reset_index:
        df = df.reset_index(drop=True)

    return df