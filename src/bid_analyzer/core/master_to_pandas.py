import pandas as pd
from datetime import date, datetime, timedelta
from collections import defaultdict
import math
import re

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
    def append_separate_piece(parts_by_date, d, piece):
        """
        Adds a stand-alone calendar item with a space between it
        and any other content on the same date.
        """
        if d is None or not piece:
            return

        piece = str(piece).strip()

        if not piece:
            return

        if parts_by_date[d] == "*":
            parts_by_date[d] = ""

        if parts_by_date[d] and not parts_by_date[d].endswith(" "):
            parts_by_date[d] += " "

        parts_by_date[d] += piece

    def render_trip(assignment):
        trip_id = assignment.get("trip_id")
        flights = assignment.get("flights") or []
        parts_by_date = defaultdict(str)

        previous_arrival = None
        previous_end_date = None
        previous_rest = None
        previous_was_code = False

        def get_flight_code(flight):
            code = flight.get("code")

            if code is None:
                return None

            code = str(code).strip()

            if not code or code.lower() in {"none", "nan"}:
                return None

            return code

        coded_only_assignment = (
            bool(flights)
            and all(
                get_flight_code(flight) is not None
                for flight in flights
            )
        )
        # Only normal flights receive the trip-opening and closing braces.
        normal_flight_indexes = [
            index
            for index, flight in enumerate(flights)
            if get_flight_code(flight) is None
        ]

        first_normal_index = (
            normal_flight_indexes[0]
            if normal_flight_indexes
            else None
        )

        last_normal_index = (
            normal_flight_indexes[-1]
            if normal_flight_indexes
            else None
        )

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

            flight_code = get_flight_code(flight)
            rest = flight.get("rest")

            # -------------------------------------------------------------
            # SBA, SBG, or any other coded flight
            # -------------------------------------------------------------
            if flight_code is not None:
                formatted_rest = (
                    format_rest(rest)
                    if rest not in (None, "")
                    else ""
                )

                if formatted_rest:
                    # Example: SBG3@[DFW16]
                    code_text = (
                        f"{flight_code}"
                        f"@[{arr or dep}{formatted_rest}]"
                    )
                else:
                    # Example: SBA@SDF
                    code_text = (
                        f"{flight_code}"
                        f"@{dep or arr}"
                    )

                # Standalone coded assignments receive their own trip braces.
                # Example: {1964 SBA@SDF}
                if coded_only_assignment and trip_id is not None:
                    code_piece = f"{{{trip_id} {code_text}}}"
                else:
                    code_piece = code_text

                append_separate_piece(
                    parts_by_date,
                    start_date,
                    code_piece,
                )

                previous_arrival = arr
                previous_end_date = end_date
                previous_rest = rest
                previous_was_code = True

                continue

            # -------------------------------------------------------------
            # Normal flight
            # -------------------------------------------------------------
            route_flags = format_route_flags(
                flight.get("route_flags")
            )

            is_first_normal_flight = index == first_normal_index
            is_last_normal_flight = index == last_normal_index

            trip_open = (
                f"{{{trip_id}"
                if is_first_normal_flight
                else ""
            )

            arrival = arrival_text(
                arr,
                rest=rest,
                close_trip=is_last_normal_flight,
            )

            can_compress_departure = (
                not is_first_normal_flight
                and not previous_was_code
                and previous_arrival == dep
                and previous_end_date == start_date
                and previous_rest is None
                and parts_by_date[start_date]
            )

            if start_date == end_date:
                if can_compress_departure:
                    piece = f"-{route_flags}{arrival}"
                else:
                    piece = (
                        f"{trip_open}"
                        f"{dep}-"
                        f"{route_flags}"
                        f"{arrival}"
                    )

                append_piece(
                    parts_by_date,
                    start_date,
                    piece,
                )

            else:
                if can_compress_departure:
                    departure_piece = "-"
                else:
                    departure_piece = f"{trip_open}{dep}-"

                arrival_piece = f"{route_flags}{arrival}"

                append_piece(
                    parts_by_date,
                    start_date,
                    departure_piece,
                )

                append_piece(
                    parts_by_date,
                    end_date,
                    arrival_piece,
                )

            # Add a dash on completely empty dates between normal flights.
            if index < len(flights) - 1:
                next_flight = flights[index + 1]
                next_start = to_date(
                    next_flight.get("start_date")
                )

                if next_start is not None:
                    gap_start = end_date + timedelta(days=1)
                    gap_end = next_start - timedelta(days=1)

                    for gap_day in date_range_strict(
                        gap_start,
                        gap_end,
                    ):
                        append_piece(
                            parts_by_date,
                            gap_day,
                            "*",
                        )

            previous_arrival = arr
            previous_end_date = end_date
            previous_rest = rest
            previous_was_code = False

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
    bid_period_info,
    *,
    date_col_format="%Y-%m-%d",
    start_bid_off = False,
    end_bid_off = False,
):

    bid_start = bid_period_info["bid_period_date_range"]["start"]
    bid_end = bid_period_info["bid_period_date_range"]["end"]
    bid_dates = _date_range_inclusive(bid_start, bid_end)

    include_start_bid_off = start_bid_off or any(
        "bid_start_days_off" in line_data
        for line_data in master_lines.values()
    )

    include_end_bid_off = end_bid_off or any(
        "bid_end_days_off" in line_data
        for line_data in master_lines.values()
    )

    rows = []

    for line_number, line_data in master_lines.items():

        date_values = build_line_calendar_values(
            line_data,
            bid_dates,
            off_value="",
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
            "Line Type": line_data.get("line_type_preference_score",0),
            "Training in Days On": line_data.get("training_fit_score", 0),
            "On/Off Blocks": line_data.get("blockiness_score", 0),  #Blockiness name alts: Work/Off Continuity, Block Quality, Clean blocks, Choppiness Avoidance
            "Total DO": line_data.get("tot_DO", 0),
            "Total CT": line_data.get("tot_CT",0),
            "Avg # of Legs": line_data.get("avg_legs_per_work_day",math.nan),
            "Avg CT": line_data.get("avg_CT", 0),
            "Avg DT": line_data.get("avg_DT",0),
            "Avg Rest": line_data.get("avg_rest",0),
            "Avg TAFB": line_data.get("avg_tafb",0),
            r"% of Tickets Paid": line_data.get("pct_company_tickets", 0),
            "% International Destinations": line_data.get("pct_dest_int",math.nan),
            "% Asia Destinations": line_data.get("pct_dest_AS",math.nan),
            r"% Europoean Destinations": line_data.get("pct_dest_EU",math.nan),
            "% South American Destinations": line_data.get("pct_dest_SA",math.nan),
            "% Weekends Off": line_data.get("pct_weekends_off",0),
            r"% of Days Off Requested": line_data.get("pct_requested_days_off",0),
            "Pay": line_data.get("tot_pay",0),
            "Tax-Free Pay":line_data.get("pay_per_diem")
        })

        if include_start_bid_off:
            row.update({"Start bid off":line_data.get("bid_start_days_off",0)})
        
        if include_end_bid_off:
            row.update({"End bid off":line_data.get("bid_end_days_off",0)})

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

#--------------------------------------------------------------------------------------------------------------------
#New Sorting
#New sorting helper functions
def normalize_sort_direction(direction):
    """
    Returns True for ascending, False for descending.
    """

    if isinstance(direction, bool):
        return direction

    direction = str(direction).lower().strip()

    ascending_words = {
        "asc",
        "ascending",
        "low_to_high",
        "small_to_large",
        "smallest_to_largest",
        "lower_is_better",
    }

    descending_words = {
        "desc",
        "descending",
        "high_to_low",
        "large_to_small",
        "largest_to_smallest",
        "higher_is_better",
    }

    if direction in ascending_words:
        return True

    if direction in descending_words:
        return False

    raise ValueError(
        f"Invalid sort direction: {direction}. "
        "Use 'asc', 'desc', 'low_to_high', or 'high_to_low'."
    )
def normalize_sort_mode(mode):
    """
    Returns:
        ("strict", None)
        ("weighted", "new")
        ("weighted", "equal")
    """

    mode = str(mode).lower().strip()

    strict_words = {
        "strict",
        "fixed",
        "priority",
        "tie_breaker",
        "tiebreaker",
        "tie-breaker",
    }

    weighted_words = {
        "weighted",
        "flexible",
        "score",
        "combined",
        "combined_score",
        "new_weight",
        "new_level",
    }

    equal_words = {
        "equal",
        "same",
        "same_weight",
        "same_level",
    }

    if mode in strict_words:
        return "strict", None

    if mode in weighted_words:
        return "weighted", "new"

    if mode in equal_words:
        return "weighted", "equal"

    raise ValueError(
        f"Invalid sort mode: {mode}. "
        "Use 'strict', 'weighted', or 'equal'."
    )
def normalize_sort_conditions(sort_conditions, *, default_mode="strict"):
    """
    Normalizes sort conditions into dictionaries.

    Accepted formats:

        ("Training", "high_to_low")

        ("Training", "high_to_low", "weighted")

        {
            "column": "Training",
            "direction": "high_to_low",
            "mode": "weighted",
        }
    """

    normalized = []

    for item in sort_conditions:

        if isinstance(item, dict):
            col = item.get("column")
            direction = item.get("direction", item.get("order", "desc"))
            mode = item.get("mode", default_mode)

        else:
            if len(item) == 2:
                col, direction = item
                mode = default_mode

            elif len(item) == 3:
                col, direction, mode = item

            else:
                raise ValueError(
                    "Each sort condition must be one of these formats:\n"
                    "    (column, direction)\n"
                    "    (column, direction, mode)\n"
                    f"Got: {item}"
                )

        if col is None:
            raise ValueError(f"Sort condition is missing a column name: {item}")

        normalized_mode, weight_role = normalize_sort_mode(mode)

        normalized.append({
            "column": col,
            "ascending": normalize_sort_direction(direction),
            "direction_text": str(direction),
            "mode": normalized_mode,
            "weight_role": weight_role,
            "original_mode": str(mode),
        })

    return normalized
def is_date_column_name(col,*,date_col_format="%Y-%m-%d",extra_date_formats=None):
    """
    Returns True if the column name looks like a date.

    Date columns should never be dropped.
    """

    if extra_date_formats is None:
        extra_date_formats = []

    if isinstance(col, (datetime, date, pd.Timestamp)):
        return True

    text = str(col).strip()

    formats_to_try = [
        date_col_format,
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%Y/%m/%d",
        "%Y%m%d",
    ]

    no_year_formats = [
        "%b %d",
        "%a %b %d",
        "%a, %b %d",
        "%m/%d",
    ]

    formats_to_try.extend(extra_date_formats)

    for fmt in formats_to_try:
        try:
            datetime.strptime(text, fmt)
            return True
        except ValueError:
            pass

    # Avoid deprecation warning by injecting a leap year.
    for fmt in no_year_formats:
        try:
            datetime.strptime(f"2000 {text}", f"%Y {fmt}")
            return True
        except ValueError:
            pass

    return False
def numeric_series(df, col, *, fill_missing=False):
    """
    Converts a DataFrame column to numeric.

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

    result = pd.to_numeric(cleaned, errors="coerce")

    if fill_missing:
        result = result.fillna(0)

    return result
def column_is_all_empty_or_zero(df, col):
    """
    Returns True if every value in the column is one of:

        0
        0.0
        "0"
        "0.0"
        "$0"
        "0%"
        None
        NaN
        ""
        "nan"
        "None"

    Important:
        Non-empty text values like "SBA", "66:00", "RFD", "0h00"
        are treated as real data and will prevent the column from being dropped.
    """

    s = df[col]

    text = s.astype(str).str.strip()

    empty_mask = (
        s.isna()
        | text.str.lower().isin({
            "",
            "none",
            "nan",
            "nat",
            "null",
            "<na>",
        })
    )

    numeric_text = (
        text
        .str.replace("%", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )

    numeric_values = pd.to_numeric(numeric_text, errors="coerce")

    zero_mask = numeric_values.eq(0)

    # Non-empty, non-numeric text means the column has real data.
    # Examples:
    #   "SBA"
    #   "66:00"
    #   "RFD"
    #   "0h00"
    non_empty_text_mask = ~empty_mask & numeric_values.isna()

    if non_empty_text_mask.any():
        return False

    return (empty_mask | zero_mask).all()
def make_unique_temp_col(df, base_name):
    temp_col = base_name

    while temp_col in df.columns:
        temp_col += "_"

    return temp_col
def remove_existing_sort_prefix(col):
    """
    Turns:
        '1. Training'
    back into:
        'Training'

    This prevents repeated renaming like:
        '1. 1. Training'
    """

    return re.sub(r"^\d+\.\s+", "", str(col)).strip()

#Sort percentage based off sort order
def get_sort_percent_contributions(
    sort_order,
    *,
    weighting_style="soft",
    soft_max_weight=3.0,
    soft_min_weight=1.0,
    round_digits=2,
):
    """
    Returns a list of percentage contributions for each sort criteria.

    Example
    -------
    sort_order = [
        ("Extra Vacation Days", "high_to_low", "strict"),
        ("Training", "high_to_low", "weighted"),
        ("Blockiness", "high_to_low", "weighted"),
        ("Total DO", "high_to_low", "weighted"),
        ("Premium", "high_to_low", "weighted"),
        ("% tickets paid", "high_to_low", "weighted"),
    ]

    returns something like:

        [nan, 30.0, 25.0, 20.0, 15.0, 10.0]

    Rules
    -----
    strict:
        Not part of weighted score.
        Returns math.nan.
        Resets the current weighted group.

    weighted:
        Starts a new weight level.

    equal:
        Uses the same weight level as the previous weighted/equal item.

    weighting_style:
        "hard":
            Stronger weighting by position.
            Example with 5 levels:
                5, 4, 3, 2, 1

        "soft":
            Softer weighting by position.
            Example with 5 levels and defaults:
                3.0, 2.5, 2.0, 1.5, 1.0

        "equal":
            Every weighted item has the same contribution.
    """

    def normalize_mode(mode):
        mode = str(mode).lower().strip()

        strict_words = {
            "strict",
            "hard_priority",
            "hard priority",
            "priority",
            "tie_breaker",
            "tiebreaker",
            "tie-breaker",
        }

        weighted_words = {
            "weighted",
            "new_weight",
            "new weight",
            "new_weighted_factor",
            "new weighted factor",
            "new_level",
            "new level",
        }

        equal_words = {
            "equal",
            "same",
            "same_weight",
            "same weight",
            "same_weight_as_above",
            "same weight as above",
            "same_level",
            "same level",
        }

        if mode in strict_words:
            return "strict"

        if mode in weighted_words:
            return "weighted"

        if mode in equal_words:
            return "equal"

        raise ValueError(
            f"Invalid sort mode: {mode}. "
            "Use 'strict', 'weighted', or 'equal'."
        )

    def get_mode(item):
        if isinstance(item, dict):
            return item.get("mode", "strict")

        if len(item) >= 3:
            return item[2]

        return "strict"

    def make_level_weights(max_level):
        style = str(weighting_style).lower().strip()

        if style in {"equal", "flat", "none"}:
            return {
                level: 1.0
                for level in range(1, max_level + 1)
            }

        if style in {"hard", "auto_hard"}:
            return {
                level: float(max_level - level + 1)
                for level in range(1, max_level + 1)
            }

        if style in {"soft", "auto_soft"}:
            if max_level == 1:
                return {1: 1.0}

            weight_range = soft_max_weight - soft_min_weight

            return {
                level: float(
                    soft_max_weight
                    - weight_range * (level - 1) / (max_level - 1)
                )
                for level in range(1, max_level + 1)
            }

        raise ValueError(
            f"Invalid weighting_style: {weighting_style}. "
            "Use 'soft', 'hard', or 'equal'."
        )

    def finalize_group(group_indexes, group_levels, result):
        """
        Converts one weighted group into percent contributions.
        """

        if not group_indexes:
            return

        max_level = max(group_levels)
        level_weights = make_level_weights(max_level)

        raw_weights = [
            level_weights[level]
            for level in group_levels
        ]

        total_weight = sum(raw_weights)

        for index, raw_weight in zip(group_indexes, raw_weights):
            percent = raw_weight / total_weight * 100

            if round_digits is not None:
                percent = round(percent, round_digits)

            result[index] = percent

    result = [math.nan] * len(sort_order)

    group_indexes = []
    group_levels = []
    current_level = 0

    for index, item in enumerate(sort_order):
        mode = normalize_mode(get_mode(item))

        if mode == "strict":
            finalize_group(group_indexes, group_levels, result)

            group_indexes = []
            group_levels = []
            current_level = 0

            result[index] = math.nan

        elif mode == "weighted":
            current_level += 1

            group_indexes.append(index)
            group_levels.append(current_level)

        elif mode == "equal":
            if current_level == 0:
                current_level = 1

            group_indexes.append(index)
            group_levels.append(current_level)

    finalize_group(group_indexes, group_levels, result)

    return result

#Drop empty columns
def drop_empty_sort_columns(
    df,
    *,
    columns_to_check=None,
    sort_conditions=None,
    always_check_cols=("Extra Vacation Days", "Training"),
    never_drop_cols=("Line Number",),
    date_col_format="%Y-%m-%d",
    extra_date_formats=None,
    default_mode="strict",
    check_all_columns=False,
    return_dropped=False,
):
    """
    Drops columns where every value is 0, 0.0, None, blank, or NaN.

    Important:
        - Date columns are NEVER dropped.
        - Line Number is NEVER dropped by default.
        - Non-numeric text values are treated as real data.
        - If check_all_columns=True, every DataFrame column is checked.
    """

    df = df.copy()

    cols_to_check = []

    if check_all_columns:
        cols_to_check.extend(df.columns)

    elif columns_to_check == "all":
        cols_to_check.extend(df.columns)

    else:
        if columns_to_check:
            cols_to_check.extend(columns_to_check)

        if always_check_cols:
            cols_to_check.extend(always_check_cols)

        if sort_conditions:
            normalized = normalize_sort_conditions(
                sort_conditions,
                default_mode=default_mode,
            )

            cols_to_check.extend([
                rule["column"]
                for rule in normalized
            ])

    # Remove duplicates while preserving order
    seen = set()
    cols_to_check = [
        col for col in cols_to_check
        if not (col in seen or seen.add(col))
    ]

    dropped_columns = []

    for col in cols_to_check:

        if col not in df.columns:
            continue

        if col in never_drop_cols:
            continue

        # Absolute rule:
        # date columns are never deleted.
        if is_date_column_name(
            col,
            date_col_format=date_col_format,
            extra_date_formats=extra_date_formats,
        ):
            continue

        if column_is_all_empty_or_zero(df, col):
            df = df.drop(columns=[col])
            dropped_columns.append(col)

    if return_dropped:
        return df, dropped_columns

    return df

#New Sort by condition
def sort_dataframe_by_conditions(
    df,
    sort_conditions,
    *,
    fixed_start_cols=("Line Number", "Extra Vacation Days"),
    date_col_format="%Y-%m-%d",
    extra_date_formats=None,

    # Sorting behavior
    default_mode="strict",
    weighting_style="soft",
    soft_max_weight=3.0,
    soft_min_weight=1.0,
    missing_score=0.0,
    score_round_digits=None,

    # Output behavior
    reset_index=True,
    missing_col_action="ignore",
    reorder_columns=True,
    add_sort_numbers=True,
    strip_existing_sort_prefixes=True,
    return_sort_details=False,
):
    """
    Sorts, reorders columns, and optionally adds sort numbers to sorted columns.

    This function does NOT drop empty columns.
    Use drop_empty_sort_columns() before this function.

    Supported sort condition formats:

        ("Training", "high_to_low")

        ("Training", "high_to_low", "weighted")

        ("Blockiness", "high_to_low", "equal")

        {
            "column": "Training",
            "direction": "high_to_low",
            "mode": "weighted",
        }

    Modes
    -----
    strict:
        Hard priority / tie-breaker sorting.

    weighted:
        Starts a new weighted level.

    equal:
        Same weighted level as the previous weighted/equal item.

    Example
    -------
    sort_conditions = [
        ("Extra Vacation Days", "high_to_low", "strict"),
        ("Training", "high_to_low", "weighted"),
        ("Blockiness", "high_to_low", "weighted"),
        ("Total DO", "high_to_low", "equal"),
        ("Premium", "high_to_low", "weighted"),
        ("% tickets paid", "high_to_low", "weighted"),
    ]

    Numbered columns become:
        1. Extra Vacation Days
        2. Training
        3. Blockiness
        3. Total DO
        4. Premium
        5. % tickets paid
    """

    df = df.copy()

    if extra_date_formats is None:
        extra_date_formats = []

    # ------------------------------------------------------------
    # Helper functions
    # ------------------------------------------------------------

    def remove_existing_sort_prefix(col):
        """
        Turns:
            '1. Training'
        back into:
            'Training'
        """

        return re.sub(r"^\d+\.\s+", "", str(col)).strip()

    def normalize_sort_direction(direction):
        """
        Returns True for ascending, False for descending.
        """

        if isinstance(direction, bool):
            return direction

        direction = str(direction).lower().strip()

        ascending_words = {
            "asc",
            "ascending",
            "low_to_high",
            "small_to_large",
            "smallest_to_largest",
            "lower_is_better",
        }

        descending_words = {
            "desc",
            "descending",
            "high_to_low",
            "large_to_small",
            "largest_to_smallest",
            "higher_is_better",
        }

        if direction in ascending_words:
            return True

        if direction in descending_words:
            return False

        raise ValueError(
            f"Invalid sort direction: {direction}. "
            "Use 'asc', 'desc', 'low_to_high', or 'high_to_low'."
        )

    def normalize_sort_mode(mode):
        """
        Returns:
            ("strict", None)
            ("weighted", "new")
            ("weighted", "equal")
        """

        mode = str(mode).lower().strip()

        strict_words = {
            "strict",
            "priority",
            "hard_priority",
            "hard priority",
            "fixed",
            "tie_breaker",
            "tiebreaker",
            "tie-breaker",
        }

        weighted_words = {
            "weighted",
            "new_weight",
            "new weight",
            "new_weighted_factor",
            "new weighted factor",
            "new_level",
            "new level",
            "score",
            "combined",
            "combined_score",
        }

        equal_words = {
            "equal",
            "same",
            "same_weight",
            "same weight",
            "same_weight_as_above",
            "same weight as above",
            "same_level",
            "same level",
        }

        if mode in strict_words:
            return "strict", None

        if mode in weighted_words:
            return "weighted", "new"

        if mode in equal_words:
            return "weighted", "equal"

        raise ValueError(
            f"Invalid sort mode: {mode}. "
            "Use 'strict', 'weighted', or 'equal'."
        )

    def normalize_sort_conditions(sort_conditions):
        """
        Converts tuple/list/dict conditions into normalized dictionaries.
        """

        normalized = []

        for item in sort_conditions:

            if isinstance(item, dict):
                col = item.get("column")
                direction = item.get("direction", item.get("order", "desc"))
                mode = item.get("mode", default_mode)

            else:
                if len(item) == 2:
                    col, direction = item
                    mode = default_mode

                elif len(item) == 3:
                    col, direction, mode = item

                else:
                    raise ValueError(
                        "Each sort condition must be one of these formats:\n"
                        "    (column, direction)\n"
                        "    (column, direction, mode)\n"
                        f"Got: {item}"
                    )

            if col is None:
                raise ValueError(f"Sort condition is missing a column name: {item}")

            normalized_mode, weight_role = normalize_sort_mode(mode)

            normalized.append({
                "column": col,
                "ascending": normalize_sort_direction(direction),
                "mode": normalized_mode,
                "weight_role": weight_role,
                "original_mode": str(mode),
            })

        return normalized

    def is_date_column_name(col):
        """
        Returns True if the column name looks like a date.

        Date/calendar columns should not be renamed with sort numbers.
        """

        if isinstance(col, (datetime, date, pd.Timestamp)):
            return True

        text = str(col).strip()

        formats_to_try = [
            date_col_format,
            "%Y-%m-%d",
            "%m/%d/%Y",
            "%m/%d/%y",
            "%Y/%m/%d",
            "%Y%m%d",
        ]

        no_year_formats = [
            "%b %d",
            "%a %b %d",
            "%a, %b %d",
            "%m/%d",
        ]

        formats_to_try.extend(extra_date_formats)

        for fmt in formats_to_try:
            try:
                datetime.strptime(text, fmt)
                return True
            except ValueError:
                pass

        # Avoid Python's warning about parsing dates without a year.
        # Use 2000 because it is a leap year.
        for fmt in no_year_formats:
            try:
                datetime.strptime(f"2000 {text}", f"%Y {fmt}")
                return True
            except ValueError:
                pass

        return False

    def numeric_series(col, *, fill_missing=False):
        """
        Converts a column to numeric for sorting/scoring.

        Handles:
            10
            "10"
            "10%"
            "$10"
            ""
            None
            NaN
        """

        cleaned = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace("$", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
            .replace({
                "": None,
                "None": None,
                "nan": None,
                "NaN": None,
                "NaT": None,
                "<NA>": None,
            })
        )

        result = pd.to_numeric(cleaned, errors="coerce")

        if fill_missing:
            result = result.fillna(0)

        return result

    def make_unique_temp_col(base_name):
        temp_col = base_name

        while temp_col in df.columns:
            temp_col += "_"

        return temp_col

    def make_weight_by_level(max_level):
        """
        Creates raw weights by weighted level.
        """

        style = str(weighting_style).lower().strip()

        if style in {"equal", "none", "flat"}:
            return {
                level: 1.0
                for level in range(1, max_level + 1)
            }

        if style in {"hard", "auto", "auto_hard"}:
            return {
                level: float(max_level - level + 1)
                for level in range(1, max_level + 1)
            }

        if style in {"soft", "auto_soft"}:
            if max_level == 1:
                return {1: 1.0}

            weight_range = soft_max_weight - soft_min_weight

            return {
                level: float(
                    soft_max_weight
                    - weight_range * (level - 1) / (max_level - 1)
                )
                for level in range(1, max_level + 1)
            }

        raise ValueError(
            f"Invalid weighting_style: {weighting_style}. "
            "Use 'soft', 'hard', or 'equal'."
        )

    def get_group_weights_from_rules(weighted_rules):
        """
        Creates automatic weights for one weighted group.

        weighted = new level
        equal    = same level as previous weighted/equal item
        """

        current_level = 0
        level_by_column = {}

        for rule in weighted_rules:
            col = rule["column"]

            if rule["weight_role"] == "new":
                current_level += 1

            elif rule["weight_role"] == "equal":
                if current_level == 0:
                    current_level = 1

            level_by_column[col] = current_level

        max_level = max(level_by_column.values())

        weight_by_level = make_weight_by_level(max_level)

        weight_by_column = {
            col: weight_by_level[level]
            for col, level in level_by_column.items()
        }

        return weight_by_column, level_by_column

    def split_rules_into_sort_stages(active_rules):
        """
        Consecutive weighted/equal rules are combined into one weighted group.
        Strict rules become their own sort stage.
        """

        stages = []
        weighted_group = []

        for rule in active_rules:

            if rule["mode"] == "weighted":
                weighted_group.append(rule)

            elif rule["mode"] == "strict":
                if weighted_group:
                    stages.append({
                        "type": "weighted_group",
                        "rules": weighted_group,
                    })
                    weighted_group = []

                stages.append({
                    "type": "strict",
                    "rules": [rule],
                })

        if weighted_group:
            stages.append({
                "type": "weighted_group",
                "rules": weighted_group,
            })

        return stages

    def build_sort_number_map(normalized_rules):
        """
        Builds the numbering used for column renaming.

        strict:
            gets a new number.

        weighted:
            gets a new number.

        equal:
            gets the same number as the previous weighted/equal item.
        """

        sort_number_by_column = {}

        sort_number = 0
        current_weight_number = None

        for rule in normalized_rules:
            col = rule["column"]

            if rule["mode"] == "strict":
                sort_number += 1
                current_weight_number = None
                sort_number_by_column[col] = sort_number

            else:
                if rule["weight_role"] == "new":
                    sort_number += 1
                    current_weight_number = sort_number

                elif rule["weight_role"] == "equal":
                    if current_weight_number is None:
                        sort_number += 1
                        current_weight_number = sort_number

                sort_number_by_column[col] = current_weight_number

        return sort_number_by_column

    # ------------------------------------------------------------
    # 1. Remove existing sort prefixes
    # ------------------------------------------------------------

    if strip_existing_sort_prefixes:
        rename_map = {}
        existing_cols = set(df.columns)

        for col in df.columns:
            base_col = remove_existing_sort_prefix(col)

            # Only rename if it will not collide with an existing column.
            if base_col != col and base_col not in existing_cols:
                rename_map[col] = base_col

        if rename_map:
            df = df.rename(columns=rename_map)

    # ------------------------------------------------------------
    # 2. Normalize and validate sort conditions
    # ------------------------------------------------------------

    normalized_rules = normalize_sort_conditions(sort_conditions)

    active_rules = []
    missing_columns = []

    for rule in normalized_rules:
        col = rule["column"]

        if col not in df.columns:
            if missing_col_action == "raise":
                raise KeyError(f"Missing sorting column: {col}")

            elif missing_col_action == "ignore":
                missing_columns.append(col)
                continue

            else:
                raise ValueError(
                    "missing_col_action must be 'raise' or 'ignore'."
                )

        active_rules.append(rule)

    # ------------------------------------------------------------
    # 3. Build sort stages
    # ------------------------------------------------------------

    stages = split_rules_into_sort_stages(active_rules)

    sort_by_cols = []
    ascending_values = []
    temp_cols_to_drop = []
    stage_details = []

    weighted_group_count = 0

    for stage in stages:

        if stage["type"] == "strict":
            rule = stage["rules"][0]
            col = rule["column"]

            temp_col = make_unique_temp_col(f"__sort_{col}")
            df[temp_col] = numeric_series(col, fill_missing=True)

            sort_by_cols.append(temp_col)
            ascending_values.append(rule["ascending"])
            temp_cols_to_drop.append(temp_col)

            stage_details.append({
                "type": "strict",
                "column": col,
                "ascending": rule["ascending"],
                "temp_column": temp_col,
            })

        elif stage["type"] == "weighted_group":
            weighted_group_count += 1
            group_rules = stage["rules"]

            score_col = make_unique_temp_col(
                f"__weighted_score_{weighted_group_count}"
            )

            group_weights, group_levels = get_group_weights_from_rules(group_rules)

            total_weight = sum(
                group_weights[rule["column"]]
                for rule in group_rules
                if group_weights[rule["column"]] > 0
            )

            if total_weight == 0:
                raise ValueError(
                    "A weighted group has total weight 0. "
                    "At least one weighted column must have a positive weight."
                )

            combined_score = pd.Series(0.0, index=df.index)

            for rule in group_rules:
                col = rule["column"]
                weight = group_weights[col]

                if weight == 0:
                    continue

                # high_to_low / desc:
                #     larger raw value gets a higher percentile.
                #
                # low_to_high / asc:
                #     smaller raw value gets a higher percentile.
                rank_ascending = not rule["ascending"]

                rank_score = numeric_series(
                    col,
                    fill_missing=False,
                ).rank(
                    method="average",
                    pct=True,
                    ascending=rank_ascending,
                ).fillna(missing_score)

                combined_score += rank_score * weight

            df[score_col] = combined_score / total_weight

            if score_round_digits is not None:
                df[score_col] = df[score_col].round(score_round_digits)

            sort_by_cols.append(score_col)
            ascending_values.append(False)
            temp_cols_to_drop.append(score_col)

            stage_details.append({
                "type": "weighted_group",
                "columns": [rule["column"] for rule in group_rules],
                "score_column": score_col,
                "weights": group_weights,
                "levels": group_levels,
            })

    # ------------------------------------------------------------
    # 4. Sort
    # ------------------------------------------------------------

    if sort_by_cols:
        df = df.sort_values(
            by=sort_by_cols,
            ascending=ascending_values,
            kind="mergesort",
        )

    df = df.drop(
        columns=[
            col for col in temp_cols_to_drop
            if col in df.columns
        ]
    )

    if reset_index:
        df = df.reset_index(drop=True)

    # ------------------------------------------------------------
    # 5. Reorder columns
    # ------------------------------------------------------------

    if reorder_columns:
        current_cols = list(df.columns)

        fixed_cols = [
            col for col in fixed_start_cols
            if col in current_cols
        ]

        calendar_cols = [
            col for col in current_cols
            if (
                col not in fixed_cols
                and is_date_column_name(col)
            )
        ]

        sort_order_cols = []

        for rule in active_rules:
            col = rule["column"]

            if (
                col in df.columns
                and col not in fixed_cols
                and col not in calendar_cols
                and col not in sort_order_cols
            ):
                sort_order_cols.append(col)

        remaining_cols = [
            col for col in current_cols
            if (
                col not in fixed_cols
                and col not in calendar_cols
                and col not in sort_order_cols
            )
        ]

        final_col_order = (
            fixed_cols
            + calendar_cols
            + sort_order_cols
            + remaining_cols
        )

        df = df[final_col_order]

    # ------------------------------------------------------------
    # 6. Add sort numbers to relevant columns
    # ------------------------------------------------------------

    if add_sort_numbers:
        # Use only columns that still exist and are actually part of the active sort.
        # This prevents dropped columns from consuming sort numbers.
        active_numbering_rules = [
            rule for rule in active_rules
            if rule["column"] in df.columns
        ]

        sort_number_by_column = build_sort_number_map(active_numbering_rules)

        rename_map = {}

        for col in df.columns:
            base_col = remove_existing_sort_prefix(col)

            if base_col == "Line Number":
                continue

            if is_date_column_name(base_col):
                continue

            if base_col in sort_number_by_column:
                rename_map[col] = f"{sort_number_by_column[base_col]}. {base_col}"

        if rename_map:
            df = df.rename(columns=rename_map)

    # ------------------------------------------------------------
    # 7. Return
    # ------------------------------------------------------------

    if return_sort_details:
        details = {
            "stages": stage_details,
            "sort_by_columns": sort_by_cols,
            "ascending_values": ascending_values,
            "missing_columns": missing_columns,
        }

        return df, details

    return df