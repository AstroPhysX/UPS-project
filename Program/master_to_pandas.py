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
            "Training": line_data.get("training_fit_score", 0),
            "Blockiness": line_data.get("blockiness_score", 0),
            "Total DO": line_data.get("tot_DO", 0),
            "% tickets paid": line_data.get("company_ticket_pct", 0),
            "Avg # of legs": line_data.get("avg_legs_per_work_day",0),
            "Total CT": line_data.get("tot_CT",0),
            "Premium": int(line_data.get("tot_Premium", line_data.get("tot_premium", 0))),
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


def sort_dataframe_by_conditions(
    df,
    sort_conditions,
    *,
    fixed_start_cols=("Line Number", "Extra Vacation Days"),
    date_col_format="%Y-%m-%d",
    drop_all_zero_or_none_cols=True,
    drop_fixed_cols_if_all_zero=True,
    reset_index=True,
    missing_col_action="raise",

    # New options
    default_mode="strict",
    weighting_style="soft",
    weights=None,
    soft_max_weight=3.0,
    soft_min_weight=1.0,
    missing_score=0.0,
    keep_score_columns=False,
    score_col_prefix="sort score",
    score_round_digits=None,
    return_sort_details=False,
):
    """
    Sorts, scores, and reorders a DataFrame using strict and/or weighted
    sorting conditions.

    sort_conditions accepts either:

        Old/simple format:
            [
                ("Training", "desc"),
                ("Blockiness", "desc"),
                ("Avg # of legs", "asc"),
            ]

        Hybrid format:
            [
                ("Extra Vacation Days", "high_to_low", "strict"),
                ("Training", "high_to_low", "weighted"),
                ("Blockiness", "high_to_low", "weighted"),
                ("Total DO", "high_to_low", "strict"),
                ("% tickets paid", "high_to_low", "weighted"),
            ]

    Modes:
        "strict":
            Used as a normal priority / tie-breaker sort.

        "weighted":
            Consecutive weighted conditions are blended into a percentile-rank
            combined score.

    weighting_style:
        "equal":
            Every weighted item in a group gets weight 1.

        "hard":
            Position-based weights within each weighted group.
            Example with 4 weighted items:
                [4, 3, 2, 1]

        "soft":
            Softer position-based weights within each weighted group.
            Example with 4 weighted items and defaults:
                [3.0, 2.33, 1.67, 1.0]

    weights:
        Optional manual weight overrides by column name.

        Example:
            {
                "Blockiness": 2,
                "Training": 1,
            }

        Manual weights override the automatic/equal weight for that column.

    default_mode:
        Used when a condition only has two items: (column, direction).

        Default is "strict", so old sort_orders behave like your original
        tie-breaker sorting function.

        Use default_mode="weighted" if you want all 2-item conditions to be
        treated as weighted.

    score_round_digits:
        If not None, weighted score columns are rounded before sorting.

        This can make later strict/tie-breaker stages matter more.
        Example:
            score_round_digits=3

    keep_score_columns:
        If True, keeps the generated combined score columns in the DataFrame.
        If False, uses them for sorting and then removes them.

    return_sort_details:
        If True, returns:
            df, details

        instead of just:
            df
    """

    df = df.copy()

    if weights is None:
        weights = {}
    elif not isinstance(weights, dict):
        raise ValueError("weights must be None or a dictionary.")

    # ------------------------------------------------------------
    # Helper functions
    # ------------------------------------------------------------

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
        Returns either 'strict' or 'weighted'.
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
        }

        if mode in strict_words:
            return "strict"

        if mode in weighted_words:
            return "weighted"

        raise ValueError(
            f"Invalid sort mode: {mode}. "
            "Use 'strict' or 'weighted'."
        )

    def is_calendar_col(col):
        """
        Returns True if the column name looks like a calendar date column.
        """

        try:
            datetime.strptime(str(col), date_col_format)
            return True
        except ValueError:
            return False

    def numeric_series(col, *, fill_missing=False):
        """
        Converts a column to numeric.

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

    def make_unique_temp_col(base_name):
        """
        Makes a temporary column name that does not already exist.
        """

        temp_col = base_name

        while temp_col in df.columns:
            temp_col += "_"

        return temp_col

    def normalize_condition(item):
        """
        Converts tuple/list/dict sort condition into a standard dictionary.
        """

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
                    "Each sort condition must be either "
                    "(column, direction) or (column, direction, mode). "
                    f"Got: {item}"
                )

        if col is None:
            raise ValueError(f"Sort condition is missing a column name: {item}")

        return {
            "column": col,
            "ascending": normalize_sort_direction(direction),
            "mode": normalize_sort_mode(mode),
        }

    def make_group_weights(weighted_group):
        """
        Creates weights for one consecutive weighted group.
        Manual weights override automatic/equal weights.
        """

        columns = [rule["column"] for rule in weighted_group]
        n = len(columns)

        style = str(weighting_style).lower().strip()

        if style in {"equal", "none", "flat"}:
            base_weights = {
                col: 1.0
                for col in columns
            }

        elif style in {"hard", "auto", "auto_hard"}:
            base_weights = {
                col: float(n - index)
                for index, col in enumerate(columns)
            }

        elif style in {"soft", "auto_soft"}:
            if n == 1:
                base_weights = {
                    columns[0]: 1.0
                }
            else:
                weight_range = soft_max_weight - soft_min_weight

                base_weights = {
                    col: float(
                        soft_max_weight
                        - (weight_range * index / (n - 1))
                    )
                    for index, col in enumerate(columns)
                }

        else:
            raise ValueError(
                f"Invalid weighting_style: {weighting_style}. "
                "Use 'equal', 'hard', or 'soft'."
            )

        # Manual weights override automatic/equal weights.
        for col in columns:
            if col in weights:
                base_weights[col] = float(weights[col])

            if base_weights[col] < 0:
                raise ValueError(
                    f"Weight for column '{col}' cannot be negative."
                )

        return base_weights

    def split_into_stages(active_rules):
        """
        Weighted rules are grouped until a strict rule appears.
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

    # ------------------------------------------------------------
    # 1. Normalize, validate, and activate sorting columns
    # ------------------------------------------------------------

    normalized_conditions = [
        normalize_condition(item)
        for item in sort_conditions
    ]

    active_rules = []
    dropped_columns = []
    inactive_all_zero_columns = []
    missing_columns = []

    for rule in normalized_conditions:
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

        col_numeric = numeric_series(col, fill_missing=True)
        all_zero_or_none = col_numeric.eq(0).all()

        col_is_fixed = col in fixed_start_cols or is_calendar_col(col)

        should_drop_col = (
            drop_all_zero_or_none_cols
            and all_zero_or_none
            and (
                not col_is_fixed
                or drop_fixed_cols_if_all_zero
            )
        )

        if should_drop_col:
            df = df.drop(columns=[col])
            dropped_columns.append(col)
            continue

        # Keep inactive all-zero protected columns, but do not use them
        # as active sorting criteria.
        if all_zero_or_none:
            inactive_all_zero_columns.append(col)
            continue

        active_rules.append(rule)

    # ------------------------------------------------------------
    # 2. Build staged sort columns
    # ------------------------------------------------------------

    stages = split_into_stages(active_rules)

    sort_by_cols = []
    ascending_values = []
    temp_cols_to_drop = []
    score_cols_created = []
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
                f"{score_col_prefix}_{weighted_group_count}"
            )

            group_weights = make_group_weights(group_rules)

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

            combined_score = 0

            for rule in group_rules:
                col = rule["column"]
                weight = group_weights[col]

                if weight == 0:
                    continue

                # For score ranking:
                # - high_to_low / desc means bigger raw value is better
                # - low_to_high / asc means smaller raw value is better
                #
                # pandas rank ascending=True gives the largest raw value
                # the highest percentile.
                # Therefore rank_ascending is the opposite of sort ascending.
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

            score_cols_created.append(score_col)

            if not keep_score_columns:
                temp_cols_to_drop.append(score_col)

            stage_details.append({
                "type": "weighted_group",
                "columns": [rule["column"] for rule in group_rules],
                "score_column": score_col,
                "weights": {
                    rule["column"]: group_weights[rule["column"]]
                    for rule in group_rules
                },
            })

    # ------------------------------------------------------------
    # 3. Sort
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
    # 4. Reorder columns
    # ------------------------------------------------------------

    current_cols = list(df.columns)

    fixed_cols = [
        col for col in fixed_start_cols
        if col in current_cols
    ]

    calendar_cols = [
        col for col in current_cols
        if col not in fixed_cols and is_calendar_col(col)
    ]

    sort_order_cols = []

    for rule in normalized_conditions:
        col = rule["column"]

        if (
            col in df.columns
            and col not in fixed_cols
            and col not in calendar_cols
            and col not in sort_order_cols
        ):
            sort_order_cols.append(col)

    visible_score_cols = [
        col for col in score_cols_created
        if col in df.columns
    ]

    remaining_cols = [
        col for col in current_cols
        if (
            col not in fixed_cols
            and col not in calendar_cols
            and col not in sort_order_cols
            and col not in visible_score_cols
        )
    ]

    final_col_order = (
        fixed_cols
        + calendar_cols
        + sort_order_cols
        + visible_score_cols
        + remaining_cols
    )

    df = df[final_col_order]

    if return_sort_details:
        details = {
            "active_rules": active_rules,
            "stages": stage_details,
            "dropped_columns": dropped_columns,
            "inactive_all_zero_columns": inactive_all_zero_columns,
            "missing_columns": missing_columns,
            "sort_by_columns": sort_by_cols,
            "ascending_values": ascending_values,
        }

        return df, details

    return df