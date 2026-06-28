from datetime import date, datetime, timedelta
import re

def add_blockiness_scores(master_lines, bid_start, bid_end, pp_length_days=28):
    """
    Adds one top-level key to each line:

        line["blockiness_score"]

    Scoring:
        - Each PP gets its own score.
        - Final line score = average of PP scores.

    Preference order:
        TRIP > VTO > RB > RA > SB > SA > VOR

    Method:
        PP score = category_base_score + blockiness_bonus

    The category base keeps your preference order dominant.
    The blockiness bonus ranks lines within the same category.
    """

    if isinstance(bid_start, str):
        bid_start = date.fromisoformat(bid_start)

    if isinstance(bid_end, str):
        bid_end = date.fromisoformat(bid_end)

    category_base_scores = {
        "TRIP": 700,
        "VTO": 600,
        "RB": 500,
        "RA": 400,
        "SB": 300,
        "SA": 200,
        "VOR": 100,
        "UNKNOWN": 0,
    }

    code_preference_order = ["VTO", "RB", "RA", "SB", "SA", "VOR"]
    measurable_codes = {"RB", "RA", "SB", "SA", "VOR"}

    for line_number, line in master_lines.items():

        pp_scores = []

        for pp_index, pp in enumerate(line["PPs"]):

            pp_start = bid_start + timedelta(days=pp_index * pp_length_days)
            pp_end = pp_start + timedelta(days=pp_length_days - 1)

            if pp_end > bid_end:
                pp_end = bid_end

            trip_blocks = []
            code_dates = {}

            # --------------------------------------------------------
            # 1. Read assignments using your actual master_lines format
            # --------------------------------------------------------
            for assignment in pp["assignments"]:

                # Normal trip assignment
                if "flights" in assignment:

                    start_dates = []
                    end_dates = []

                    for flight in assignment["flights"]:
                        start_dates.append(date.fromisoformat(flight["start_date"]))
                        end_dates.append(date.fromisoformat(flight["end_date"]))

                    trip_start = min(start_dates)
                    trip_end = max(end_dates)

                    trip_blocks.append({
                        "start_date": trip_start,
                        "end_date": trip_end,
                        "days_gone": assignment["total_days_gone"],
                    })

                # Code assignment: {'code': 'VTO', 'date': '2023-05-21'}
                elif "code" in assignment:

                    code = assignment["code"]
                    code_date = date.fromisoformat(assignment["date"])

                    code_dates.setdefault(code, []).append(code_date)

            # --------------------------------------------------------
            # 2. Determine PP category and work blocks
            # --------------------------------------------------------
            if trip_blocks:
                pp_category = "TRIP"
                work_blocks = trip_blocks

            else:
                pp_category = "UNKNOWN"

                for code in code_preference_order:
                    if code in code_dates:
                        pp_category = code
                        break

                work_blocks = []

                # VTO is time off, so it has no work blocks.
                if pp_category == "VTO":
                    pass

                # RB / RA / SB / SA / VOR can be measured as blocks.
                elif pp_category in measurable_codes:

                    all_code_work_dates = []

                    for code, dates in code_dates.items():
                        if code in measurable_codes:
                            all_code_work_dates.extend(dates)

                    all_code_work_dates = sorted(set(all_code_work_dates))

                    # Group consecutive code dates into blocks.
                    # Example:
                    #   Jun 01, Jun 02, Jun 03 becomes one 3-day block.
                    if all_code_work_dates:
                        block_start = all_code_work_dates[0]
                        previous_date = all_code_work_dates[0]

                        for current_date in all_code_work_dates[1:]:

                            if current_date == previous_date + timedelta(days=1):
                                previous_date = current_date
                            else:
                                days_gone = (previous_date - block_start).days + 1

                                work_blocks.append({
                                    "start_date": block_start,
                                    "end_date": previous_date,
                                    "days_gone": days_gone,
                                })

                                block_start = current_date
                                previous_date = current_date

                        days_gone = (previous_date - block_start).days + 1

                        work_blocks.append({
                            "start_date": block_start,
                            "end_date": previous_date,
                            "days_gone": days_gone,
                        })

            base_score = category_base_scores.get(pp_category, 0)

            # --------------------------------------------------------
            # 3. Special handling for VTO
            # --------------------------------------------------------
            if pp_category == "VTO":
                # VTO usually covers the whole PP.
                # Treat it as one big clean off block.
                vto_fixed_score = 60

                pp_scores.append(base_score + vto_fixed_score) # Change this line to multiplication if desired
                continue

            # --------------------------------------------------------
            # 4. Calculate blockiness bonus
            # --------------------------------------------------------
            work_blocks.sort(key=lambda block: block["start_date"])

            if not work_blocks:
                pp_scores.append(base_score)
                continue

            days_gone_list = [
                block["days_gone"]
                for block in work_blocks
            ]

            days_between_trips_list = []

            # PP start to first work block
            first_gap = (work_blocks[0]["start_date"] - pp_start).days
            days_between_trips_list.append(max(first_gap, 0))

            # Between work blocks
            for i in range(1, len(work_blocks)):
                previous_end = work_blocks[i - 1]["end_date"]
                next_start = work_blocks[i]["start_date"]

                gap = (next_start - previous_end).days
                days_between_trips_list.append(max(gap, 0))

            # Last work block to PP end
            last_gap = (pp_end - work_blocks[-1]["end_date"]).days
            days_between_trips_list.append(max(last_gap, 0))

            # Trips can use official DD/DO.
            # Codes use calculated DD/DO because package DD/DO can be incoherent.
            if pp_category == "TRIP":
                dd_denominator = int(pp["DD"])
                do_denominator = int(pp["DO"])
            else:
                dd_denominator = sum(days_gone_list)
                do_denominator = sum(days_between_trips_list)

            if dd_denominator > 0:
                days_gone_component = (
                    sum(day ** 2 for day in days_gone_list) / dd_denominator
                )
            else:
                days_gone_component = 0

            if do_denominator > 0:
                days_between_component = (
                    sum(day ** 2 for day in days_between_trips_list) / do_denominator
                )
            else:
                days_between_component = 0

            blockiness_bonus = (
                0.5 * days_gone_component
                + 0.5 * days_between_component
            )

            pp_scores.append(base_score + blockiness_bonus) # Change this line to multiplication if desired

        if pp_scores:
            line["blockiness_score"] = sum(pp_scores) / len(pp_scores)
        else:
            line["blockiness_score"] = 0




def add_company_ticket_percentages(master_lines):
    """
    Adds company-paid ticket percentage to each pay period and each line.

    Logic:
        - Look only at the first and last flight of each trip.
        - If first flight has DH + airline, except DH UPS, count as company-paid ticket to work.
        - If last flight has DH + airline, except DH UPS, count as company-paid ticket from work.
        - Non-trip assignments like VTO, VOR, RA, RB, SA, SB are ignored.

    Adds to each line:
        line["company_ticket_pct"]
    """

    dh_pattern = re.compile(r"\bDH\s+([A-Z0-9]+)\b", re.IGNORECASE)

    for line_num, line in master_lines.items():

        line_to_work = 0
        line_from_work = 0
        line_ticket_count = 0
        line_ticket_possible = 0

        for pp in line.get("PPs", []):

            pp_to_work = 0
            pp_from_work = 0
            pp_ticket_count = 0
            pp_ticket_possible = 0

            for assignment in pp.get("assignments", []):

                flights = assignment.get("flights")

                # Skip VTO / VOR / RA / RB / SA / SB / anything that is not a trip
                if not flights:
                    continue

                first_flight = flights[0]
                last_flight = flights[-1]

                # Each trip has 2 possible ticket positions:
                #   1. ticket to work
                #   2. ticket from work
                pp_ticket_possible += 2

                # -------------------------
                # Check first flight
                # -------------------------
                first_flags = first_flight.get("route_flags") or []

                if isinstance(first_flags, str):
                    first_flags = [first_flags]

                first_has_company_ticket = False

                for flag in first_flags:
                    flag = str(flag).upper()
                    matches = dh_pattern.findall(flag)

                    for carrier in matches:
                        if carrier.upper() != "UPS":
                            first_has_company_ticket = True
                            break

                    if first_has_company_ticket:
                        break

                if first_has_company_ticket:
                    pp_to_work += 1
                    pp_ticket_count += 1

                # -------------------------
                # Check last flight
                # -------------------------
                last_flags = last_flight.get("route_flags") or []

                if isinstance(last_flags, str):
                    last_flags = [last_flags]

                last_has_company_ticket = False

                for flag in last_flags:
                    flag = str(flag).upper()
                    matches = dh_pattern.findall(flag)

                    for carrier in matches:
                        if carrier.upper() != "UPS":
                            last_has_company_ticket = True
                            break

                    if last_has_company_ticket:
                        break

                if last_has_company_ticket:
                    pp_from_work += 1
                    pp_ticket_count += 1

            if pp_ticket_possible:
                pp_ticket_pct = round((pp_ticket_count / pp_ticket_possible) * 100, 1)
            else:
                pp_ticket_pct = 0.0

            line_to_work += pp_to_work
            line_from_work += pp_from_work
            line_ticket_count += pp_ticket_count
            line_ticket_possible += pp_ticket_possible

        if line_ticket_possible:
            line_ticket_pct = round((line_ticket_count / line_ticket_possible) * 100, 1)
        else:
            line_ticket_pct = 0.0
        

        line["company_ticket_pct"] = line_ticket_pct

from datetime import date, datetime, timedelta

def to_date(value):
    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, str):
        return date.fromisoformat(value)

    raise TypeError(f"Unsupported date value: {value!r}")

def add_training_fit_score(
    master_lines,
    training_start,
    training_end,
    bid_start,
    bid_end,
    trip_weight=0.80,
    off_edge_weight=0.20,
    category_base_scores=None,
):
    """
    Adds training_fit_score to each line.

    Final score:
        training_fit_score = category_base_score + fit_score

    Category base score:
        Based on the best/highest category touched by the training dates.

    Fit score:
        0 to 100 score that rewards training replacing work/reserve days
        and penalizes training falling in the middle of true days-off blocks.
    """

    if category_base_scores is None:
        category_base_scores = {
            "TRIP": 700,
            "VTO": 200,
            "RB": 600,
            "RA": 500,
            "SB": 400,
            "SA": 300,
            "VOR": 100,
            "UNKNOWN": 0,
        }

    # Codes that should count as "on/work" days for training replacement.
    # VTO is intentionally excluded because it is time off.
    work_codes = {"TRIP", "RB", "RA", "SB", "SA", "VOR"}

    def date_range_inclusive(start, end):
        current = start
        while current <= end:
            yield current
            current += timedelta(days=1)

    def add_day_category(day_categories, day, category):
        """
        Saves the highest-value category for a given day.

        Example:
            If a day somehow has both TRIP and RB,
            TRIP wins because 700 > 500.
        """
        current_category = day_categories.get(day, "UNKNOWN")

        if category_base_scores.get(category, 0) > category_base_scores.get(current_category, 0):
            day_categories[day] = category

    def get_day_categories_for_line(line):
        """
        Builds a dictionary like:

            {
                date(2023, 7, 6): "TRIP",
                date(2023, 7, 7): "RB",
                date(2023, 7, 8): "VTO",
            }

        True days off will simply not appear and later become UNKNOWN.
        """

        day_categories = {}

        for pp in line.get("PPs", []):
            for assignment in pp.get("assignments", []):

                flights = assignment.get("flights")

                # Case 1: real trip
                if flights:
                    start_dates = [
                        to_date(flight["start_date"])
                        for flight in flights
                        if flight.get("start_date")
                    ]

                    end_dates = [
                        to_date(flight["end_date"])
                        for flight in flights
                        if flight.get("end_date")
                    ]

                    if not start_dates or not end_dates:
                        continue

                    trip_start = min(start_dates)
                    trip_end = max(end_dates)

                    for day in date_range_inclusive(trip_start, trip_end):
                        add_day_category(day_categories, day, "TRIP")

                    continue

                # Case 2: coded assignment such as VTO, RB, RA, SB, SA, VOR
                code = assignment.get("code")
                assignment_date = assignment.get("date")

                if code and assignment_date:
                    code = str(code).strip().upper()

                    if code in category_base_scores:
                        add_day_category(
                            day_categories,
                            to_date(assignment_date),
                            code,
                        )

        return day_categories

    def build_off_blocks(bid_days, work_days):
        off_blocks = []

        current_start = None
        previous_day = None

        for day in sorted(bid_days):
            is_off_day = day not in work_days

            if is_off_day:
                if current_start is None:
                    current_start = day

                previous_day = day

            else:
                if current_start is not None:
                    off_blocks.append((current_start, previous_day))
                    current_start = None
                    previous_day = None

        if current_start is not None:
            off_blocks.append((current_start, previous_day))

        return off_blocks

    training_start = to_date(training_start)
    training_end = to_date(training_end)
    bid_start = to_date(bid_start)
    bid_end = to_date(bid_end)

    if training_end < training_start:
        raise ValueError("training_end must be on or after training_start")

    if bid_end < bid_start:
        raise ValueError("bid_end must be on or after bid_start")

    training_days = list(date_range_inclusive(training_start, training_end))
    bid_days = set(date_range_inclusive(bid_start, bid_end))
    training_total_days = len(training_days)

    for line_num, line in master_lines.items():

        day_categories = get_day_categories_for_line(line)

        # Determine the category for each training day.
        training_day_categories = [
            day_categories.get(day, "UNKNOWN")
            for day in training_days
        ]

        # Pick the best/highest category touched during training.
        best_training_category = max(
            training_day_categories,
            key=lambda category: category_base_scores.get(category, 0),
        )

        category_base_score = category_base_scores.get(best_training_category, 0)

        # Work days are TRIP, RB, RA, SB, SA, VOR.
        # VTO and UNKNOWN are treated as off days.
        work_days = {
            day
            for day, category in day_categories.items()
            if category in work_codes
        }

        # 1. Reward training that overlaps work/reserve days.
        training_work_days = [
            day for day in training_days
            if day in work_days
        ]

        work_overlap_pct = (
            len(training_work_days) / training_total_days
        ) * 100

        # 2. Penalize training that falls in the middle of true days-off blocks.
        training_off_days = [
            day for day in training_days
            if day not in work_days
        ]

        off_blocks = build_off_blocks(bid_days, work_days)

        off_day_to_block = {}

        for block_start, block_end in off_blocks:
            for day in date_range_inclusive(block_start, block_end):
                off_day_to_block[day] = (block_start, block_end)

        middle_penalty_values = []

        for day in training_off_days:
            block = off_day_to_block.get(day)

            if block is None:
                middle_penalty_values.append(1.0)
                continue

            block_start, block_end = block
            block_length = (block_end - block_start).days + 1

            if block_length <= 1:
                middle_penalty_values.append(0.0)
                continue

            days_from_left_edge = (day - block_start).days
            days_from_right_edge = (block_end - day).days

            edge_distance = min(days_from_left_edge, days_from_right_edge)
            max_possible_edge_distance = (block_length - 1) / 2

            middle_penalty = edge_distance / max_possible_edge_distance
            middle_penalty_values.append(middle_penalty)

        if middle_penalty_values:
            off_middle_penalty_pct = (
                sum(middle_penalty_values) / len(middle_penalty_values)
            ) * 100
        else:
            off_middle_penalty_pct = 0.0

        off_edge_score = 100 - off_middle_penalty_pct

        fit_score = (
            trip_weight * work_overlap_pct
            + off_edge_weight * off_edge_score
        )

        line["training_fit_score"] = round(category_base_score + fit_score, 1)



def count_days_off_around_date(assignments, target_date, before_or_after, bid_start, bid_end):
    """
    Counts consecutive days off immediately before or after a given date.

    Parameters:
        assignments:
            The assignments list from a line or pay period.

        target_date:
            Date to count around. Can be 'YYYY-MM-DD' or date object.

        before_or_after:
            Either 'before' or 'after'.

        bid_start:
            Start date boundary. Can be 'YYYY-MM-DD' or date object.

        bid_end:
            End date boundary. Can be 'YYYY-MM-DD' or date object.

    Returns:
        Integer number of consecutive days off.
    """

    target_date = to_date(target_date)
    bid_start = to_date(bid_start)
    bid_end = to_date(bid_end)

    if before_or_after not in {"before", "after"}:
        raise ValueError("before_or_after must be 'before' or 'after'")

    busy_dates = set()

    for assignment in assignments:

        # Trip assignment
        if "flights" in assignment:
            flight_dates = []

            for flight in assignment["flights"]:
                flight_dates.append(to_date(flight["start_date"]))
                flight_dates.append(to_date(flight["end_date"]))

            trip_start = min(flight_dates)
            trip_end = max(flight_dates)

            current = trip_start
            while current <= trip_end:
                busy_dates.add(current)
                current += timedelta(days=1)

        # Single-day code assignment: RA, RB, SA, SB, VOR, VTO, etc.
        elif "date" in assignment:
            assignment_date = to_date(assignment["date"])
            code = assignment.get("code")

            # RA, RB, SA, SB, VOR, VTO, etc. are not counted as normal days off.
            if code is not None:
                busy_dates.add(assignment_date)

    if before_or_after == "before":
        current = target_date - timedelta(days=1)
        step = -1
    else:
        current = target_date + timedelta(days=1)
        step = 1

    days_off = 0

    while bid_start <= current <= bid_end:
        if current in busy_dates:
            break

        days_off += 1
        current += timedelta(days=step)

    return days_off


def get_all_assignments(line_data):
    assignments = []

    for pp in line_data.get("PPs", []):
        assignments.extend(pp.get("assignments", []))

    return assignments


def add_vacation_days_off_score(
    master_lines,
    vacation_ranges,
    bid_period_info,
    pp_drop_threshold_days=14,
    save_details=True,
):
    """
    Adds a vacation days-off score to each master line.

    The score counts ONLY the extra line-dependent days off directly connected
    to the vacation or dropped pay period.

    It does NOT count:
        - the vacation days themselves
        - the days inside a dropped pay period

    UPS rule:
        If vacation days in a pay period are >= pp_drop_threshold_days,
        that full pay period is treated as protected/dropped.

    Edge cases handled:
        - Vacation causing current PP1 or PP2 to drop.
        - Vacation causing the previous pay period to drop.
        - Vacation causing the next pay period to drop.
        - Multiple separate vacation blocks.
          The function focuses on the largest protected vacation/drop block.
    """
    score_field="extra_vacation_days"

    def make_range(start, end):
        start = to_date(start)
        end = to_date(end)
        return {"start": start, "end": end}

    def count_overlap_days(range_a, range_b):
        start = max(range_a["start"], range_b["start"])
        end = min(range_a["end"], range_b["end"])

        if start > end:
            return 0

        return (end - start).days + 1

    def range_length(date_range):
        return (date_range["end"] - date_range["start"]).days + 1

    def merge_blocks(blocks):
        """
        Merges protected blocks that touch or overlap.

        Example:
            Jun 1-Jun 7 and Jun 8-Jun 14 become Jun 1-Jun 14.
        """

        if not blocks:
            return []

        blocks = sorted(blocks, key=lambda b: b["start"])

        merged = [blocks[0].copy()]

        for block in blocks[1:]:
            last = merged[-1]

            if block["start"] <= last["end"] + timedelta(days=1):
                last["end"] = max(last["end"], block["end"])
                last["reason"] += " + " + block["reason"]
            else:
                merged.append(block.copy())

        return merged

    # Use pay period dates as the real bid boundaries.
    # This avoids accidentally counting the extra day if bid_period end is one day after PP2.
    pp_ranges = {
        pp_name: make_range(pp_info["start"], pp_info["end"])
        for pp_name, pp_info in bid_period_info["pay_period_date_ranges"].items()
    }

    sorted_pps = sorted(pp_ranges.items(), key=lambda item: item[1]["start"])

    bid_start = min(pp["start"] for pp in pp_ranges.values())
    bid_end = max(pp["end"] for pp in pp_ranges.values())

    # Assume pay periods are 28 days unless the actual PP length says otherwise.
    first_pp = sorted_pps[0][1]
    last_pp = sorted_pps[-1][1]
    pp_length = range_length(first_pp)

    previous_pp = {
        "start": first_pp["start"] - timedelta(days=pp_length),
        "end": first_pp["start"] - timedelta(days=1),
    }

    next_pp = {
        "start": last_pp["end"] + timedelta(days=1),
        "end": last_pp["end"] + timedelta(days=pp_length),
    }

    all_pps_to_check = {
        "PREVIOUS_PP": previous_pp,
        **pp_ranges,
        "NEXT_PP": next_pp,
    }

    vacation_blocks = [
        make_range(vac["start"], vac["end"])
        for vac in vacation_ranges
    ]

    protected_blocks = []

    # First add the actual vacation ranges.
    for vac in vacation_blocks:
        protected_blocks.append({
            "start": vac["start"],
            "end": vac["end"],
            "reason": "VACATION",
        })

    # Then apply PP-drop logic.
    for pp_name, pp_range in all_pps_to_check.items():
        vacation_days_in_pp = 0

        for vac in vacation_blocks:
            vacation_days_in_pp += count_overlap_days(vac, pp_range)

        if vacation_days_in_pp >= pp_drop_threshold_days:
            protected_blocks.append({
                "start": pp_range["start"],
                "end": pp_range["end"],
                "reason": f"{pp_name}_DROPPED",
            })

    protected_blocks = merge_blocks(protected_blocks)

    new_vacation_ranges = [
        {
            "start": block["start"].isoformat(),
            "end": block["end"].isoformat(),
        }
        for block in protected_blocks
    ]

    for line_num, line_data in master_lines.items():
        assignments = get_all_assignments(line_data)

        block_scores = []

        for block in protected_blocks:
            days_before = count_days_off_around_date(
                assignments=assignments,
                target_date=block["start"],
                before_or_after="before",
                bid_start=bid_start,
                bid_end=bid_end,
            )

            days_after = count_days_off_around_date(
                assignments=assignments,
                target_date=block["end"],
                before_or_after="after",
                bid_start=bid_start,
                bid_end=bid_end,
            )

            extra_days_off = days_before + days_after
            protected_days = range_length(block)

            block_scores.append({
                "block_start": block["start"],
                "block_end": block["end"],
                "reason": block["reason"],
                "protected_days": protected_days,
                "days_off_before": days_before,
                "days_off_after": days_after,
                "extra_days_off": extra_days_off,
            })

        if block_scores:
            # Important:
            # First prefer the larger vacation/drop block.
            # Then, among similar blocks, prefer more extra days off.
            best_block = max(
                block_scores,
                key=lambda b: (b["protected_days"], b["extra_days_off"])
            )

            line_data[score_field] = best_block["extra_days_off"]

            if save_details:
                line_data[f"{score_field}_details"] = {
                    "selected_block_start": best_block["block_start"].isoformat(),
                    "selected_block_end": best_block["block_end"].isoformat(),
                    "reason": best_block["reason"],
                    "protected_days_not_counted_in_score": best_block["protected_days"],
                    "days_off_before": best_block["days_off_before"],
                    "days_off_after": best_block["days_off_after"],
                    "score": best_block["extra_days_off"],
                }

        else:
            line_data[score_field] = 0

            if save_details:
                line_data[f"{score_field}_details"] = None
    return new_vacation_ranges

from datetime import date, datetime, timedelta


def add_bid_edge_days_off(
    master_lines,
    bid_period_info,
    edge="both",
    start_field="bid_start_days_off",
    end_field="bid_end_days_off",
):
    """
    Adds days-off counts at the start and/or end of the bid period.

    Uses count_days_off_around_date().

    Parameters:
        master_lines:
            Dictionary of master lines.

        bid_period_info:
            Dictionary containing:
                bid_period_info["bid_period_date_range"]["start"]
                bid_period_info["bid_period_date_range"]["end"]

        edge:
            "start", "end", or "both"

        start_field:
            Field name saved in each line for days off at bid start.

        end_field:
            Field name saved in each line for days off at bid end.

    Returns:
        master_lines, modified in place.
    """

    if edge not in {"start", "end", "both"}:
        raise ValueError("edge must be 'start', 'end', or 'both'")

    bid_start = to_date(bid_period_info["bid_period_date_range"]["start"])
    bid_end = to_date(bid_period_info["bid_period_date_range"]["end"])

    for line_number, line_data in master_lines.items():

        assignments = get_all_assignments(line_data)

        if edge in {"start", "both"}:
            line_data[start_field] = count_days_off_around_date(
                assignments=assignments,
                target_date=bid_start - timedelta(days=1),
                before_or_after="after",
                bid_start=bid_start,
                bid_end=bid_end,
            )

        if edge in {"end", "both"}:
            line_data[end_field] = count_days_off_around_date(
                assignments=assignments,
                target_date=bid_end + timedelta(days=1),
                before_or_after="before",
                bid_start=bid_start,
                bid_end=bid_end,
            )
