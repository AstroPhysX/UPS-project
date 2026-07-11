from datetime import date, datetime, timedelta
import re
import math
from collections import Counter
import pandas as pd
from copy import deepcopy

#Blockiness score using Red flag without category scores--------------------------------------------------------------------------------------------------------------------------------------
def weighted_block_average(lengths):
    """
    Rewards larger blocks.
    A 14-day block is better than two 7-day blocks.
    No ideal block length is used.
    """
    lengths = [length for length in lengths if length > 0]

    if not lengths:
        return 0

    return sum(length ** 2 for length in lengths) / sum(lengths)

def harmonic_block_average(lengths):
    """
    Punishes tiny blocks.
    A single 1-day block drags this down hard.
    """
    lengths = [length for length in lengths if length > 0]

    if not lengths:
        return 0

    return len(lengths) / sum(1 / length for length in lengths)

def block_quality(lengths):
    """
    Combines:
        - reward for big blocks
        - punishment for tiny blocks

    No ideal block length is used.
    """
    lengths = [length for length in lengths if length > 0]

    if not lengths:
        return 0

    weighted = weighted_block_average(lengths)
    harmonic = harmonic_block_average(lengths)

    return (weighted * harmonic) ** 0.5

def merge_touching_work_blocks(work_blocks):
    """
    Merges work blocks that have no real day off between them.

    Example:
        5 on, 0 off, 2 on
        becomes:
        7 on
    """
    if not work_blocks:
        return []

    work_blocks = sorted(work_blocks, key=lambda block: block["start_date"])

    merged = [work_blocks[0].copy()]

    for block in work_blocks[1:]:
        previous = merged[-1]

        days_off_between = (
            block["start_date"] - previous["end_date"]
        ).days - 1

        if days_off_between <= 0:
            previous["end_date"] = max(
                previous["end_date"],
                block["end_date"]
            )

            previous["days_gone"] = (
                previous["end_date"] - previous["start_date"]
            ).days + 1

        else:
            merged.append(block.copy())

    return merged

def calculate_red_flag_penalty(
    work_blocks,
    edge_off_gaps,
    internal_off_gaps,
):
    """
    Penalizes the things that usually make a line feel ugly manually:

        - 1-day or 2-day work islands
        - 1-day or 2-day off gaps between work blocks
        - too many separate work blocks
        - a short work block surrounded by large off blocks

    These are not ideal block lengths.
    They are anti-choppiness penalties.
    """

    penalty = 0

    work_lengths = [
        block["days_gone"]
        for block in work_blocks
        if block["days_gone"] > 0
    ]

    # ------------------------------------------------------------
    # 1. Penalize short work blocks
    # ------------------------------------------------------------
    # 1-day work block = big penalty
    # 2-day work block = medium penalty
    # 3-day work block = small penalty
    for length in work_lengths:
        if length == 1:
            penalty += 28
        elif length == 2:
            penalty += 18
        elif length == 3:
            penalty += 8

    # ------------------------------------------------------------
    # 2. Penalize short internal off gaps
    # ------------------------------------------------------------
    # These are days off between work blocks.
    # Edge days off are not punished here because they may connect
    # to another pay period.
    for gap in internal_off_gaps:
        if gap == 1:
            penalty += 24
        elif gap == 2:
            penalty += 14
        elif gap == 3:
            penalty += 6

    # ------------------------------------------------------------
    # 3. Penalize too many separate work blocks
    # ------------------------------------------------------------
    # One or two work blocks in a PP can still be clean.
    # Three starts to feel chopped up.
    # Four or more is usually ugly.
    if len(work_blocks) > 2:
        penalty += (len(work_blocks) - 2) * 10

    # ------------------------------------------------------------
    # 4. Penalize isolated work islands inside off time
    # ------------------------------------------------------------
    # Example:
    #     10 off, 1 on, 12 off
    #
    # This is worse than simply having a 1-day work block.
    # It breaks what would otherwise be a large off block.
    for i, block in enumerate(work_blocks):
        length = block["days_gone"]

        left_off = None
        right_off = None

        if i == 0:
            left_off = edge_off_gaps[0]
        else:
            left_off = internal_off_gaps[i - 1]

        if i == len(work_blocks) - 1:
            right_off = edge_off_gaps[-1]
        else:
            right_off = internal_off_gaps[i]

        if left_off is None or right_off is None:
            continue

        surrounding_off = left_off + right_off

        if surrounding_off >= 7:
            if length == 1:
                penalty += 30
            elif length == 2:
                penalty += 20
            elif length == 3:
                penalty += 10

    return penalty

def safe_harmonic_average(values):
    """
    Harmonic average for PP bonuses.

    This makes one ugly PP drag the final line score down more
    than a normal average would.
    """
    values = [value for value in values if value > 0]

    if not values:
        return 0

    return len(values) / sum(1 / value for value in values)

def add_blockiness_scores(
    master_lines,
    bid_period_info,
    vto_fixed_score=10,
    vor_fixed_score=0,
    round_to_nearest=5,
):
    """
    Adds:

        line["blockiness_score"]

    This version removes category_base_scores entirely.

    Scoring:
        - VTO pay periods receive vto_fixed_score.
        - VOR pay periods receive vor_fixed_score.
        - TRIP/RB/RA/SB/SA pay periods receive only the calculated
          red-flag blockiness score.
        - Final line score = average of PP scores, bucketed by 5 points.

    Example:
        add_blockiness_scores(
            master_lines,
            bid_period_info,
            vto_fixed_score=95,
            vor_fixed_score=0,
        )
    """

    pay_period_ranges = bid_period_info["pay_period_date_ranges"]

    code_preference_order = ["VTO", "RB", "RA", "SB", "SA", "VOR"]
    measurable_codes = {"RB", "RA", "SB", "SA", "VOR"}

    for line_number, line in master_lines.items():

        pp_scores = []

        for pp_index, pp in enumerate(line["PPs"]):

            pp_name = pp.get("pp", f"PP{pp_index + 1}")

            if pp_name not in pay_period_ranges:
                pp_scores.append(0)
                continue

            pp_start = date.fromisoformat(pay_period_ranges[pp_name]["start"])
            pp_end = date.fromisoformat(pay_period_ranges[pp_name]["end"])

            trip_blocks = []
            code_dates = {}

            # --------------------------------------------------------
            # Read assignments
            # --------------------------------------------------------
            for assignment in pp["assignments"]:

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
                        "days_gone": (trip_end - trip_start).days + 1,
                    })

                elif "code" in assignment:

                    code = assignment["code"]
                    code_date = date.fromisoformat(assignment["date"])

                    code_dates.setdefault(code, []).append(code_date)

            # --------------------------------------------------------
            # Determine PP category and work blocks
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

                if pp_category == "VTO":
                    pass

                elif pp_category in measurable_codes:

                    all_code_work_dates = []

                    for code, dates in code_dates.items():
                        if code in measurable_codes:
                            all_code_work_dates.extend(dates)

                    all_code_work_dates = sorted(set(all_code_work_dates))

                    if all_code_work_dates:
                        block_start = all_code_work_dates[0]
                        previous_date = all_code_work_dates[0]

                        for current_date in all_code_work_dates[1:]:

                            if current_date == previous_date + timedelta(days=1):
                                previous_date = current_date
                            else:
                                work_blocks.append({
                                    "start_date": block_start,
                                    "end_date": previous_date,
                                    "days_gone": (previous_date - block_start).days + 1,
                                })

                                block_start = current_date
                                previous_date = current_date

                        work_blocks.append({
                            "start_date": block_start,
                            "end_date": previous_date,
                            "days_gone": (previous_date - block_start).days + 1,
                        })

            # --------------------------------------------------------
            # Fixed-score categories
            # --------------------------------------------------------
            if pp_category == "VTO":
                pp_scores.append(vto_fixed_score)
                continue

            if pp_category == "VOR":
                pp_scores.append(vor_fixed_score)
                continue

            # --------------------------------------------------------
            # No work blocks
            # --------------------------------------------------------
            work_blocks = merge_touching_work_blocks(work_blocks)
            work_blocks.sort(key=lambda block: block["start_date"])

            if not work_blocks:
                pp_scores.append(0)
                continue

            # --------------------------------------------------------
            # Build off gaps
            # --------------------------------------------------------
            edge_off_gaps = []
            internal_off_gaps = []

            first_gap = (work_blocks[0]["start_date"] - pp_start).days
            edge_off_gaps.append(max(first_gap, 0))

            for i in range(1, len(work_blocks)):
                previous_end = work_blocks[i - 1]["end_date"]
                next_start = work_blocks[i]["start_date"]

                gap = (next_start - previous_end).days - 1
                internal_off_gaps.append(max(gap, 0))

            last_gap = (pp_end - work_blocks[-1]["end_date"]).days
            edge_off_gaps.append(max(last_gap, 0))

            all_off_gaps = edge_off_gaps + internal_off_gaps

            work_lengths = [
                block["days_gone"]
                for block in work_blocks
                if block["days_gone"] > 0
            ]

            off_lengths = [
                gap
                for gap in all_off_gaps
                if gap > 0
            ]

            # --------------------------------------------------------
            # Positive block quality
            # --------------------------------------------------------
            work_quality = block_quality(work_lengths)
            off_quality = block_quality(off_lengths)

            raw_bonus = (
                0.50 * work_quality
                + 0.50 * off_quality
            )

            raw_bonus = raw_bonus * 7

            # --------------------------------------------------------
            # Red-flag penalties
            # --------------------------------------------------------
            penalty = calculate_red_flag_penalty(
                work_blocks=work_blocks,
                edge_off_gaps=edge_off_gaps,
                internal_off_gaps=internal_off_gaps,
            )

            pp_score = raw_bonus - penalty

            # Keep score controlled.
            pp_score = max(0, min(pp_score, 99))

            pp_scores.append(pp_score)

        # ------------------------------------------------------------
        # Final line score
        # ------------------------------------------------------------
        if pp_scores:
            blockiness_score = sum(pp_scores) / len(pp_scores)
        else:
            blockiness_score = 0

        bucketed_blockiness_score = int(blockiness_score // round_to_nearest) * round_to_nearest

        line["blockiness_score"] = bucketed_blockiness_score

#% of tickets paid by Company--------------------------------------------------------------------------------------------------------------------------------------
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
        

        line["pct_company_tickets"] = line_ticket_pct

#Training in Days on score removed the category scores--------------------------------------------------------------------------------------------------------------------------------------
def add_training_fit_score(
    master_lines,
    training_start,
    training_end,
    bid_start,
    bid_end,
    vto_score=30,
    vor_score=30,
    true_off_max_score=20,
):
    """
    Adds training_fit_score to each line.

    Final score:
        0 to 100 percent.

    Higher score = better.

    Per-day scoring:
        TRIP / RB / RA / SB / SA = 100
        VTO                      = 30
        VOR                      = 20
        true day off              = 0 to true_off_max_score

    True day off logic:
        - Off day at the edge of an off block gets true_off_max_score.
        - Off day in the middle of an off block gets close to 0.
        - This helps avoid training falling in the middle of clean days off.

    Date logic:
        Uses inclusive calendar dates.

        Example:
            training_start = "2023-07-06"
            training_end   = "2023-07-10"

        Means:
            Jul 06, Jul 07, Jul 08, Jul 09, Jul 10

    Adds:
        line["training_fit_score"]

    Returns:
        master_lines
    """

    normal_work_codes = {"TRIP", "RB", "RA", "SB", "SA"}

    def to_date(value):
        if isinstance(value, datetime):
            return value.date()

        if isinstance(value, date):
            return value

        if isinstance(value, str):
            return date.fromisoformat(value)

        raise TypeError(f"Unsupported date value: {value!r}")

    def date_range_inclusive(start, end):
        current = start
        while current <= end:
            yield current
            current += timedelta(days=1)

    def add_day_category(day_categories, day, category):
        """
        Saves the strongest category for a given day.

        Priority:
            TRIP/RB/RA/SB/SA = strongest
            VTO
            VOR
            UNKNOWN = weakest

        This protects you if a day somehow appears in more than one assignment.
        """

        priority = {
            "TRIP": 100,
            "RB": 100,
            "RA": 100,
            "SB": 100,
            "SA": 100,
            "VTO": vto_score,
            "VOR": vor_score,
            "UNKNOWN": 0,
        }

        current_category = day_categories.get(day, "UNKNOWN")

        if priority.get(category, 0) > priority.get(current_category, 0):
            day_categories[day] = category

    def get_day_categories_for_line(line):
        """
        Builds a dictionary like:

            {
                date(2023, 7, 6): "TRIP",
                date(2023, 7, 7): "RB",
                date(2023, 7, 8): "VTO",
                date(2023, 7, 9): "VOR",
            }

        True days off are not stored and later become UNKNOWN.
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

                # Case 2: coded assignment such as VTO, VOR, RB, RA, SB, SA
                code = assignment.get("code")
                assignment_date = assignment.get("date")

                if code and assignment_date:
                    code = str(code).strip().upper()

                    if code in {"VTO", "VOR", "RB", "RA", "SB", "SA"}:
                        add_day_category(
                            day_categories,
                            to_date(assignment_date),
                            code,
                        )

        return day_categories

    def build_true_off_blocks(bid_days, day_categories):
        """
        Builds blocks of true days off.

        True days off are days that do not have:
            - TRIP
            - RB / RA / SB / SA
            - VTO
            - VOR

        In other words, only UNKNOWN days count as true off days.
        """

        off_blocks = []

        current_start = None
        previous_day = None

        for day in sorted(bid_days):
            category = day_categories.get(day, "UNKNOWN")
            is_true_off_day = category == "UNKNOWN"

            if is_true_off_day:
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

    def true_off_day_score(day, off_day_to_block):
        """
        Scores a true day off from 0 to true_off_max_score.

        Edge of off block:
            score = true_off_max_score

        Middle of off block:
            score approaches 0
        """

        block = off_day_to_block.get(day)

        if block is None:
            return 0

        block_start, block_end = block
        block_length = (block_end - block_start).days + 1

        if block_length <= 1:
            return true_off_max_score

        days_from_left_edge = (day - block_start).days
        days_from_right_edge = (block_end - day).days

        edge_distance = min(days_from_left_edge, days_from_right_edge)
        max_possible_edge_distance = (block_length - 1) / 2

        middle_penalty = edge_distance / max_possible_edge_distance

        edge_score_fraction = 1 - middle_penalty

        return true_off_max_score * edge_score_fraction

    def score_training_day(day, day_categories, off_day_to_block):
        category = day_categories.get(day, "UNKNOWN")

        if category in normal_work_codes:
            return 100

        if category == "VTO":
            return vto_score

        if category == "VOR":
            return vor_score

        return true_off_day_score(day, off_day_to_block)

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

    for line_num, line in master_lines.items():

        day_categories = get_day_categories_for_line(line)

        true_off_blocks = build_true_off_blocks(bid_days, day_categories)

        off_day_to_block = {}

        for block_start, block_end in true_off_blocks:
            for day in date_range_inclusive(block_start, block_end):
                off_day_to_block[day] = (block_start, block_end)

        daily_scores = [
            score_training_day(day, day_categories, off_day_to_block)
            for day in training_days
        ]

        training_fit_score = sum(daily_scores) / len(daily_scores)

        line["training_fit_score"] = round(training_fit_score, 1)


#Util functions--------------------------------------------------------------------------------------------------------------------------------------
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

#New vacation score with ocv toggle--------------------------------------------------------------------------------------------------------------------------------------
def add_vacation_days_off_score(
    master_lines,
    vacation_ranges,
    bid_period_info,
    pp_drop_threshold_days=14,
    save_details=False,
):
    score_field = "extra_vacation_days"

    def to_date(value):
        if isinstance(value, date):
            return value
        return datetime.strptime(value, "%Y-%m-%d").date()

    def make_range(start, end, pp_drop=True):
        return {
            "start": to_date(start),
            "end": to_date(end),
            "pp_drop": pp_drop,
        }

    def count_overlap_days(range_a, range_b):
        start = max(range_a["start"], range_b["start"])
        end = min(range_a["end"], range_b["end"])

        if start > end:
            return 0

        return (end - start).days + 1

    def range_length(date_range):
        return (date_range["end"] - date_range["start"]).days + 1

    def get_all_assignments(line_data):
        assignments = []

        for pp in line_data.get("PPs", []):
            assignments.extend(pp.get("assignments", []))

        return assignments

    def merge_blocks(blocks):
        if not blocks:
            return []

        blocks = sorted(blocks, key=lambda b: b["start"])
        merged = [blocks[0].copy()]

        for block in blocks[1:]:
            last = merged[-1]

            if block["start"] <= last["end"] + timedelta(days=1):
                last["end"] = max(last["end"], block["end"])
            else:
                merged.append(block.copy())

        return merged

    pp_ranges = {
        pp_name: make_range(pp_info["start"], pp_info["end"])
        for pp_name, pp_info in bid_period_info["pay_period_date_ranges"].items()
    }

    sorted_pps = sorted(pp_ranges.items(), key=lambda item: item[1]["start"])

    bid_start = min(pp["start"] for pp in pp_ranges.values())
    bid_end = max(pp["end"] for pp in pp_ranges.values())

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
        make_range(
            vac["start"],
            vac["end"],
            pp_drop=vac.get("pp_drop", True),
        )
        for vac in vacation_ranges
    ]

    protected_blocks = []

    # Add the actual vacation ranges first.
    # This was missing in your current version.
    for vac in vacation_blocks:
        protected_blocks.append({
            "start": vac["start"],
            "end": vac["end"],
        })

    # Then apply PP-drop logic.
    for pp_name, pp_range in all_pps_to_check.items():

        pp_drop_check_enabled = False

        for vac in vacation_blocks:
            if vac["pp_drop"] and count_overlap_days(vac, pp_range) > 0:
                pp_drop_check_enabled = True
                break

        if not pp_drop_check_enabled:
            continue

        vacation_days_in_pp = 0

        for vac in vacation_blocks:
            vacation_days_in_pp += count_overlap_days(vac, pp_range)

        if vacation_days_in_pp >= pp_drop_threshold_days:
            protected_blocks.append({
                "start": pp_range["start"],
                "end": pp_range["end"],
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
                "protected_days": protected_days,
                "days_off_before": days_before,
                "days_off_after": days_after,
                "extra_days_off": extra_days_off,
            })

        if block_scores:
            best_block = max(
                block_scores,
                key=lambda b: (b["protected_days"], b["extra_days_off"])
            )

            line_data[score_field] = best_block["extra_days_off"]

            if save_details:
                line_data[f"{score_field}_details"] = {
                    "selected_block_start": best_block["block_start"].isoformat(),
                    "selected_block_end": best_block["block_end"].isoformat(),
                    "protected_days_not_counted_in_score": best_block["protected_days"],
                    "days_off_before": best_block["days_off_before"],
                    "days_off_after": best_block["days_after"] if "days_after" in best_block else best_block["days_off_after"],
                    "score": best_block["extra_days_off"],
                }

        else:
            line_data[score_field] = 0

            if save_details:
                line_data[f"{score_field}_details"] = None

    return new_vacation_ranges

#End or start bid off--------------------------------------------------------------------------------------------------------------------------------------
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

#Average number of legs per rest--------------------------------------------------------------------------------------------------------------------------------------
def add_avg_legs_per_work_day(
    master_lines,
    score_field="avg_legs_per_work_day",
    round_digits=1,
    no_score_categories=None,
):
    """
    Adds this top-level key to each line:

        line_data[score_field]

    Meaning:
        Average number of counted flight legs between rests.

    Rules:
        - DH flights are not counted.
        - BUS positioning is not counted.
        - Flights with code SBG or SBA are not counted.
          Example: SBG2, SBA1, SBG5
        - Rest is detected when flight["rest"] has a real value.
        - VTO, RB, RA, SB, SA, VOR lines/PPs/assignments receive NaN.
        - If no qualifying flight blocks exist, score is NaN, not 0.
    """

    if no_score_categories is None:
        no_score_categories = {"VTO", "RB", "RA", "SB", "SA", "VOR"}
    else:
        no_score_categories = {
            str(value).strip().upper()
            for value in no_score_categories
        }

    def normalize_category(value):
        if value is None:
            return None
        return str(value).strip().upper()

    def get_category(obj):
        if not isinstance(obj, dict):
            return None

        possible_keys = [
            "category",
            "type",
            "assignment_type",
            "line_type",
            "pp_type",
            "status",
        ]

        for key in possible_keys:
            value = normalize_category(obj.get(key))
            if value in no_score_categories:
                return value

        for category in no_score_categories:
            if obj.get(category) is True:
                return category

        return None

    def has_excluded_route_flag(flight):
        flags = flight.get("route_flags")

        if not flags:
            return False

        if isinstance(flags, str):
            flags = [flags]

        for flag in flags:
            flag_text = str(flag).strip().upper()

            if flag_text.startswith("DH"):
                return True

            if "BUS" in flag_text:
                return True

        return False

    def has_excluded_code(flight):
        code = flight.get("code")

        if not code:
            return False

        code_text = str(code).strip().upper()

        return (
            code_text.startswith("SBG")
            or code_text.startswith("SBA")
        )

    def should_count_as_leg(flight):
        if has_excluded_route_flag(flight):
            return False

        if has_excluded_code(flight):
            return False

        return True

    def has_rest_after_flight(flight):
        rest = flight.get("rest")
        return rest not in (None, "", "-")

    for line_number, line_data in master_lines.items():

        # Default every line to NaN first.
        # This prevents missing values from later becoming 0 through .get(..., 0).
        line_data[score_field] = math.nan

        # Pure VTO/RB/RA/SB/SA/VOR line: leave as NaN.
        if get_category(line_data) in no_score_categories:
            continue

        blocks_between_rests = []

        for pp in line_data.get("PPs", []):

            # Ignore full non-trip pay periods.
            if get_category(pp) in no_score_categories:
                continue

            for assignment in pp.get("assignments", []):

                # Ignore non-trip assignments.
                if get_category(assignment) in no_score_categories:
                    continue

                flights = assignment.get("flights", [])

                if not flights:
                    continue

                current_block_count = 0

                for flight in flights:

                    if should_count_as_leg(flight):
                        current_block_count += 1

                    if has_rest_after_flight(flight):
                        if current_block_count > 0:
                            blocks_between_rests.append(current_block_count)

                        current_block_count = 0

                if current_block_count > 0:
                    blocks_between_rests.append(current_block_count)

        # If there are no real counted blocks, keep NaN.
        if not blocks_between_rests:
            continue

        avg_value = sum(blocks_between_rests) / len(blocks_between_rests)

        if round_digits is not None:
            avg_value = round(avg_value, round_digits)

        line_data[score_field] = avg_value

    return master_lines

#Function should maybe be moved to df
def line_numbers_to_bid_string(df, number_of_lines, column="Line Number"):
    """
    Takes the first `number_of_lines` entries from the line number column
    and returns a compressed bid string.

    Example:
        [3, 5, 1, 7, 8, 9, 10, 19]
        -> "3 5 1 7=10 19"
    """

    # Get only the requested number of line numbers
    line_numbers = (
        df[column]
        .head(number_of_lines)
        .dropna()
        .astype(int)
        .tolist()
    )

    if not line_numbers:
        return ""

    parts = []
    start = line_numbers[0]
    previous = line_numbers[0]

    for number in line_numbers[1:]:
        if number == previous + 2:
            # Continue the current range
            previous = number
        else:
            # Close the previous range
            if start == previous:
                parts.append(str(start))
            else:
                parts.append(f"{start}={previous}")

            # Start a new range
            start = number
            previous = number

    # Close the final range
    if start == previous:
        parts.append(str(start))
    else:
        parts.append(f"{start}={previous}")

    return " ".join(parts)

#Requested days off--------------------------------------------------------------------------------------------------------------------------------------
def add_requested_days_off_scores(
    master_lines,
    bid_period_info,
    requested_dates,
    *,
    score_key="pct_requested_days_off",

    # Percentage scoring
    true_off_percent=100.0,
    trip_percent=0.0,
    vto_vor_percent=40.0,
    default_code_percent=0.0,

    # Optional custom scores for other codes
    code_percent_scores=None,

    # Final cap
    max_score=100.0,

    # What to store if every requested date is outside the bid period
    no_valid_dates_score=0,
):
    """
    Adds a requested-days-off percentage score to each line in master_lines.

    New scoring logic:

        The function only cares about the percentage of requested dates that are off.

        True off day:
            100%

        VTO or VOR day:
            40%

        Trip day:
            0%

        RA/RB/SA/SB/etc.:
            0% by default, unless custom code_percent_scores are passed.

    Important behavior:

        - Dates outside the bid period are ignored completely.
        - There is no day-before bonus.
        - There is no day-after bonus.
        - Maximum score is 100.
        - VTO/VOR never count as true off, but they still receive 40%.

    requested_dates accepts:

        "2026-08-15"

        ["2026-08-15", "2026-08-22"]

        ("2026-08-15", "2026-08-22")

        [
            ("2026-08-15", "2026-08-22"),
            "2026-09-01",
        ]

    Mutates master_lines in place and also returns master_lines.
    """

    # -----------------------------
    # Internal helpers
    # -----------------------------

    def to_date(value):
        if isinstance(value, datetime):
            return value.date()

        if isinstance(value, date):
            return value

        if isinstance(value, str):
            text = value.strip()

            for fmt in (
                "%Y-%m-%d",
                "%m/%d/%Y",
                "%m/%d/%y",
                "%Y/%m/%d",
            ):
                try:
                    return datetime.strptime(text, fmt).date()
                except ValueError:
                    pass

        raise ValueError(f"Could not convert {value!r} to a date.")

    def date_range_inclusive(start, end):
        start = to_date(start)
        end = to_date(end)

        if end < start:
            raise ValueError(f"Date range end {end} is before start {start}.")

        result = []
        current = start

        while current <= end:
            result.append(current)
            current += timedelta(days=1)

        return result

    def normalize_requested_dates(value):
        """
        Converts requested_dates into a flat list of date objects.

        This intentionally flattens everything because the score is now based
        on the percentage of requested days that are off, not on separate groups.
        """

        dates = []

        # Tuple means one inclusive date range.
        if isinstance(value, tuple) and len(value) == 2:
            return date_range_inclusive(value[0], value[1])

        # Optional dictionary support.
        if isinstance(value, dict):
            if "start" in value and "end" in value:
                return date_range_inclusive(value["start"], value["end"])

            if "date" in value:
                return [to_date(value["date"])]

            raise ValueError(
                "Dictionary date request must contain either "
                "{'date': ...} or {'start': ..., 'end': ...}."
            )

        # List means multiple requests.
        # Each item can be a singular date, tuple range, or dict.
        if isinstance(value, list):
            for item in value:
                if isinstance(item, tuple) and len(item) == 2:
                    dates.extend(date_range_inclusive(item[0], item[1]))
                elif isinstance(item, dict):
                    dates.extend(normalize_requested_dates(item))
                else:
                    dates.append(to_date(item))

            return dates

        # Anything else is one singular date.
        return [to_date(value)]

    def expand_dates(start, end):
        start = to_date(start)
        end = to_date(end)

        current = start

        while current <= end:
            yield current
            current += timedelta(days=1)

    def clean_code(code):
        if code is None:
            return ""

        return str(code).strip().upper()

    def date_is_inside_bid_period(d):
        return bid_start_date <= d <= bid_end_date

    def get_code_percent(code):
        code = clean_code(code)

        if code in code_percent_scores:
            return code_percent_scores[code]

        return default_code_percent

    def build_day_status_map_for_line(line_data):
        """
        Builds a map of assigned days.

        If a date is missing from this map, and the date is inside the bid period,
        it means the date is truly off.

        Returns:

            {
                date_obj: {
                    "kind": "TRIP" / "CODE",
                    "percent": float,
                    "codes": [...],
                    "trip_ids": [...],
                }
            }
        """

        day_map = {}

        for pp in line_data.get("PPs", []):
            for assignment in pp.get("assignments", []):

                # -----------------------------
                # Trip assignment
                # -----------------------------
                if "flights" in assignment:
                    trip_id = assignment.get("trip_id")

                    for flight in assignment.get("flights", []):
                        start = flight.get("start_date")
                        end = flight.get("end_date")

                        if not start or not end:
                            continue

                        for d in expand_dates(start, end):
                            existing = day_map.setdefault(d, {
                                "kind": None,
                                "percent": default_code_percent,
                                "codes": [],
                                "trip_ids": [],
                            })

                            # Trip day dominates everything else.
                            existing["kind"] = "TRIP"
                            existing["percent"] = trip_percent

                            if trip_id is not None and trip_id not in existing["trip_ids"]:
                                existing["trip_ids"].append(trip_id)

                # -----------------------------
                # Code assignment: VTO, VOR, RA, RB, etc.
                # -----------------------------
                elif "code" in assignment:
                    code = clean_code(assignment.get("code"))

                    if assignment.get("date"):
                        code_dates = [to_date(assignment["date"])]
                    elif assignment.get("start_date") and assignment.get("end_date"):
                        code_dates = list(expand_dates(
                            assignment["start_date"],
                            assignment["end_date"],
                        ))
                    else:
                        continue

                    for d in code_dates:
                        existing = day_map.setdefault(d, {
                            "kind": None,
                            "percent": default_code_percent,
                            "codes": [],
                            "trip_ids": [],
                        })

                        if code not in existing["codes"]:
                            existing["codes"].append(code)

                        # If a trip already exists on this date, keep the trip result.
                        if existing["kind"] == "TRIP":
                            existing["percent"] = trip_percent
                            continue

                        existing["kind"] = "CODE"

                        # If multiple codes somehow exist on the same date,
                        # use the best score among those codes.
                        existing["percent"] = max(
                            existing["percent"],
                            get_code_percent(code),
                        )

        return day_map

    def get_day_status(day_map, d):
        """
        Returns status for a requested date.

        This function assumes d is inside the bid period.

        Inside bid period:
            Missing from day_map = true day off.
            Existing in day_map = trip or code assignment.
        """

        if d not in day_map:
            return {
                "kind": "OFF",
                "percent": true_off_percent,
                "codes": [],
                "trip_ids": [],
                "is_true_off": True,
                "is_vto_or_vor": False,
            }

        status = day_map[d].copy()

        codes = status.get("codes", [])
        is_vto_or_vor = bool(codes) and all(
            clean_code(code) in ("VTO", "VOR")
            for code in codes
        )

        status["is_true_off"] = False
        status["is_vto_or_vor"] = is_vto_or_vor

        return status

    # -----------------------------
    # Normalize settings
    # -----------------------------

    bid_start_date = to_date(bid_period_info["bid_period_date_range"]["start"])
    bid_end_date = to_date(bid_period_info["bid_period_date_range"]["end"])

    requested_date_list = normalize_requested_dates(requested_dates)

    # Remove duplicate requested dates so overlapping ranges do not double-count.
    requested_date_list = sorted(set(requested_date_list))

    valid_requested_dates = [
        d for d in requested_date_list
        if date_is_inside_bid_period(d)
    ]

    ignored_dates_outside_bid_period = [
        d for d in requested_date_list
        if not date_is_inside_bid_period(d)
    ]

    if code_percent_scores is None:
        code_percent_scores = {
            "VTO": vto_vor_percent,
            "VOR": vto_vor_percent,
        }
    else:
        code_percent_scores = {
            clean_code(code): float(percent)
            for code, percent in code_percent_scores.items()
        }

        # Guarantee VTO/VOR exist unless the user explicitly included them.
        code_percent_scores.setdefault("VTO", vto_vor_percent)
        code_percent_scores.setdefault("VOR", vto_vor_percent)

    # -----------------------------
    # Score every line
    # -----------------------------

    for line_number, line_data in master_lines.items():
        day_map = build_day_status_map_for_line(line_data)

        day_details = []
        day_percents = []

        true_off_count = 0
        vto_vor_count = 0
        trip_count = 0
        other_code_count = 0

        for d in valid_requested_dates:
            status = get_day_status(day_map, d)

            day_percents.append(status["percent"])

            if status["kind"] == "OFF":
                true_off_count += 1
            elif status["kind"] == "TRIP":
                trip_count += 1
            elif status["is_vto_or_vor"]:
                vto_vor_count += 1
            elif status["kind"] == "CODE":
                other_code_count += 1

            day_details.append({
                "date": d,
                "kind": status["kind"],
                "percent": status["percent"],
                "codes": status["codes"],
                "trip_ids": status["trip_ids"],
                "is_true_off": status["is_true_off"],
                "is_vto_or_vor": status["is_vto_or_vor"],
            })

        if day_percents:
            final_score = sum(day_percents) / len(day_percents)
        else:
            # This means every requested date was outside the bid period.
            final_score = no_valid_dates_score

        final_score = min(final_score, max_score)

        line_data[score_key] = round(final_score)

#line type preference--------------------------------------------------------------------------------------------------------------------------------------
def add_line_type_preference_scores(
    master_lines,
    preference_order,
    *,
    score_key="line_type_preference_score",
    save_details=False,
    counts_key="line_type_counts",
    scoring_percentages_key="line_type_power_adjusted_percentage",
    preference_score_map_key="line_type_preference_score_map",
    clear_existing_details=True,
    top_score=100,
    bottom_score=0,
    power_law_coeff=1.0,
    unknown_score=None,
    overlay_types=("VTO", "VOR"),
    round_digits=1,
):
    """
    Adds a line-type preference score to each line.

    Default behavior:
        Only saves:
            line_data["line_type_preference_score"]

    If save_details=True, also saves:
        line_data["line_type_counts"]
        line_data["line_type_power_adjusted_percentage"]
        line_data["line_type_preference_score_map"]

    Scoring philosophy:

        1. Pure Trips:
            {"TRIPS": 10}

            Scoring percentage:
                TRIPS = 100%

        2. Trips mixed with VTO:
            {"TRIPS": 10, "VTO": 28}

            Actual counts do NOT matter for the VTO amount.

            Scoring percentage:
                TRIPS = 50%
                VTO   = 50%

            So if:
                TRIPS = 100
                VTO   = 80

            Final score:
                90

        3. Trips mixed with VOR:
            Same behavior as VTO.

        4. Trips mixed with RA/RB/SA/SB/SBA/SBG:
            Uses actual percentages.

        5. DH/BUS:
            Counts as TRIPS by default.

            Exception:
                If a trip starts with DH/BUS, ends with DH/BUS,
                and every non-DH/BUS day inside is the same SBA/SBG type,
                the whole trip counts as SBA/SBG.

                Example:
                    DH + SBG + SBG + SBG + DH

                Counts as:
                    SBG = total_days_gone

    power_law_coeff:
        Controls the power-law falloff of the preference order.

        1.0 = linear
        >1.0 = lower preferences fall off faster
        <1.0 = lower preferences stay closer to the top
    """

    if power_law_coeff <= 0:
        raise ValueError("power_law_coeff must be greater than 0.")

    def normalize_line_type(value):
        if value is None:
            return None

        text = str(value).strip().upper()

        if not text:
            return None

        aliases = {
            "TRIP": "TRIPS",
            "TRIPS": "TRIPS",
            "FLY": "TRIPS",
            "FLYING": "TRIPS",

            "VTO": "VTO",
            "RB": "RB",
            "RA": "RA",
            "SB": "SB",
            "SA": "SA",
            "VOR": "VOR",

            "SBA": "SBA",
            "SBG": "SBG",
        }

        if text in aliases:
            return aliases[text]

        match = re.fullmatch(r"(SBA|SBG)\d*", text)

        if match:
            return match.group(1)

        return text

    def route_flags_have_dh_or_bus(route_flags):
        if not route_flags:
            return False

        if isinstance(route_flags, str):
            route_flags = [route_flags]

        for flag in route_flags:
            flag_text = str(flag).upper()

            if "DH" in flag_text or "BUS" in flag_text:
                return True

        return False

    def safe_int(value, default=0):
        try:
            return int(value)
        except (TypeError, ValueError):
            try:
                return int(float(value))
            except (TypeError, ValueError):
                return default

    def get_flight_date_key(flight, fallback_index):
        return (
            flight.get("start_date")
            or flight.get("date")
            or flight.get("end_date")
            or f"NO_DATE_{fallback_index}"
        )

    def detect_sba_sbg_sandwich_trip(flights):
        """
        Detects:

            DH/BUS + SBA/SBG days + DH/BUS

        If all non-DH/BUS rows are the same SBA/SBG type,
        then the entire trip counts as that SBA/SBG type.
        """

        if not flights:
            return None

        starts_with_dh_or_bus = route_flags_have_dh_or_bus(
            flights[0].get("route_flags")
        )

        ends_with_dh_or_bus = route_flags_have_dh_or_bus(
            flights[-1].get("route_flags")
        )

        if not starts_with_dh_or_bus or not ends_with_dh_or_bus:
            return None

        coded_types = set()

        for flight in flights:
            flight_code = normalize_line_type(flight.get("code"))
            is_dh_or_bus = route_flags_have_dh_or_bus(
                flight.get("route_flags")
            )

            if flight_code in {"SBA", "SBG"}:
                coded_types.add(flight_code)
                continue

            if is_dh_or_bus:
                continue

            return None

        if len(coded_types) == 1:
            return next(iter(coded_types))

        return None

    # ------------------------------------------------------------
    # Normalize overlay types
    # ------------------------------------------------------------

    overlay_types = {
        normalize_line_type(item)
        for item in overlay_types
        if normalize_line_type(item) is not None
    }

    # ------------------------------------------------------------
    # Build power-law preference score map
    # ------------------------------------------------------------

    normalized_order = []

    for item in preference_order:
        normalized = normalize_line_type(item)

        if normalized and normalized not in normalized_order:
            normalized_order.append(normalized)

    if not normalized_order:
        raise ValueError("preference_order must contain at least one valid line type.")

    if len(normalized_order) == 1:
        preference_score_map = {
            normalized_order[0]: float(top_score)
        }
    else:
        max_index = len(normalized_order) - 1
        preference_score_map = {}

        for index, line_type in enumerate(normalized_order):
            rank_position = 1 - (index / max_index)
            curved_position = rank_position ** power_law_coeff

            score = bottom_score + (top_score - bottom_score) * curved_position
            preference_score_map[line_type] = score

    if unknown_score is None:
        unknown_score = bottom_score

    # ------------------------------------------------------------
    # Score each line
    # ------------------------------------------------------------

    for line_number, line_data in master_lines.items():
        counts = Counter()

        for pp in line_data.get("PPs", []):
            for assignment in pp.get("assignments", []):

                # ------------------------------------------------
                # Non-trip assignment:
                # VTO / RB / RA / SB / SA / VOR
                # ------------------------------------------------
                if "flights" not in assignment:
                    code = assignment.get("code")

                    if code:
                        normalized_code = normalize_line_type(code)
                        counts[normalized_code] += 1

                    continue

                # ------------------------------------------------
                # Trip assignment
                # ------------------------------------------------
                flights = assignment.get("flights") or []

                total_trip_days = safe_int(
                    assignment.get("total_days_gone"),
                    default=0,
                )

                if total_trip_days <= 0:
                    unique_dates = {
                        get_flight_date_key(flight, index)
                        for index, flight in enumerate(flights)
                    }
                    total_trip_days = len(unique_dates)

                sandwich_type = detect_sba_sbg_sandwich_trip(flights)

                if sandwich_type in {"SBA", "SBG"}:
                    counts[sandwich_type] += total_trip_days
                    continue

                coded_days = {}

                for index, flight in enumerate(flights):
                    flight_code = normalize_line_type(flight.get("code"))

                    if flight_code in {"SBA", "SBG"}:
                        date_key = get_flight_date_key(flight, index)
                        coded_days[date_key] = flight_code

                for coded_type in coded_days.values():
                    counts[coded_type] += 1

                normal_trip_days = max(total_trip_days - len(coded_days), 0)

                if normal_trip_days > 0:
                    counts["TRIPS"] += normal_trip_days

        # --------------------------------------------------------
        # Build scoring percentages
        # --------------------------------------------------------

        present_overlay_types = {
            line_type
            for line_type, count in counts.items()
            if count > 0 and line_type in overlay_types
        }

        non_overlay_counts = Counter({
            line_type: count
            for line_type, count in counts.items()
            if count > 0 and line_type not in overlay_types
        })

        scoring_weights = {}

        if present_overlay_types and non_overlay_counts:
            # VTO/VOR mixed with something else:
            #
            # Overlay group gets 50%.
            # Non-overlay group gets 50%.
            #
            # Example:
            #   TRIPS 10, VTO 28
            #
            # Becomes:
            #   TRIPS 50%
            #   VTO   50%

            overlay_share = 0.5
            non_overlay_share = 0.5

            # Split overlay share equally among present overlay types.
            # Usually this is just VTO or VOR.
            overlay_each = overlay_share / len(present_overlay_types)

            for overlay_type in present_overlay_types:
                scoring_weights[overlay_type] = overlay_each

            # Split non-overlay share by actual non-overlay percentages.
            non_overlay_total = sum(non_overlay_counts.values())

            for line_type, count in non_overlay_counts.items():
                scoring_weights[line_type] = (
                    non_overlay_share * count / non_overlay_total
                )

        else:
            # No VTO/VOR overlay issue.
            # Use actual percentages.
            total_counted = sum(counts.values())

            if total_counted > 0:
                for line_type, count in counts.items():
                    scoring_weights[line_type] = count / total_counted

        # --------------------------------------------------------
        # Calculate final score
        # --------------------------------------------------------

        if not scoring_weights:
            final_score = 0
            scoring_percentages = {}
        else:
            final_score = 0

            for line_type, weight in scoring_weights.items():
                line_type_score = preference_score_map.get(
                    line_type,
                    unknown_score,
                )

                final_score += line_type_score * weight

            scoring_percentages = {
                line_type: round(weight * 100, round_digits)
                for line_type, weight in scoring_weights.items()
            }

        # --------------------------------------------------------
        # Save result
        # --------------------------------------------------------

        line_data[score_key] = round(final_score, round_digits)

        if save_details:
            line_data[counts_key] = dict(counts)
            line_data[scoring_percentages_key] = scoring_percentages
            line_data[preference_score_map_key] = {
                key: round(value, round_digits)
                for key, value in preference_score_map.items()
            }
        
#% international flights and continents--------------------------------------------------------------------------------------------------------------------------------------
def normalize_airport_code(value):
    """
    Normalizes airport codes like SDF, DFW, CGN, KSDF, etc.
    """

    if value is None:
        return None

    text = str(value).strip().upper()

    if not text or text in {"NONE", "NAN", "NULL", "-", "0"}:
        return None

    text = re.sub(r"[^A-Z0-9]", "", text)

    return text or None

def is_sba_sbg_flight(flight):
    """
    Returns True for SBA/SBG entries such as:
        SBG3
        SBA1

    Handles both trips_dict style:
        flight["flight"] = "SBG3"

    and master_lines style:
        flight["code"] = "SBG3"
    """

    if not isinstance(flight, dict):
        return False

    possible_values = [
        flight.get("flight"),
        flight.get("code"),
    ]

    for value in possible_values:
        if value is None:
            continue

        text = str(value).strip().upper()

        if re.match(r"^(SBA|SBG)\d*\b", text):
            return True

    return False

def collect_unique_arrival_destinations_from_trips(
    trips,
    *,
    ignore_sba_sbg=True,
):
    """
    Collects unique arrival/destination airport codes from the trips dictionary.

    This should be run BEFORE creating_master_lines.

    It only collects arrivals because you want destination percentages,
    not leg percentages.

    Parameters
    ----------
    trips:
        The trips dictionary.

    ignore_sba_sbg:
        If True, ignores SBA/SBG same-airport entries because they are not
        real travel destinations.

    Returns
    -------
    set
        Example:
            {"SDF", "DFW", "ATL", "CGN"}
    """

    destinations = set()

    for trip_id, trip_data in trips.items():
        if not isinstance(trip_data, dict):
            continue

        for block in trip_data.get("blocks", []):
            if not isinstance(block, dict):
                continue

            for flight in block.get("flights", []):
                if not isinstance(flight, dict):
                    continue

                if ignore_sba_sbg and is_sba_sbg_flight(flight):
                    continue

                arrival = normalize_airport_code(flight.get("arrival"))

                if arrival:
                    destinations.add(arrival)

    return destinations

def build_bid_period_airport_lookup(
    airports_csv_path,
    destination_codes,
    *,
    allowed_types=("large_airport", "medium_airport"),
):
    """
    Builds a small airport lookup table for only the destination codes
    found in the bid period.

    Uses OurAirports airports.csv.

    It matches against:
        iata_code
        local_code
        gps_code
        ident

    Returns
    -------
    airport_lookup:
        Dictionary keyed by the bid package destination code.

    unmatched_codes:
        Sorted list of destination codes that were not found.

    matched_airports_df:
        Small DataFrame useful for debugging.
    """

    destination_codes = {
        normalize_airport_code(code)
        for code in destination_codes
        if normalize_airport_code(code)
    }

    if not destination_codes:
        return {}, [], pd.DataFrame()

    columns_needed = [
        "ident",
        "type",
        "name",
        "continent",
        "iso_country",
        "gps_code",
        "iata_code",
        "local_code",
    ]

    airports = pd.read_csv(
        airports_csv_path,
        usecols=lambda col: col in columns_needed,
        dtype=str,
        keep_default_na=False,
    )

    if allowed_types is not None:
        airports = airports[airports["type"].isin(allowed_types)].copy()

    code_columns = [
        "iata_code",
        "local_code",
        "gps_code",
        "ident",
    ]

    matches = []

    for code_column in code_columns:
        temp = airports.copy()

        temp["bid_destination_code"] = temp[code_column].apply(
            normalize_airport_code
        )
        temp["matched_code_column"] = code_column

        temp = temp[temp["bid_destination_code"].isin(destination_codes)]

        if not temp.empty:
            matches.append(temp)

    if not matches:
        return {}, sorted(destination_codes), pd.DataFrame()

    matched_airports = pd.concat(matches, ignore_index=True)

    # If the same code matches multiple rows, prefer IATA first.
    match_priority = {
        "iata_code": 0,
        "local_code": 1,
        "gps_code": 2,
        "ident": 3,
    }

    type_priority = {
        "large_airport": 0,
        "medium_airport": 1,
        "small_airport": 2,
        "heliport": 3,
        "seaplane_base": 4,
        "closed": 5,
    }

    matched_airports["match_priority"] = (
        matched_airports["matched_code_column"]
        .map(match_priority)
        .fillna(99)
    )

    matched_airports["type_priority"] = (
        matched_airports["type"]
        .map(type_priority)
        .fillna(99)
    )

    matched_airports = matched_airports.sort_values(
        by=[
            "bid_destination_code",
            "match_priority",
            "type_priority",
        ]
    )

    matched_airports = matched_airports.drop_duplicates(
        subset=["bid_destination_code"],
        keep="first",
    )

    airport_lookup = {}

    for _, row in matched_airports.iterrows():
        code = row["bid_destination_code"]

        airport_lookup[code] = {
            "local_code": row.get("local_code", ""),
            "iata_code": row.get("iata_code", ""),
            "gps_code": row.get("gps_code", ""),
            "ident": row.get("ident", ""),
            "iso_country": row.get("iso_country", ""),
            "continent": row.get("continent", ""),
            "type": row.get("type", ""),
            "name": row.get("name", ""),
            "matched_code_column": row.get("matched_code_column", ""),
        }

    unmatched_codes = sorted(destination_codes - set(airport_lookup.keys()))

    return airport_lookup, unmatched_codes, matched_airports

def add_international_destination_scores(
    master_lines,
    airport_lookup,
    *,
    home_country="US",
    ignore_sba_sbg=True,
    continent_codes=("EU", "AS", "SA", "AF", "OC"),
    continent_percent_denominator="all_known_destinations",
):
    """
    Adds international destination scoring to each line in master_lines.

    This counts ARRIVALS only.

    Percentage fields are set to NaN when the line has no countable trip
    destinations at all.
    """

    if continent_percent_denominator not in {
        "all_known_destinations",
        "international_destinations",
    }:
        raise ValueError(
            "continent_percent_denominator must be either "
            "'all_known_destinations' or 'international_destinations'"
        )

    for line_number, line_data in master_lines.items():
        international_count = 0
        domestic_count = 0
        known_count = 0

        has_countable_destination = False
        unknown_destinations = set()

        continent_counts = {
            continent: 0 for continent in continent_codes
        }

        for pp in line_data.get("PPs", []):
            if not isinstance(pp, dict):
                continue

            for assignment in pp.get("assignments", []):
                if not isinstance(assignment, dict):
                    continue

                for flight in assignment.get("flights", []):
                    if not isinstance(flight, dict):
                        continue

                    if ignore_sba_sbg and is_sba_sbg_flight(flight):
                        continue

                    arrival = normalize_airport_code(flight.get("arrival"))

                    if not arrival:
                        continue

                    # This means the line has at least one real destination
                    # arrival, even if that airport is not found in the lookup.
                    has_countable_destination = True

                    airport_info = airport_lookup.get(arrival)

                    if airport_info is None:
                        unknown_destinations.add(arrival)
                        continue

                    known_count += 1

                    iso_country = airport_info.get("iso_country", "")
                    continent = airport_info.get("continent", "")

                    if iso_country == home_country:
                        domestic_count += 1
                    else:
                        international_count += 1

                        if continent in continent_counts:
                            continent_counts[continent] += 1

        # ---------------------------------------------------------
        # If there are no real trip destinations, set pct fields NaN
        # ---------------------------------------------------------

        if not has_countable_destination:
            line_data["pct_dest_int"] = float("nan")

            for continent in continent_counts:
                line_data[f"pct_dest_{continent}"] = float("nan")

            continue

        # ---------------------------------------------------------
        # If there are destinations but none matched the airport lookup,
        # the percentage cannot be calculated reliably.
        # ---------------------------------------------------------

        if known_count == 0:
            line_data["pct_dest_int"] = float("nan")

            for continent in continent_counts:
                line_data[f"pct_dest_{continent}"] = float("nan")

            continue

        # ---------------------------------------------------------
        # Normal percentage calculation
        # ---------------------------------------------------------

        line_data["pct_dest_int"] = round(
            international_count / known_count * 100,
            2,
        )

        for continent, count in continent_counts.items():
            percent_field = f"pct_dest_{continent}"

            if continent_percent_denominator == "all_known_destinations":
                denominator = known_count
            else:
                denominator = international_count

            if denominator:
                line_data[percent_field] = round(count / denominator * 100, 2)
            else:
                line_data[percent_field] = 0.0

#Calculates pay per line--------------------------------------------------------------------------------------------------------------------------------------
def add_pay_to_master_lines(
    master_lines,
    hourly_rate,
    *,
    default_pp_guarantee_hours=75.0,
    pp_guarantees=None,
    mutate=True,
    round_digits=2,
    add_flat_fields=True,
    save_details = False
):
    """
    Adds estimated pay information to each line in master_lines.

    This does NOT overwrite:
        CT
        BT
        DT
        tot_CT
        tot_BT
        tot_DT

    Instead it adds:
        line_data["pay"]

    Formula per pay period:
        extracted_credit_hours = PP["CT"]
        paid_credit_hours = max(extracted_credit_hours, guarantee_hours)
        base_pay = paid_credit_hours * hourly_rate

    Line formula:
        total_base_pay = sum(PP base pay)
        total_premium = sum(trip assignment premium)
        total_per_diem = sum(trip assignment per_diem)

        taxable_pay_estimate = total_base_pay + total_premium
        cash_pay_estimate = taxable_pay_estimate + total_per_diem

    Parameters
    ----------
    master_lines:
        Dictionary keyed by line number.

    hourly_rate:
        Hourly pay rate for seat/year.
        Example: 284.29 for 2025 15th-year FO.

    default_pp_guarantee_hours:
        Usually 75.0 for a 28-day pay period.
        Use 96.0 for a 35-day pay period, or pass pp_guarantees.

    pp_guarantees:
        Optional dictionary by PP name.
        Example:
            {"PP1": 75.0, "PP2": 75.0}

    mutate:
        If True, modifies master_lines in place.
        If False, returns a deep copy.

    add_flat_fields:
        If True, also adds simple top-level numeric fields useful for DataFrame sorting:
            pay_cash_estimate
            pay_taxable_estimate
            pay_base
            pay_premium
            pay_per_diem
            paid_CT
            guarantee_credit_added
    """
    def time_to_hours(value):
        """
        Converts UPS-style time strings to decimal hours.

        Handles:
            '74:24'   -> 74.4
            '26:09'   -> 26.15
            '15h37'   -> 15.6167
            '41h08T'  -> 41.1333
            '33h16D'  -> 33.2667
            74.4      -> 74.4
            None      -> 0.0
        """

        if value is None:
            return 0.0

        if isinstance(value, (int, float)):
            return float(value)

        text = str(value).strip()

        if not text:
            return 0.0

        # Remove trailing UPS credit markers such as T, D, M, etc.
        # Example: 41h08T -> 41h08
        text = re.sub(r"[A-Za-z]+$", "", text)

        # Format: HH:MM
        if ":" in text:
            hours, minutes = text.split(":", 1)
            return int(hours) + int(minutes) / 60

        # Format: HHhMM
        match = re.match(r"^(\d+)h(\d+)$", text)
        if match:
            hours = int(match.group(1))
            minutes = int(match.group(2))
            return hours + minutes / 60

        # Format: HHh
        match = re.match(r"^(\d+)h$", text)
        if match:
            return float(match.group(1))

        # Fallback for strings like '74.4'
        try:
            return float(text)
        except ValueError:
            return 0.0


    def hours_to_hhmm(hours):
        """
        Converts decimal hours to H:MM string.

        Example:
            75.0 -> '75:00'
            74.4 -> '74:24'
        """

        if hours is None:
            hours = 0.0

        total_minutes = int(round(float(hours) * 60))
        h = total_minutes // 60
        m = total_minutes % 60
        return f"{h}:{m:02d}"


    def safe_money(value):
        """
        Converts money-like values to float.

        Handles:
            431.89
            '431.89'
            '$431.89'
            None
        """

        if value is None:
            return 0.0

        if isinstance(value, (int, float)):
            return float(value)

        text = str(value).strip().replace("$", "").replace(",", "")

        if not text:
            return 0.0

        try:
            return float(text)
        except ValueError:
            return 0.0


    def get_pp_guarantee_hours(pp, default_guarantee_hours=75.0, pp_guarantees=None):
        """
        Returns the pay-period guarantee for a PP.

        default_guarantee_hours:
            Usually 75.0 for a normal 28-day pay period.

        pp_guarantees:
            Optional dictionary if you need custom guarantees.

            Example:
                {
                    "PP1": 75.0,
                    "PP2": 96.0,
                }
        """

    pp_name = pp.get("pp")

    if pp_guarantees and pp_name in pp_guarantees:
        return float(pp_guarantees[pp_name])

    return float(default_guarantee_hours)

    if not mutate:
        master_lines = deepcopy(master_lines)

    hourly_rate = float(hourly_rate)

    for line_number, line_data in master_lines.items():
        pp_pay_details = []

        line_guarantee_applied = False

        total_extracted_credit_hours = 0.0
        total_paid_credit_hours = 0.0
        total_guarantee_credit_added_hours = 0.0

        total_base_pay = 0.0
        total_premium = 0.0
        total_per_diem = 0.0

        for pp_index, pp in enumerate(line_data.get("PPs", []), start=1):
            pp_name = pp.get("pp", f"PP{pp_index}")

            extracted_credit_hours = time_to_hours(pp.get("CT"))
            guarantee_hours = get_pp_guarantee_hours(
                pp,
                default_guarantee_hours=default_pp_guarantee_hours,
                pp_guarantees=pp_guarantees,
            )

            paid_credit_hours = max(extracted_credit_hours, guarantee_hours)
            guarantee_credit_added_hours = max(
                0.0,
                paid_credit_hours - extracted_credit_hours,
            )

            pp_guarantee_applied = paid_credit_hours > extracted_credit_hours
            line_guarantee_applied = line_guarantee_applied or pp_guarantee_applied

            pp_base_pay = paid_credit_hours * hourly_rate

            pp_premium = 0.0
            pp_per_diem = 0.0
            pp_trip_ids = []

            for assignment in pp.get("assignments", []):
                # Trip assignments have trip_id, premium, per_diem.
                # VTO/RA/RB/SA/SB/VOR day assignments usually only have code/date.
                if "trip_id" in assignment:
                    pp_trip_ids.append(assignment.get("trip_id"))

                pp_premium += safe_money(assignment.get("premium"))
                pp_per_diem += safe_money(assignment.get("per_diem"))

            pp_taxable_pay = pp_base_pay + pp_premium
            pp_cash_pay = pp_taxable_pay + pp_per_diem

            total_extracted_credit_hours += extracted_credit_hours
            total_paid_credit_hours += paid_credit_hours
            total_guarantee_credit_added_hours += guarantee_credit_added_hours

            total_base_pay += pp_base_pay
            total_premium += pp_premium
            total_per_diem += pp_per_diem
            if save_details:
                pp_pay_details.append({
                    "pp": pp_name,

                    "guarantee_applied": pp_guarantee_applied,
                    "guarantee_credit_added_hours": round(guarantee_credit_added_hours, 4),

                    # Money
                    "base_pay": round(pp_base_pay, round_digits),
                    "taxable_pay_estimate": round(pp_taxable_pay, round_digits),
                    "cash_pay_estimate": round(pp_cash_pay, round_digits),

                    # Debug / traceability
                    "trip_ids": pp_trip_ids,
                })

        taxable_pay_estimate = total_base_pay + total_premium
        cash_pay_estimate = taxable_pay_estimate + total_per_diem
        if save_details:
            line_data["pay"] = {
                "guarantee_credit_added_hours": round(total_guarantee_credit_added_hours, 4),

                "base_pay": round(total_base_pay, round_digits),
                "premium": round(total_premium, round_digits),
                "per_diem": round(total_per_diem, round_digits),

                "taxable_pay_estimate": round(taxable_pay_estimate, round_digits),
                "cash_pay_estimate": round(cash_pay_estimate, round_digits),

                "PPs": pp_pay_details,
            }

        if add_flat_fields:
            line_data["pay_guarantee_applied"] = line_guarantee_applied
            line_data["tot_pay"] = round(cash_pay_estimate, round_digits)
            line_data["pay_taxable"] = round(taxable_pay_estimate, round_digits)
            line_data["pay_premium"] = round(total_premium, round_digits)
            line_data["pay_per_diem"] = round(total_per_diem, round_digits)

#% of weekends off--------------------------------------------------------------------------------------------------------------------------------------
def add_weekends_off_percentage(
    master_lines,
    bid_period_info=None,
    *,
    round_digits=0,
    save_details=False,
):
    """
    Adds the percentage of complete Saturday/Sunday weekends off
    to each line in master_lines.

    A weekend counts as off only when:
        - Saturday is off
        - Sunday is off

    VTO and VOR are ignored completely:
        - They do not count as working.
        - They do not count as off.
        - A weekend containing VTO or VOR is excluded from the calculation.

    Normal trips count as working for every calendar date from the
    beginning through the end of the trip, including layover/rest days.

    Other dated codes, such as RA, RB, SA, and SB, count as working.

    Adds:
        line_data["weekends_off_percent"]

    When save_details=True, also adds:
        line_data["weekends_off_count"]
        line_data["weekends_worked_count"]
        line_data["weekends_ignored_count"]
        line_data["weekends_counted"]
    """

    ignored_codes = {"VTO", "VOR"}

    def parse_date(value):
        if value is None:
            return None

        if isinstance(value, datetime):
            return value.date()

        if isinstance(value, date):
            return value

        return date.fromisoformat(str(value).strip())

    if bid_period_info is not None:
        bid_range = bid_period_info["bid_period_date_range"]

        bid_start = parse_date(bid_range["start"])
        bid_end = parse_date(bid_range["end"])

    else:
        all_dates = []

        for line_data in master_lines.values():
            for pp in line_data.get("PPs", []):
                for assignment in pp.get("assignments", []):
                    for flight in assignment.get("flights") or []:
                        flight_start = parse_date(
                            flight.get("start_date")
                        )
                        flight_end = parse_date(
                            flight.get("end_date")
                        )

                        if flight_start is not None:
                            all_dates.append(flight_start)

                        if flight_end is not None:
                            all_dates.append(flight_end)

                    assignment_date = parse_date(
                        assignment.get("date")
                    )

                    if assignment_date is not None:
                        all_dates.append(assignment_date)

        if not all_dates:
            raise ValueError(
                "No dates were found in master_lines."
            )

        bid_start = min(all_dates)
        bid_end = max(all_dates)

    def add_date_range(target_set, start_date, end_date):
        current_date = start_date

        while current_date <= end_date:
            if bid_start <= current_date <= bid_end:
                target_set.add(current_date)

            current_date += timedelta(days=1)

    def get_line_dates(line_data):
        """
        Returns:
            work_dates
            ignored_dates
        """

        work_dates = set()
        ignored_dates = set()

        for pp in line_data.get("PPs", []):
            for assignment in pp.get("assignments", []):
                flights = assignment.get("flights") or []

                # Normal trip assignment
                if flights:
                    trip_dates = []

                    for flight in flights:
                        flight_start = parse_date(
                            flight.get("start_date")
                        )
                        flight_end = parse_date(
                            flight.get("end_date")
                        )

                        if flight_start is not None:
                            trip_dates.append(flight_start)

                        if flight_end is not None:
                            trip_dates.append(flight_end)

                    if trip_dates:
                        add_date_range(
                            work_dates,
                            min(trip_dates),
                            max(trip_dates),
                        )

                    continue

                # Dated code assignment
                code = str(
                    assignment.get("code") or ""
                ).strip().upper()

                assignment_date = parse_date(
                    assignment.get("date")
                )

                if assignment_date is not None:
                    if not bid_start <= assignment_date <= bid_end:
                        continue

                    if code in ignored_codes:
                        ignored_dates.add(assignment_date)
                    else:
                        work_dates.add(assignment_date)

                    continue

                # Support an assignment-level date range if one exists
                assignment_start = parse_date(
                    assignment.get("start_date")
                )
                assignment_end = parse_date(
                    assignment.get("end_date")
                )

                if (
                    assignment_start is None
                    and assignment_end is None
                ):
                    continue

                assignment_start = (
                    assignment_start or assignment_end
                )
                assignment_end = (
                    assignment_end or assignment_start
                )

                target_set = (
                    ignored_dates
                    if code in ignored_codes
                    else work_dates
                )

                add_date_range(
                    target_set,
                    assignment_start,
                    assignment_end,
                )

        # If conflicting data places both work and VTO/VOR on a date,
        # treat the actual work assignment as controlling.
        ignored_dates.difference_update(work_dates)

        return work_dates, ignored_dates

    # Find the first Saturday in the bid period
    days_until_saturday = (5 - bid_start.weekday()) % 7
    saturday = bid_start + timedelta(days=days_until_saturday)

    weekends = []

    while saturday + timedelta(days=1) <= bid_end:
        sunday = saturday + timedelta(days=1)

        weekends.append((saturday, sunday))
        saturday += timedelta(days=7)

    for line_data in master_lines.values():
        if not isinstance(line_data, dict):
            continue

        work_dates, ignored_dates = get_line_dates(line_data)

        weekends_off = 0
        weekends_worked = 0
        weekends_ignored = 0

        for saturday, sunday in weekends:
            # The entire weekend is ignored if either day is VTO or VOR.
            if (
                saturday in ignored_dates
                or sunday in ignored_dates
            ):
                weekends_ignored += 1
                continue

            if (
                saturday in work_dates
                or sunday in work_dates
            ):
                weekends_worked += 1
            else:
                weekends_off += 1

        weekends_counted = (
            weekends_off + weekends_worked
        )

        if weekends_counted:
            percentage = (
                weekends_off
                / weekends_counted
                * 100
            )
        else:
            percentage = 0.0

        line_data["pct_weekends_off"] = round(
            percentage,
            round_digits,
        )

        if save_details:
            line_data["weekends_off_count"] = weekends_off
            line_data["weekends_worked_count"] = weekends_worked
            line_data["weekends_ignored_count"] = weekends_ignored
            line_data["weekends_counted"] = weekends_counted