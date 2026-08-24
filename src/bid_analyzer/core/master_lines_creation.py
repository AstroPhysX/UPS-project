#Updated new master lines
#run creating_master_line
from datetime import datetime, timedelta
import re


TRIP_TYPES = {"trip", "trips"}

#new master lines
def normalize_sba_sbg_code(value):
    """
    Detects:
        SBA
        SBA2
        SBG
        SBG5

    Returns the normalized code or None.
    """
    if value is None:
        return None

    text = str(value).upper().strip()

    match = re.match(r"^(SBA|SBG)\d*$", text)

    if match:
        return text

    return None


def find_sba_sbg_block(trip, line_code=None, block_index=0):
    """
    Finds the block to use for an SBA/SBG assignment.

    Prefer a block whose flight code matches the line_code.
    If line_code is only SBA/SBG without a number, use block_index.
    """
    blocks = trip.get("blocks", [])

    normalized_line_code = normalize_sba_sbg_code(line_code)

    # If the line_code has a number, prefer exact matching flight code.
    if normalized_line_code and re.match(r"^(SBA|SBG)\d+$", normalized_line_code):
        for block in blocks:
            for flight in block.get("flights", []):
                flight_code = extract_flight_code(flight)

                if flight_code == normalized_line_code:
                    return block

    # Otherwise, use the block index supplied by creating_master_line().
    if 0 <= block_index < len(blocks):
        return blocks[block_index]

    # Fallback to first block if something is off.
    if blocks:
        return blocks[0]

    return None


def build_sba_sbg_master_assignment(assignment, trip):
    """
    Builds a synthetic one-flight assignment for SBA/SBG line-code trips
    that do not have a start_time in the Lines package.
    """
    assignment_date = assignment.get("date")
    line_code = normalize_sba_sbg_code(assignment.get("line_code"))
    block_index = safe_int(assignment.get("_sba_sbg_block_index"))

    block = find_sba_sbg_block(
        trip,
        line_code=line_code,
        block_index=block_index
    )

    if block is None:
        return {
            "trip_id": assignment.get("value"),
            "date": assignment_date,
            "code": line_code,
            "error": "SBA/SBG assignment has no usable block in trip",
            "flights": []
        }

    flights = block.get("flights", [])
    first_flight = flights[0] if flights else {}

    # Prefer the actual SBA/SBG code from the trip package.
    # This allows line_code='SBA' to become code='SBA2' if the trip says SBA2.
    trip_code = extract_flight_code(first_flight)
    code = trip_code or line_code

    record = {
        "start_date": assignment_date,
        "departure": first_flight.get("departure"),

        "end_date": assignment_date,
        "arrival": first_flight.get("arrival"),

        "route_flags": extract_route_flags(first_flight),
        "code": code,

        "rest": block.get("rest") if block.get("rest") != "-" else None,
        "block": block.get("block"),
        "credit": block.get("credit"),
        "duty": block.get("duty"),
    }

    return {
        "trip_id": assignment.get("value"),

        # For this edge case, use the selected block as the assignment summary.
        # That prevents counting the whole trip every time one SBA/SBG day appears.
        "tafb": None,
        "block_time": block.get("block"),
        "credit_time": block.get("credit"),
        "duty_time": block.get("duty"),
        "total_blocks": 1,

        "total_days_gone": 1,
        "flights": [record]
    }

def normalize_line_code(value):
    """
    Detects line codes like:
        SBA
        SBA3
        SBG
        SBG5

    Returns:
        'SBA'
        'SBA3'
        'SBG'
        'SBG5'
        None
    """
    if value is None:
        return None

    text = str(value).upper().strip()

    match = re.match(r"^(SBA|SBG)\d*$", text)

    if match:
        return text

    return None

def trip_duration_to_minutes(value):
    """
    Converts:
        '32h19'  -> 1939
        '44h10T' -> 2650
        '7h19D'  -> 439
        '4h00M'  -> 240
        '-'      -> 0
        None     -> 0
    """
    if value is None:
        return 0

    text = str(value).strip()

    if text == "-":
        return 0

    match = re.match(r"^(\d+)h(\d{2})([A-Z])?$", text, re.IGNORECASE)

    if not match:
        return 0

    hours = int(match.group(1))
    minutes = int(match.group(2))

    return hours * 60 + minutes

def datetime_to_master_string(value):
    """
    Stores datetimes in master_lines in a simple parseable format.

    Example:
        2026-09-06T18:20
    """
    if value is None:
        return None

    return value.isoformat(timespec="minutes")

def minutes_to_decimal_hours(total_minutes, decimals=2):
    return round(total_minutes / 60, decimals)


def safe_int(value, default=0):
    if value is None or value == "-":
        return default

    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def safe_float(value, default=0.0):
    if value is None or value == "-":
        return default

    try:
        return float(value)
    except (TypeError, ValueError):
        return default

def parse_duration(value):
    """
    Converts:
        '20h55' -> timedelta(hours=20, minutes=55)
        '66h51' -> timedelta(hours=66, minutes=51)
        '-'     -> None
    """
    if value is None or value == "-":
        return None

    match = re.match(r"^(\d+)h(\d{2})$", str(value).strip())

    if not match:
        return None

    hours = int(match.group(1))
    minutes = int(match.group(2))

    return timedelta(hours=hours, minutes=minutes)


def datetime_at_or_after(reference_dt, time_value):
    if reference_dt is None:
        raise ValueError("reference_dt is None")

    if time_value is None:
        raise ValueError("time_value is None")

    hour, minute = map(int, time_value.split(":"))

    candidate = reference_dt.replace(
        hour=hour,
        minute=minute,
        second=0,
        microsecond=0
    )

    while candidate < reference_dt:
        candidate += timedelta(days=1)

    return candidate

def unique_preserve_order(items):
    result = []

    for item in items:
        if item not in result:
            result.append(item)

    return result


def extract_route_flags(flight):
    flags = []

    if flight.get("route_flags"):
        flags.extend(flight["route_flags"])

    if flight.get("route_flag"):
        flags.append(flight["route_flag"])

    flight_text = str(flight.get("flight", "")).upper().strip()
    route_raw = str(flight.get("route_raw", "")).upper().strip()

    # Deadhead detection
    # Examples:
    # 'DH AA194'    -> 'DH AA'
    # 'DH UPS396'   -> 'DH UPS'
    # 'DH LH1844'   -> 'DH LH'
    # 'DH WN1782-2' -> 'DH WN'
    dh_match = re.match(r"^DH\s+([A-Z]+)", flight_text)

    if dh_match:
        carrier = dh_match.group(1)
        flags.append(f"DH {carrier}")

        if "BUS" in flight_text or "BUS" in route_raw:
            flags.append("BUS")

    return unique_preserve_order(flags)
    
def extract_flight_code(flight):
    """
    Detects special flight codes like:
        SBA1
        SBA 1
        SBG5
        SBG 5
        SBG5 EXTRA_TEXT

    Returns:
        'SBA1'
        'SBG5'
        'SBG12'
        None
    """
    flight_text = str(flight.get("flight", "")).upper().strip()

    match = re.match(r"^(SBA|SBG)\s*(\d+)", flight_text)

    if match:
        code_type = match.group(1)
        code_number = match.group(2)
        return f"{code_type}{code_number}"

    return None


def build_master_assignment(assignment, trips):
    assignment_type = assignment.get("type")

    # -------------------------
    # Non-trip assignment
    # -------------------------
    if assignment_type not in TRIP_TYPES:
        code = assignment.get("value")

        if code is None:
            code = assignment_type

        return {
            "date": assignment.get("date"),
            "code": code
        }

    # -------------------------
    # Trip assignment
    # -------------------------
    trip_id = assignment.get("value")
    trip = trips.get(trip_id)

    if trip is None:
        return {
            "trip_id": trip_id,
            "date": assignment.get("date"),
            "error": f"Trip {trip_id} not found in trips dictionary",
            "flights": []
        }

    assignment_date = assignment.get("date")
    assignment_start_time = assignment.get("start_time")

    # -------------------------
    # SBA / SBG no-start-time edge case
    # -------------------------
    if (
        assignment_start_time is None
        and normalize_sba_sbg_code(assignment.get("line_code")) is not None
    ):
        return build_sba_sbg_master_assignment(assignment, trip)

    if assignment_date is None or assignment_start_time is None:
        return {
            "trip_id": trip_id,

            "premium": safe_float(trip.get("premium")),
            "tafb": trip.get("tafb"),

            "per_diem": safe_float(trip.get("per_diem")),
            "ldgs": safe_int(trip.get("ldgs")),

            "credit_time": trip.get("credit_time"),
            "duty_time": trip.get("duty_time"),
            "block_time": trip.get("block_time"),
            "total_blocks": safe_int(trip.get("total_blocks")),

            "total_days_gone": None,
            "error": "Missing assignment date or start_time",
            "flights": []
        }

    if assignment_date is None or assignment_start_time is None:
        return {
            "trip_id": trip_id,

            "premium": safe_float(trip.get("premium")),
            "tafb": trip.get("tafb"),

            "per_diem": safe_float(trip.get("per_diem")),
            "ldgs": safe_int(trip.get("ldgs")),

            "credit_time": trip.get("credit_time"),
            "duty_time": trip.get("duty_time"),
            "block_time": trip.get("block_time"),
            "total_blocks": safe_int(trip.get("total_blocks")),

            "total_days_gone": None,
            "error": "Missing assignment date or start_time",
            "flights": []
        }

    current_dt = datetime.strptime(
        f"{assignment_date} {assignment_start_time}",
        "%Y-%m-%d %H:%M"
    )

    flight_records = []

    trip_start_dt = None
    trip_end_dt = None

    blocks = trip.get("blocks", [])

    for block_index, block in enumerate(blocks, start=1):
        block_start_time = block.get("start")

        if current_dt is None:
            raise ValueError(
                f"current_dt is None before block {block_index} "
                f"of trip {trip_id}. Previous block probably had missing end time."
            )

        if block_index == 1:
            block_start_dt = current_dt
        else:
            block_start_dt = datetime_at_or_after(current_dt, block_start_time)

        last_flight_end_dt = block_start_dt
        flights = block.get("flights", [])

        for flight_index, flight in enumerate(flights):
            is_first_flight_in_block = flight_index == 0
            is_last_flight_in_block = flight_index == len(flights) - 1

            if is_first_flight_in_block:
                flight_start_dt = datetime_at_or_after(
                    block_start_dt,
                    flight.get("start")
                )
            else:
                flight_start_dt = datetime_at_or_after(
                    last_flight_end_dt,
                    flight.get("start")
                )

            flight_end_dt = datetime_at_or_after(
                flight_start_dt,
                flight.get("end")
            )

            if trip_start_dt is None:
                trip_start_dt = flight_start_dt

            trip_end_dt = flight_end_dt

            record = {
                "start_datetime": datetime_to_master_string(flight_start_dt),
                "start_date": flight_start_dt.date().isoformat(),
                "departure": flight.get("departure"),

                "end_datetime": datetime_to_master_string(flight_end_dt),
                "end_date": flight_end_dt.date().isoformat(),
                "arrival": flight.get("arrival"),

                # DH / BUS stay as route flags.
                "route_flags": extract_route_flags(flight),

                # SBG / SBA are saved as code, not route_flags.
                "code": extract_flight_code(flight),

                # Block-level info is stored only on the last flight of the block,
                # same idea as rest.
                "rest": block.get("rest")
                    if is_last_flight_in_block and block.get("rest") != "-"
                    else None,

                "block": block.get("block")
                    if is_last_flight_in_block
                    else None,

                "credit": block.get("credit")
                    if is_last_flight_in_block
                    else None,

                "duty": block.get("duty")
                    if is_last_flight_in_block
                    else None,
            }

            flight_records.append(record)
            last_flight_end_dt = flight_end_dt

        block_end_time = block.get("end")

        if block_end_time is None:
            raise ValueError(
                f"Missing block end time in trip {trip_id}, block {block_index}"
            )

        block_end_dt = datetime_at_or_after(
            last_flight_end_dt,
            block_end_time
        )

        rest_td = parse_duration(block.get("rest"))

        if rest_td is not None:
            current_dt = block_end_dt + rest_td
        else:
            current_dt = block_end_dt

    if trip_start_dt is not None and trip_end_dt is not None:
        total_days_gone = (
            trip_end_dt.date() - trip_start_dt.date()
        ).days + 1
    else:
        total_days_gone = None

    return {
        "trip_id": trip_id,

        "trip_start": datetime_to_master_string(trip_start_dt),
        "trip_end": datetime_to_master_string(trip_end_dt),

        "premium": safe_float(trip.get("premium")),
        "tafb": trip.get("tafb"),

        "per_diem": safe_float(trip.get("per_diem")),
        "ldgs": safe_int(trip.get("ldgs")),

        "credit_time": trip.get("credit_time"),
        "duty_time": trip.get("duty_time"),
        "block_time": trip.get("block_time"),
        "total_blocks": safe_int(trip.get("total_blocks")),

        "total_days_gone": total_days_gone,
        "flights": flight_records
    }

def minutes_to_decimal_hours(total_minutes, decimals=2):
    """
    Converts:
        4350 -> 72.5
        2559 -> 42.65
    """
    return round(total_minutes / 60, decimals)

def hhmm_to_minutes(value):
    """
    Converts:
        '42:39' -> 2559 minutes
        None    -> 0
    """
    if value is None:
        return 0

    hours, minutes = value.split(":")
    return int(hours) * 60 + int(minutes)


def minutes_to_hhmm(total_minutes):
    """
    Converts:
        2559 -> '42:39'
    """
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours}:{minutes:02d}"

def creating_master_line(trips, lines):
    master_lines = {}
    sba_sbg_run_state = {}
    
    for line in lines["lines"]:

        total_BT_minutes = 0
        total_CT_minutes = 0
        total_DT_minutes = 0
        total_tafb_minutes = 0

        total_DD = 0
        total_DO = 0

        total_trip_count = 0

        # Used only for avg_BT / avg_CT / avg_DT.
        # Not saved in final master_lines.
        total_blocks = 0

        total_trip_block_minutes = 0
        total_trip_credit_minutes = 0
        total_trip_duty_minutes = 0

        # Used only for avg_rest.
        total_rest_minutes = 0
        total_rest_count = 0

        line_num = line["line_number"]

        master_lines[line_num] = {
            "tot_BT": None,
            "tot_CT": None,
            "tot_DT": None,
            "tot_tafb": None,

            "tot_DD": None,
            "tot_DO": None,

            "avg_BT": None,
            "avg_CT": None,
            "avg_DT": None,
            "avg_tafb": None,
            "avg_rest": None,

            "PPs": []
        }

        for pp in line["pay_periods"]:
            pp_DD = safe_int(pp.get("DD"))

            if pp.get("DO") is None:
                pp_DO = 28 - pp_DD
            else:
                pp_DO = safe_int(pp.get("DO"))

            # These totals come from the Lines package.
            total_BT_minutes += hhmm_to_minutes(pp.get("BT"))
            total_CT_minutes += hhmm_to_minutes(pp.get("CT"))
            total_DD += pp_DD
            total_DO += pp_DO

            master_pp = {
                "pp": pp.get("pp"),
                "BT": pp.get("BT"),
                "CT": pp.get("CT"),
                "DD": pp_DD,
                "DO": pp_DO,
                "assignments": []
            }

            for assignment in pp["assignments"]:
                master_assignment = build_master_assignment(assignment, trips)
                master_pp["assignments"].append(master_assignment)

                # Skip VTO / RA / RB / SA / SB / VOR / etc.
                if "trip_id" not in master_assignment or "flights" not in master_assignment:
                    continue

                # Skip missing-trip or bad trip assignments.
                if master_assignment.get("error"):
                    continue

                total_trip_count += 1

                trip_blocks = safe_int(master_assignment.get("total_blocks"))
                total_blocks += trip_blocks

                block_minutes = trip_duration_to_minutes(
                    master_assignment.get("block_time")
                )

                credit_minutes = trip_duration_to_minutes(
                    master_assignment.get("credit_time")
                )

                duty_minutes = trip_duration_to_minutes(
                    master_assignment.get("duty_time")
                )

                tafb_minutes = trip_duration_to_minutes(
                    master_assignment.get("tafb")
                )

                total_trip_block_minutes += block_minutes
                total_trip_credit_minutes += credit_minutes
                total_trip_duty_minutes += duty_minutes

                total_DT_minutes += duty_minutes
                total_tafb_minutes += tafb_minutes

                # Average rest is based on the rest values stored in flights.
                # Since only the last flight of each block carries rest,
                # this naturally averages rest per rest period.
                for flight in master_assignment.get("flights", []):
                    rest_value = flight.get("rest")

                    if rest_value is not None:
                        total_rest_minutes += trip_duration_to_minutes(rest_value)
                        total_rest_count += 1

            master_lines[line_num]["PPs"].append(master_pp)

        # Final line totals.
        master_lines[line_num]["tot_BT"] = minutes_to_decimal_hours(total_BT_minutes)
        master_lines[line_num]["tot_CT"] = minutes_to_decimal_hours(total_CT_minutes)
        master_lines[line_num]["tot_DT"] = minutes_to_decimal_hours(total_DT_minutes)
        master_lines[line_num]["tot_tafb"] = minutes_to_decimal_hours(total_tafb_minutes)

        master_lines[line_num]["tot_DD"] = total_DD
        master_lines[line_num]["tot_DO"] = total_DO

        # Averages per block.
        if total_blocks > 0:
            master_lines[line_num]["avg_BT"] = minutes_to_decimal_hours(
                total_trip_block_minutes / total_blocks
            )
            master_lines[line_num]["avg_CT"] = minutes_to_decimal_hours(
                total_trip_credit_minutes / total_blocks
            )
            master_lines[line_num]["avg_DT"] = minutes_to_decimal_hours(
                total_trip_duty_minutes / total_blocks
            )
        else:
            master_lines[line_num]["avg_BT"] = 0
            master_lines[line_num]["avg_CT"] = 0
            master_lines[line_num]["avg_DT"] = 0

        # Average TAFB per trip.
        if total_trip_count > 0:
            master_lines[line_num]["avg_tafb"] = minutes_to_decimal_hours(
                total_tafb_minutes / total_trip_count
            )
        else:
            master_lines[line_num]["avg_tafb"] = 0

        # Average rest per rest period.
        # If there are no rests, set to 0.
        if total_rest_count > 0:
            master_lines[line_num]["avg_rest"] = minutes_to_decimal_hours(
                total_rest_minutes / total_rest_count
            )
        else:
            master_lines[line_num]["avg_rest"] = 0

    return master_lines