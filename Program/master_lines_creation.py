from datetime import datetime, timedelta
import re


TRIP_TYPES = {"trip", "trips"}


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

        # Bus detection
        if "BUS" in flight_text or "BUS" in route_raw:
            flags.append("BUS")

        return unique_preserve_order(flags)


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

    if assignment_date is None or assignment_start_time is None:
        return {
            "trip_id": trip_id,
            "premium": trip.get("premium"),
            "tafb": trip.get("tafb"),
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
                "start_date": flight_start_dt.date().isoformat(),
                "departure": flight.get("departure"),

                "end_date": flight_end_dt.date().isoformat(),
                "arrival": flight.get("arrival"),

                "route_flags": extract_route_flags(flight),

                "rest": block.get("rest") if is_last_flight_in_block and block.get("rest") != "-" else None
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
        "premium": trip.get("premium"),
        "tafb": trip.get("tafb"),
        "total_days_gone": total_days_gone,
        "flights": flight_records
    }

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
    for line in lines['lines']:
        
        total_BT_minutes = 0
        total_CT_minutes = 0
        total_DD = 0
        total_DO = 0
        total_premium = 0.0

        line_num = line['line_number']
        master_lines[line_num] = {'tot_BT':None, 'tot_CT':None, 'tot_DD':None, 'tot_DO':None, 'tot_Premium':None, 'PPs':[]}

        for pp in line["pay_periods"]:
            total_BT_minutes += hhmm_to_minutes(pp.get("BT"))
            total_CT_minutes += hhmm_to_minutes(pp.get("CT"))
            total_DD += int(pp.get("DD"))
            total_DO += int(pp.get("DO") if pp.get("DO") is not None else (28-int(pp.get("DD"))))

            master_pp = {
                "pp": pp.get("pp"),
                "BT": pp.get("BT"),
                "CT": pp.get("CT"),
                "DD": int(pp.get("DD")),
                "DO": pp.get("DO") if pp.get("DO") is not None else (28-int(pp.get("DD"))),
                "assignments": []
            }

            for assignment in pp["assignments"]:
                master_assignment = build_master_assignment(assignment, trips)
                master_pp["assignments"].append(master_assignment)
                
                if "premium" in master_assignment:
                    total_premium += float(master_assignment.get("premium"))

            master_lines[line_num]["PPs"].append(master_pp)

        master_lines[line_num]["tot_BT"] = minutes_to_hhmm(total_BT_minutes)
        master_lines[line_num]["tot_BT_mins"] = total_BT_minutes
        master_lines[line_num]["tot_CT"] = minutes_to_hhmm(total_CT_minutes)
        master_lines[line_num]["tot_CT_mins"] = total_CT_minutes
        master_lines[line_num]["tot_DD"] = total_DD
        master_lines[line_num]["tot_DO"] = total_DO
        master_lines[line_num]["tot_Premium"] = total_premium
    
    return master_lines