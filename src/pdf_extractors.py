#simply run: trips = extract_trips_from_pdf(pdf_path,first_page=2)
#Simply run: lines = parse_line_report_pdf(pdf_path, first_calendar_page=3)
#Simply run bid_period = matching_bid_period(lines_pdf_path, trips_pdf_path)

import re
import pdfplumber
from datetime import datetime, timedelta

#------------------------------------------------------------------------------------------------

def get_bid_period_from_pdf(pdf_path):
    """
    Opens a PDF and returns the number after 'Bid Period:'.

    Example match:
        Bid Period: 2604

    Returns:
        str | None
    """

    pattern = re.compile(r"Bid\s*Period\s*:\s*(\d+)", re.IGNORECASE)

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            match = pattern.search(text)
            if match:
                return match.group(1)

    return None


def matching_bid_period(pdf_path_1, pdf_path_2):
    """
    Returns the bid period number if both PDFs have the same Bid Period.
    Otherwise returns None.
    """

    bid_period_1 = get_bid_period_from_pdf(pdf_path_1)
    bid_period_2 = get_bid_period_from_pdf(pdf_path_2)

    if bid_period_1 is not None and bid_period_1 == bid_period_2:
        return bid_period_1

    return None

#---------------------------------------------------------------------
#simply run: trips = extract_trips_from_pdf(pdf_path,first_page=2)
#TRIPS working Progress report + EXTRA INFO
TIME_RE = r"\(\d{2}\)\d{2}:\d{2}"
DUR_RE = r"\d+h\d{2}"

def get_first_match(pattern, text, default=None):
    match = re.search(pattern, text)
    return match.group(1) if match else default


def get_first_float(pattern, text, default=None):
    value = get_first_match(pattern, text, default=None)
    return float(value) if value is not None else default


def get_first_int(pattern, text, default=None):
    value = get_first_match(pattern, text, default=None)
    return int(value) if value is not None else default

def clean_time(value):
    """
    Converts:
        (16)20:43 -> 20:43
        (00)04:09 -> 04:09
    """
    if value is None:
        return None

    return re.sub(r"^\(\d{2}\)", "", value)


def group_words_by_line(words, tolerance=2):
    lines = []

    for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
        for line in lines:
            if abs(line[0]["top"] - word["top"]) <= tolerance:
                line.append(word)
                break
        else:
            lines.append([word])

    return [sorted(line, key=lambda w: w["x0"]) for line in lines]


def find_trip_anchors(page):
    """
    Finds each 'Trip Id: ###' on the page.
    Needed because each page can contain several trip tables.
    """
    words = page.extract_words(x_tolerance=1, y_tolerance=3) or []
    anchors = []

    for line in group_words_by_line(words):
        text_parts = [w["text"] for w in line]

        for i in range(len(text_parts) - 2):
            if (
                text_parts[i] == "Trip"
                and text_parts[i + 1] == "Id:"
                and text_parts[i + 2].isdigit()
            ):
                anchors.append({
                    "trip_id": int(text_parts[i + 2]),
                    "x0": line[i]["x0"],
                    "top": line[i]["top"],
                })

    return sorted(anchors, key=lambda a: (a["top"], a["x0"]))


def make_trip_crops(page):
    """
    Builds a crop box for each trip table.

    The crop is necessary because the PDF pages can have left/right tables.
    Without cropping, pdfplumber may mix text from different tables.
    """
    anchors = find_trip_anchors(page)
    page_middle = page.width / 2

    for anchor in anchors:
        anchor["column"] = 0 if anchor["x0"] < page_middle else 1

    crops = []

    for anchor in anchors:
        next_trip_same_column = [
            other
            for other in anchors
            if other["column"] == anchor["column"]
            and other["top"] > anchor["top"] + 5
        ]

        bottom = min(
            [other["top"] for other in next_trip_same_column],
            default=page.height - 10,
        )

        if anchor["column"] == 0:
            x0, x1 = 0, page_middle - 2
        else:
            x0, x1 = page_middle + 2, page.width

        crops.append((
            x0,
            max(0, anchor["top"] - 3),
            x1,
            min(page.height, bottom - 2),
        ))

    return crops


def split_route(route_raw):
    """
    Examples:
        SDF-PHL         -> departure SDF, arrival PHL, route_flags []
        SDF-IRO-PHL     -> departure SDF, arrival PHL, route_flags ['IRO']
        SDF-BDL(C)      -> departure SDF, arrival BDL, route_flags ['C']
        SDF-IRO-BDL(C)  -> departure SDF, arrival BDL, route_flags ['IRO', 'C']
    """
    parts = route_raw.split("-")

    airports = []
    route_flags = []

    for part in parts:
        if part == "IRO":
            route_flags.append("IRO")
            continue

        # Handles airport with parenthetical flag, like BDL(C)
        match = re.match(r"^([A-Z]{3})(?:\(([A-Z]+)\))?$", part)

        if match:
            airport = match.group(1)
            flag = match.group(2)

            airports.append(airport)

            if flag:
                route_flags.append(flag)
        else:
            airports.append(part)

    departure = airports[0] if airports else None
    arrival = airports[-1] if airports else None

    return departure, arrival, route_flags


def parse_flight_line(line):
    """
    Extracts only:
    - flight
    - route_raw
    - departure
    - arrival
    - route_flags, such as IRO
    - start
    - end
    """

    match = re.match(r"^\d+\s+\([^)]*\)[A-Za-z]{0,2}\s+(.*)$", line)

    if not match:
        return None

    body = match.group(1)

    route_match = re.search(
        r"[A-Z]{3}(?:\([A-Z]\))?(?:-(?:IRO|[A-Z]{3}(?:\([A-Z]\))?))+",
        body,
    )

    if not route_match:
        return None

    flight = body[:route_match.start()].strip()
    route_raw = route_match.group(0)
    after_route = body[route_match.end():].strip()

    time_match = re.match(
        rf"(?P<start>{TIME_RE})\s+(?P<end>{TIME_RE})",
        after_route,
    )

    if not time_match:
        return None

    departure, arrival, route_flags = split_route(route_raw)

    return {
        "flight": flight,
        "route_raw": route_raw,
        "departure": departure,
        "arrival": arrival,
        "route_flags": route_flags,
        "start": clean_time(time_match.group("start")),
        "end": clean_time(time_match.group("end")),
    }


def parse_trip_text(text):
    """
    Parses one trip table into a dictionary.

    Extracts:
        Trip-level:
            - trip_id
            - lines
            - total_blocks
            - tafb
            - premium
            - duty_time
            - block_time
            - credit_time
            - per_diem
            - ldgs

        Block-level:
            - start
            - end
            - duty
            - block
            - rest
            - credit
            - flights
    """

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    full_text = "\n".join(lines)

    trip_id = int(re.search(r"Trip Id:\s*(\d+)", full_text).group(1))

    lines_match = re.search(r"Lines:\s*([^\n]+)", full_text)
    line_numbers = []

    if lines_match:
        line_numbers = [int(x) for x in re.findall(r"\d+", lines_match.group(1))]

    trip = {
        "trip_id": trip_id,
        "lines": line_numbers,
        "total_blocks": 0,

        # Trip summary values
        "tafb": get_first_match(r"TAFB:\s*(\d+h\d{2})", full_text),
        "premium": get_first_float(r"Premium\s+([\d.]+)", full_text),
        "duty_time": get_first_match(r"Duty Time:\s*(\d+h\d{2})", full_text),
        "block_time": get_first_match(r"Block Time:\s*(\d+h\d{2})", full_text),
        "credit_time": get_first_match(r"Credit Time:\s*(\d+h\d{2}[A-Z]?)", full_text),
        "per_diem": get_first_float(r"per Diem\s+([\d.]+)", full_text),
        "ldgs": get_first_int(r"LDGS\s+(\d+)", full_text),

        # Block list
        "blocks": [],
    }

    current_block = None

    for line in lines:
        # New block starts with something like:
        # (15)19:43 1h00 Duty 8h26
        # (16)21:00 0h00 Duty 0h00
        duty_match = re.match(
            rf"^(?P<block_start>{TIME_RE})\s+"
            rf"(?P<report_time>{DUR_RE})\s+"
            rf"Duty\s+(?P<duty>{DUR_RE})",
            line,
        )

        if duty_match:
            current_block = {
                "start": clean_time(duty_match.group("block_start")),
                "end": None,

                # New block-level fields
                "duty": duty_match.group("duty"),
                "block": None,
                "rest": None,
                "credit": None,

                "flights": [],
            }
            trip["blocks"].append(current_block)
            continue

        if current_block is None:
            continue

        # Block ends can look like:
        # (00)04:09 0h15 Credit 4h13D
        # (03)07:00 0h00 Rest -
        # (20)00:51 0h15 Credit 4h00M
        #
        # But this should NOT count as a block end:
        # (16)21:00 0h00 Duty 0h00
        block_end_match = re.match(
            rf"^(?P<block_end>{TIME_RE})\s+"
            rf"(?P<ground_time>{DUR_RE})"
            rf"(?:\s+(?P<after>.*))?$",
            line,
        )

        if block_end_match:
            after = block_end_match.group("after") or ""

            if not after.startswith("Duty"):
                current_block["end"] = clean_time(block_end_match.group("block_end"))

        flight = parse_flight_line(line)

        if flight:
            # For the first flight in each block, use the duty/block start time,
            # not the actual flight departure time.
            if not current_block["flights"]:
                flight["start"] = current_block["start"]

            current_block["flights"].append(flight)

        # Block-level Block.
        # Important: this should match "Block 3h19",
        # but NOT "Block Time: 3h19".
        block_match = re.search(
            r"\bBlock\s+(?!Time:)(\d+h\d{2})",
            line,
        )

        if block_match:
            current_block["block"] = block_match.group(1)

        # Block-level Rest.
        rest_match = re.search(
            r"\bRest\s+(-|\d+h\d{2})",
            line,
        )

        if rest_match:
            current_block["rest"] = rest_match.group(1)

        # Block-level Credit.
        # Important: this should match "Credit 5h30M",
        # but NOT "Credit Time: 5h30D".
        credit_match = re.search(
            r"\bCredit\s+(?!Time:)(\d+h\d{2}[A-Z]?)",
            line,
        )

        if credit_match:
            current_block["credit"] = credit_match.group(1)

    trip["total_blocks"] = len(trip["blocks"])

    return trip

def extract_trips_from_pdf(
    pdf_path,
    first_page=2,
    last_page=None,
    stop_after_empty_pages=4,
    progress_callback=None,
):
    """
    Returns a dictionary keyed by Trip ID.

    progress_callback:
        Optional function that receives a progress dictionary.

    Example progress data:
        {
            "current": 10,
            "total": 150,
            "page": 11,
            "trips_on_page": 8,
            "total_trips": 72,
            "status": "running",
            "message": "Extracting page 11 of 150",
        }
    """

    def send_progress(
        current,
        total,
        page_number=None,
        trips_on_page=0,
        total_trips=0,
        status="running",
        message=None,
    ):
        if progress_callback is None:
            return

        progress_callback({
            "current": current,
            "total": total,
            "page": page_number,
            "trips_on_page": trips_on_page,
            "total_trips": total_trips,
            "status": status,
            "message": message,
        })

    trips = {}
    empty_pages_in_a_row = 0

    with pdfplumber.open(pdf_path) as pdf:
        total_pdf_pages = len(pdf.pages)

        start_index = first_page - 1
        end_index = last_page if last_page is not None else total_pdf_pages

        start_index = max(0, start_index)
        end_index = min(end_index, total_pdf_pages)

        total_pages_to_process = end_index - start_index

        if total_pages_to_process <= 0:
            send_progress(
                current=0,
                total=0,
                status="done",
                message="No pages to process.",
            )
            return trips

        send_progress(
            current=0,
            total=total_pages_to_process,
            status="starting",
            message="Starting trip extraction...",
        )

        for page_index in range(start_index, end_index):
            page_number = page_index + 1
            current_progress = page_index - start_index + 1

            page = pdf.pages[page_index]
            trips_found_on_page = 0

            try:
                for bbox in make_trip_crops(page):
                    cropped = page.crop(bbox)
                    text = cropped.extract_text(x_tolerance=1, y_tolerance=3) or ""

                    if "Trip Id:" not in text:
                        continue

                    trip = parse_trip_text(text)
                    trips[trip["trip_id"]] = trip
                    trips_found_on_page += 1

            finally:
                page.close()

            if trips_found_on_page == 0:
                empty_pages_in_a_row += 1
            else:
                empty_pages_in_a_row = 0

            send_progress(
                current=current_progress,
                total=total_pages_to_process,
                page_number=page_number,
                trips_on_page=trips_found_on_page,
                total_trips=len(trips),
                status="running",
                message=(
                    f"Extracting trips: page {page_number} "
                    f"({current_progress} of {total_pages_to_process})"
                ),
            )

            if (
                stop_after_empty_pages is not None
                and empty_pages_in_a_row >= stop_after_empty_pages
            ):
                send_progress(
                    current=current_progress,
                    total=total_pages_to_process,
                    page_number=page_number,
                    trips_on_page=trips_found_on_page,
                    total_trips=len(trips),
                    status="stopped",
                    message=(
                        f"Stopped at page {page_number}: "
                        f"{empty_pages_in_a_row} empty pages in a row."
                    ),
                )
                break

        send_progress(
            current=min(current_progress, total_pages_to_process),
            total=total_pages_to_process,
            page_number=page_number,
            total_trips=len(trips),
            status="done",
            message=f"Finished extracting {len(trips)} trips.",
        )

    return trips


#-------------------------------------------------------------------------------------------------------------------------

#Simply run: lines = parse_line_report_pdf(pdf_path, first_calendar_page=3)

def find_trip_number_under_special_code(
    words,
    code_word,
    x_tolerance=20,
    min_y_distance=2,
    max_y_distance=28,
):
    """
    Finds the trip number printed underneath a special line-report code
    such as SBA, SBA3, SBG, SBG3.

    Visual example:

        SBA
        1964

    or:

        SBA3
        1965
    """

    code_x = word_center(code_word)

    candidates = []

    for w in words:
        token = w["text"].strip()

        if not re.fullmatch(r"\d+", token):
            continue

        y_distance = w["top"] - code_word["top"]

        # Must be below the SBA/SBG word.
        if y_distance < min_y_distance or y_distance > max_y_distance:
            continue

        x_distance = abs(word_center(w) - code_x)

        if x_distance > x_tolerance:
            continue

        candidates.append({
            "word": w,
            "token": token,
            "x_distance": x_distance,
            "y_distance": y_distance,
        })

    if not candidates:
        return None

    candidates.sort(
        key=lambda c: (
            c["y_distance"],
            c["x_distance"],
        )
    )

    return candidates[0]

def parse_bid_range(page_text):
    m = re.search(
        r"Bid Period Date Range:\s*(\d{1,2}[A-Za-z]{3}\d{4})\s*-\s*(\d{1,2}[A-Za-z]{3}\d{4})",
        page_text,
    )
    if not m:
        raise ValueError("Could not find Bid Period Date Range")

    start = datetime.strptime(m.group(1), "%d%b%Y").date()
    end = datetime.strptime(m.group(2), "%d%b%Y").date()
    return start, end


def parse_domicile(page_text):
    m = re.search(r"Domicile:\s*([A-Z]{3})", page_text)
    if not m:
        raise ValueError("Could not find domicile")
    return m.group(1)


def words_on_same_line(words, top, tolerance=2.0):
    return [w for w in words if abs(w["top"] - top) <= tolerance]

def build_package_metadata(bid_start, bid_end):
    return {
        "bid_period_date_range": {
            "start": bid_start.isoformat(),
            "end": bid_end.isoformat(),
        },
        "pay_period_date_ranges": {
            "PP1": {
                "start": bid_start.isoformat(),
                "end": (bid_start + timedelta(days=27)).isoformat(),
            },
            "PP2": {
                "start": (bid_start + timedelta(days=28)).isoformat(),
                "end": (bid_start + timedelta(days=55)).isoformat(),
            },
        },
    }

PP_TOP_FROM_CT_OFFSET = 9.55


def find_pp_anchors(block_words):
    """
    Finds PP1 / PP2 sections using the CT: rows instead of the PP1/PP2 labels.

    This is more reliable because in VTO/VOR/RA/etc. pages,
    pdfplumber can merge hidden assignment text with the PP label.
    """
    ct_words = [
        w for w in block_words
        if w["text"] == "CT:"
        and 35 <= w["x0"] <= 100
    ]

    ct_words = sorted(ct_words, key=lambda w: w["top"])

    pp_anchors = []

    for i, ct_word in enumerate(ct_words[:2], start=1):
        pp_anchors.append({
            "pp_index": i,
            "top": ct_word["top"] - PP_TOP_FROM_CT_OFFSET,
        })

    return pp_anchors

def find_line_blocks(words, domicile):
    """
    Finds line blocks such as:
        SDF 1
        SDF 17
        SDF 18

    Returns y-ranges for each line block.
    """
    starts = []

    for w in words:
        if w["text"] == domicile and w["x0"] < 80 and w["top"] > 120:
            same_line = words_on_same_line(words, w["top"], tolerance=1.5)

            possible_numbers = [
                x for x in same_line
                if x["x0"] > w["x1"]
                and x["x0"] < 100
                and re.fullmatch(r"\d+", x["text"])
            ]

            if possible_numbers:
                starts.append({
                    "line_number": int(possible_numbers[0]["text"]),
                    "top": w["top"],
                })

    starts = sorted(starts, key=lambda x: x["top"])

    for i, block in enumerate(starts):
        if i + 1 < len(starts):
            block["bottom"] = starts[i + 1]["top"] - 5
        else:
            block["bottom"] = 99999

    return starts


def get_metric(words, pp_top, label):
    """
    Gets CT, BT, DO, DD from the left side of each PP row.
    """
    for w in words:
        if (
            w["text"] == label
            and w["x0"] < 90
            and pp_top <= w["top"] <= pp_top + 45
        ):
            same_line = sorted(
                words_on_same_line(words, w["top"], tolerance=1.5),
                key=lambda x: x["x0"],
            )

            after_label = [x for x in same_line if x["x0"] > w["x1"]]
            if after_label:
                return after_label[0]["text"]

    return None


def get_date_columns(words, pp_top):
    """
    The date numbers are on the line just above the PP label.
    We take only the first 28 date columns.

    Example:
        PP1: 21, 22, 23 ... 17
        PP2: 18, 19, 20 ... 15

    We ignore the extra '-- Mon 19' or '-- Mon 17' text after the 28-day grid.
    """
    date_words = [
        w for w in words
        if pp_top - 15 <= w["top"] <= pp_top - 5
        and re.fullmatch(r"\d{1,2}", w["text"])
        and w["x0"] > 80
    ]

    date_words = sorted(date_words, key=lambda w: w["x0"])
    return date_words[:28]

"""
def get_weekday_columns(words, pp_top):
    weekday_words = [
        w for w in words
        if pp_top - 28 <= w["top"] <= pp_top - 17
        and w["text"] in DOWS
        and w["x0"] > 80
    ]

    weekday_words = sorted(weekday_words, key=lambda w: w["x0"])
    return weekday_words[:28]
"""

def nearest_column_index(x, columns, max_distance=12.5):
    best_idx = None
    best_distance = None

    for i, col in enumerate(columns):
        distance = abs(col["center"] - x)

        if best_distance is None or distance < best_distance:
            best_idx = i
            best_distance = distance

    if best_distance is not None and best_distance <= max_distance:
        return best_idx

    return None


NORMAL_CODE_PATTERN = r"(?:VTO|VOR|RA|SA|RB|SB)"
SPECIAL_TRIP_CODE_PATTERN = r"(?:SBA\d*|SBG\d*)"

ASSIGNMENT_PATTERN = rf"(?:\d+|{NORMAL_CODE_PATTERN}|{SPECIAL_TRIP_CODE_PATTERN})"
TIME_PATTERN = r"(?:[01]\d|2[0-3])[0-5]\d"


def word_center(w):
    return (w["x0"] + w["x1"]) / 2


def format_hhmm(token):
    return f"{token[:2]}:{token[2:]}"


def is_valid_time_token(token):
    if not re.fullmatch(r"\d{4}", token):
        return False

    hh = int(token[:2])
    mm = int(token[2:])

    return 0 <= hh <= 23 and 0 <= mm <= 59


def time_minutes(token):
    return int(token[:2]) * 60 + int(token[2:])

def nearest_date_boundary(x, columns):
    """
    Finds the nearest boundary between two adjacent date columns.

    Returns:
        {
            "left_index": i,
            "right_index": i + 1,
            "boundary_x": ...,
            "distance": ...
        }
    """

    best = None

    for i in range(len(columns) - 1):
        left_center = columns[i]["center"]
        right_center = columns[i + 1]["center"]

        boundary_x = (left_center + right_center) / 2
        distance = abs(x - boundary_x)

        if best is None or distance < best["distance"]:
            best = {
                "left_index": i,
                "right_index": i + 1,
                "boundary_x": boundary_x,
                "distance": distance,
            }

    return best 

def choose_trip_column_by_time(
    x,
    start_time_token,
    columns,
    fallback_index=None,
    boundary_tolerance=8,
    noon_cutoff_minutes=12 * 60,
):
    """
    Chooses the correct date column for a trip.

    Boundary case:
        If the trip number is printed near the boundary between two dates:
            1200-2359 -> left/previous date
            0000-1159 -> right/next date

    Normal case:
        Use nearest date column.

    Safety:
        If nearest_column_index() returns None, fall back to the original
        assignment column so the parser does not crash.
    """

    boundary = nearest_date_boundary(x, columns)

    if boundary is not None and boundary["distance"] <= boundary_tolerance:
        if start_time_token is not None and is_valid_time_token(start_time_token):
            if time_minutes(start_time_token) >= noon_cutoff_minutes:
                return boundary["left_index"]
            else:
                return boundary["right_index"]

    nearest_idx = nearest_column_index(x, columns)

    if nearest_idx is not None:
        return nearest_idx

    if fallback_index is not None:
        return fallback_index

    return None


def find_assignment_words(words, target_top, columns, y_tolerance=3):
    assignments = []

    min_center = min(col["center"] for col in columns) - 15
    max_center = max(col["center"] for col in columns) + 15

    for w in words:
        token = w["text"].strip().upper()

        if not re.fullmatch(ASSIGNMENT_PATTERN, token):
            continue

        if abs(w["top"] - target_top) > y_tolerance:
            continue

        x = word_center(w)

        if x < min_center or x > max_center:
            continue

        idx = nearest_column_index(x, columns)

        if idx is None:
            continue

        assignments.append({
            "token": token,
            "word": w,
            "column_index": idx,
        })

    assignments.sort(
        key=lambda item: (
            item["word"]["x0"],
            item["word"]["top"],
        )
    )

    return assignments


def find_start_time_for_trip(words, trip_word, x_tolerance=30, y_window=60):
    """
    Finds the start time associated with a numeric trip ID.

    Visual stack usually looks like:

        310
        RDU RDU RDU
        2310
        33:16

    So we prefer a valid HHMM time below the trip number.
    """

    trip_x = word_center(trip_word)

    candidates = []

    for w in words:
        token = w["text"].strip()

        if not is_valid_time_token(token):
            continue

        x_distance = abs(word_center(w) - trip_x)
        y_distance = abs(w["top"] - trip_word["top"])

        if x_distance > x_tolerance:
            continue

        if y_distance > y_window:
            continue

        is_below_trip = w["top"] > trip_word["top"]

        candidates.append({
            "word": w,
            "token": token,
            "x_distance": x_distance,
            "y_distance": y_distance,
            "is_below_trip": is_below_trip,
        })

    if not candidates:
        return None

    candidates.sort(
        key=lambda c: (
            0 if c["is_below_trip"] else 1,
            c["y_distance"],
            c["x_distance"],
        )
    )

    return candidates[0]

def parse_pp(words, pp_anchor, bid_start):
    pp_index = pp_anchor["pp_index"]
    pp_top = pp_anchor["top"]

    date_words = get_date_columns(words, pp_top)

    if len(date_words) < 28:
        raise ValueError(f"Only found {len(date_words)} date columns for PP{pp_index}")

    pp_start = bid_start + timedelta(days=28 * (pp_index - 1))

    columns = []

    for i, date_word in enumerate(date_words):
        actual_date = pp_start + timedelta(days=i)

        columns.append({
            "index": i,
            "date": actual_date.isoformat(),
            "center": word_center(date_word),
        })

    assignment_words = find_assignment_words(
        words=words,
        target_top=pp_top,
        columns=columns,
    )

    assignments = []

    for item in assignment_words:
        token = item["token"]
        assignment_word = item["word"]
        date_column_index = item["column_index"]

        if token.isdigit():
            start_time_info = find_start_time_for_trip(
                words=words,
                trip_word=assignment_word,
            )

            start_time = None

            if start_time_info is not None:
                start_time_token = start_time_info["token"]
                start_time = format_hhmm(start_time_token)

                date_column_index = choose_trip_column_by_time(
                    x=word_center(assignment_word),
                    start_time_token=start_time_token,
                    columns=columns,
                    fallback_index=item["column_index"],
                    boundary_tolerance=8,
                )

            if date_column_index is None:
                date_column_index = item["column_index"]

            assignments.append({
                "date": columns[date_column_index]["date"],
                "start_time": start_time,
                "type": "trip",
                "value": int(token),
            })

        elif re.fullmatch(SPECIAL_TRIP_CODE_PATTERN, token):
            trip_number_info = find_trip_number_under_special_code(
                words=words,
                code_word=assignment_word,
            )

            if trip_number_info is not None:
                assignments.append({
                    "date": columns[date_column_index]["date"],
                    "start_time": None,
                    "type": "trip",
                    "value": int(trip_number_info["token"]),
                    "line_code": token,
                })
            else:
                # Fallback: preserve the SBA/SBG code even if the trip number
                # could not be found underneath it.
                assignments.append({
                    "date": columns[date_column_index]["date"],
                    "type": "special_trip_code",
                    "value": token,
                    "trip_id": None,
                })

        else:
            assignments.append({
                "date": columns[date_column_index]["date"],
                "type": "code",
                "value": token,
            })

    return {
        "pp": f"PP{pp_index}",
        "CT": get_metric(words, pp_top, "CT:"),
        "BT": get_metric(words, pp_top, "BT:"),
        "DO": get_metric(words, pp_top, "DO:"),
        "DD": get_metric(words, pp_top, "DD:"),
        "assignments": assignments,
    }
    
def parse_line_report_page(page, bid_start, domicile):
    words = page.extract_words(
        x_tolerance=1,
        y_tolerance=2,
        keep_blank_chars=False
    )

    line_blocks = find_line_blocks(words, domicile)

    parsed_lines = []

    for line_block in line_blocks:
        block_words = [
            w for w in words
            if line_block["top"] - 5 <= w["top"] < line_block["bottom"]
        ]

        pp_anchors = find_pp_anchors(block_words)

        pp_data = []

        for pp_anchor in pp_anchors:
            pp_data.append(parse_pp(block_words, pp_anchor, bid_start))

        parsed_lines.append({
            "line_number": line_block["line_number"],
            "pay_periods": pp_data,
        })

    return parsed_lines


def parse_line_report_pdf(pdf_path, first_calendar_page=3):
    """
    first_calendar_page uses normal PDF page numbering.

    Example:
        first_calendar_page=5 means:
        skip pages 1-4, start extracting line calendar data on page 5.
    """

    first_calendar_index = first_calendar_page - 1

    with pdfplumber.open(pdf_path) as pdf:
        if first_calendar_index >= len(pdf.pages):
            raise ValueError(
                f"first_calendar_page={first_calendar_page} is beyond the end of the PDF. "
                f"The PDF only has {len(pdf.pages)} pages."
            )

        # Read metadata from the first actual calendar page, not PDF page 1.
        metadata_text = pdf.pages[first_calendar_index].extract_text(
            x_tolerance=1,
            y_tolerance=2
        ) or ""

        bid_start, bid_end = parse_bid_range(metadata_text)
        domicile = parse_domicile(metadata_text)

        result = build_package_metadata(bid_start, bid_end)
        result["lines"] = []

        for page in pdf.pages[first_calendar_index:]:
            page_lines = parse_line_report_page(
                page=page,
                bid_start=bid_start,
                domicile=domicile,
            )
            result["lines"].extend(page_lines)

    return result