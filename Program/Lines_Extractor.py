#Simply run: lines = parse_line_report_pdf(pdf_path, first_calendar_page=3)

import re
import pdfplumber
from datetime import datetime, timedelta


#DOWS = {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"}


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


ASSIGNMENT_PATTERN = r"(?:\d+|VTO|VOR|RA|SA|RB|SB)"
TIME_PATTERN = r"(?:[01]\d|2[0-3])[0-5]\d"


TIME_PATTERN = r"(?:[01]\d|2[0-3])[0-5]\d"
ASSIGNMENT_PATTERN = r"(?:\d+|VTO|VOR|RA|SA|RB|SB)"


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

        if token.isdigit():
            start_time_info = find_start_time_for_trip(
                words=words,
                trip_word=assignment_word,
            )

            start_time = None
            date_column_index = item["column_index"]

            if start_time_info is not None:
                start_time_token = start_time_info["token"]
                start_time_word = start_time_info["word"]

                start_time = format_hhmm(start_time_token)

                # Important:
                # Use the time to choose left/right date only when near a date boundary.
                date_column_index = choose_trip_column_by_time(
                            x=word_center(assignment_word),
                            start_time_token=start_time_token,
                            columns=columns,
                            fallback_index=item["column_index"],
                            boundary_tolerance=8,
                            )

            assignments.append({
                "date": columns[date_column_index]["date"],
                "start_time": start_time,
                "type": "trip",
                "value": int(token),
            })

        else:
            # VTO / VOR / RA / SA / RB / SB do not need time logic.
            assignments.append({
                "date": columns[item["column_index"]]["date"],
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