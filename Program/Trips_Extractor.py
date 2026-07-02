#simply run: trips = extract_trips_from_pdf(pdf_path,first_page=2)
#Working

import re
import pdfplumber


TIME_RE = r"\(\d{2}\)\d{2}:\d{2}"
DUR_RE = r"\d+h\d{2}"

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
    """

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    full_text = "\n".join(lines)

    trip_id = int(re.search(r"Trip Id:\s*(\d+)", full_text).group(1))

    lines_match = re.search(r"Lines:\s*([^\n]+)", full_text)
    line_numbers = []

    if lines_match:
        line_numbers = [int(x) for x in re.findall(r"\d+", lines_match.group(1))]

    tafb_match = re.search(r"TAFB:\s*(\d+h\d{2})", full_text)
    premium_match = re.search(r"Premium\s+([\d.]+)", full_text)

    trip = {
        "trip_id": trip_id,
        "lines": line_numbers,
        "total_blocks": 0,
        "tafb": tafb_match.group(1) if tafb_match else None,
        "premium": float(premium_match.group(1)) if premium_match else None,
        "blocks": [],
    }

    current_block = None

    for line in lines:
        # New block starts with something like:
        # (15)19:43 1h00 Duty 8h26
        duty_match = re.match(
            rf"^(?P<block_start>{TIME_RE})\s+{DUR_RE}\s+Duty\s+{DUR_RE}",
            line,
        )

        if duty_match:
            current_block = {
                "start": clean_time(duty_match.group("block_start")),
                "end": None,
                "rest": None,
                "flights": [],
            }
            trip["blocks"].append(current_block)
            continue

        if current_block is None:
            continue

        # Block ends with something like:
        # (00)04:09 0h15 Credit 4h13D
        block_end_match = re.match(
            rf"^(?P<block_end>{TIME_RE})\s+"
            rf"(?P<ground_time>{DUR_RE})"
            rf"(?:\s+(?P<after>.*))?$",
            line,
        )

        if block_end_match:
            after = block_end_match.group("after") or ""

            # If the word immediately after the duration is Duty,
            # this is a new block start, not a block end.
            if not after.startswith("Duty"):
                current_block["end"] = clean_time(block_end_match.group("block_end"))

        flight = parse_flight_line(line)

        if flight:
            # For the first flight in each block, use the block/duty start time,
            # not the actual flight departure time.
            if not current_block["flights"]:
                flight["start"] = current_block["start"]

            current_block["flights"].append(flight)

        rest_match = re.search(r"Rest\s+(-|\d+h\d{2})", line)

        if rest_match:
            current_block["rest"] = rest_match.group(1)

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