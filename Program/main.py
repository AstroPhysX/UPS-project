from Trips_Extractor import extract_trips_from_pdf
from Lines_Extractor import parse_line_report_pdf
from master_lines_creation import creating_master_line
from master_to_pandas import master_lines_to_dataframe, sort_dataframe_by_conditions
from export_to_excel import export_master_lines_to_excel_table
import Processing_fucntions as pf
from concurrent.futures import ThreadPoolExecutor
import json
from pathlib import Path
from datetime import datetime

#-------------Config save and load

CONFIG_PATH = Path("bid_config.json")


def load_saved_config():
    """
    Loads previously saved user inputs.

    If the config file does not exist yet, returns an empty dictionary.
    """
    if not CONFIG_PATH.exists():
        return {}

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config):
    """
    Saves user inputs for future program runs.
    """
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)

#---------------General yes no prompts

def prompt_yes_no(message, default=None):
    """
    Prompts user for y/n.

    default can be:
        True  -> pressing Enter means yes
        False -> pressing Enter means no
        None  -> user must type y or n
    """
    while True:
        value = input(message).strip().lower()

        if value == "" and default is not None:
            return default

        if value in ("y", "yes"):
            return True

        if value in ("n", "no"):
            return False

        print("Please enter y or n.")

def prompt_date(message):
    while True:
        value = input(message).strip()

        try:
            datetime.strptime(value, "%Y-%m-%d")
            return value
        except ValueError:
            print("Please enter the date as YYYY-MM-DD.")

#-------------------Vacation prompting

def get_vacation_ranges_from_user_or_saved(config):
    """
    Uses saved vacation ranges if still accurate.
    Otherwise prompts user for new vacation ranges and saves them.
    """

    def prompt_vacation_ranges():
        vacation_ranges = []

        has_vacation = prompt_yes_no("Do you have vacation ranges? (y/n): ")

        if not has_vacation:
            return vacation_ranges

        while True:
            start = prompt_date("Vacation start date (YYYY-MM-DD): ")
            end = prompt_date("Vacation end date (YYYY-MM-DD): ")

            vacation_ranges.append({
                "start": start,
                "end": end,
            })

            more = prompt_yes_no("Add another vacation range? (y/n): ")
            if not more:
                break

        return vacation_ranges

    saved_vacation_ranges = config.get("vacation_ranges", [])

    if saved_vacation_ranges:
        print("\nSaved vacation ranges:")

        for i, vacation in enumerate(saved_vacation_ranges, start=1):
            print(f"  {i}. {vacation['start']} to {vacation['end']}")

        still_accurate = prompt_yes_no(
            "Are these vacation dates still accurate? (y/n): ",
            default=True,
        )

        if still_accurate:
            return saved_vacation_ranges

    vacation_ranges = prompt_vacation_ranges()

    config["vacation_ranges"] = vacation_ranges
    save_config(config)

    return vacation_ranges

#--------------------Training dates prompt

def get_training_dates_from_user_or_saved(config):
    """
    Uses saved training dates if still accurate.
    Otherwise prompts user for new training dates and saves them.
    """
    saved_training_start = config.get("training_start")
    saved_training_end = config.get("training_end")

    if saved_training_start and saved_training_end:
        print("\nSaved training dates:")
        print(f"  {saved_training_start} to {saved_training_end}")

        still_accurate = prompt_yes_no(
            "Are these training dates still accurate? (y/n): ",
            default=True,
        )

        if still_accurate:
            return saved_training_start, saved_training_end

    has_training = prompt_yes_no("Do you have training dates? (y/n): ")

    if has_training:
        training_start = prompt_date("Training start date YYYY-MM-DD: ")
        training_end = prompt_date("Training end date YYYY-MM-DD: ")
    else:
        training_start = None
        training_end = None

    config["training_start"] = training_start
    config["training_end"] = training_end
    save_config(config)

    return training_start, training_end

#------------Preferred start/end bid period off

def prompt_bid_period_days_off_preference():
    """
    Asks the user whether they want the bid period to start and/or end with days off.

    Returns:
        "start" -> prefer days off at the beginning of the bid period
        "end"   -> prefer days off at the end of the bid period
        "both"  -> prefer days off at both beginning and end
        "none"  -> no preference
    """

    print("\nBid period days-off preference:")
    print("  1. Start bid period with days off")
    print("  2. End bid period with days off")
    print("  3. Both start and end with days off")
    print("  4. No preference")

    choices = {
        "1": "start",
        "2": "end",
        "3": "both",
        "4": "none",
        "start": "start",
        "end": "end",
        "both": "both",
        "none": "none",
        "no": "none",
        "n": "none",
    }

    while True:
        answer = input("Choose 1, 2, 3, or 4: ").strip().lower()

        if answer in choices:
            return choices[answer]

        print("Invalid choice. Please enter 1, 2, 3, or 4.")



#-------------------------------sorting prompting
from datetime import date, datetime
import pandas as pd


from datetime import date, datetime
import pandas as pd


def is_calendar_date_column(col):
    """
    Returns True if the DataFrame column looks like a calendar date column.

    Handles:
        date/datetime objects
        2026-07-12
        07/12/2026
        07/12/26
        Wed, May 27
        Wednesday, May 27

    Avoids Python's warning about parsing month/day without a year.
    """

    if isinstance(col, (date, datetime, pd.Timestamp)):
        return True

    if not isinstance(col, str):
        return False

    text = col.strip()

    formats_with_year = [
        "%Y-%m-%d",      # 2026-07-12
        "%m/%d/%Y",      # 07/12/2026
        "%m/%d/%y",      # 07/12/26
    ]

    formats_without_year = [
        "%a, %b %d",     # Wed, May 27
        "%A, %b %d",     # Wednesday, May 27
        "%a, %B %d",     # Wed, May 27
        "%A, %B %d",     # Wednesday, May 27
    ]

    # Normal date formats that already include a year
    for fmt in formats_with_year:
        try:
            datetime.strptime(text, fmt)
            return True
        except ValueError:
            pass

    # Month/day formats without a year:
    # Add a safe dummy year to avoid Python's deprecation warning.
    # Use 2000 because it is a leap year, so Feb 29 remains valid.
    for fmt in formats_without_year:
        try:
            datetime.strptime(f"{text} 2000", f"{fmt} %Y")
            return True
        except ValueError:
            pass

    return False

def get_sortable_columns_from_df(df, include_text_columns=False):
    """
    Returns columns that can be offered to the user for sorting.

    By default:
        - excludes calendar date columns
        - excludes fully non-numeric text columns

    If include_text_columns=True:
        - excludes only calendar date columns
    """

    sortable_columns = []

    for col in df.columns:

        if is_calendar_date_column(col):
            continue

        if include_text_columns:
            sortable_columns.append(col)
            continue

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

        numeric_values = pd.to_numeric(cleaned, errors="coerce")

        if numeric_values.notna().any():
            sortable_columns.append(col)

    return sortable_columns

def prompt_sort_order_from_df(df):
    """
    Prompts the user for sorting priority based on actual DataFrame columns.

    Returns:
        [
            ["Blockiness", "desc"],
            ["Extra Vacation Days", "desc"],
            ["Total DO", "desc"],
        ]

    List-of-lists is JSON-friendly.
    Your sorting function can still use it because:
        for col, direction in sort_order:
            ...
    """

    sortable_columns = get_sortable_columns_from_df(df)

    if not sortable_columns:
        print("No sortable columns found.")
        return []

    print("\nAvailable sorting columns:")

    for i, col in enumerate(sortable_columns, start=1):
        print(f"  {i}. {col}")

    print("\nEnter the columns in the order you want to sort.")
    print("Example: 3,1,5")
    print("Leave blank for no custom sorting.")

    while True:
        answer = input("Sorting priority: ").strip()

        if answer == "":
            return []

        selected_numbers = [
            item.strip()
            for item in answer.split(",")
            if item.strip()
        ]

        try:
            selected_indexes = [int(item) for item in selected_numbers]
        except ValueError:
            print("Please enter numbers separated by commas.")
            continue

        invalid_indexes = [
            index for index in selected_indexes
            if index < 1 or index > len(sortable_columns)
        ]

        if invalid_indexes:
            print(f"Invalid choice(s): {invalid_indexes}")
            continue

        # Remove duplicates while preserving order
        seen = set()
        selected_columns = []

        for index in selected_indexes:
            col = sortable_columns[index - 1]

            if col in seen:
                continue

            seen.add(col)
            selected_columns.append(col)

        break

    sort_order = []

    for col in selected_columns:
        while True:
            direction = input(
                f"Sort '{col}' high-to-low or low-to-high? [high]: "
            ).strip().lower()

            if direction == "":
                direction = "high"

            if direction in ["high", "h", "desc", "descending", "high_to_low"]:
                sort_order.append([col, "desc"])
                break

            if direction in ["low", "l", "asc", "ascending", "low_to_high"]:
                sort_order.append([col, "asc"])
                break

            print("Please enter 'high' or 'low'.")

    return sort_order

def clean_saved_sort_order_for_df(saved_sort_order, df):
    """
    Removes saved sorting columns that are not valid for the current DataFrame.
    """

    sortable_columns = set(get_sortable_columns_from_df(df))

    cleaned_sort_order = []
    skipped_columns = []

    for rule in saved_sort_order:
        if len(rule) != 2:
            continue

        col, direction = rule

        if col in sortable_columns:
            cleaned_sort_order.append([col, direction])
        else:
            skipped_columns.append(col)

    return cleaned_sort_order, skipped_columns

def get_sort_order_from_user_or_saved(config, df):
    """
    Uses saved sort order if still acceptable.
    Otherwise prompts the user for a new one and saves it.
    """

    saved_sort_order = config.get("sort_order", [])

    if saved_sort_order:
        cleaned_sort_order, skipped_columns = clean_saved_sort_order_for_df(
            saved_sort_order,
            df,
        )

        if cleaned_sort_order:
            print("\nSaved sorting order:")

            for i, (col, direction) in enumerate(cleaned_sort_order, start=1):
                label = "low-to-high" if str(direction).lower() in ["asc", "ascending", "true"] else "high-to-low"
                print(f"  {i}. {col} ({label})")

            if skipped_columns:
                print("\nSkipped saved sorting columns not found in this DataFrame:")
                for col in skipped_columns:
                    print(f"  - {col}")

            use_saved = prompt_yes_no(
                "Use this saved sorting order? y/n [y]: ",
                default=True,
            )

            if use_saved:
                return cleaned_sort_order

    sort_order = prompt_sort_order_from_df(df)

    config["sort_order"] = sort_order
    save_config(config)

    return sort_order

#-------------------Output

from pathlib import Path
import platform
import re


def get_os_name():
    """
    Returns:
        'Windows'
        'Linux'
        'Darwin' for macOS
    """
    return platform.system()


def clean_filename(filename):
    """
    Makes filename safe for Windows/Linux.
    Removes extension if user accidentally provides one.
    """

    filename = filename.strip().strip('"').strip("'")

    # Remove extension if user types Bid_Results.xlsx
    filename = Path(filename).stem

    # Use Windows-safe rules because they are stricter than Linux
    filename = re.sub(r'[<>:"/\\|?*]', "_", filename)

    # Remove trailing spaces/dots, which Windows does not like
    filename = filename.strip().rstrip(".")

    return filename


def prompt_output_filename(default_name="Bid_Results"):
    """
    Always asks for filename.
    User should not provide extension.
    """

    while True:
        filename = input(
            f"Output filename without extension [{default_name}]: "
        ).strip()

        if filename == "":
            filename = default_name

        filename = clean_filename(filename)

        if filename:
            return filename

        print("Please enter a valid filename.")


def prompt_output_folder(config):
    """
    Prompts for output folder.

    Saves output folders separately by operating system.

    Config example:
        {
            "output_paths": {
                "Windows": "C:\\Users\\Jerome\\Documents\\Bid Results",
                "Linux": "/home/jerome/Bid Results"
            }
        }
    """

    os_name = get_os_name()

    output_paths = config.setdefault("output_paths", {})
    saved_output_path = output_paths.get(os_name)

    if saved_output_path:
        saved_folder = Path(saved_output_path).expanduser()

        print(f"\nSaved output folder for {os_name}:")
        print(f"  {saved_folder}")

        if saved_folder.exists() and saved_folder.is_dir():
            use_saved = prompt_yes_no(
                "Use this output folder? y/n [y]: ",
                default=True,
            )

            if use_saved:
                return saved_folder

        else:
            print("Saved output folder does not exist on this computer.")

    while True:
        folder_text = input("Enter output folder path: ").strip().strip('"').strip("'")

        if folder_text == "":
            folder = Path.cwd()
            print(f"Using current folder: {folder}")
        else:
            folder = Path(folder_text).expanduser()

        if folder.exists() and folder.is_dir():
            output_paths[os_name] = str(folder)
            save_config(config)
            return folder

        create_folder = prompt_yes_no(
            "That folder does not exist. Create it? y/n [y]: ",
            default=True,
        )

        if create_folder:
            folder.mkdir(parents=True, exist_ok=True)

            output_paths[os_name] = str(folder)
            save_config(config)

            return folder

def get_output_file_from_user_or_saved(config, extension=".xlsx"):
    """
    Always asks for filename.
    Reuses saved output folder for the current operating system.
    Returns the full output file path as a string.
    """

    filename = prompt_output_filename()
    output_folder = prompt_output_folder(config)

    output_file = output_folder / f"{filename}{extension}"

    return str(output_file)

def main():
    config = load_saved_config()

    with ThreadPoolExecutor(max_workers=3) as executor:

        trips_pdf_path = input("Enter the file path of TRIPS pdf (can be dragged and dropped onto this window): ").strip().strip('"\'')
        extracted_trips = executor.submit(extract_trips_from_pdf, trips_pdf_path, first_page=2)

        lines_pdf_path = input("Enter the file path of LINES pdf (can be dragged and dropped onto this window): ").strip().strip('"\'')
        extracted_lines = executor.submit(parse_line_report_pdf, lines_pdf_path, first_calendar_page=3)

        vacation_ranges = get_vacation_ranges_from_user_or_saved(config)
        
        training_start, training_end = get_training_dates_from_user_or_saved(config)

        bid_period_DO_preference = prompt_bid_period_days_off_preference()

        if not extracted_lines.done() or not extracted_trips.done():
            print("Please wait for PDF Extraction to complete, this may take a few minutes.....")

        lines = extracted_lines.result()
        trips = extracted_trips.result()
    
    bid_period_info = {x: lines[x] for x in ('bid_period_date_range','pay_period_date_ranges')}

    master_lines = creating_master_line(trips,lines)

    pf.add_blockiness_scores(master_lines,bid_period_info)
    pf.add_company_ticket_percentages(master_lines)
    new_vacation_range = pf.add_vacation_days_off_score(master_lines,vacation_ranges,bid_period_info,save_details=False)
    pf.add_training_fit_score(master_lines,training_start,training_end,bid_period_info)
    if bid_period_DO_preference != "none":
        pf.add_bid_edge_days_off(master_lines, bid_period_info,edge=bid_period_DO_preference)

    df = master_lines_to_dataframe(master_lines,bid_period_info)

    sort_order = get_sort_order_from_user_or_saved(config, df)

    df = sort_dataframe_by_conditions(df, sort_order)

    output_path = get_output_file_from_user_or_saved(config)

    print(output_path)

    export_master_lines_to_excel_table(df,output_path, training_start=training_start, training_end=training_end, vacation_ranges=new_vacation_range)

    while True:
        exit_program = prompt_yes_no("\nDo you want to exit the program? (y/n): ",default=False,)

        if exit_program:
            print("Exiting program.")
            break


        vacation_ranges = get_vacation_ranges_from_user_or_saved(config)
        
        training_start, training_end = get_training_dates_from_user_or_saved(config)

        bid_period_DO_preference = prompt_bid_period_days_off_preference()

        master_lines = creating_master_line(trips,lines)

        pf.add_blockiness_scores(master_lines,bid_period_info)
        pf.add_company_ticket_percentages(master_lines)
        new_vacation_range = pf.add_vacation_days_off_score(master_lines,vacation_ranges,bid_period_info,save_details=False)
        pf.add_training_fit_score(master_lines,training_start,training_end,bid_period_info)
        if bid_period_DO_preference != "none":
            pf.add_bid_edge_days_off(master_lines, bid_period_info,edge=bid_period_DO_preference)

        df = master_lines_to_dataframe(master_lines,bid_period_info)

        sort_order = get_sort_order_from_user_or_saved(config, df)

        df = sort_dataframe_by_conditions(df, sort_order)

        output_path = get_output_file_from_user_or_saved(config)

        print(output_path)

        export_master_lines_to_excel_table(df,output_path, training_start=training_start, training_end=training_end, vacation_ranges=new_vacation_range)


if __name__ == "__main__":
    main()