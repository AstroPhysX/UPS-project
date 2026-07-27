"""
Backend analysis pipeline for the UPS Bid Analyzer.

Step 1 refactor:
    - Keeps the existing project files in the same folder.
    - Owns PDF extraction and extraction caching.
    - Owns airport lookup creation and caching.
    - Creates master_lines.
    - Applies all processing functions.
    - Creates and optionally sorts the final DataFrame.
    - Contains no PySide6 imports and does not manipulate GUI widgets.

The GUI communicates with this class through ordinary callbacks:
    log_callback(message)
    trip_progress_callback(progress_dictionary)
"""

from __future__ import annotations

import re
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any, Callable

import pandas as pd

from pdf_extractors import extract_trips_from_pdf, parse_line_report_pdf
from master_lines_creation import creating_master_line
from master_to_pandas import (
    master_lines_to_dataframe,
    drop_empty_sort_columns,
    sort_dataframe_by_conditions,
)
import processing_functions as pf


LogCallback = Callable[[str], None]
TripProgressCallback = Callable[[dict[str, Any]], None]


DEFAULT_SORTING_SETTINGS = {
    "default_mode": "weighted",
    "weighting_style": "soft",
    "soft_max_weight": 3.0,
    "soft_min_weight": 1.0,
    "keep_score_columns": True,
}


def resource_path(relative_path: str) -> Path:
    """
    Return a resource path that works when running normally and in a
    PyInstaller/auto-py-to-exe build.

    This temporary step-1 helper assumes analysis_pipeline.py and airports.csv
    are in the same folder. It can be replaced by importlib.resources after the
    package/folder reorganization.
    """
    try:
        base_path = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    except AttributeError:
        base_path = Path(__file__).resolve().parent

    return base_path / relative_path


class AnalysisPipeline:
    """
    Coordinate the complete backend analysis process.

    The class intentionally knows nothing about PySide6. The callbacks supplied
    by the GUI are ordinary Python callables, so this pipeline can later be used
    by a CLI, tests, or a different interface.
    """

    def __init__(
        self,
        *,
        log_callback: LogCallback | None = None,
        trip_progress_callback: TripProgressCallback | None = None,
        airports_csv_path: str | Path | None = None,
    ) -> None:
        self._log_callback = log_callback
        self._trip_progress_callback = trip_progress_callback
        self._airports_csv_path = (
            Path(airports_csv_path)
            if airports_csv_path is not None
            else resource_path("airports.csv")
        )

        self._cached_lines: dict[str, Any] | None = None
        self._cached_trips: dict[str, Any] | None = None
        self._cached_pdf_key: tuple[str, str] | None = None

        self._cached_airport_lookup_key: tuple[str, str] | None = None
        self._cached_airport_lookup: dict[str, Any] | None = None
        self._cached_unmatched_airports: Any = None
        self._cached_matched_airports_df: pd.DataFrame | None = None

    @property
    def cached_pdf_key(self) -> tuple[str, str] | None:
        return self._cached_pdf_key

    @property
    def has_cached_pdf_data(self) -> bool:
        return (
            self._cached_pdf_key is not None
            and self._cached_trips is not None
            and self._cached_lines is not None
        )

    def has_cached_pdf_data_for(
        self,
        trips_pdf_path: str,
        lines_pdf_path: str,
    ) -> bool:
        return (
            self.has_cached_pdf_data
            and self._cached_pdf_key == (trips_pdf_path, lines_pdf_path)
        )

    def invalidate_if_paths_changed(
        self,
        trips_pdf_path: str,
        lines_pdf_path: str,
    ) -> bool:
        """
        Clear backend caches when the selected PDF paths no longer match them.

        Returns True when an existing cache was invalidated.
        """
        current_key = (trips_pdf_path, lines_pdf_path)

        if self._cached_pdf_key is None or current_key == self._cached_pdf_key:
            return False

        self.clear_cache()
        return True

    def clear_cache(self) -> None:
        """Clear extracted-PDF and airport-lookup caches."""
        self._cached_lines = None
        self._cached_trips = None
        self._cached_pdf_key = None

        self._cached_airport_lookup_key = None
        self._cached_airport_lookup = None
        self._cached_unmatched_airports = None
        self._cached_matched_airports_df = None

    def _log(self, message: str) -> None:
        if self._log_callback is not None:
            self._log_callback(message)

    def _report_trip_progress(self, progress_data: dict[str, Any]) -> None:
        if self._trip_progress_callback is not None:
            self._trip_progress_callback(progress_data)

    def _extract_trips_with_progress(
        self,
        trips_pdf_path: str,
    ) -> dict[str, Any]:
        try:
            return extract_trips_from_pdf(
                trips_pdf_path,
                first_page=2,
                progress_callback=self._report_trip_progress,
            )
        except TypeError as exc:
            if "progress_callback" not in str(exc):
                raise

            self._log(
                "Trip extractor does not support progress updates; "
                "loading trips without page progress."
            )
            self._report_trip_progress({
                "current": 0,
                "total": 0,
                "status": "running",
                "message": "Trip extractor does not support progress updates.",
            })

            trips = extract_trips_from_pdf(
                trips_pdf_path,
                first_page=2,
            )

            self._report_trip_progress({
                "current": 1,
                "total": 1,
                "status": "done",
                "message": f"Finished extracting {len(trips)} trips.",
                "total_trips": len(trips),
            })
            return trips

    def _extract_pdfs(
        self,
        inputs: dict[str, Any],
    ) -> tuple[dict[str, Any], dict[str, Any]]:
        pdf_key = (
            inputs["trips_pdf_path"],
            inputs["lines_pdf_path"],
        )

        if (
            self._cached_pdf_key == pdf_key
            and self._cached_trips is not None
            and self._cached_lines is not None
        ):
            self._log("Using already-loaded PDF data.")
            self._report_trip_progress({
                "current": 1,
                "total": 1,
                "status": "cached",
                "message": "Using already-loaded trip data.",
                "total_trips": len(self._cached_trips),
            })
            return self._cached_trips, self._cached_lines

        self._log("Extracting PDFs...")
        self._report_trip_progress({
            "current": 0,
            "total": 1,
            "status": "starting",
            "message": "Starting trip extraction...",
            "total_trips": 0,
        })

        with ThreadPoolExecutor(max_workers=2) as executor:
            trips_future = executor.submit(
                self._extract_trips_with_progress,
                inputs["trips_pdf_path"],
            )
            lines_future = executor.submit(
                parse_line_report_pdf,
                inputs["lines_pdf_path"],
                first_calendar_page=3,
            )

            trips = trips_future.result()
            lines = lines_future.result()

        self._cached_pdf_key = pdf_key
        self._cached_trips = trips
        self._cached_lines = lines

        self._log("PDF extraction complete.")
        return trips, lines

    def _get_airport_lookup(
        self,
        *,
        inputs: dict[str, Any],
        trips: dict[str, Any],
    ) -> dict[str, Any]:
        pdf_key = (
            inputs["trips_pdf_path"],
            inputs["lines_pdf_path"],
        )

        if (
            self._cached_airport_lookup_key == pdf_key
            and self._cached_airport_lookup is not None
        ):
            self._log("Using cached bid-period airport lookup...")
            return self._cached_airport_lookup

        if not self._airports_csv_path.exists():
            raise FileNotFoundError(
                "airports.csv was not found next to the program. "
                "Add it to the project folder and to the PyInstaller data files."
            )

        self._log("Collecting bid-period destination airports...")
        destination_codes = pf.collect_unique_arrival_destinations_from_trips(
            trips,
            ignore_sba_sbg=True,
        )

        self._log("Building bid-period airport lookup...")
        (
            airport_lookup,
            unmatched_airports,
            matched_airports_df,
        ) = pf.build_bid_period_airport_lookup(
            str(self._airports_csv_path),
            destination_codes,
            allowed_types=("large_airport", "medium_airport"),
        )

        self._cached_airport_lookup_key = pdf_key
        self._cached_airport_lookup = airport_lookup
        self._cached_unmatched_airports = unmatched_airports
        self._cached_matched_airports_df = matched_airports_df

        try:
            unmatched_count = len(unmatched_airports)
        except TypeError:
            unmatched_count = 0

        if unmatched_count:
            self._log(
                f"Airport lookup completed with {unmatched_count} "
                "unmatched destination code(s)."
            )
        else:
            self._log(
                "Airport lookup completed with all destination codes matched."
            )

        return airport_lookup

    @staticmethod
    def _pay_hours(value: Any) -> float:
        """Convert common UPS time formats to decimal hours."""
        if value is None:
            return 0.0

        if isinstance(value, (int, float)):
            return float(value)

        text = str(value).strip()
        if not text or text in {"-", "None", "nan", "NaN"}:
            return 0.0

        hhmm_match = re.fullmatch(r"(\d+):(\d{1,2})", text)
        if hhmm_match:
            return (
                int(hhmm_match.group(1))
                + int(hhmm_match.group(2)) / 60.0
            )

        trip_time_match = re.fullmatch(
            r"(\d+)h(\d{1,2})(?:[A-Za-z])?",
            text,
        )
        if trip_time_match:
            return (
                int(trip_time_match.group(1))
                + int(trip_time_match.group(2)) / 60.0
            )

        try:
            return float(
                text.replace(",", "").replace("$", "")
            )
        except ValueError:
            return 0.0

    @staticmethod
    def _pay_number(value: Any) -> float:
        if value is None:
            return 0.0

        if isinstance(value, (int, float)):
            return float(value)

        text = (
            str(value)
            .strip()
            .replace(",", "")
            .replace("$", "")
        )

        if not text or text == "-":
            return 0.0

        try:
            return float(text)
        except ValueError:
            return 0.0

    def _add_pay_fallback(
        self,
        master_lines: dict[Any, dict[str, Any]],
        hourly_rate: float,
        default_pp_guarantee_hours: float = 75.0,
    ) -> None:
        """
        Apply the documented pay formula when the current processing function
        encounters its local-variable ``pp`` bug.
        """
        hourly_rate = float(hourly_rate)

        for line_data in master_lines.values():
            pay_period_details: list[dict[str, Any]] = []
            total_extracted_credit = 0.0
            total_paid_credit = 0.0
            total_base_pay = 0.0
            total_premium = 0.0
            total_per_diem = 0.0

            for pp_index, pay_period in enumerate(
                line_data.get("PPs") or [],
                start=1,
            ):
                extracted_credit = self._pay_hours(
                    pay_period.get("CT")
                )
                guarantee_hours = float(
                    default_pp_guarantee_hours
                )
                paid_credit = max(
                    extracted_credit,
                    guarantee_hours,
                )
                guarantee_added = max(
                    0.0,
                    paid_credit - extracted_credit,
                )
                base_pay = paid_credit * hourly_rate

                total_extracted_credit += extracted_credit
                total_paid_credit += paid_credit
                total_base_pay += base_pay

                for assignment in (
                    pay_period.get("assignments") or []
                ):
                    total_premium += self._pay_number(
                        assignment.get("premium")
                    )
                    total_per_diem += self._pay_number(
                        assignment.get("per_diem")
                    )

                pay_period_details.append({
                    "pp": pay_period.get("pp") or f"PP{pp_index}",
                    "extracted_credit_hours": round(extracted_credit, 2),
                    "guarantee_hours": round(guarantee_hours, 2),
                    "paid_credit_hours": round(paid_credit, 2),
                    "guarantee_credit_added": round(guarantee_added, 2),
                    "base_pay": round(base_pay, 2),
                })

            guarantee_credit_added = max(
                0.0,
                total_paid_credit - total_extracted_credit,
            )
            taxable_pay = total_base_pay + total_premium
            cash_pay = taxable_pay + total_per_diem

            line_data["pay"] = {
                "hourly_rate": round(hourly_rate, 2),
                "pay_periods": pay_period_details,
                "extracted_CT": round(total_extracted_credit, 2),
                "paid_CT": round(total_paid_credit, 2),
                "guarantee_credit_added": round(guarantee_credit_added, 2),
                "total_base_pay": round(total_base_pay, 2),
                "total_premium": round(total_premium, 2),
                "total_per_diem": round(total_per_diem, 2),
                "taxable_pay_estimate": round(taxable_pay, 2),
                "cash_pay_estimate": round(cash_pay, 2),
            }

            line_data["pay_cash_estimate"] = round(cash_pay, 2)
            line_data["pay_taxable_estimate"] = round(taxable_pay, 2)
            line_data["pay_base"] = round(total_base_pay, 2)
            line_data["pay_premium"] = round(total_premium, 2)
            line_data["pay_per_diem"] = round(total_per_diem, 2)
            line_data["paid_CT"] = round(total_paid_credit, 2)
            line_data["guarantee_credit_added"] = round(
                guarantee_credit_added,
                2,
            )

    def build_dataframe(
        self,
        inputs: dict[str, Any],
        *,
        apply_sort: bool,
    ) -> tuple[pd.DataFrame, list[dict[str, Any]] | None]:
        """
        Extract/cache PDFs, build master_lines, apply processing, create a
        DataFrame, and optionally sort it.
        """
        trips, lines = self._extract_pdfs(inputs)

        bid_period_info = {
            key: lines[key]
            for key in (
                "bid_period_date_range",
                "pay_period_date_ranges",
            )
        }

        airport_lookup = self._get_airport_lookup(
            inputs=inputs,
            trips=trips,
        )

        self._log("Creating master lines...")
        master_lines = creating_master_line(
            trips,
            lines,
        )

        self._log("Adding blockiness scores...")
        pf.add_blockiness_scores(
            master_lines,
            bid_period_info,
        )

        self._log("Adding company-ticket percentages...")
        pf.add_company_ticket_percentages(master_lines)

        self._log("Adding line-type preference scores...")
        pf.add_line_type_preference_scores(
            master_lines,
            inputs["line_type_preference_order"],
            power_law_coeff=3,
        )

        self._log("Adding estimated pay...")
        try:
            pf.add_pay(
                master_lines,
                inputs["hourly_rate"],
            )
        except UnboundLocalError as exc:
            error_text = str(exc)
            if (
                "local variable 'pp'" not in error_text
                and 'local variable "pp"' not in error_text
            ):
                raise

            self._log(
                "The current add_pay function hit its local-variable "
                "'pp' bug. Applying the same documented pay formula "
                "with the pipeline fallback."
            )
            self._add_pay_fallback(
                master_lines,
                inputs["hourly_rate"],
            )

        self._log("Adding complete-weekends-off percentages...")
        pf.add_weekends_off_percentage(master_lines)

        self._log("Adding international-destination scores...")
        pf.add_international_destination_scores(
            master_lines,
            airport_lookup,
        )

        if inputs["requested_dates"] is not None:
            self._log("Adding requested-days-off scores...")
            pf.add_requested_days_off_scores(
                master_lines,
                bid_period_info,
                inputs["requested_dates"],
            )

        if inputs["vacation_ranges"] is not None:
            self._log("Adding vacation scores...")
            new_vacation_ranges = pf.add_vacation_days_off_score(
                master_lines,
                inputs["vacation_ranges"],
                bid_period_info,
                save_details=False,
            )
        else:
            new_vacation_ranges = None

        if inputs["training_start"] is not None:
            self._log("Adding training-fit scores...")
            pf.add_training_fit_score(
                master_lines,
                inputs["training_start"],
                inputs["training_end"],
                bid_period_info,
            )

        self._log("Adding average legs per work day...")
        pf.add_avg_legs_per_work_day(master_lines)

        if inputs["bid_edge"] != "none":
            self._log("Adding bid-edge days-off scores...")
            pf.add_bid_edge_days_off(
                master_lines,
                bid_period_info,
                edge=inputs["bid_edge"],
            )

        self._log("Creating DataFrame...")
        df = master_lines_to_dataframe(
            master_lines,
            bid_period_info,
        )

        if apply_sort and inputs["sort_order"]:
            sorting_settings = (
                inputs.get("sorting_settings")
                or DEFAULT_SORTING_SETTINGS
            )

            self._log(
                "Sorting DataFrame "
                f"({sorting_settings['default_mode']}, "
                f"{sorting_settings['weighting_style']}, "
                f"soft weights "
                f"{sorting_settings['soft_min_weight']}–"
                f"{sorting_settings['soft_max_weight']})."
            )

            df = drop_empty_sort_columns(
                df,
                check_all_columns=True,
            )

            df = sort_dataframe_by_conditions(
                df,
                inputs["sort_order"],
                default_mode=sorting_settings["default_mode"],
                weighting_style=sorting_settings["weighting_style"],
                soft_max_weight=sorting_settings["soft_max_weight"],
                soft_min_weight=sorting_settings["soft_min_weight"],
            )

        return df, new_vacation_ranges
