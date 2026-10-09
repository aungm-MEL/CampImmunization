from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font
from openpyxl.styles import PatternFill


SUMMARY_COLUMNS = [
    "Year",
    "period",
    "Organization",
    "Project Name",
    "District (EHO)",
    "Township_EHO",
    "Twp_MIMU",
    "Clinic Name",
    "ALOD_U1",
    "ALOD_U5",
    "ALOD_>5",
    "BCG_U1",
    "BCG_U5",
    "BCG_>5",
    "OPV1_U1",
    "OPV1_U5",
    "OPV1_>5",
    "OPV2_U1",
    "OPV2_U5",
    "OPV2_>5",
    "OPV3_U1",
    "OPV3_U5",
    "OPV3_>5",
    "Penta1_U1",
    "Penta1_U5",
    "Penta1_>5",
    "Penta2_U1",
    "Penta2_U5",
    "Penta2_>5",
    "Penta3_U1",
    "Penta3_U5",
    "Penta3_>5",
    "MMR1_U1",
    "MMR1_U5",
    "MMR1_>5",
    "MMR2_U1",
    "MMR2_U5",
    "MMR2_>5",
    "JE_U1",
    "JE_U5",
    "JE_>5",
    "IPV_U1",
    "IPV_U5",
    "IPV_>5",
    "CD_U1",
    "CD_U5",
    "CD_>5",
    "Td1",
    "Td2",
    "Td At least one dose",
]

YEARLY_CUMMU_SUMMARY_COLUMNS = [
    "Year",
    "Organization",
    "Project Name",
    "District (EHO)",
    "Township_EHO",
    "Twp_MIMU",
    "Clinic Name",
    "ALOD_U1",
    "ALOD_U5",
    "ALOD_>5",
]

INDICATOR_COLUMNS = [
    "Period",
    "Organization",
    "Project Name",
    "indicator",
    "Q1 Target",
    "Q1 U1 Male",
    "Q1 U1 Female",
    "Q1 1-5 Male",
    "Q1 1-5 Female",
    "Q1 Total",
    "Q2 Target",
    "Q2 U1 Male",
    "Q2 U1 Female",
    "Q2 1-5 Male",
    "Q2 1-5 Female",
    "Q2 Total",
    "Q3 Target",
    "Q3 U1 Male",
    "Q3 U1 Female",
    "Q3 1-5 Male",
    "Q3 1-5 Female",
    "Q3 Total",
    "Q4 Target",
    "Q4 U1 Male",
    "Q4 U1 Female",
    "Q4 1-5 Male",
    "Q4 1-5 Female",
    "Q4 Total",
]

ALOD_CUMMU_INDICATOR_COLUMNS = [
    "Period",
    "Organization",
    "Project Name",
    "indicator",
    "Annual U1 Male",
    "Annual U1 Female",
    "Annual 1-5 Male",
    "Annual 1-5 Female",
    "Annual Total",
    "S1 U1 Male",
    "S1 U1 Female",
    "S1 1-5 Male",
    "S1 1-5 Female",
    "S1 Total",
    "uptoQ3 U1 Male",
    "uptoQ3 U1 Female",
    "uptoQ3 1-5 Male",
    "uptoQ3 1-5 Female",
    "uptoQ3 Total",
    "S2 U1 Male",
    "S2 U1 Female",
    "S2 1-5 Male",
    "S2 1-5 Female",
    "S2 Total",
]

INDICATOR_NAMES = [
    "Penta3 under 1-yr-old",
    "MMR1 under 1-yr-old",
    "Penta1 under 5-yr-old",
    "Penta3 under 5-yr-old",
    "MMR1 under 5-yr-old",
    "MMR2 under 5-yr-old",
    "Full dose under 5-yr-old",
    "At least one dose under 5-yr-old",
    "Td ALOD",
    "Td Two Doses",
]

UNPIVOT_EXCLUDED_VACCINES = {
    "rv2-1",
    "rv2-2",
    "rv3-1",
    "rv3-2",
    "rv3-3",
    "rabies vaccine 1",
    "rabies vaccine 2",
    "rabies vaccine 3",
    "rabies vaccine 4",
    "tt1",
    "ap",
    "dt1",
    "dt2",
    "dt3",
    "dt5",
}


def convert_thai_buddhist_dates(frame: pd.DataFrame) -> pd.DataFrame:
    date_pattern = r"^\d{4}-\d{2}-\d{2}$"
    converted_cells = 0
    converted_columns: list[str] = []

    for column in frame.columns:
        if not (pd.api.types.is_object_dtype(frame[column]) or pd.api.types.is_string_dtype(frame[column])):
            continue

        values = frame[column].astype("string").str.strip()
        non_empty_mask = values.notna() & values.ne("")
        if not non_empty_mask.any():
            continue

        date_like_mask = values.str.match(date_pattern, na=False)
        date_like_ratio = float(date_like_mask.sum()) / float(non_empty_mask.sum())

        normalized_name = str(column).strip().lower()
        likely_date_column = "date" in normalized_name or "วันที่" in normalized_name
        if not likely_date_column and date_like_ratio < 0.7:
            continue

        years = pd.to_numeric(values.str.slice(0, 4), errors="coerce")
        thai_mask = date_like_mask & years.ge(2400)
        if not thai_mask.any():
            continue

        new_years = (years - 543).astype("Int64").astype("string").str.zfill(4)
        values.loc[thai_mask] = (
            new_years.loc[thai_mask]
            + "-"
            + values.str.slice(5, 7).loc[thai_mask]
            + "-"
            + values.str.slice(8, 10).loc[thai_mask]
        )

        frame[column] = values
        converted_cells += int(thai_mask.sum())
        converted_columns.append(str(column))

    if converted_cells > 0:
        print(
            "INFO: Converted "
            f"{converted_cells} Thai Buddhist date value(s) to Gregorian in column(s): "
            f"{', '.join(converted_columns)}"
        )
    else:
        print("INFO: No Thai Buddhist calendar dates found to convert.")

    return frame


def normalize_sex_column(frame: pd.DataFrame) -> pd.DataFrame:
    normalized_columns = {str(col).strip().lower(): col for col in frame.columns}
    sex_column = normalized_columns.get("sex") or normalized_columns.get("เพศ")

    if not sex_column:
        print("WARNING: Sex column not found, skipping sex normalization.")
        return frame

    values = frame[sex_column].astype("string").str.strip().str.lower()
    mapping = {
        "ชาย": "M",
        "หญิง": "F",
        "m": "M",
        "male": "M",
        "f": "F",
        "female": "F",
    }

    normalized = values.map(mapping)
    normalized = normalized.where(values.notna() & values.ne(""), "NA")
    normalized = normalized.fillna("NA")
    frame[sex_column] = normalized

    print(f"INFO: Normalized sex values in column '{sex_column}' to M/F/NA")
    return frame


def standardize_vaccine_names(frame: pd.DataFrame) -> pd.DataFrame:
    normalized_columns = {str(col).strip().lower(): col for col in frame.columns}
    vaccine_column = normalized_columns.get("vaccine")

    if not vaccine_column:
        print("WARNING: Vaccine column not found, skipping vaccine name standardization.")
        return frame

    def map_vaccine_name(value: object) -> str:
        text = str(value).strip()
        compact = " ".join(text.lower().split())

        if compact.startswith("dtp-hb-hib"):
            suffix = compact.removeprefix("dtp-hb-hib").strip().replace(" ", "")
            return f"Penta{suffix}" if suffix else "Penta"

        if compact.startswith("dtp-ipv-hb-hib"):
            suffix = compact.removeprefix("dtp-ipv-hb-hib").strip().replace(" ", "")
            return f"hexa{suffix}" if suffix else "hexa"

        if compact in {"measles/mmr", "mmrv1", "mmrs"}:
            return "MMR1"

        if compact == "mmrv2":
            return "MMR2"

        return text

    frame[vaccine_column] = frame[vaccine_column].map(map_vaccine_name)
    print("INFO: Standardized vaccine names in Combined sheet for requested mappings")
    return frame


def add_age_at_dose_column(frame: pd.DataFrame) -> pd.DataFrame:
    normalized_columns = {str(col).strip().lower().replace(" ", "_"): col for col in frame.columns}
    birth_column = normalized_columns.get("birth_date")
    vaccine_column = normalized_columns.get("vaccine_date")
    visit_column = normalized_columns.get("visit_number") or normalized_columns.get("visited_number")

    if not birth_column or not vaccine_column:
        print("WARNING: age_at_dose not added because birth_date or vaccine_date column is missing.")
        return frame

    birth_dates = pd.to_datetime(frame[birth_column], errors="coerce")
    vaccine_dates = pd.to_datetime(frame[vaccine_column], errors="coerce")

    month_diff = (vaccine_dates.dt.year - birth_dates.dt.year) * 12 + (vaccine_dates.dt.month - birth_dates.dt.month)
    month_diff = month_diff - (vaccine_dates.dt.day < birth_dates.dt.day).astype("Int64")

    earlier_than_birth_mask = vaccine_dates < birth_dates

    if earlier_than_birth_mask.any():
        if visit_column:
            visit_values = frame.loc[earlier_than_birth_mask, visit_column].dropna().astype(str).str.strip()
            unique_visits = sorted({value for value in visit_values if value})
            if unique_visits:
                preview_limit = 30
                preview_visits = unique_visits[:preview_limit]
                preview_text = ", ".join(preview_visits)
                if len(unique_visits) > preview_limit:
                    preview_text = f"{preview_text}, ..."
                print(
                    "WARNING: visited_number "
                    f"{preview_text} were vaccination date is earlier than birth date"
                )
            else:
                print("WARNING: Found rows where vaccination date is earlier than birth date")
        else:
            print("WARNING: Found rows where vaccination date is earlier than birth date")

    invalid_mask = birth_dates.isna() | vaccine_dates.isna() | earlier_than_birth_mask | (month_diff < 0)
    month_diff = month_diff.mask(invalid_mask, 99999)

    frame["age_at_dose"] = month_diff.astype("Int64")
    print("INFO: Added age_at_dose column in completed months using vaccine_date - birth_date (errors set to 99999)")
    return frame


def add_visit_age_column(frame: pd.DataFrame) -> pd.DataFrame:
    normalized_columns = {str(col).strip().lower().replace(" ", "_"): col for col in frame.columns}
    birth_column = normalized_columns.get("birth_date")
    visit_date_column = normalized_columns.get("visit_date")
    visit_column = normalized_columns.get("visit_number") or normalized_columns.get("visited_number")

    if not birth_column or not visit_date_column:
        print("WARNING: visit_age not added because birth_date or visit_date column is missing.")
        return frame

    birth_dates = pd.to_datetime(frame[birth_column], errors="coerce")
    visit_dates = pd.to_datetime(frame[visit_date_column], errors="coerce")

    month_diff = (visit_dates.dt.year - birth_dates.dt.year) * 12 + (visit_dates.dt.month - birth_dates.dt.month)
    month_diff = month_diff - (visit_dates.dt.day < birth_dates.dt.day).astype("Int64")

    earlier_than_birth_mask = visit_dates < birth_dates
    if earlier_than_birth_mask.any():
        if visit_column:
            visit_values = frame.loc[earlier_than_birth_mask, visit_column].dropna().astype(str).str.strip()
            unique_visits = sorted({value for value in visit_values if value})
            if unique_visits:
                preview_limit = 30
                preview_visits = unique_visits[:preview_limit]
                preview_text = ", ".join(preview_visits)
                if len(unique_visits) > preview_limit:
                    preview_text = f"{preview_text}, ..."
                print(
                    "WARNING: visited_number "
                    f"{preview_text} were visit date is earlier than birth date"
                )
            else:
                print("WARNING: Found rows where visit date is earlier than birth date")
        else:
            print("WARNING: Found rows where visit date is earlier than birth date")

    invalid_mask = birth_dates.isna() | visit_dates.isna() | earlier_than_birth_mask | (month_diff < 0)
    month_diff = month_diff.mask(invalid_mask, 99999)

    frame["visit_age"] = month_diff.astype("Int64")
    print("INFO: Added visit_age column in completed months using visit_date - birth_date (errors set to 99999)")
    return frame


def map_vaccine_group(vaccine_value: object) -> str | None:
    vaccine = str(vaccine_value).strip().lower()
    vaccine_compact = " ".join(vaccine.split())

    if vaccine_compact == "bcg":
        return "BCG"
    if vaccine_compact == "opv1":
        return "OPV1"
    if vaccine_compact == "opv2":
        return "OPV2"
    if vaccine_compact == "opv3":
        return "OPV3"
    if vaccine_compact in {"dtp-hb-hib 1", "hbv1"}:
        return "Penta1"
    if vaccine_compact in {"dtp-hb-hib 2", "dtp-ipv-hb-hib2", "hbv2"}:
        return "Penta2"
    if vaccine_compact == "dtp-hb-hib 3":
        return "Penta3"
    if vaccine_compact == "penta1":
        return "Penta1"
    if vaccine_compact in {"penta2", "hexa2"}:
        return "Penta2"
    if vaccine_compact == "penta3":
        return "Penta3"
    if vaccine_compact in {"measles/mmr", "mmrv1", "mmrs"}:
        return "MMR1"
    if vaccine_compact in {"mmr2", "mmrv2"}:
        return "MMR2"
    if vaccine_compact == "mmr1":
        return "MMR1"
    if vaccine_compact.startswith("je"):
        return "JE"
    if vaccine_compact == "ipv1":
        return "IPV"
    if vaccine_compact in {"dt1", "tt1"}:
        return "Td1"
    if vaccine_compact == "dt2":
        return "Td2"
    if vaccine_compact.startswith("dt") or vaccine_compact.startswith("tt"):
        return "TdAny"
    return None


def create_summary_sheet(frame: pd.DataFrame) -> pd.DataFrame:
    normalized_columns = {str(col).strip().lower().replace(" ", "_"): col for col in frame.columns}
    vaccine_date_column = normalized_columns.get("vaccine_date")
    vaccine_column = normalized_columns.get("vaccine")
    beneficiary_column = normalized_columns.get("beneficiary_code") or normalized_columns.get("hn")

    if not vaccine_date_column or not vaccine_column:
        print("WARNING: Summary sheet could not be generated because vaccine_date or vaccine column is missing.")
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    if beneficiary_column != normalized_columns.get("beneficiary_code"):
        print("WARNING: beneficiary_code column not found. Using 'hn' as unique beneficiary identifier.")

    work = frame.copy()
    work[vaccine_date_column] = pd.to_datetime(work[vaccine_date_column], errors="coerce")
    work = work.dropna(subset=[vaccine_date_column, beneficiary_column])

    if work.empty:
        print("WARNING: Summary sheet generated with no rows because required values are missing.")
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    work["beneficiary_id"] = work[beneficiary_column].astype(str).str.strip()
    work = work[work["beneficiary_id"] != ""]
    work["Year"] = work[vaccine_date_column].dt.year.astype("Int64")
    work["period"] = (
        "Q" + work[vaccine_date_column].dt.quarter.astype("Int64").astype(str) + "_" + work["Year"].astype(str)
    )

    work["vaccine_group"] = work[vaccine_column].map(map_vaccine_group)
    work["age_months"] = pd.to_numeric(work.get("age_at_dose"), errors="coerce")
    summary_rows: list[dict[str, object]] = []

    u1_rules: dict[str, tuple[int, int]] = {
        "BCG": (0, 11),
        "Penta1": (2, 11),
        "Penta2": (3, 11),
        "Penta3": (4, 11),
        "MMR1": (9, 11),
    }
    default_u1_rule = (0, 11)
    # U1 includes month 11; 1-5 uses 12-59.
    u5_rule = (12, 59)

    for period_value in sorted(work["period"].dropna().unique()):
        period_frame = work[work["period"] == period_value]
        year_value = int(period_frame["Year"].dropna().iloc[0])

        row: dict[str, object] = {
            "Year": year_value,
            "period": period_value,
            "Organization": "PRF",
            "Project Name": "REACH-KK",
            "District (EHO)": "",
            "Township_EHO": "",
            "Twp_MIMU": "Phop Phra",
            "Clinic Name": "Umphang Camp",
        }

        def unique_count_by_age(min_age: int, max_age: int, vaccine_group: str | None = None) -> int:
            subset = period_frame[period_frame["age_months"].between(min_age, max_age, inclusive="both")]
            if vaccine_group is not None:
                subset = subset[subset["vaccine_group"] == vaccine_group]
            return int(subset["beneficiary_id"].nunique())

        def unique_count_over_5(vaccine_group: str | None = None) -> int:
            subset = period_frame[period_frame["age_months"] >= 60]
            if vaccine_group is not None:
                subset = subset[subset["vaccine_group"] == vaccine_group]
            return int(subset["beneficiary_id"].nunique())

        row["ALOD_U1"] = unique_count_by_age(0, 11)
        row["ALOD_U5"] = unique_count_by_age(11, 59)
        row["ALOD_>5"] = unique_count_over_5()

        for antigen in ["BCG", "OPV1", "OPV2", "OPV3", "Penta1", "Penta2", "Penta3", "MMR1", "MMR2", "JE", "IPV", "CD"]:
            u1_min, u1_max = u1_rules.get(antigen, default_u1_rule)
            row[f"{antigen}_U1"] = unique_count_by_age(u1_min, u1_max, antigen)
            row[f"{antigen}_U5"] = unique_count_by_age(u5_rule[0], u5_rule[1], antigen)
            row[f"{antigen}_>5"] = unique_count_over_5(antigen)

        row["Td1"] = int(period_frame[period_frame["vaccine_group"] == "Td1"]["beneficiary_id"].nunique())
        row["Td2"] = int(period_frame[period_frame["vaccine_group"] == "Td2"]["beneficiary_id"].nunique())
        td_any_mask = period_frame["vaccine_group"].isin(["Td1", "Td2", "TdAny"])
        row["Td At least one dose"] = int(period_frame[td_any_mask]["beneficiary_id"].nunique())

        summary_rows.append(row)

    summary = pd.DataFrame(summary_rows)
    summary = summary.reindex(columns=SUMMARY_COLUMNS, fill_value=0)
    print(f"INFO: Created Summary sheet with {len(summary)} period row(s)")
    return summary


def create_yearly_cummu_summary_sheet(frame: pd.DataFrame) -> pd.DataFrame:
    normalized_columns = {str(col).strip().lower().replace(" ", "_"): col for col in frame.columns}
    vaccine_date_column = normalized_columns.get("vaccine_date")
    beneficiary_column = normalized_columns.get("beneficiary_code") or normalized_columns.get("hn")

    if not vaccine_date_column:
        print("WARNING: yearly_cummu_summary could not be generated because vaccine_date column is missing.")
        return pd.DataFrame(columns=YEARLY_CUMMU_SUMMARY_COLUMNS)

    if beneficiary_column != normalized_columns.get("beneficiary_code"):
        print("WARNING: beneficiary_code column not found for yearly_cummu_summary. Using 'hn' as unique beneficiary identifier.")

    work = frame.copy()
    work[vaccine_date_column] = pd.to_datetime(work[vaccine_date_column], errors="coerce")
    work = work.dropna(subset=[vaccine_date_column, beneficiary_column])

    if work.empty:
        print("WARNING: yearly_cummu_summary generated with no rows because required values are missing.")
        return pd.DataFrame(columns=YEARLY_CUMMU_SUMMARY_COLUMNS)

    work["beneficiary_id"] = work[beneficiary_column].astype(str).str.strip()
    work = work[work["beneficiary_id"] != ""]
    work["Year"] = work[vaccine_date_column].dt.year.astype("Int64")
    work["age_months"] = pd.to_numeric(work.get("age_at_dose"), errors="coerce")

    yearly_rows: list[dict[str, object]] = []

    for year_value in sorted(work["Year"].dropna().astype(int).unique()):
        year_frame = work[work["Year"] == year_value]

        def unique_count_by_age(min_age: int, max_age: int) -> int:
            subset = year_frame[year_frame["age_months"].between(min_age, max_age, inclusive="both")]
            return int(subset["beneficiary_id"].nunique())

        def unique_count_over_5() -> int:
            subset = year_frame[year_frame["age_months"] >= 60]
            return int(subset["beneficiary_id"].nunique())

        row: dict[str, object] = {
            "Year": year_value,
            "Organization": "PRF",
            "Project Name": "REACH-KK",
            "District (EHO)": "",
            "Township_EHO": "",
            "Twp_MIMU": "Phop Phra",
            "Clinic Name": "Umphang Camp",
            "ALOD_U1": unique_count_by_age(0, 11),
            "ALOD_U5": unique_count_by_age(11, 59),
            "ALOD_>5": unique_count_over_5(),
        }
        yearly_rows.append(row)

    yearly_summary = pd.DataFrame(yearly_rows)
    yearly_summary = yearly_summary.reindex(columns=YEARLY_CUMMU_SUMMARY_COLUMNS, fill_value=0)
    print(f"INFO: Created yearly_cummu_summary sheet with {len(yearly_summary)} year row(s)")
    return yearly_summary


def create_indicator_sheet(frame: pd.DataFrame) -> pd.DataFrame:
    normalized_columns = {str(col).strip().lower().replace(" ", "_"): col for col in frame.columns}
    vaccine_date_column = normalized_columns.get("vaccine_date")
    vaccine_column = normalized_columns.get("vaccine")
    sex_column = normalized_columns.get("sex")
    beneficiary_column = normalized_columns.get("hn")

    if not vaccine_date_column or not vaccine_column or not sex_column or not beneficiary_column:
        print("WARNING: Indicator sheet could not be generated because required columns are missing.")
        return pd.DataFrame(columns=INDICATOR_COLUMNS)

    work = frame.copy()
    work[vaccine_date_column] = pd.to_datetime(work[vaccine_date_column], errors="coerce")
    work = work.dropna(subset=[vaccine_date_column, beneficiary_column])
    if work.empty:
        print("WARNING: Indicator sheet generated with no rows because required values are missing.")
        return pd.DataFrame(columns=INDICATOR_COLUMNS)

    work["beneficiary_id"] = work[beneficiary_column].astype(str).str.strip()
    work = work[work["beneficiary_id"] != ""]
    work["Year"] = work[vaccine_date_column].dt.year.astype("Int64")
    work["Quarter"] = work[vaccine_date_column].dt.quarter.astype("Int64")
    work["vaccine_group"] = work[vaccine_column].map(map_vaccine_group)
    work["age_months"] = pd.to_numeric(work.get("age_at_dose"), errors="coerce")
    work["sex_norm"] = work[sex_column].astype(str).str.strip().str.upper()

    # Same age logic used in Summary sheet.
    u1_rules: dict[str, tuple[int, int]] = {
        "BCG": (0, 11),
        "Penta1": (2, 11),
        "Penta2": (3, 11),
        "Penta3": (4, 11),
        "MMR1": (9, 11),
    }
    default_u1_rule = (0, 11)
    u5_rule = (12, 59)

    def indicator_mask(indicator: str) -> pd.Series:
        if indicator == "Penta3 under 1-yr-old":
            return work["vaccine_group"] == "Penta3"
        if indicator == "MMR1 under 1-yr-old":
            return work["vaccine_group"] == "MMR1"
        if indicator == "Penta1 under 5-yr-old":
            return work["vaccine_group"] == "Penta1"
        if indicator == "Penta3 under 5-yr-old":
            return work["vaccine_group"] == "Penta3"
        if indicator == "MMR1 under 5-yr-old":
            return work["vaccine_group"] == "MMR1"
        if indicator == "MMR2 under 5-yr-old":
            return work["vaccine_group"] == "MMR2"
        if indicator == "Full dose under 5-yr-old":
            # Full dose approximated by Penta3 in this dataset.
            return work["vaccine_group"] == "Penta3"
        if indicator == "At least one dose under 5-yr-old":
            return work["vaccine_group"].notna()
        if indicator == "Td ALOD":
            return work["vaccine_group"].isin(["Td1", "Td2", "TdAny"])
        if indicator == "Td Two Doses":
            return work["vaccine_group"] == "Td2"
        return pd.Series(False, index=work.index)

    def u1_range_for_indicator(indicator: str) -> tuple[int, int]:
        if "Penta3" in indicator:
            return u1_rules["Penta3"]
        if "Penta1" in indicator:
            return u1_rules["Penta1"]
        if "MMR1" in indicator:
            return u1_rules["MMR1"]
        if "MMR2" in indicator:
            return default_u1_rule
        return default_u1_rule

    rows: list[dict[str, object]] = []
    years = sorted(work["Year"].dropna().astype(int).unique())
    for year_value in years:
        for indicator in INDICATOR_NAMES:
            row: dict[str, object] = {
                "Period": year_value,
                "Organization": "PRF",
                "Project Name": "REACH-KK",
                "indicator": indicator,
            }

            base_mask = indicator_mask(indicator) & work["Year"].eq(year_value)
            u1_min, u1_max = u1_range_for_indicator(indicator)

            for quarter in [1, 2, 3, 4]:
                quarter_mask = base_mask & work["Quarter"].eq(quarter)

                if indicator in {"Td ALOD", "Td Two Doses"}:
                    # Requested layout: Td values are reported in 1-5 Female column only.
                    td_female_count = int(
                        work.loc[quarter_mask & work["sex_norm"].eq("F"), "beneficiary_id"].nunique()
                    )
                    q_u1_m = 0
                    q_u1_f = 0
                    q_u5_m = 0
                    q_u5_f = td_female_count
                    q_total = q_u5_f
                else:
                    u1_mask = quarter_mask & work["age_months"].between(u1_min, u1_max, inclusive="both")
                    u5_mask = quarter_mask & work["age_months"].between(u5_rule[0], u5_rule[1], inclusive="both")

                    q_u1_m = int(work.loc[u1_mask & work["sex_norm"].eq("M"), "beneficiary_id"].nunique())
                    q_u1_f = int(work.loc[u1_mask & work["sex_norm"].eq("F"), "beneficiary_id"].nunique())
                    if "under 1-yr-old" in indicator.lower():
                        q_u5_m = ""
                        q_u5_f = ""
                        q_total = q_u1_m + q_u1_f
                    else:
                        q_u5_m = int(work.loc[u5_mask & work["sex_norm"].eq("M"), "beneficiary_id"].nunique())
                        q_u5_f = int(work.loc[u5_mask & work["sex_norm"].eq("F"), "beneficiary_id"].nunique())
                        q_total = q_u1_m + q_u1_f + q_u5_m + q_u5_f

                row[f"Q{quarter} Target"] = 0
                row[f"Q{quarter} U1 Male"] = q_u1_m
                row[f"Q{quarter} U1 Female"] = q_u1_f
                row[f"Q{quarter} 1-5 Male"] = q_u5_m
                row[f"Q{quarter} 1-5 Female"] = q_u5_f
                row[f"Q{quarter} Total"] = q_total

            rows.append(row)

    indicator_df = pd.DataFrame(rows).reindex(columns=INDICATOR_COLUMNS, fill_value=0)
    print(f"INFO: Created Indicator sheet with {len(indicator_df)} row(s)")
    return indicator_df


def create_alod_cummu_indicator_sheet(frame: pd.DataFrame) -> pd.DataFrame:
    normalized_columns = {str(col).strip().lower().replace(" ", "_"): col for col in frame.columns}
    vaccine_date_column = normalized_columns.get("vaccine_date")
    vaccine_column = normalized_columns.get("vaccine")
    sex_column = normalized_columns.get("sex")
    beneficiary_column = normalized_columns.get("hn")

    if not vaccine_date_column or not vaccine_column or not sex_column or not beneficiary_column:
        print("WARNING: alod_cummu_indicator could not be generated because required columns are missing.")
        return pd.DataFrame(columns=ALOD_CUMMU_INDICATOR_COLUMNS)

    work = frame.copy()
    work[vaccine_date_column] = pd.to_datetime(work[vaccine_date_column], errors="coerce")
    work = work.dropna(subset=[vaccine_date_column, beneficiary_column])
    if work.empty:
        print("WARNING: alod_cummu_indicator generated with no rows because required values are missing.")
        return pd.DataFrame(columns=ALOD_CUMMU_INDICATOR_COLUMNS)

    work["beneficiary_id"] = work[beneficiary_column].astype(str).str.strip()
    work = work[work["beneficiary_id"] != ""]
    work["Year"] = work[vaccine_date_column].dt.year.astype("Int64")
    work["Month"] = work[vaccine_date_column].dt.month.astype("Int64")
    work["vaccine_group"] = work[vaccine_column].map(map_vaccine_group)
    work["age_months"] = pd.to_numeric(work.get("age_at_dose"), errors="coerce")
    work["sex_norm"] = work[sex_column].astype(str).str.strip().str.upper()

    # Same base indicator logic as "At least one dose under 5-yr-old" in Indicator sheet.
    work = work[work["vaccine_group"].notna()].copy()

    def segment_counts(segment_mask: pd.Series) -> tuple[int, int, int, int, int]:
        u1_mask = segment_mask & work["age_months"].between(0, 11, inclusive="both")
        u5_mask = segment_mask & work["age_months"].between(12, 59, inclusive="both")

        u1_m = int(work.loc[u1_mask & work["sex_norm"].eq("M"), "beneficiary_id"].nunique())
        u1_f = int(work.loc[u1_mask & work["sex_norm"].eq("F"), "beneficiary_id"].nunique())
        u5_m = int(work.loc[u5_mask & work["sex_norm"].eq("M"), "beneficiary_id"].nunique())
        u5_f = int(work.loc[u5_mask & work["sex_norm"].eq("F"), "beneficiary_id"].nunique())
        total = u1_m + u1_f + u5_m + u5_f
        return u1_m, u1_f, u5_m, u5_f, total

    rows: list[dict[str, object]] = []
    years = sorted(work["Year"].dropna().astype(int).unique())

    for year_value in years:
        year_mask = work["Year"].eq(year_value)

        annual = segment_counts(year_mask)
        s1 = segment_counts(year_mask & work["Month"].between(1, 6, inclusive="both"))
        upto_q3 = segment_counts(year_mask & work["Month"].between(1, 9, inclusive="both"))
        s2 = segment_counts(year_mask & work["Month"].between(7, 12, inclusive="both"))

        row: dict[str, object] = {
            "Period": year_value,
            "Organization": "PRF",
            "Project Name": "REACH-KK",
            "indicator": "At least one dose under 5-yr-old",
            "Annual U1 Male": annual[0],
            "Annual U1 Female": annual[1],
            "Annual 1-5 Male": annual[2],
            "Annual 1-5 Female": annual[3],
            "Annual Total": annual[4],
            "S1 U1 Male": s1[0],
            "S1 U1 Female": s1[1],
            "S1 1-5 Male": s1[2],
            "S1 1-5 Female": s1[3],
            "S1 Total": s1[4],
            "uptoQ3 U1 Male": upto_q3[0],
            "uptoQ3 U1 Female": upto_q3[1],
            "uptoQ3 1-5 Male": upto_q3[2],
            "uptoQ3 1-5 Female": upto_q3[3],
            "uptoQ3 Total": upto_q3[4],
            "S2 U1 Male": s2[0],
            "S2 U1 Female": s2[1],
            "S2 1-5 Male": s2[2],
            "S2 1-5 Female": s2[3],
            "S2 Total": s2[4],
        }
        rows.append(row)

    alod_cummu_indicator = pd.DataFrame(rows).reindex(columns=ALOD_CUMMU_INDICATOR_COLUMNS, fill_value=0)
    print(f"INFO: Created alod_cummu_indicator sheet with {len(alod_cummu_indicator)} year row(s)")
    return alod_cummu_indicator


def create_unpivot_sheet(frame: pd.DataFrame) -> pd.DataFrame:
    required_columns = ["child_name", "visit_number", "mother_name", "vaccine", "vaccine_date"]
    normalized_columns = {str(col).strip().lower().replace(" ", "_"): col for col in frame.columns}
    missing = [col for col in required_columns if col not in normalized_columns]

    if missing:
        print(f"WARNING: Unpivot sheet could not be generated. Missing column(s): {', '.join(missing)}")
        return pd.DataFrame()

    ordered_columns = list(frame.columns)
    child_col = normalized_columns["child_name"]
    mother_col = normalized_columns["mother_name"]
    vaccine_col = normalized_columns["vaccine"]
    vaccine_date_col = normalized_columns["vaccine_date"]

    child_idx = ordered_columns.index(child_col)
    mother_idx = ordered_columns.index(mother_col)
    base_columns = ordered_columns[child_idx : mother_idx + 1]

    age_filter_column = normalized_columns.get("visit_age") or normalized_columns.get("age_at_dose")
    age_at_dose_column = normalized_columns.get("age_at_dose")
    selected_columns = base_columns + [vaccine_col, vaccine_date_col]
    if age_filter_column:
        selected_columns.append(age_filter_column)
    if age_at_dose_column and age_at_dose_column not in selected_columns:
        selected_columns.append(age_at_dose_column)

    work = frame[selected_columns].copy()
    work = work.dropna(subset=[normalized_columns["visit_number"], vaccine_col])

    vaccine_norm = work[vaccine_col].astype(str).str.strip().str.lower().str.split().str.join(" ")
    excluded_vaccine_mask = vaccine_norm.isin(UNPIVOT_EXCLUDED_VACCINES)

    age_filter_mask = pd.Series(False, index=work.index)
    if age_filter_column:
        age_months = pd.to_numeric(work[age_filter_column], errors="coerce")
        age_filter_mask = age_months >= 72

    work = work.loc[~(excluded_vaccine_mask | age_filter_mask)].copy()

    if age_filter_column and age_filter_column in work.columns and age_filter_column != age_at_dose_column:
        work = work.drop(columns=[age_filter_column])

    # Ensure one value per visit_number + vaccine before pivoting to wide format.
    work = work.sort_values(vaccine_date_col).drop_duplicates(subset=base_columns + [vaccine_col], keep="first")

    date_wide = work.pivot(index=base_columns, columns=vaccine_col, values=vaccine_date_col)
    date_wide.columns = [str(col) for col in date_wide.columns]
    unpivot = date_wide.reset_index()

    if age_at_dose_column and age_at_dose_column in work.columns:
        age_wide = work.pivot(index=base_columns, columns=vaccine_col, values=age_at_dose_column)
        age_wide.columns = [str(col) for col in age_wide.columns]
        age_wide = age_wide.reset_index()
        age_wide = age_wide.rename(columns={col: f"{col}_age" for col in age_wide.columns if col not in base_columns})

        unpivot = unpivot.merge(age_wide, on=base_columns, how="left")

        vaccine_columns = [str(col) for col in date_wide.columns]

        def vaccine_sort_key(name: str) -> tuple[int, int, str]:
            normalized = "".join(name.lower().split())

            # Requested display order: BCG, Penta, OPV, IPV, MMR, JE, then the rest.
            if normalized.startswith("bcg"):
                return (0, 0, normalized)
            if normalized.startswith("penta"):
                suffix = normalized.removeprefix("penta")
                seq = int(suffix) if suffix.isdigit() else 0
                return (1, seq, normalized)
            if normalized.startswith("opv"):
                suffix = normalized.removeprefix("opv")
                seq = int(suffix) if suffix.isdigit() else 0
                return (2, seq, normalized)
            if normalized.startswith("ipv"):
                suffix = normalized.removeprefix("ipv")
                seq = int(suffix) if suffix.isdigit() else 0
                return (3, seq, normalized)
            if normalized.startswith("mmr"):
                suffix = normalized.removeprefix("mmr")
                seq = int(suffix) if suffix.isdigit() else 0
                return (4, seq, normalized)
            if normalized.startswith("je"):
                suffix = normalized.removeprefix("je")
                seq = int(suffix) if suffix.isdigit() else 0
                return (5, seq, normalized)
            return (6, 0, normalized)

        vaccine_columns = sorted(vaccine_columns, key=vaccine_sort_key)
        ordered_columns = list(base_columns)
        for vaccine_name in vaccine_columns:
            ordered_columns.append(vaccine_name)
            age_name = f"{vaccine_name}_age"
            if age_name in unpivot.columns:
                ordered_columns.append(age_name)
        unpivot = unpivot.reindex(columns=ordered_columns)
    else:
        print("WARNING: age_at_dose column not found, Unpivot age columns were not created.")

    unpivot.columns.name = None
    print(f"INFO: Created Unpivot sheet with {len(unpivot)} row(s)")
    return unpivot


def highlight_invalid_age_hn_cells(writer: pd.ExcelWriter, combined: pd.DataFrame, sheet_name: str) -> None:
    normalized_columns = {str(col).strip().lower().replace(" ", "_"): col for col in combined.columns}
    hn_column = normalized_columns.get("hn")
    age_at_dose_column = normalized_columns.get("age_at_dose")
    sex_column = normalized_columns.get("sex")

    if not hn_column or not age_at_dose_column:
        print("WARNING: HN highlight skipped because hn or age_at_dose column is missing.")
        return

    invalid_age_mask = pd.to_numeric(combined[age_at_dose_column], errors="coerce").eq(99999)
    if not invalid_age_mask.any():
        print("INFO: No age_at_dose=99999 rows found for HN highlight.")

    worksheet = writer.sheets.get(sheet_name)
    if worksheet is None:
        print(f"WARNING: HN highlight skipped because sheet '{sheet_name}' was not found.")
        return

    hn_col_index = combined.columns.get_loc(hn_column) + 1
    red_font = Font(color="00FF0000")
    highlighted = 0

    for row_index in combined.index[invalid_age_mask]:
        excel_row = int(row_index) + 2  # Header is row 1 in Excel.
        worksheet.cell(row=excel_row, column=hn_col_index).font = red_font
        highlighted += 1

    print(f"INFO: Highlighted {highlighted} HN cell(s) in red where age_at_dose is 99999.")

    if not sex_column:
        print("WARNING: Sex NA fill skipped because sex column is missing.")
        return

    na_sex_mask = combined[sex_column].astype(str).str.strip().str.upper().eq("NA")
    if not na_sex_mask.any():
        print("INFO: No sex='NA' rows found for red fill.")
        return

    sex_col_index = combined.columns.get_loc(sex_column) + 1
    red_fill = PatternFill(fill_type="solid", start_color="00FF0000", end_color="00FF0000")
    filled = 0

    for row_index in combined.index[na_sex_mask]:
        excel_row = int(row_index) + 2  # Header is row 1 in Excel.
        worksheet.cell(row=excel_row, column=sex_col_index).fill = red_fill
        filled += 1

    print(f"INFO: Applied red fill to {filled} sex cell(s) where value is NA.")


def combine_sheets(input_path: Path, output_path: Path, sheet_name: str = "Combined") -> None:
    workbook = pd.ExcelFile(input_path)
    frames: list[pd.DataFrame] = []

    for source_sheet in workbook.sheet_names:
        frame = pd.read_excel(input_path, sheet_name=source_sheet)
        frame = frame.loc[:, ~frame.columns.astype(str).str.match(r"^Unnamed")]
        frames.append(frame)

    combined = pd.concat(frames, ignore_index=True)
    combined = convert_thai_buddhist_dates(combined)
    combined = normalize_sex_column(combined)
    combined = standardize_vaccine_names(combined)

    normalized_columns = {str(col).strip().lower(): col for col in combined.columns}
    hn_column = normalized_columns.get("hn")
    vaccine_column = normalized_columns.get("vaccine")

    if hn_column and vaccine_column:
        # Normalize values so same pairs with different case/spacing are treated as duplicates.
        combined["__hn_key"] = combined[hn_column].astype(str).str.strip().str.lower()
        combined["__vaccine_key"] = combined[vaccine_column].astype(str).str.strip().str.lower()

        duplicate_mask = combined.duplicated(subset=["__hn_key", "__vaccine_key"], keep="first")
        duplicate_count = int(duplicate_mask.sum())

        if duplicate_count > 0:
            print(
                "WARNING: "
                f"{duplicate_count} duplicate row(s) found based on '{hn_column}' + '{vaccine_column}'. "
                "Keeping first occurrence and deleting the rest."
            )
            combined = combined.loc[~duplicate_mask].reset_index(drop=True)
        else:
            print(f"No duplicates found based on '{hn_column}' + '{vaccine_column}'.")

        combined = combined.drop(columns=["__hn_key", "__vaccine_key"], errors="ignore")
    else:
        print(
            "WARNING: Duplicate check skipped because required columns were not found: "
            "hn and vaccine."
        )

    date_column_targets = {"birth_date", "visit_date", "vaccine_date"}
    date_columns_to_format: list[str] = []

    for column in combined.columns:
        normalized_column = str(column).strip().lower().replace(" ", "_")
        if normalized_column in date_column_targets:
            combined[column] = pd.to_datetime(combined[column], format="%Y-%m-%d", errors="coerce")
            date_columns_to_format.append(str(column))

    if date_columns_to_format:
        print(
            "INFO: Saved as Excel date format (yyyy-mm-dd) for column(s): "
            f"{', '.join(date_columns_to_format)}"
        )
    else:
        print("WARNING: Date-format columns not found: birth_date, visit_date, vaccine_date")

    combined = add_age_at_dose_column(combined)
    combined = add_visit_age_column(combined)

    summary = create_summary_sheet(combined)
    yearly_cummu_summary = create_yearly_cummu_summary_sheet(combined)
    indicator = create_indicator_sheet(combined)
    alod_cummu_indicator = create_alod_cummu_indicator_sheet(combined)
    unpivot = create_unpivot_sheet(combined)

    with pd.ExcelWriter(output_path, engine="openpyxl", date_format="yyyy-mm-dd", datetime_format="yyyy-mm-dd") as writer:
        combined.to_excel(writer, index=False, sheet_name=sheet_name)
        summary.to_excel(writer, index=False, sheet_name="Summary")
        yearly_cummu_summary.to_excel(writer, index=False, sheet_name="yearly_cummu_summary")
        indicator.to_excel(writer, index=False, sheet_name="Indicator")
        alod_cummu_indicator.to_excel(writer, index=False, sheet_name="alod_cummu_indicator")
        unpivot.to_excel(writer, index=False, sheet_name="Unpivot")
        highlight_invalid_age_hn_cells(writer, combined, sheet_name)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Combine all sheets from an Excel workbook into one sheet.")
    parser.add_argument(
        "input_path",
        nargs="?",
        default="PRF_EPI_data.xlsx",
        help="Path to the source workbook.",
    )
    parser.add_argument(
        "output_path",
        nargs="?",
        default=f"Umpium_EPI_Quarterly_report_{date.today():%Y%m%d}.xlsx",
        help="Path to the output workbook.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input_path)
    output_path = Path(args.output_path)
    combine_sheets(input_path, output_path)
    print(f"Created {output_path} from {input_path}")


if __name__ == "__main__":
    main()