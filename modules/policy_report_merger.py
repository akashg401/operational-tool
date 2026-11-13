import io
import re

import pandas as pd
from openpyxl import Workbook, load_workbook


COLUMN_ALIASES = {
    "Policy Number": ["Policy Number", "Assistance Programme Number"],
    "Issue Date": ["Issue Date"],
    "Asego Partner": ["Asego Partner"],
    "Traveller Name": ["Traveller Name"],
    "Policy Amount": ["Policy Amount", "Assistance Programme Amount"],
    "Policy Path": ["Policy Path", "Assistance Programme Path"],
    "Passport": ["Passport"],
    "Hub Name": ["Hub Name"],
}

SUPPORTED_SHEETS = ["Assistance Programme Details", "Policy Details"]

REQUIRED_COLUMNS = [
    "Policy Number",
    "Traveller Name",
    "Policy Amount",
    "Policy Path",
    "Passport",
]


def _normalize_header(value):
    """Normalize whitespace, line breaks, and case for header matching."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().casefold()


def read_policy_report(uploaded_file):
    wb = load_workbook(uploaded_file, data_only=True)

    ws = next((wb[name] for name in SUPPORTED_SHEETS if name in wb.sheetnames), None)
    if ws is None:
        raise ValueError(
            "Required worksheet not found. Expected 'Assistance Programme Details' "
            "or 'Policy Details'."
        )

    alias_to_internal = {
        _normalize_header(alias): internal_name
        for internal_name, aliases in COLUMN_ALIASES.items()
        for alias in aliases
    }
    # The current Assistance Programme Details portal export does not include
    # Passport. Keep it in the output as blank for that schema; the older
    # Policy Details format still requires its existing Passport column.
    required_columns = REQUIRED_COLUMNS
    if ws.title == "Assistance Programme Details":
        required_columns = [name for name in REQUIRED_COLUMNS if name != "Passport"]

    header_row = None
    column_positions = {}
    # A valid header row must contain every required column, so titles and
    # partially populated rows in the report preamble are not mistaken for it.
    for row_idx in range(1, min(ws.max_row, 5) + 1):
        row_headers = {}
        for col_idx in range(1, ws.max_column + 1):
            normalized = _normalize_header(ws.cell(row_idx, col_idx).value)
            if normalized in alias_to_internal:
                row_headers.setdefault(alias_to_internal[normalized], col_idx)
        if all(name in row_headers for name in required_columns):
            header_row = row_idx
            column_positions = row_headers
            break

    if header_row is None:
        raise ValueError(
            f"Could not find a header row containing all required columns in the "
            f"first five rows of '{ws.title}'. Required: {', '.join(required_columns)}."
        )

    records = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        def get_value(column_name):
            col_idx = column_positions.get(column_name)
            if col_idx is None:
                return ""
            value = ws.cell(row_idx, col_idx).value
            return value if value is not None else ""

        policy_number = get_value("Policy Number")
        traveller_name = get_value("Traveller Name")
        if not policy_number and not traveller_name:
            continue

        path_cell = ws.cell(row_idx, column_positions["Policy Path"])
        policy_url = ""
        if path_cell.hyperlink:
            policy_url = path_cell.hyperlink.target
        elif isinstance(path_cell.value, str) and path_cell.value.startswith(("http://", "https://")):
            policy_url = path_cell.value

        records.append({
            "Policy Number": policy_number,
            "Issue Date": get_value("Issue Date"),
            "Asego Partner": get_value("Asego Partner"),
            "Traveller Name": traveller_name,
            "Policy Amount": get_value("Policy Amount"),
            "Policy Path": "🔗 Click Here",
            "Policy Path URL": policy_url,
            "Passport": get_value("Passport"),
            "Hub Name": get_value("Hub Name"),
        })

    return pd.DataFrame(records, columns=[
        "Policy Number", "Issue Date", "Asego Partner", "Traveller Name",
        "Policy Amount", "Policy Path", "Policy Path URL", "Passport", "Hub Name",
    ])


def merge_policy_reports(files):
    all_frames = []
    total_rows = 0

    for file in files:
        if hasattr(file, "seek"):
            file.seek(0)
        df = read_policy_report(file)
        total_rows += len(df)
        all_frames.append(df)

    if not all_frames:
        return pd.DataFrame(), {
            "files_uploaded": 0,
            "rows_read": 0,
            "duplicates_removed": 0,
            "final_records": 0,
        }

    merged_df = pd.concat(all_frames, ignore_index=True)
    before = len(merged_df)
    merged_df = merged_df.drop_duplicates(subset=["Policy Number"], keep="first")
    summary = {
        "files_uploaded": len(all_frames),
        "rows_read": total_rows,
        "duplicates_removed": before - len(merged_df),
        "final_records": len(merged_df),
    }
    return merged_df, summary


def export_merged_report(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Policy Details"

    export_df = df.drop(columns=["Policy Path URL"], errors="ignore")
    for col_idx, column in enumerate(export_df.columns, start=1):
        ws.cell(row=1, column=col_idx, value=column)

    for row_idx, row in enumerate(export_df.itertuples(index=False, name=None), start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    if "Policy Path" in export_df.columns and "Policy Path URL" in df.columns:
        policy_path_col = export_df.columns.get_loc("Policy Path") + 1
        for row_idx, url in enumerate(df["Policy Path URL"], start=2):
            if isinstance(url, str) and url.startswith(("http://", "https://")):
                ws.cell(row=row_idx, column=policy_path_col).hyperlink = url

    wb.save(output)
    output.seek(0)
    return output.getvalue()
