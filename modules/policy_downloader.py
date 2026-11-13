import csv
import re
import shutil
import tempfile
import threading
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import urlsplit

import pandas as pd
import requests
from openpyxl import load_workbook


SHEET_ALIASES = {
    "Policy Details": "old",
    "Assistance Programme Details": "new",
}

COLUMN_ALIASES = {
    "policy_number": ["Policy Number", "Assistance Programme Number"],
    "traveller_name": ["Traveller Name"],
    "url": ["Policy Path", "Assistance Programme Path"],
    "hub_name": ["Hub", "Hub Name", "Hub_Name"],
}

MAX_DOWNLOAD_WORKERS = 8
DEFAULT_TIMEOUT = (10, 60)
DEFAULT_RETRIES = 2


def _normalize_header(value):
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def _select_sheet(workbook):
    for sheet_name in SHEET_ALIASES:
        if sheet_name in workbook.sheetnames:
            return workbook[sheet_name]
    expected = " or ".join(f"'{name}'" for name in SHEET_ALIASES)
    raise ValueError(f"Required worksheet not found. Expected {expected}.")


def _detect_columns(ws):
    alias_map = {
        _normalize_header(alias): internal
        for internal, aliases in COLUMN_ALIASES.items()
        for alias in aliases
    }
    required = {"policy_number", "traveller_name", "url"}
    for header_row in range(1, min(ws.max_row, 5) + 1):
        positions = {}
        for col in range(1, ws.max_column + 1):
            normalized = _normalize_header(ws.cell(header_row, col).value)
            if normalized in alias_map:
                positions.setdefault(alias_map[normalized], col)
        if required.issubset(positions):
            return header_row, positions
    raise ValueError(
        f"Could not find required Policy Number, Traveller Name, and Policy Path "
        f"headers in the first five rows of '{ws.title}'."
    )


def read_policy_file(uploaded_file):
    """Read the selected supported worksheet for the page summary and preview."""
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    workbook = load_workbook(uploaded_file, read_only=True, data_only=True)
    ws = _select_sheet(workbook)
    sheet_name = ws.title
    workbook.close()
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def get_policy_summary(df):
    total_policies = len(df)
    charge_col = next(
        (
            col for col in df.columns
            if any(term in str(col).casefold() for term in ("charge", "premium", "amount"))
        ),
        None,
    )
    if charge_col is None:
        return total_policies, 0
    charges = (
        df[charge_col]
        .fillna("")
        .astype(str)
        .str.replace(r"[^\d.\-]", "", regex=True)
        .str.strip()
    )
    total_charges = float(pd.to_numeric(charges, errors="coerce").fillna(0).sum())
    return total_policies, total_charges


def _text_value(value):
    if value is None:
        return ""
    return str(value).strip()


def read_policy_links(uploaded_file):
    """Extract links only from the column identified as Policy Path."""
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    workbook = load_workbook(uploaded_file, data_only=True, read_only=False)
    ws = _select_sheet(workbook)
    header_row, positions = _detect_columns(ws)
    records = []

    for row_idx in range(header_row + 1, ws.max_row + 1):
        policy_number = _text_value(ws.cell(row_idx, positions["policy_number"]).value)
        traveller_name = _text_value(ws.cell(row_idx, positions["traveller_name"]).value)
        path_cell = ws.cell(row_idx, positions["url"])
        # The visible label is not a URL. Only trust the hyperlink on the
        # header-matched Policy Path cell, never a nearby traveller hyperlink.
        url = _text_value(path_cell.hyperlink.target if path_cell.hyperlink else "")
        hub_name = (
            _text_value(ws.cell(row_idx, positions["hub_name"]).value)
            if "hub_name" in positions else ""
        )

        if not any((policy_number, traveller_name, url, hub_name)):
            continue

        invalid_reasons = []
        if not policy_number:
            invalid_reasons.append("Missing policy number")
        if not traveller_name:
            invalid_reasons.append("Missing traveller name")
        if not url:
            invalid_reasons.append("Missing PDF hyperlink")
        else:
            parsed = urlsplit(url)
            if parsed.scheme.casefold() not in ("http", "https") or not parsed.netloc:
                invalid_reasons.append("Invalid PDF hyperlink (expected an HTTP or HTTPS URL)")

        records.append({
            "policy_number": policy_number,
            "traveller_name": traveller_name,
            "url": url,
            "hub_name": hub_name,
            "is_valid": not invalid_reasons,
            "validation_error": "; ".join(invalid_reasons),
        })

    workbook.close()
    return pd.DataFrame(records, columns=[
        "policy_number", "traveller_name", "url", "hub_name", "is_valid", "validation_error",
    ])


def clean_filename(text):
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(text))
    text = text.strip().rstrip(" .")
    if not text:
        text = "unnamed"
    if text.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
        text = f"_{text}"
    return text


def _session_for_worker(thread_sessions, worker_sessions, session_lock):
    session = getattr(thread_sessions, "session", None)
    if session is None:
        session = requests.Session()
        thread_sessions.session = session
        with session_lock:
            worker_sessions.append(session)
    return session


def _is_expired_response(response):
    body = response.raw.read(8192, decode_content=True)
    lowered = body.decode("utf-8", errors="ignore").casefold()
    return any(marker in lowered for marker in (
        "requestexpired", "expiredtoken", "request has expired", "signature expired",
    ))


def _download_one(record, filepath, timeout, retries, thread_sessions, worker_sessions, session_lock):
    url = record["url"]
    part_path = filepath.with_suffix(filepath.suffix + ".part")
    for attempt in range(retries + 1):
        try:
            session = _session_for_worker(thread_sessions, worker_sessions, session_lock)
            with session.get(url, stream=True, timeout=timeout) as response:
                status = response.status_code
                if status in (400, 403) and _is_expired_response(response):
                    return False, "Signed URL expired"
                if status in (429, 500, 502, 503, 504) and attempt < retries:
                    time.sleep(0.4 * (attempt + 1))
                    continue
                if status != 200:
                    return False, f"HTTP {status} response"

                prefix = b""
                byte_count = 0
                try:
                    with part_path.open("wb") as output:
                        for chunk in response.iter_content(chunk_size=64 * 1024):
                            if not chunk:
                                continue
                            if len(prefix) < 5:
                                prefix += chunk[: 5 - len(prefix)]
                            output.write(chunk)
                            byte_count += len(chunk)
                except Exception:
                    part_path.unlink(missing_ok=True)
                    raise

                if byte_count == 0:
                    part_path.unlink(missing_ok=True)
                    return False, "Empty response"
                if not prefix.startswith(b"%PDF-"):
                    part_path.unlink(missing_ok=True)
                    return False, "Response is not a PDF"
                part_path.replace(filepath)
                return True, ""
        except (requests.Timeout, requests.ConnectionError) as exc:
            part_path.unlink(missing_ok=True)
            if attempt < retries:
                time.sleep(0.4 * (attempt + 1))
                continue
            return False, f"Network {type(exc).__name__} after {retries + 1} attempts"
        except requests.RequestException as exc:
            part_path.unlink(missing_ok=True)
            if isinstance(
                exc,
                (requests.exceptions.ChunkedEncodingError, requests.exceptions.ContentDecodingError),
            ) and attempt < retries:
                time.sleep(0.4 * (attempt + 1))
                continue
            return False, f"Request failed ({type(exc).__name__})"
        except OSError as exc:
            part_path.unlink(missing_ok=True)
            return False, f"File write failed ({type(exc).__name__})"
    return False, "Download failed after retries"


def _unique_pdf_name(traveller, policy_number, seen):
    stem = clean_filename(f"{traveller} [{policy_number}]")
    candidate = f"{stem}.pdf"
    suffix = 2
    while candidate.casefold() in seen:
        candidate = f"{stem}_{suffix}.pdf"
        suffix += 1
    seen.add(candidate.casefold())
    return candidate


def download_policy_pdfs(
    link_df,
    progress_callback=None,
    hub_wise=False,
    max_workers=MAX_DOWNLOAD_WORKERS,
    timeout=DEFAULT_TIMEOUT,
    retries=DEFAULT_RETRIES,
):
    """Download validated rows concurrently and build an integrity-checked ZIP.

    Returns (zip_path, downloaded_count, failed_count, elapsed_seconds, failures).
    Failure rows contain policy_number, traveller_name, and reason; URLs are omitted.
    """
    if link_df.empty:
        raise ValueError("No policy records were found in the workbook.")

    base_dir = Path(tempfile.mkdtemp(prefix="policy_download_"))
    valid_rows = []
    failures = []
    for _, row in link_df.iterrows():
        record = row.to_dict()
        if bool(record.get("is_valid", False)):
            valid_rows.append(record)
        else:
            failures.append({
                "policy_number": _text_value(record.get("policy_number")),
                "traveller_name": _text_value(record.get("traveller_name")),
                "reason": _text_value(record.get("validation_error")) or "Invalid record",
            })

    folder_record = next(iter(link_df.to_dict("records")), {})
    first_traveller = clean_filename(_text_value(folder_record.get("traveller_name")) or "Policy Downloads")
    first_policy = clean_filename(_text_value(folder_record.get("policy_number")) or "Batch")
    folder_name = f"{first_traveller} [{first_policy}]"
    downloads_dir = base_dir / folder_name
    downloads_dir.mkdir(parents=True, exist_ok=True)

    hub_counts = {}
    if hub_wise:
        for record in valid_rows:
            hub = _text_value(record.get("hub_name")) or "Unknown Hub"
            hub_counts[hub] = hub_counts.get(hub, 0) + 1

    seen_names = set()
    jobs = []
    for record in valid_rows:
        policy_number = _text_value(record.get("policy_number"))
        traveller = _text_value(record.get("traveller_name"))
        filename = _unique_pdf_name(traveller, policy_number, seen_names)
        destination_dir = downloads_dir
        if hub_wise:
            hub = _text_value(record.get("hub_name")) or "Unknown Hub"
            destination_dir = downloads_dir / clean_filename(f"{hub}_{hub_counts[hub]}_Policies")
            destination_dir.mkdir(parents=True, exist_ok=True)
        jobs.append((record, destination_dir / filename))

    start_time = time.monotonic()
    completed = len(failures)
    downloaded = 0
    total = len(link_df)
    if progress_callback and completed:
        elapsed = time.monotonic() - start_time
        progress_callback(completed, total, elapsed, 0)

    max_workers = max(1, min(int(max_workers), MAX_DOWNLOAD_WORKERS, len(jobs) or 1))
    thread_sessions = threading.local()
    session_lock = threading.Lock()
    worker_sessions = []
    try:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(
                    _download_one,
                    record,
                    filepath,
                    timeout,
                    retries,
                    thread_sessions,
                    worker_sessions,
                    session_lock,
                ): (record, filepath)
                for record, filepath in jobs
            }
            for future in as_completed(futures):
                record, filepath = futures[future]
                try:
                    success, reason = future.result()
                except Exception as exc:
                    success, reason = False, f"Download failed ({type(exc).__name__})"
                if success:
                    downloaded += 1
                else:
                    failures.append({
                        "policy_number": _text_value(record.get("policy_number")),
                        "traveller_name": _text_value(record.get("traveller_name")),
                        "reason": reason,
                    })
                completed += 1
                if progress_callback:
                    elapsed = time.monotonic() - start_time
                    remaining = max(total - completed, 0)
                    eta = int((elapsed / completed) * remaining) if completed else 0
                    progress_callback(completed, total, elapsed, eta)
    finally:
        with session_lock:
            sessions = list(worker_sessions)
        for session in sessions:
            session.close()

    csv_path = base_dir / "failure_report.csv"
    with csv_path.open("w", newline="", encoding="utf-8-sig") as report_file:
        writer = csv.DictWriter(report_file, fieldnames=["policy_number", "traveller_name", "reason"])
        writer.writeheader()
        writer.writerows(failures)

    zip_name = f"{folder_name}.zip" if total == 1 else f"{folder_name} x {total}.zip"
    zip_path = base_dir / zip_name
    try:
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for pdf_file in downloads_dir.rglob("*.pdf"):
                archive.write(pdf_file, arcname=str(pdf_file.relative_to(base_dir)))
            archive.write(csv_path, arcname="failure_report.csv")
        with zipfile.ZipFile(zip_path, "r") as archive:
            bad_member = archive.testzip()
            if bad_member is not None:
                raise zipfile.BadZipFile(f"Corrupt archive member: {bad_member}")
    except Exception:
        shutil.rmtree(base_dir, ignore_errors=True)
        raise

    return zip_path, downloaded, len(failures), time.monotonic() - start_time, failures
