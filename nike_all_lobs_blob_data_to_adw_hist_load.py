import os
import re
import json
import zipfile
import tempfile
import time
from datetime import datetime

import pandas as pd


class SkipFileException(Exception):
    """Raised when a source file/sheet should be skipped without failing the full load."""
    pass

CONFIG_FILE = os.environ.get(
    "CONFIG_FILE",
    "/mnt/fss-rodn-iad-odi-fs/POC/config/nike_blob_loader.conf"
)

DEFAULT_INPUT_DIR = "/mnt/fss-rodn-iad-odi-fs/POC/Input/Nike"
DEFAULT_TARGET_TABLE = "CUSTOM_STG_ADW.DM_STG_FORECAST_MASTER_BLOB"

HEADER_SCAN_ROWS = 50
FLUSH_EVERY_ROWS = int(os.environ.get("FLUSH_EVERY_ROWS", "100000"))

# Process all matching files sequentially by default.
# Set PROCESS_ALL_FILES=N only if you want latest-file-only behavior.
PROCESS_ALL_FILES = os.environ.get("PROCESS_ALL_FILES", "Y").upper() in ("Y", "YES", "TRUE", "1")

# =====================================================
# File Config Details
# =====================================================
# All Nike forecast file types are included here.
# Apparel is included and will be loaded from ZIP.
# BOM Usage is intentionally skipped in this forecast loader.
# Lion Brothers uses the last date from filename because the last date is run date.

FILE_CONFIGS = [
    {
        "folder_name": "Accessories",
        "brand_name": "Nike",
        "lob": "Accessories",
        "file_type": "Accessories Forecast",
        "sheet_keywords": ["page"],
        "allow_zip": False,
        "date_rule": "first"
    },
    {
        "folder_name": "Apparel",
        "brand_name": "Nike",
        "lob": "Apparel",
        "file_type": "Apparel Forecast",
        "sheet_keywords": ["ap production plan - fm", "production plan - fm", "fm"],
        "allow_zip": True,
        "date_rule": "first"
    },
    {
        "folder_name": "Footwear",
        "brand_name": "Nike",
        "lob": "Footwear",
        "file_type": "Footwear Forecast",
        "sheet_keywords": ["FW Production Plan - FM", "production plan - fm", "finished material", "fm"],
        "allow_zip": False,
        "date_rule": "first"
    },
    {
        "folder_name": "Lion_Brothers",
        "brand_name": "Nike",
        "lob": "Lion Brothers",
        "file_type": "Lion Brothers Forecast",
        "sheet_keywords": ["ap production plan - fm", "production plan - fm", "finished material", "fm"],
        "allow_zip": False,
        "date_rule": "last"
    }
]


def read_config_file(config_file):
    config = {}

    if not os.path.exists(config_file):
        print("Config file not found, using environment/default values: " + config_file)
        return config

    with open(config_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()

            if line == "" or line.startswith("#") or "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()

            if len(value) >= 2 and value[0] == '"' and value[-1] == '"':
                value = value[1:-1]
            elif len(value) >= 2 and value[0] == "'" and value[-1] == "'":
                value = value[1:-1]

            config[key] = value

    return config


CONFIG = read_config_file(CONFIG_FILE)

INPUT_DIR = (
    os.environ.get("NIKE_INPUT_DIR")
    or os.environ.get("INPUT_DIR")
    or CONFIG.get("NIKE_INPUT_DIR")
    or CONFIG.get("INPUT_DIR")
    or DEFAULT_INPUT_DIR
)

# If someone passes the Apparel subfolder from the earlier Apparel-only test,
# move one level up so all Nike folders are available.
if INPUT_DIR.rstrip("/").endswith("/Apparel"):
    INPUT_DIR = os.path.dirname(INPUT_DIR.rstrip("/"))

TARGET_TABLE = os.environ.get("TARGET_TABLE") or CONFIG.get("TARGET_TABLE") or DEFAULT_TARGET_TABLE

DB_USER = os.environ.get("ORACLE_USER") or CONFIG.get("ORACLE_USER")
DB_PASSWORD = os.environ.get("ORACLE_PASSWORD")
DB_DSN = os.environ.get("ORACLE_DSN") or CONFIG.get("ORACLE_DSN")
TNS_ADMIN = os.environ.get("TNS_ADMIN") or CONFIG.get("TNS_ADMIN")

if TNS_ADMIN:
    os.environ["TNS_ADMIN"] = TNS_ADMIN


def get_cx_oracle():
    try:
        import cx_Oracle
        return cx_Oracle
    except ImportError:
        raise Exception(
            "cx_Oracle is not installed/visible to this Python process. "
            "Check PYTHONPATH in shell script/backend session."
        )


def validate_db_config():
    missing = []

    if not DB_USER:
        missing.append("ORACLE_USER")

    if not DB_PASSWORD:
        missing.append("ORACLE_PASSWORD")

    if not DB_DSN:
        missing.append("ORACLE_DSN")

    if TNS_ADMIN and not os.path.exists(os.path.join(TNS_ADMIN, "tnsnames.ora")):
        raise Exception("tnsnames.ora not found under TNS_ADMIN: " + TNS_ADMIN)

    if missing:
        raise Exception("Missing DB config values: " + ", ".join(missing))


def get_connection():
    validate_db_config()
    cx_Oracle = get_cx_oracle()

    print("Oracle Python driver used: cx_Oracle " + str(cx_Oracle.__version__))
    print("Oracle Client version: " + str(cx_Oracle.clientversion()))

    conn = cx_Oracle.connect(
        user=DB_USER,
        password=DB_PASSWORD,
        dsn=DB_DSN
    )

    return conn, cx_Oracle


def build_insert_sql():
    return """
        INSERT INTO {target_table}
        (
            BRAND_NAME,
            LOB,
            FILE_TYPE,
            SOURCE_FILE_NAME,
            DATA_INGESTION_DT,
            FILE_CREATION_DATE,
            FILE_CONTENT_JSON,
            SOURCE_COLUMN_LIST_JSON
        )
        VALUES
        (
            :brand_name,
            :lob,
            :file_type,
            :source_file_name,
            SYSTIMESTAMP,
            TO_DATE(:file_creation_date, 'DD-MM-YYYY'),
            :file_content_json,
            :source_column_list_json
        )
    """.format(target_table=TARGET_TABLE)


MONTH_TOKEN_MAP = {
    "jan": "01", "january": "01",
    "feb": "02", "february": "02",
    "mar": "03", "march": "03",
    "apr": "04", "april": "04",
    "may": "05",
    "jun": "06", "june": "06",
    "jul": "07", "july": "07",
    "aug": "08", "august": "08",
    "sep": "09", "sept": "09", "september": "09",
    "oct": "10", "october": "10",
    "nov": "11", "november": "11",
    "dec": "12", "december": "12"
}


def normalize_date_parts(year, month, day):
    year = str(year)
    day = str(day).zfill(2)
    month_text = str(month).strip().lower()

    if month_text.isdigit():
        month_num = month_text.zfill(2)
    else:
        month_num = MONTH_TOKEN_MAP.get(month_text[:3])

    if not month_num:
        return None

    # Validate actual calendar date before returning.
    # This prevents time tokens like 012039 from becoming invalid dates like 01-13-2039.
    try:
        parsed_date = datetime(int(year), int(month_num), int(day))
    except Exception:
        return None

    # Safety check: file/run date should not be far future.
    if parsed_date.year > datetime.now().year + 1:
        return None

    return day + "-" + month_num + "-" + year


def date_to_sort_key(date_text):
    try:
        return datetime.strptime(date_text, "%d-%m-%Y")
    except Exception:
        return datetime.min


def extract_date_candidates(file_name):
    candidates = []
    text = os.path.basename(file_name)

    # Example:
    #   2024May20
    #   2025June21
    #   2025July5
    #   2025Sep13
    pattern1 = re.compile(
        r"((?:19|20)\d{2})[\s_\-]*"
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|"
        r"January|February|March|April|May|June|July|August|September|October|November|December)"
        r"[\s_\-]*(\d{1,2})",
        re.IGNORECASE
    )

    for match in pattern1.finditer(text):
        normalized = normalize_date_parts(match.group(1), match.group(2), match.group(3))
        if normalized:
            candidates.append({"date": normalized, "start": match.start(), "end": match.end()})

    # Example:
    #   20260427
    #   2026_04_27
    #   2026-04-27
    pattern2 = re.compile(r"(?<!\d)((?:19|20)\d{2})[\-_]?(\d{2})[\-_]?(\d{2})(?!\d)")
    for match in pattern2.finditer(text):
        normalized = normalize_date_parts(match.group(1), match.group(2), match.group(3))
        if normalized:
            candidates.append({"date": normalized, "start": match.start(), "end": match.end()})

    # Example:
    #   27_04_2026
    #   27-04-2026
    # Separators are intentionally mandatory.
    # Do NOT match compact 6-digit values like 012039 because those are often time values in LB files.
    pattern3 = re.compile(r"(?<!\d)(\d{2})[\-_](\d{2})[\-_]((?:19|20)\d{2})(?!\d)")
    for match in pattern3.finditer(text):
        normalized = normalize_date_parts(match.group(3), match.group(2), match.group(1))
        if normalized:
            candidates.append({"date": normalized, "start": match.start(), "end": match.end()})

    deduped = []
    seen = set()

    for item in sorted(candidates, key=lambda x: x["start"]):
        key = (item["date"], item["start"], item["end"])
        if key not in seen:
            deduped.append(item)
            seen.add(key)

    return deduped


def extract_run_date_from_filename(file_name):
    """
    Explicitly handles Lion Brothers run date patterns:
      run date 2025June21
      run_date_2025July5
      run date_2025Sep13 (AD - Confidential)

    Returns DD-MM-YYYY or None.
    """
    text = os.path.basename(file_name)

    run_date_pattern = re.compile(
        r"run\s*[_\-]?\s*date\s*[_\-]?\s*"
        r"((?:19|20)\d{2})[\s_\-]*"
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|"
        r"January|February|March|April|May|June|July|August|September|October|November|December)"
        r"[\s_\-]*(\d{1,2})",
        re.IGNORECASE
    )

    match = run_date_pattern.search(text)

    if match:
        return normalize_date_parts(match.group(1), match.group(2), match.group(3))

    return None


def choose_file_creation_date(file_name, date_rule):
    # For Lion Brothers date_rule='last':
    # Prefer explicit run date when present. This avoids accidentally picking time suffixes.
    if date_rule == "last":
        run_date = extract_run_date_from_filename(file_name)

        if run_date:
            return run_date

    candidates = extract_date_candidates(file_name)

    if len(candidates) == 0:
        return datetime.now().strftime("%d-%m-%Y")

    if date_rule == "last":
        return candidates[-1]["date"]

    return candidates[0]["date"]


def clean_text(value):
    if pd.isna(value):
        return ""

    return str(value).strip().replace("\n", " ").replace("\r", " ").replace("|", " ")


def clean_value(value):
    if pd.isna(value):
        return None

    value = str(value).strip().replace("\n", " ").replace("\r", " ").replace("|", " ")

    if value == "":
        return None

    return value


def get_excel_engine(file_path):
    lower_file = file_path.lower()

    if lower_file.endswith((".xlsx", ".xlsm")):
        return "openpyxl"

    if lower_file.endswith(".xls"):
        return "xlrd"

    return None


def make_unique_columns(header_values):
    final_cols = []
    seen = {}

    for i, value in enumerate(header_values):
        col = clean_text(value)

        if col == "":
            col = "UNNAMED_COL_" + str(i + 1)

        if col in seen:
            seen[col] += 1
            col = col + "_" + str(seen[col])
        else:
            seen[col] = 1

        final_cols.append(col)

    return final_cols


def detect_header_row(raw_df):
    max_score = 0
    header_row_index = None
    scan_limit = min(HEADER_SCAN_ROWS, len(raw_df))

    if raw_df is None or len(raw_df) == 0:
        raise SkipFileException("Blank selected sheet. No rows available to detect header.")

    for idx in range(scan_limit):
        row = raw_df.iloc[idx]
        non_empty_count = 0

        for value in row:
            if clean_text(value) != "":
                non_empty_count += 1

        if non_empty_count > max_score:
            max_score = non_empty_count
            header_row_index = idx

    if header_row_index is None or max_score == 0:
        raise SkipFileException("Blank selected sheet. Could not detect header row.")

    return header_row_index


def remove_empty_unnamed_columns(df):
    columns_to_keep = []

    for col in df.columns:
        if not str(col).startswith("UNNAMED_COL_"):
            columns_to_keep.append(col)
        else:
            if not df[col].isna().all():
                columns_to_keep.append(col)

    return df[columns_to_keep]


def normalize_sheet_name(value):
    return re.sub(r"\s+", " ", str(value).strip().lower())


def select_excel_sheet(file_path, config):
    engine = get_excel_engine(file_path)
    xl = pd.ExcelFile(file_path, engine=engine)

    sheet_names = xl.sheet_names
    sheet_keywords = config.get("sheet_keywords", [])

    normalized_sheets = [
        {"original": sheet_name, "normalized": normalize_sheet_name(sheet_name)}
        for sheet_name in sheet_names
    ]

    normalized_keywords = [normalize_sheet_name(keyword) for keyword in sheet_keywords]

    for keyword in normalized_keywords:
        for sheet in normalized_sheets:
            if sheet["normalized"] == keyword:
                return sheet["original"]

    for keyword in normalized_keywords:
        for sheet in normalized_sheets:
            if keyword in sheet["normalized"]:
                return sheet["original"]

    raise Exception(
        "Required sheet not found. File: "
        + os.path.basename(file_path)
        + ", Expected sheet keywords: "
        + str(sheet_keywords)
        + ", Available sheets: "
        + str(sheet_names)
    )


def read_excel_source(file_path, config):
    engine = get_excel_engine(file_path)

    if engine is None:
        raise Exception("Unsupported Excel file extension: " + file_path)

    selected_sheet = select_excel_sheet(file_path, config)

    raw_df = pd.read_excel(
        file_path,
        sheet_name=selected_sheet,
        header=None,
        engine=engine,
        dtype=object
    )

    if raw_df is None or raw_df.empty:
        raise SkipFileException(
            "Blank selected sheet. File="
            + os.path.basename(file_path)
            + " | Sheet="
            + str(selected_sheet)
        )

    header_row_index = detect_header_row(raw_df)
    headers = make_unique_columns(raw_df.iloc[header_row_index].tolist())

    df = raw_df.iloc[header_row_index + 1:].copy()
    df.columns = headers

    df = df.dropna(how="all")
    df = remove_empty_unnamed_columns(df)

    if df.empty:
        raise SkipFileException(
            "No data rows found after header. File="
            + os.path.basename(file_path)
            + " | Sheet="
            + str(selected_sheet)
        )

    return df, selected_sheet


def is_bom_usage_file_name(file_name):
    lower_file = os.path.basename(file_name).lower()

    return (
        "bom usage" in lower_file
        or "bom_usage" in lower_file
        or "bom-usage" in lower_file
    )


def is_supported_file(file_name, allow_zip):
    lower_file = file_name.lower()

    if lower_file.startswith("~$"):
        return False

    if is_bom_usage_file_name(file_name):
        return False

    if allow_zip and lower_file.endswith(".zip"):
        return True

    return lower_file.endswith((".xlsx", ".xls", ".xlsm"))


def is_apparel_fm_file(file_name):
    lower_file = os.path.basename(file_name).lower()

    return lower_file.endswith((".xlsx", ".xls", ".xlsm")) \
        and "fm" in lower_file \
        and not is_bom_usage_file_name(lower_file)


def get_files_for_config(config):
    folder_path = os.path.join(INPUT_DIR, config["folder_name"])

    if not os.path.exists(folder_path):
        return []

    candidates = []

    for file_name in os.listdir(folder_path):
        full_path = os.path.join(folder_path, file_name)

        if os.path.isdir(full_path):
            continue

        if not is_supported_file(file_name, config.get("allow_zip", False)):
            continue

        file_creation_date = choose_file_creation_date(file_name, config.get("date_rule", "first"))

        candidates.append({
            "path": full_path,
            "file_name": file_name,
            "file_creation_date": file_creation_date,
            "sort_date": date_to_sort_key(file_creation_date),
            "mtime": os.path.getmtime(full_path)
        })

    if len(candidates) == 0:
        return []

    # Sequential processing order: older file creation date first, newer later.
    candidates = sorted(
        candidates,
        key=lambda x: (x["sort_date"], x["mtime"])
    )

    print("Available files for " + config["lob"] + " / " + config["file_type"] + ":")
    for item in candidates:
        print("  " + item["file_name"] + " | FileCreationDate=" + item["file_creation_date"])

    if PROCESS_ALL_FILES:
        return candidates

    # Optional latest-file-only mode.
    latest_candidates = sorted(
        candidates,
        key=lambda x: (x["sort_date"], x["mtime"]),
        reverse=True
    )

    return [latest_candidates[0]]


def insert_rows(cursor, conn, rows):
    if not rows:
        return 0

    cursor.executemany(build_insert_sql(), rows)
    conn.commit()
    return len(rows)


def insert_dataframe_as_blob(conn, cursor, df, config, source_file_name, file_creation_date):
    total_inserted = 0
    pending_rows = []
    columns = list(df.columns)
    source_columns = list(df.columns)
    source_column_list_json_bytes = json.dumps(source_columns, ensure_ascii=False).encode("utf-8")

    last_flush_time = time.time()

    for values in df.itertuples(index=False, name=None):
        row_payload = {}

        for i, col in enumerate(columns):
            row_payload[col] = clean_value(values[i])

        pending_rows.append({
            "brand_name": config["brand_name"],
            "lob": config["lob"],
            "file_type": config["file_type"],
            "source_file_name": source_file_name,
            "file_content_json": json.dumps(row_payload, ensure_ascii=False).encode("utf-8"),
            "source_column_list_json": source_column_list_json_bytes,
            "file_creation_date": file_creation_date
        })

        if FLUSH_EVERY_ROWS > 0 and len(pending_rows) >= FLUSH_EVERY_ROWS:
            flush_start = time.time()
            chunk_size = len(pending_rows)
            total_inserted += insert_rows(cursor, conn, pending_rows)
            flush_seconds = round(time.time() - flush_start, 2)
            total_seconds = round(time.time() - last_flush_time, 2)

            print(
                "Committed chunk | RowsInChunk="
                + str(chunk_size)
                + " | TotalInsertedForFile="
                + str(total_inserted)
                + " | InsertSeconds="
                + str(flush_seconds)
                + " | ChunkTotalSeconds="
                + str(total_seconds)
            )

            last_flush_time = time.time()
            pending_rows = []

    if pending_rows:
        flush_start = time.time()
        chunk_size = len(pending_rows)
        total_inserted += insert_rows(cursor, conn, pending_rows)
        flush_seconds = round(time.time() - flush_start, 2)

        print(
            "Committed final chunk | RowsInChunk="
            + str(chunk_size)
            + " | TotalInsertedForFile="
            + str(total_inserted)
            + " | InsertSeconds="
            + str(flush_seconds)
        )

    return total_inserted


def process_excel_file(conn, cursor, file_path, config, source_file_name, file_creation_date):
    print("--------------------------------------------------")
    print("Processing Forecast file:")
    print("  LOB             : " + config["lob"])
    print("  File Type       : " + config["file_type"])
    print("  Source Name     : " + source_file_name)
    print("  FileCreationDate: " + file_creation_date)

    read_start = time.time()
    df, selected_sheet = read_excel_source(file_path, config)
    read_seconds = round(time.time() - read_start, 2)

    print(
        "Read complete | Rows="
        + str(len(df))
        + " | Columns="
        + str(len(df.columns))
        + " | Sheet="
        + str(selected_sheet)
        + " | ReadSeconds="
        + str(read_seconds)
    )

    insert_start = time.time()
    inserted = insert_dataframe_as_blob(
        conn=conn,
        cursor=cursor,
        df=df,
        config=config,
        source_file_name=source_file_name,
        file_creation_date=file_creation_date
    )
    insert_seconds = round(time.time() - insert_start, 2)

    print("Inserted rows: " + str(inserted) + " | InsertTotalSeconds=" + str(insert_seconds))
    return inserted


def process_zip_file(conn, cursor, zip_path, config, file_creation_date):
    zip_file_name = os.path.basename(zip_path)
    total_inserted = 0

    zip_start = time.time()

    with tempfile.TemporaryDirectory() as temp_dir:
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            members = zip_ref.namelist()

            target_members = [
                member for member in members
                if not member.endswith("/")
                and is_apparel_fm_file(member)
            ]

            if len(target_members) == 0:
                raise Exception("No Apparel FM Excel files found inside ZIP: " + zip_file_name)

            print("Target forecast files inside ZIP:")
            for member in sorted(target_members):
                print("  " + member)

            for member in sorted(target_members):
                extracted_path = zip_ref.extract(member, temp_dir)
                source_file_name = zip_file_name + "/" + os.path.basename(member)

                try:
                    total_inserted += process_excel_file(
                        conn=conn,
                        cursor=cursor,
                        file_path=extracted_path,
                        config=config,
                        source_file_name=source_file_name,
                        file_creation_date=file_creation_date
                    )
                except SkipFileException as e:
                    print(
                        "SKIPPED_FILE | Reason="
                        + str(e)
                        + " | Source="
                        + source_file_name
                    )
                    continue

    zip_seconds = round(time.time() - zip_start, 2)
    print("Total ZIP processing seconds: " + str(zip_seconds))

    return total_inserted


def process_config(conn, cursor, config):
    selected_files = get_files_for_config(config)

    if len(selected_files) == 0:
        return 0

    config_total = 0

    for selected_file in selected_files:
        print("==================================================")
        print("Selected file for " + config["lob"] + " / " + config["file_type"] + ":")
        print("  " + selected_file["path"])
        print("  FileCreationDate=" + selected_file["file_creation_date"])

        try:
            if selected_file["file_name"].lower().endswith(".zip"):
                config_total += process_zip_file(
                    conn=conn,
                    cursor=cursor,
                    zip_path=selected_file["path"],
                    config=config,
                    file_creation_date=selected_file["file_creation_date"]
                )
            else:
                config_total += process_excel_file(
                    conn=conn,
                    cursor=cursor,
                    file_path=selected_file["path"],
                    config=config,
                    source_file_name=selected_file["file_name"],
                    file_creation_date=selected_file["file_creation_date"]
                )

        except SkipFileException as e:
            print(
                "SKIPPED_FILE | Reason="
                + str(e)
                + " | Source="
                + selected_file["file_name"]
            )
            continue

    print(
        "Total inserted from "
        + config["lob"]
        + " / "
        + config["file_type"]
        + ": "
        + str(config_total)
    )

    return config_total


def print_runtime_config():
    print("==================================================")
    print("Nike Forecast Only BLOB Loader")
    print("CONFIG_FILE: " + str(CONFIG_FILE))
    print("INPUT_DIR: " + str(INPUT_DIR))
    print("TARGET_TABLE: " + str(TARGET_TABLE))
    print("ORACLE_USER: " + str(DB_USER))
    print("ORACLE_DSN: " + str(DB_DSN))
    print("TNS_ADMIN: " + str(TNS_ADMIN))
    print("FLUSH_EVERY_ROWS: " + str(FLUSH_EVERY_ROWS))
    print("PROCESS_ALL_FILES: " + str(PROCESS_ALL_FILES))
    print("==================================================")


def main():
    print_runtime_config()

    if not os.path.exists(INPUT_DIR):
        raise Exception("Input folder not found: " + INPUT_DIR)

    conn, cx_Oracle = get_connection()
    cursor = conn.cursor()

    cursor.setinputsizes(
        file_content_json=cx_Oracle.BLOB,
        source_column_list_json=cx_Oracle.BLOB
    )

    grand_total = 0

    try:
        for config in FILE_CONFIGS:
            grand_total += process_config(conn, cursor, config)

        print("==================================================")
        print("Nike Forecast Only BLOB load completed.")
        print("Grand total rows inserted into " + TARGET_TABLE + ": " + str(grand_total))

    except Exception as e:
        conn.rollback()
        print("ERROR occurred. Transaction rolled back for current connection/batch.")
        print(str(e))
        raise

    finally:
        cursor.close()
        conn.close()


if __name__ == "__main__":
    main()
