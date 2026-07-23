#!/bin/bash

# ============================================================
# NIKE ALL LOBS - ODI EXCEL/ZIP TO CSV RUNNER
# ============================================================

export NIKE_INPUT_BASE=/mnt/fss-rodn-iad-odi-fs/project_conch/nike
export NIKE_CSV_OUTPUT_BASE=/mnt/fss-rodn-iad-odi-fs/project_conch/nike/file_conversion_to_csv
# Keep BOM Usage isolated from Forecast so the existing Forecast Groovy loader
# cannot ingest BOM CSVs into DM_STG_FORECAST_MASTER.
export NIKE_BOM_CSV_OUTPUT_BASE=/mnt/fss-rodn-iad-odi-fs/project_conch/nike/file_conversion_to_csv/bom_usage

# Use Y only for the current clean recovery run.
# Change this back to N after the ODI scenario succeeds.
export OVERWRITE_EXISTING=N

export LOB_FILTER=ALL
export MAX_FILES_PER_LOB=0
export ADD_METADATA_COLUMNS=Y
export ENABLE_BOM_USAGE=Y
export LOG_DIR=/mnt/fss-rodn-iad-odi-fs/project_conch/nike/logs

PYTHON_COMMAND=/usr/bin/python3
SCRIPT_PATH=/mnt/fss-rodn-iad-odi-fs/project_conch/nike/file_conversion_to_csv/nike_all_lobs_excel_to_csv_conversion.py

mkdir -p "${LOG_DIR}" || exit 1

# Create a separate log for every execution.
RUN_TIMESTAMP=$(date +%Y%m%d_%H%M%S)
RUN_LOG="${LOG_DIR}/nike_excel_to_csv_odi_run_${RUN_TIMESTAMP}_$$.log"

exec >"${RUN_LOG}" 2>&1

echo "============================================================"
echo "Nike Excel-to-CSV ODI execution started"
echo "Start time: $(date)"
echo "User: $(id)"
echo "Host: $(hostname)"
echo "Python: ${PYTHON_COMMAND}"
echo "Script: ${SCRIPT_PATH}"
echo "Log: ${RUN_LOG}"
echo "OVERWRITE_EXISTING: ${OVERWRITE_EXISTING}"
echo "ENABLE_BOM_USAGE: ${ENABLE_BOM_USAGE}"
echo "BOM CSV output: ${NIKE_BOM_CSV_OUTPUT_BASE}"
echo "============================================================"

# Prevent two conversion processes from running simultaneously.
LOCK_FILE=/tmp/nike_excel_to_csv_conversion.lock
exec 9>"${LOCK_FILE}"

if ! /usr/bin/flock -n 9; then
    echo "ERROR: Another Nike Excel-to-CSV conversion is already running."
    echo "Execution stopped to prevent CSV file corruption."
    exit 75
fi

echo "Conversion lock acquired successfully."

if [ ! -f "${SCRIPT_PATH}" ]; then
    echo "ERROR: Python script not found: ${SCRIPT_PATH}"
    exit 1
fi

if [ ! -x "${PYTHON_COMMAND}" ]; then
    echo "ERROR: Python executable not found: ${PYTHON_COMMAND}"
    exit 1
fi

"${PYTHON_COMMAND}" "${SCRIPT_PATH}"
RETURN_CODE=$?

echo "============================================================"
echo "Python return code: ${RETURN_CODE}"
echo "Execution completed: $(date)"
echo "============================================================"

exit ${RETURN_CODE}
