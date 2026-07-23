#!/bin/bash
set -e

# =====================================================
# Nike Apparel BLOB Loader Shell
# Purpose:
#   Called from ODI OS Command.
#   Password comes from ODI variable as first argument.
#   Loads only Nike Apparel forecast ZIP using BLOB Python loader.
# =====================================================

CONFIG_FILE="/mnt/fss-rodn-iad-odi-fs/POC/config/nike_blob_loader.conf"
PY_SCRIPT="/mnt/fss-rodn-iad-odi-fs/POC/nike_all_lobs_blob_data_to_adw_hist_load.py"
LOG_DIR="/mnt/fss-rodn-iad-odi-fs/POC/logs"

# Use Friday BLOB POC table / change this to main table if required.
TARGET_TABLE="CUSTOM_STG_ADW.DM_STG_FORECAST_MASTER_BLOB"

# Apparel input folder.
INPUT_DIR="/mnt/fss-rodn-iad-odi-fs/POC/Input/Nike/Apparel"

# Oracle wallet and instant client.
TNS_ADMIN="/mnt/fss-rodn-iad-odi-fs/POC/config/wallet_nonprodidsadb"
INSTANT_CLIENT="/mnt/fss-rodn-iad-odi-fs/POC/oracle_client/instantclient_19_31"

# Python package paths for ODI runtime.
export PYTHONPATH="/home/oracle/.local/lib64/python3.6/site-packages:/home/oracle/.local/lib/python3.6/site-packages:${PYTHONPATH}"
export LD_LIBRARY_PATH="${INSTANT_CLIENT}:${LD_LIBRARY_PATH}"

export CONFIG_FILE
export TNS_ADMIN
export INPUT_DIR
export TARGET_TABLE
export FLUSH_EVERY_ROWS="100000"

# Password should come from ODI variable as first argument.
export ORACLE_PASSWORD="$1"

mkdir -p "$LOG_DIR"

RUN_TS=$(date +%Y%m%d_%H%M%S)
LOG_FILE="${LOG_DIR}/nike_apparel_blob_loader_${RUN_TS}.log"

if [ ! -f "$CONFIG_FILE" ]; then
    echo "ERROR: Config file not found: $CONFIG_FILE" | tee -a "$LOG_FILE"
    exit 1
fi

if [ ! -f "$PY_SCRIPT" ]; then
    echo "ERROR: Python script not found: $PY_SCRIPT" | tee -a "$LOG_FILE"
    exit 1
fi

if [ ! -f "$TNS_ADMIN/tnsnames.ora" ]; then
    echo "ERROR: tnsnames.ora not found under TNS_ADMIN=$TNS_ADMIN" | tee -a "$LOG_FILE"
    exit 1
fi

if [ ! -d "$INSTANT_CLIENT" ]; then
    echo "ERROR: Instant Client folder not found: $INSTANT_CLIENT" | tee -a "$LOG_FILE"
    exit 1
fi

if [ -z "$ORACLE_PASSWORD" ]; then
    echo "ERROR: Oracle password not received from ODI variable." | tee -a "$LOG_FILE"
    exit 1
fi

echo "Nike Apparel BLOB Load Started: $(date)" | tee -a "$LOG_FILE"
echo "CONFIG_FILE=$CONFIG_FILE" | tee -a "$LOG_FILE"
echo "PY_SCRIPT=$PY_SCRIPT" | tee -a "$LOG_FILE"
echo "INPUT_DIR=$INPUT_DIR" | tee -a "$LOG_FILE"
echo "TARGET_TABLE=$TARGET_TABLE" | tee -a "$LOG_FILE"
echo "TNS_ADMIN=$TNS_ADMIN" | tee -a "$LOG_FILE"
echo "LD_LIBRARY_PATH=$LD_LIBRARY_PATH" | tee -a "$LOG_FILE"
echo "PYTHONPATH=$PYTHONPATH" | tee -a "$LOG_FILE"
echo "Password length: ${#ORACLE_PASSWORD}" | tee -a "$LOG_FILE"

echo "Checking cx_Oracle..." | tee -a "$LOG_FILE"
python3 - <<'PY' 2>&1 | tee -a "$LOG_FILE"
import cx_Oracle
print("cx_Oracle OK:", cx_Oracle.__version__)
print("Oracle Client:", cx_Oracle.clientversion())
PY

echo "Starting Python script..." | tee -a "$LOG_FILE"

python3 -u "$PY_SCRIPT" 2>&1 | tee -a "$LOG_FILE"

EXIT_STATUS=${PIPESTATUS[0]}

echo "Nike Apparel BLOB Load Ended: $(date)" | tee -a "$LOG_FILE"
echo "Exit Status: $EXIT_STATUS" | tee -a "$LOG_FILE"

exit $EXIT_STATUS

