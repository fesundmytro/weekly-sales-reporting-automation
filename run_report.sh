#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
mkdir -p logs
python3 weekly_report.py >> logs/manual_run.log 2>&1
