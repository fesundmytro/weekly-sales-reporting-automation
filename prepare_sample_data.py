#!/usr/bin/env python3
"""Create two sample weekly CSV files from sales_clean.csv."""

from pathlib import Path
import re
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
SOURCE_FILE = BASE_DIR / "sales_clean.csv"
INPUT_DIR = BASE_DIR / "input"


def normalize_column_name(column: object) -> str:
    text = str(column).replace("\ufeff", "").replace("ï»¿", "")
    text = text.strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


if not SOURCE_FILE.exists():
    raise FileNotFoundError(
        f"Place sales_clean.csv next to this script: {SOURCE_FILE}"
    )

INPUT_DIR.mkdir(parents=True, exist_ok=True)

df = pd.read_csv(SOURCE_FILE, encoding="utf-8-sig")
df.columns = [normalize_column_name(column) for column in df.columns]

df["order_date"] = pd.to_datetime(
    df["order_date"],
    dayfirst=True,
    format="mixed",
    errors="raise",
)

df["week_end"] = (
    df["order_date"]
    .dt.to_period("W-SUN")
    .dt.end_time
    .dt.normalize()
)

latest_week_end = df["week_end"].max()
max_order_date = df["order_date"].max().normalize()

# Exclude the last calendar week when the source file ends before Sunday.
if max_order_date.dayofweek != 6:
    available = sorted(df.loc[df["week_end"] < latest_week_end, "week_end"].unique())
else:
    available = sorted(df["week_end"].unique())

selected_weeks = available[-2:]
if len(selected_weeks) < 2:
    raise ValueError("The source file does not contain two complete weeks.")

for week_end in selected_weeks:
    weekly = df.loc[df["week_end"] == week_end].drop(columns="week_end")
    output = INPUT_DIR / f"sales_{pd.Timestamp(week_end):%Y-%m-%d}.csv"
    weekly.to_csv(output, index=False, encoding="utf-8-sig")
    print(f"Created {output.name}: {len(weekly):,} rows")
