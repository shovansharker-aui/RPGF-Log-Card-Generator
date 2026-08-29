"""
test_trend.py

Tests TrendHelper
"""

import pandas as pd

from trend_helper import TrendHelper

# ---------------------------------------------------
# Load Equipment
# ---------------------------------------------------

df = pd.read_excel(
    "data/Equipment.xlsx",
    sheet_name="Equipment"
)

# ---------------------------------------------------
# Test first equipment
# ---------------------------------------------------

row = df.iloc[0]

helper = TrendHelper(row)

result = helper.get_month_schedule(
    month_name="Apr",
    friday_count=4,
    current_year=2026
)

print()
print(row["Equipment Name"])
print(result)
print()