#!/usr/bin/env python3
"""
Test script to verify Section 232 Actions table data loading
"""
import sys
import sqlite3
import pandas as pd
from pathlib import Path

# Get database path
script_dir = Path(__file__).parent
DB_PATH = script_dir / "DerivativeMill" / "Resources" / "derivativemill.db"

print(f"Testing data loading from {DB_PATH}")
print(f"Database exists: {DB_PATH.exists()}")

# Test database connection and query
try:
    conn = sqlite3.connect(str(DB_PATH))
    df = pd.read_sql("""SELECT tariff_no, action, description, advalorem_rate,
                               effective_date, expiration_date, specific_rate,
                               additional_declaration, note, link
                        FROM sec_232_actions
                        ORDER BY tariff_no""", conn)
    conn.close()

    print(f"\nQuery successful!")
    print(f"Rows returned: {len(df)}")
    print(f"Columns: {df.columns.tolist()}")
    print(f"\nFirst 3 rows:")
    print(df.head(3))

except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
