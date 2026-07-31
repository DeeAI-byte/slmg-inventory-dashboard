"""
create_database.py
──────────────────
Builds secondary.duckdb from four Secondary Excel files.

Usage:
    python create_database.py

Place this script in the same folder as your app.py.
The four Excel files must be in:
    Secondary Files/Secondary1.xlsx
    Secondary Files/Secondary2.xlsx
    Secondary Files/Secondary3.xlsx
    Secondary Files/Secondary4.xlsx

Run this script ONCE (or whenever the source data changes).
The resulting secondary.duckdb is what app.py reads at runtime.
"""

import gc
import os
import sys

import duckdb
import pandas as pd

# ── CONFIG ────────────────────────────────────────────────────────────────────

SOURCE_FILES = [
    os.path.join("Secondary Files", "Secondary1.xlsx"),
    os.path.join("Secondary Files", "Secondary2.xlsx"),
    os.path.join("Secondary Files", "Secondary3.xlsx"),
    os.path.join("Secondary Files", "Secondary4.xlsx"),
]

DB_PATH    = "secondary.duckdb"
TABLE_NAME = "secondary"

# ── HELPERS ───────────────────────────────────────────────────────────────────

def validate_files() -> None:
    """Confirm all source files exist before starting."""
    missing = [f for f in SOURCE_FILES if not os.path.exists(f)]
    if missing:
        print("ERROR — the following files were not found:")
        for m in missing:
            print(f"  {m}")
        sys.exit(1)


def read_excel_stripped(path: str) -> pd.DataFrame:
    """Read one Excel file and strip column-name whitespace."""
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = df.columns.str.strip()
    return df


def validate_columns(reference: list[str], incoming: list[str], file_path: str) -> None:
    """Abort if column names do not match the first file."""
    if reference != incoming:
        missing_in_new  = set(reference) - set(incoming)
        extra_in_new    = set(incoming)  - set(reference)
        print(f"\nERROR — column mismatch in: {file_path}")
        if missing_in_new:
            print(f"  Missing columns : {sorted(missing_in_new)}")
        if extra_in_new:
            print(f"  Extra columns   : {sorted(extra_in_new)}")
        sys.exit(1)


# ── MAIN ─────────────────────────────────────────────────────────────────────

def build_database() -> None:
    validate_files()

    # Remove any existing database so we always start clean
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        print(f"Removed existing {DB_PATH}")

    con = duckdb.connect(DB_PATH)
    reference_columns: list[str] | None = None
    total_rows = 0

    for idx, path in enumerate(SOURCE_FILES, start=1):
        print(f"\n[{idx}/{len(SOURCE_FILES)}] Reading: {path}")

        # ── Read one file at a time ──────────────────────────────────────────
        df = read_excel_stripped(path)
        rows_this_file = len(df)
        print(f"  Rows read     : {rows_this_file:,}")
        print(f"  Columns       : {list(df.columns)}")

        # ── Column validation ────────────────────────────────────────────────
        if reference_columns is None:
            reference_columns = list(df.columns)
        else:
            validate_columns(reference_columns, list(df.columns), path)

        # ── Insert into DuckDB ───────────────────────────────────────────────
        # Register the DataFrame so DuckDB can reference it by name
        con.register("_chunk", df)

        if idx == 1:
            # Create the table from the first chunk — schema is inferred automatically
            con.execute(f"CREATE TABLE {TABLE_NAME} AS SELECT * FROM _chunk")
        else:
            con.execute(f"INSERT INTO {TABLE_NAME} SELECT * FROM _chunk")

        con.unregister("_chunk")

        # ── Release memory before reading the next file ──────────────────────
        del df
        gc.collect()

        total_rows += rows_this_file
        running = con.execute(f"SELECT COUNT(*) FROM {TABLE_NAME}").fetchone()[0]
        print(f"  Running total : {running:,} rows in database")

    # ── Final verification ───────────────────────────────────────────────────
    print("\n" + "─" * 60)
    final_count = con.execute(f"SELECT COUNT(*) FROM {TABLE_NAME}").fetchone()[0]
    col_count   = len(con.execute(f"DESCRIBE {TABLE_NAME}").fetchall())

    print(f"Database       : {DB_PATH}")
    print(f"Table          : {TABLE_NAME}")
    print(f"Columns        : {col_count}")
    print(f"Total rows     : {final_count:,}")

    if final_count != total_rows:
        print(f"\nWARNING — expected {total_rows:,} rows but found {final_count:,}!")
    else:
        print(f"\n✅  All {final_count:,} rows imported successfully.")

    # ── Print schema ─────────────────────────────────────────────────────────
    print("\nSchema:")
    for row in con.execute(f"DESCRIBE {TABLE_NAME}").fetchall():
        print(f"  {row[0]:<30} {row[1]}")

    # ── Sample check ─────────────────────────────────────────────────────────
    print("\nFirst 3 rows:")
    print(con.execute(f"SELECT * FROM {TABLE_NAME} LIMIT 3").df().to_string(index=False))

    con.close()
    print(f"\nDatabase saved to: {os.path.abspath(DB_PATH)}")


if __name__ == "__main__":
    build_database()
