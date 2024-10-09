import os
import pandas as pd
import sqlite3
import time
import logging
from time import strftime
from pathlib import Path

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.FileHandler("logs/file_to_sqlite.log"), logging.StreamHandler()],
)


def file_to_sqlite(file_path, sqlite_db, table_name):
    logging.info(f"Starting conversion of {file_path} to SQLite database {sqlite_db}.")

    # Determine the file type
    if file_path.endswith(".csv"):
        logging.info("File format is CSV.")
        df = pd.read_csv(file_path)
    elif file_path.endswith(".xlsx"):
        logging.info("File format is Excel.")
        df = pd.read_excel(file_path)
    else:
        logging.error("Unsupported file format. Please use .csv or .xlsx files.")
        raise ValueError("Unsupported file format. Please use .csv or .xlsx files.")

    check_for_existing_db(sqlite_db)

    # Connect to the SQLite3 database (or create it if it doesn't exist)
    conn = sqlite3.connect(sqlite_db)
    logging.info(f"Connected to SQLite database {sqlite_db}.")

    # Write the data to the SQLite3 table
    df.to_sql(table_name, conn, if_exists="replace", index=False)
    logging.info(f"Data written to table {table_name} in SQLite database {sqlite_db}.")

    # Close the database connection
    conn.close()
    logging.info(f"Connection to SQLite database {sqlite_db} closed.")


def check_for_existing_db(sqlite_db):
    existing_db = Path(sqlite_db)

    if existing_db.exists():
        stat = existing_db.stat()
        try:
            creation_time = stat.st_birthtime
        except AttributeError:
            # st_birthtime is not available, fallback to ctime
            creation_time = stat.st_ctime

        date = strftime("%Y%m%d%H%M", time.localtime(creation_time))
        new_db = existing_db.with_name(f"{date}_{existing_db.name}")

        existing_db.rename(new_db)
        logging.info(f"Existing database renamed to {new_db}.")
    else:
        logging.info(f"No existing database found. Proceeding without renaming.")


if __name__ == "__main__":
    file_to_sqlite(
        "_CSV/202408071320_gor_diva_merged_cleaned_final.xlsx", "database.db", "assets"
    )
