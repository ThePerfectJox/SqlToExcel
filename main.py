import os
from datetime import datetime
import mysql.connector
import pandas as pd

SQL_QUERY = """
SELECT *
FROM dummies;
"""

OUTPUT_FOLDER_PATH = r"C:\Users\jsubagyo\OneDrive - Philip Morris International\ExportSQL"

MYSQL_HOST = os.environ.get("MYSQL_HOST")

if not MYSQL_HOST:
    raise RuntimeError(
        "MYSQL_HOST environment variable has not been configured."
    )

MYSQL_PORT = os.environ.get("MYSQL_PORT")

if not MYSQL_PORT:
    raise RuntimeError(
        "MYSQL_PORT environment variable has not been configured."
    )

MYSQL_DATABASE = os.environ.get("MYSQL_DATABASE")

if not MYSQL_DATABASE:
    raise RuntimeError(
        "MYSQL_DATABASE environment variable has not been configured."
    )

MYSQL_USER = os.environ.get("MYSQL_USER")

if not MYSQL_USER:
    raise RuntimeError(
        "MYSQL_USER environment variable has not been configured."
    )

MYSQL_PASSWORD = os.environ.get("MYSQL_PASSWORD")

if not MYSQL_PASSWORD:
    raise RuntimeError(
        "MYSQL_PASSWORD environment variable has not been configured."
    )

def export_mysql_to_excel():
    connection = None

    try:
        os.makedirs(OUTPUT_FOLDER_PATH, exist_ok=True)

        connection = mysql.connector.connect(
            host=MYSQL_HOST,
            port=MYSQL_PORT,
            database=MYSQL_DATABASE,
            user=MYSQL_USER,
            password=MYSQL_PASSWORD
        )

        print("Connected to MySQL successfully.")

        dataframe = pd.read_sql_query(
            SQL_QUERY,
            connection
        )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"mysql_export_{timestamp}.xlsx"

        output_file_path = os.path.join(
            OUTPUT_FOLDER_PATH,
            output_filename
        )

        dataframe.to_excel(
            output_file_path,
            index=False,
            engine="openpyxl"
        )

        print(f"Export completed successfully.")
        print(f"Rows exported: {len(dataframe)}")
        print(f"Output file: {output_file_path}")

    except mysql.connector.Error as error:
        print(f"MySQL error: {error}")

    except PermissionError:
        print(
            "Permission denied. Make sure the Excel file is not open "
            "and you have access to the output folder."
        )

    except Exception as error:
        print(f"Unexpected error: {error}")

    finally:
        if connection is not None and connection.is_connected():
            connection.close()
            print("MySQL connection closed.")


if __name__ == "__main__":
    export_mysql_to_excel()