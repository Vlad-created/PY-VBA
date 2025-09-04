import pyodbc
from datetime import datetime, timedelta

# Path to Access database
db_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\New folder\Fil_November_SalesKPI.accdb"

# 🔌 Connection string for Access DB
conn_str = (
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={db_path};"
)

# Calculate the 1st day of the previous month
today = datetime.today()
first_day_prev_month = datetime(today.year, today.month, 1) - timedelta(days=1)
cutoff_date = datetime(first_day_prev_month.year, first_day_prev_month.month, 1)

# List of tables to process
table_list = [
    "M_Voice_above_PT", "M_Voice_aboveBundlePTplus", "M_Voice_Reset",
    "PT_Fix_Above", "PT_Fix_Reset", "PT_Fix_Upsell", "PT_Fix_Xsell",
    "PT_Fix_Xsell_Router", "SOHO_PT", "Stop_reset", "TV_Bundles_TA1"
]

# Function to check if a column exists in a table
def column_exists(cursor, table, column):
    try:
        cursor.execute(f"SELECT TOP 1 * FROM [{table}]")
        return column in [desc[0] for desc in cursor.description]
    except Exception:
        return False

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    for tbl in table_list:
        print(f"🧹 Checking table: {tbl}")

        if column_exists(cursor, tbl, "Available To"):
            print(f"✅ Column verified. Cleaning table: {tbl}")

            sql = f"""
            DELETE FROM [{tbl}]
            WHERE [Available To] < #{cutoff_date.strftime('%m/%d/%Y')}#
            """
            cursor.execute(sql)
            conn.commit()

            print(f"✅ Deleted records from {tbl} before {cutoff_date.strftime('%d-%b-%Y')}")
        else:
            print(f"⚠️ Skipping {tbl}: missing 'Available To' column")

except Exception as e:
    print(f"❌ Error: {e}")

finally:
    if 'cursor' in locals(): cursor.close()
    if 'conn' in locals(): conn.close()
