import pyodbc
from datetime import datetime


ACDB_FILES = [
    r"\\virgo.mobiltel.bg\Work\OpControl\2016\Digital Signature\Digital_Signature.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\smart_app_ott_watched.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\Digital Signature New.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\Port in-out.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\GAs_Fixed.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\Involuntary Churn New.accdb", 
    r"\\virgo\Work\OpControl\2016\SALES\GAs.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\Prima KPI.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\Retention_NBS.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\Equipment_All.accdb",
    r"\\virgo\Work\OpControl\Daily report_Core services_Focus products\Morning Daily New Format\Morning_Daily.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\Served_customers.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\Voluntary_churn_new.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\Churn_rate_by_shop.accdb",
    r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Monthly Result.accdb",
    r"\\virgo\Work\OpControl\2016\SALES\SC_Potential.accdb", 
    r"\\virgo\Work\OpControl\2016\SALES\Potential_Retentions.accdb",
    # r"\\virgo\Work\OpControl\2016\Rado\SWE_New_Sorse\SWE_BIO_report.accdb",
    r"\\virgo\work\OpControl\SIP\SIP3.0_PowerBI.accdb"
    
]


today = datetime.today()
current_year = today.year
current_month = today.month

print(f"Updating SwitchboardItems → sw_year={current_year}, sw_month={current_month}")

for acdb_path in ACDB_FILES:
    print(f"\nProcessing: {acdb_path}")

    try:
        # ODBC connection string
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={acdb_path};"
        )

        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        # Update SwitchboardItems table
        cursor.execute("""
            UPDATE SwitchboardItems
            SET sw_year = ?, sw_month = ?
            WHERE SwitchboardID = 1 AND ItemNumber = 0
        """, (current_year, current_month))

        conn.commit()
        conn.close()

        print("✔ Successfully updated SwitchboardItems")

    except Exception as e:
        print(f"❌ Error updating database: {e}")
