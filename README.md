# do-case-reports
Regular process to generate case reports for CFT.

## Setup
Tested on Python 3.10.10, may work for 3.8.x and 3.9.x. Also requires SAP Logon GUI for Windows and access to Production ISU.

Fields for **required** .env file:
```env
DATABRICKS_TOKEN=""
DATABRICKS_HOST=""
DATABRICKS_HTTP_PATH=""
ORACLE_DBDSN=""
ORACLE_USER=""
ORACLE_PASS=""
```

To create, activate and configure the Python environment:
```bat
python -m venv venv
.\venv\Scripts\activate.bat
pip install -r dependencies.txt
```

To generate MPSC frequent outage and non-frequent outage case reports:
```bat
python create_reports.py
```

To generate non-MPSC case summary sheet:
```bat
python create_non_mpsc_sheet.py
```
---
### On missing case premises
Insert `<case-id> <premise>` key-value pairs into `special-cases.txt` in case current sources have missing premises.
