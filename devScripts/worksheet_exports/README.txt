Export package contents

- manifest.json: list of all exported worksheets
- one CSV per worksheet, grouped by workbook
- import_exports_to_db.py: imports the CSVs into SQLite or another SQLAlchemy-supported database

Suggested SQLite command:
python /mnt/data/import_exports_to_db.py --manifest /mnt/data/worksheet_exports/manifest.json --db-url sqlite:////mnt/data/worksheet_imports.sqlite
