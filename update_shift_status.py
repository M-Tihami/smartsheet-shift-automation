import os
import re
import smartsheet
from datetime import datetime, date, timezone, timedelta

SAT = timezone(timedelta(hours=3))  # Saudi Arabia UTC+3
now = datetime.now(SAT)

client = smartsheet.Smartsheet(os.environ['SMARTSHEET_API_TOKEN'])
sheet_id = int(os.environ['SMARTSHEET_SHEET_ID'])

sheet = client.Sheets.get_sheet(sheet_id)
col_ids = {col.title: col.id for col in sheet.columns}
col_titles = [col.title for col in sheet.columns]

STATUS_COL = 'Ambulance Readiness'


def parse_row_date(value):
    """Return the date in SAT. Values with a timezone are converted;
    values without one are assumed to already be SAT."""
    s = str(value).strip()
    if len(s) == 10:  # 'YYYY-MM-DD'
        return date.fromisoformat(s)
    dt = datetime.fromisoformat(s.replace('Z', '+00:00'))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=SAT)
    return dt.astimezone(SAT).date()


def parse_shift(shift):
    """Parse '07:00 - 19:00'. Accepts -, – or — as the separator."""
    parts = re.split(r'\s*[-–—]\s*', str(shift).strip())
    start_str, end_str = parts[0].strip(), parts[1].strip()
    sh, sm = map(int, start_str.split(':'))
    eh, em = map(int, end_str.split(':'))
    return (sh, sm), (eh, em)


rows_to_update = []

for row in sheet.rows:
    vals = {col_titles[i]: c.value for i, c in enumerate(row.cells)}

    created = vals.get('Created Date')
    shift = vals.get('Shift')
    if not created or not shift:
        continue

    try:
        row_date = parse_row_date(created)
        (sh, sm), (eh, em) = parse_shift(shift)
    except Exception as e:
        print(f"Row {row.id} skipped (parse error): {e} | created={created!r} | shift={shift!r}")
        continue

    start_dt = datetime(row_date.year, row_date.month, row_date.day, sh, sm, tzinfo=SAT)
    end_dt = datetime(row_date.year, row_date.month, row_date.day, eh, em, tzinfo=SAT)
    if end_dt <= start_dt:  # overnight shift ends the next day
        end_dt += timedelta(days=1)

    status = 'On Duty' if start_dt <= now < end_dt else 'Off Duty'

    # Only write if the value actually changed
    if vals.get(STATUS_COL) == status:
        continue

    cell = smartsheet.models.Cell()
    cell.column_id = col_ids[STATUS_COL]
    cell.value = status

    new_row = smartsheet.models.Row()
    new_row.id = row.id
    new_row.cells.append(cell)
    rows_to_update.append(new_row)

    print(f"Row {row.id} | created={created!r} | shift={shift!r} | "
          f"window {start_dt:%Y-%m-%d %H:%M} -> {end_dt:%Y-%m-%d %H:%M} | {status}")

# Batched writes instead of one API call per row
BATCH = 100
for i in range(0, len(rows_to_update), BATCH):
    client.Sheets.update_rows(sheet_id, rows_to_update[i:i + BATCH])

print(f"Updated {len(rows_to_update)} rows. Done at {now:%Y-%m-%d %H:%M:%S} SAT")
