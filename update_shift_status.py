import smartsheet
import os
from datetime import datetime, timezone, timedelta

# Saudi Arabia is UTC+3
SAT = timezone(timedelta(hours=3))
now = datetime.now(SAT)
current_time = now.hour * 60 + now.minute
today = now.date()
yesterday = today - timedelta(days=1)

client = smartsheet.Smartsheet(os.environ['SMARTSHEET_API_TOKEN'])
sheet_id = int(os.environ['SMARTSHEET_SHEET_ID'])

sheet = client.Sheets.get_sheet(sheet_id)
columns = {col.title: col.id for col in sheet.columns}

for row in sheet.rows:
    cell_values = {sheet.columns[i].title: cell.value for i, cell in enumerate(row.cells)}

    created_date = cell_values.get('Created Date', '')
    if not created_date:
        continue

    try:
        row_date = datetime.fromisoformat(str(created_date)).astimezone(SAT).date()
    except:
        continue

    if row_date not in [today, yesterday]:
        continue

    shift = cell_values.get('Shift', '')
    if not shift or '-' not in shift:
        continue

    parts = shift.split('-')
    start_str = parts[0].strip()
    end_str = parts[1].strip()

    start_h, start_m = map(int, start_str.split(':'))
    end_h, end_m = map(int, end_str.split(':'))

    start_minutes = start_h * 60 + start_m
    end_minutes = end_h * 60 + end_m

    if end_minutes <= start_minutes:
        if current_time >= start_minutes or current_time < end_minutes:
            status = 'On Duty'
        else:
            status = 'Off Duty'
    else:
        if start_minutes <= current_time < end_minutes:
            status = 'On Duty'
        else:
            status = 'Off Duty'

    if row_date == yesterday and status == 'On Duty':
        status = 'Off Duty'

    new_cell = smartsheet.models.Cell()
    new_cell.column_id = columns['Ambulance Readiness']
    new_cell.value = status

    new_row = smartsheet.models.Row()
    new_row.id = row.id
    new_row.cells.append(new_cell)

    client.Sheets.update_rows(sheet_id, [new_row])
    print(f"Row {row.id} | Shift: {shift} | Status: {status}")

print(f"Done at {now.strftime('%Y-%m-%d %H:%M:%S')} SAT")
