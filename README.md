# LFLB attendance export

Exports attendance rows from `https://lflbssapp.com/admin/api/workAttendanceRest/staff-list`
for multiple employees across the last 7 days into an Excel workbook.

## Setup

```powershell
python -m pip install -r requirements.txt
```

## Employees

Edit `employees.txt`, one employee per line. You can write just the name, or `name,code`, or `name<TAB>code` if you know the staff code:

```text
Alice,ica2041
Bob
Charlie
```

## Run

Use environment variables so the password is not saved in the command history:

```powershell
$env:LFLB_USERNAME = "test"
$env:LFLB_PASSWORD = "test1230"
$env:LFLB_TENANT_CODE = "bssapp_esrdc"
python .\export_attendance.py
```

By default the script exports the 7 days ending today. To choose a specific end date:

```powershell
python .\export_attendance.py --end-date 2026-06-24
```

If login or the data call fails, rerun with `--debug` to print the server response head:

```powershell
python .\export_attendance.py --debug
```

The output is saved under `outputs/`.
