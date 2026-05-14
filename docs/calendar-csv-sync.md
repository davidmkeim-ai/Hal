# Calendar CSV Sync

HAL can keep `My Day` updated without direct Outlook access by syncing your existing `calendar_export.csv` file into the same Supabase-backed HAL state the hosted app already uses.

## What this setup does

1. Your existing process keeps writing `calendar_export.csv` into the HAL folder or a configured export path.
2. A small local script reads that CSV file.
3. The script uploads the raw CSV into HAL's Supabase-backed state.
4. The hosted HAL app reads that cloud copy when you click `Refresh calendar`.

## Files involved

- [scripts/sync-calendar-csv.js](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\scripts\sync-calendar-csv.js)
- [sync-calendar-csv.cmd](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\sync-calendar-csv.cmd)
- [package.json](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\package.json)

## Run it manually once

From the HAL folder, either:

```powershell
npm run sync:calendar-csv
```

or double-click:

[sync-calendar-csv.cmd](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\sync-calendar-csv.cmd)

If it works, the script will print a message showing how many CSV rows were uploaded.

## What it expects

- `HAL_STORAGE_PROVIDER=supabase` in [`.env`](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\.env)
- `SUPABASE_URL` and `SUPABASE_SERVICE_ROLE_KEY` configured in [`.env`](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\.env)
- `HAL_CSV_CALENDAR_PATH` pointing at your export file or export folder

If `HAL_CSV_CALENDAR_PATH` is not set, HAL defaults to:

`calendar_export.csv`

inside the app folder.

## Automate it with Windows Task Scheduler

Recommended setup:

1. Open `Task Scheduler`
2. Choose `Create Basic Task...`
3. Name it something like:
   `HAL Calendar CSV Sync`
4. Trigger:
   `Daily`
5. Repeat task every:
   `30 minutes`
6. Action:
   `Start a program`
7. Program/script:

```text
C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\sync-calendar-csv.cmd
```

8. Finish and save

That gives HAL a fresh cloud copy of your CSV roughly every 30 minutes without you manually uploading anything.

## How the hosted app uses it

The hosted HAL app now re-checks the cloud-backed CSV settings before it refreshes the calendar. That means the online version can pick up a newer synced CSV even if HAL has already been open in the browser.
