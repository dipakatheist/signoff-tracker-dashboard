# Renewal Signoff Dashboard

This local dashboard reads the `Data` worksheet from the September 2026 signoff workbook. It serves aggregate counts only; consumer numbers, names, and meter identifiers are never sent to the browser.

## Start

Double-click `start_dashboard.bat`. It opens the dashboard at `http://127.0.0.1:8765` and keeps the local server running in the console. Press Ctrl+C in that console to stop it.

The default workbook path is:

`%USERPROFILE%\Desktop\Sept'26 PKG1-2 Fresh-Resignof Signoff Sheet.xlsx`

To use a different local copy, run:

```powershell
python dashboard_server.py --workbook "C:\path\to\workbook.xlsx"
```

The browser checks for updates every 60 seconds. The server reloads the workbook only after its modified time or file size changes. If OneDrive syncs the daily workbook to the configured local path, the dashboard picks up the saved update automatically. Keep the server running on the computer that has the workbook.