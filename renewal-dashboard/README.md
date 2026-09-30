# Renewal Signoff Dashboard

This local dashboard reads the `Data` worksheets from the September 2026 PKG1-2 and PKG-7 signoff workbooks. It serves aggregate counts only; consumer numbers, names, and meter identifiers are never sent to the browser. Use the package selector to see either workbook or both combined.

The current Vercel deployment contains the dashboard UI only. It does not read workbook data until a secure SharePoint API is configured; neither workbook nor its aggregates are included in the deployment.

## Start

Double-click `start_dashboard.bat`. It opens the dashboard at `http://127.0.0.1:8765` and keeps the local server running in the console. Press Ctrl+C in that console to stop it.

The default workbook paths are:

- `%USERPROFILE%\Desktop\Sept'26 PKG1-2 Fresh-Resignof Signoff Sheet.xlsx`
- `%USERPROFILE%\Desktop\PKG-7 (September) Fresh-Resignoff Pending Signof Sheet as on 08-09-26.xlsx`

To use different local copies, run:

```powershell
python dashboard_server.py --pkg1-2 "C:\path\to\pkg1-2.xlsx" --pkg-7 "C:\path\to\pkg-7.xlsx"
```

The browser checks for updates every 60 seconds. The server checks both workbook timestamps and sizes, and only reparses them when either file changes. If OneDrive syncs the daily workbooks to those local paths, the dashboard picks up saved updates automatically. Keep the server running on the computer that has both workbooks.