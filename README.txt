MOULDING STORE — LIVE DISPATCH V1

FILES
- 01_live_dispatch_schema.ts : run once in Excel for the web > Automate > New Script.
- index.html
- config.js
- excel-api.js
- app.js
- styles.css

KEEP
- Keep your existing auth.js. The new app expects the same functions:
  initAuth(), getAccessToken(), getAccount(), signIn(), signOut().

WORKBOOK
- MouldingStoreDatabase.xlsx in OneDrive.
- The script does not delete existing columns.
- It adds/ensures the dispatch fields needed by the live scheduler.

SETUP
1. Back up MouldingStoreDatabase.xlsx.
2. Excel for web > Automate > New Script.
3. Paste 01_live_dispatch_schema.ts and Run.
4. Replace the listed web files in your GitHub Pages repository.
5. Keep auth.js.
6. Open the GitHub Pages site.
7. Sign in.
8. If the workbook is not connected, open Settings > Find workbook in OneDrive.

IMPORTANT
- This version uses Excel/Graph as the live data store.
- The browser refreshes every 15 seconds by default.
- Dispatcher changes are written back to Excel.
- Drag/drop assignment is supported; an Assign button is also available.
- The current version uses the driver's DefaultTruckID. Manual truck assignment and advanced route optimization can be added next.
- The map/GPS layer is intentionally not included in this first working version; it should be added after the scheduling workflow is verified.
