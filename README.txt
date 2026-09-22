MOULDING STORE — PREMIUM DELIVERY PORTAL V1

Purpose:
A read-only delivery information portal. The website only displays data that
already exists in MouldingStoreDatabase.xlsx.

What changed:
- Removed web-based order creation.
- Removed web-based editing.
- Removed assignment controls.
- Removed status-change controls.
- Removed Sales and PO screens.
- Removed driver mobile workflow.
- Today's Deliveries is the default screen.
- Delivery History provides search by order/customer/address/driver and date range.
- Delivery details show information and recorded status history.
- Auto refresh is enabled.
- Excel is the source of truth.

Security:
- config.js requests User.Read + Files.Read, not Files.ReadWrite.
- For real access control, regular users should have Read permission to the
  workbook in OneDrive/SharePoint and admins should have Edit permission.
- The portal itself does not contain write operations.

Files:
- index.html
- config.js
- excel-api.js
- app.js
- styles.css

Keep:
- Keep the existing auth.js from your working Microsoft Entra setup.

Excel:
- No new schema script is required for this version if your existing workbook
  already contains OrdersTable, DeliveryStatusTable, DriversTable, TrucksTable,
  DailyScheduleTable and RouteStopsTable.

Setup:
1. Back up the workbook.
2. Replace the five web files in GitHub Pages.
3. Keep auth.js.
4. Sign in.
5. If necessary, open the settings button and Find Workbook.
