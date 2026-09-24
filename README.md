# Moulding Store Fleet Command v13

A rebuild of the delivery operation as a daily dispatch + driver + proof-of-delivery system. This is a GitHub Pages frontend with local IndexedDB storage and optional Power Automate/Excel synchronization.

## Included

- Daily delivery board with KPI cards
- Alberta-only address autocomplete using Photon/OpenStreetMap data
- Leaflet + OpenStreetMap fleet map
- 4 x 5-ton, 2 x 3-ton, 2 x Sprinter fleet seeded
- Driver pool with default truck assignments and one-driver-per-truck guard
- Visible UNASSIGNED queue
- Drag deliveries onto compatible trucks
- Drag stops between routes and reorder by dropping on a stop
- Capacity guard using estimated tons
- Local route optimizer that balances distance/workload while respecting vehicle type/capacity
- OSRM road geometry and route distance/time when available
- Route share to Google Maps
- Driver mode with next-stop workflow, navigation, phone call, arrival, delivery completion and POD
- Photo POD capture stored in IndexedDB
- Signature capture
- Driver browser GPS tracking with optional webhook
- Status workflow: Scheduled, En Route, Arrived, Unloading, Delivered, Exception/Canceled
- Delivery history / event log
- CSV export/import
- Excel/Power Automate sync payload
- Automatic migration of the previous v11/v12 localStorage format (`ms_jobs`, `ms_drivers`, `ms_truck_driver_*`, `ms_excel_webhook`)
- No service worker in this build

## First setup

1. Upload the contents of this package to the root of your GitHub Pages repository.
2. Open `reset.html` once if you previously used one of the older prototypes.
3. Open `index.html`.
4. Go to Settings -> Warehouse and enter the exact warehouse coordinates. This is important for proper route start/end and optimization.
5. Go to Settings -> Drivers and replace Driver 01…08 with your real drivers.

## Excel / Microsoft 365 sync

GitHub Pages should not contain Microsoft credentials. Use this pattern:

GitHub Pages -> Power Automate -> Excel Online (Business) -> Office Script

1. Open the included `MouldingStore_Delivery_Control_Template.xlsx`.
2. Open Excel's Automate tab and create a new script, then paste `MouldingStore_Excel_Sync_OfficeScript.ts`.
3. Save it as `Moulding Store Fleet Sync`.
4. Create a Power Automate flow with `When an HTTP request is received`.
5. Use the schema in `power_automate_request_schema.json`.
6. Add Excel Online (Business) -> Run script.
7. Select the workbook and `Moulding Store Fleet Sync` script.
8. Pass the HTTP trigger's `payloadJson` field to the Office Script's `payloadJson` parameter.
9. Copy the generated HTTP URL into Settings -> Excel / workflow sync -> Power Automate webhook URL.
10. Press Sync now.

The Office Script upserts deliveries by DeliveryID and retains the DeliveryEvents log. That means a later driver reassignment does not erase the delivery record and the current driver is written alongside the address.

## Important production note

This GitHub Pages build is designed to be stable and useful immediately, but truly shared multi-device fleet management requires a server-side data layer. The optional webhook is the bridge. A future production deployment can replace the webhook adapter with SharePoint Lists, Supabase, Firebase, Azure Functions, or a dedicated API without changing the dispatch UI/data model.

Live traffic and commercial telematics are also provider-dependent. The no-key build uses OpenStreetMap/Leaflet for the map, Photon for address search and OSRM for road routing when available. For live traffic, live GPS from hardware, driver safety telemetry, ELD/HOS, and vehicle diagnostics, connect a telematics provider or a commercial routing API.
