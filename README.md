# Moulding Store Route Planner

Upload this entire folder to your GitHub repository.

## Included
- React-style single-page architecture without a build step (plain JS module)
- Premium responsive mobile UI
- 4 x 5-ton, 2 x 3-ton, 2 x Sprinter
- Add delivery and choose vehicle requirement
- Drag/drop jobs between compatible trucks
- Colored route map and numbered stops
- Google Maps share link
- CSV export
- PWA install/offline shell
- Excel Office Script to create the Lists/sheet structure

## Important
The browser prototype uses localStorage so it runs immediately. Your existing sheet/list should be connected by replacing the load/save functions with your existing API/connector.

Do not put private API keys in GitHub.

## GitHub Pages
Repository -> Settings -> Pages -> Deploy from branch -> main -> /(root).

## Production next step
Connect the exact spreadsheet/list you already have. Then replace localStorage with that API and add a secure backend for traffic-aware routing/optimization.


### Address autocomplete
The address field now requires a house/building number and street and returns Alberta-only premise-level results. The fallback uses Nominatim's `address` layer and rejects city/street-only results. For Canada Post AddressComplete-quality coverage, set `addressProvider` to `canadapost` and add an AddressComplete API key restricted to your GitHub Pages domain. Canada Post's Find/Retrieve API is designed for address autocomplete and returns fields such as BuildingNumber, Street, City, Province and PostalCode.
