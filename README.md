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


### Address autocomplete — street level only

The address field now requires a numeric house/building number and filters out city-only, town-only, province-only, postal-code-only and other broad place results. The fallback search is restricted to Alberta.

For the closest possible Canada Post AddressComplete experience, create an AddressComplete API key and set `addressProvider:"canadapost"` and `canadaPostKey:"..."` in `config.js`. Canada Post's Find service returns autocomplete candidates and Retrieve returns fields including BuildingNumber, Street, City, Province and PostalCode. 

## No-API-key address autocomplete

This version uses the public Photon geocoder backed by OpenStreetMap for address search and geocoding. Photon supports search-as-you-type and is designed for address/name search. No API key is configured or required for the public endpoint. See the Photon/OpenStreetMap documentation before putting heavy production traffic on the public instance; for larger usage, self-hosting Photon is the appropriate long-term option.

The delivery address field:
- waits until a query contains a house/building number;
- searches with an Alberta bounding box;
- keeps only Canadian/Alberta results with a house number and street/name;
- stores latitude/longitude from the selected result;
- rejects unselected free-typed addresses when saving a delivery.

The map geocoder also uses Photon, so this build no longer calls Nominatim for address autocomplete or route geocoding.
