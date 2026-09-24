# Moulding Store Route Planner — Alberta No-API Address Search

This GitHub-ready build uses **Leaflet + OpenStreetMap tiles** for the map and **Photon (OpenStreetMap data)** for search-as-you-type address autocomplete. No API key is required.

## Address search
- Canada-only request filtering.
- Alberta bounding box: `-120.0,48.9,-110.0,60.0`.
- Client-side validation requires the returned country to be Canada and province/state to be Alberta.
- House-level results are preferred.
- Photon structured search is used as a fallback for `house number + street + city` input.
- City-only and street-only results are not accepted as a delivery address.
- Selecting an address stores its latitude/longitude for the map and routing.

## Important
Photon is a public OSM-based demo service. Its project says the public demo server can be used for projects when requests remain reasonable. For a production fleet with substantial usage, self-hosting Photon or using a dedicated OSM-derived geocoder is recommended.

OpenStreetMap tiles are used with visible attribution. The OSM tile service is best-effort and has usage requirements; do not bulk-download tiles.

## GitHub Pages
Upload the contents of this folder to your repository and enable GitHub Pages. No build step is required.
