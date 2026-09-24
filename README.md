# Moulding Store — Leaflet / OpenStreetMap / Alberta / No API

## Current build
- Leaflet map with OpenStreetMap tiles
- Photon/OpenStreetMap address autocomplete
- Alberta-only validation
- No paid geocoding API key
- Compact ultra-luxury light interface
- Alberta address suggestions show readable primary + secondary text
- 8-vehicle fleet, drag/drop compatible jobs, route lines, Google Maps sharing, CSV export

## IMPORTANT: one-time cache reset
The previous version used a PWA service worker that could cache the old black interface. After replacing your GitHub repository with this version, open:

`https://YOUR-GITHUB-PAGES-URL/reset.html`

Wait for it to redirect to the app. This unregisters the old service worker and clears only browser caches. **It does not delete your saved deliveries from localStorage.**

After that, the new light version will register its new service worker.

## GitHub Pages
Upload all files and folders in this ZIP to the repository root. Keep `assets/icon.svg` inside the `assets` folder.

## Address search
Start with a house number and street, for example:
`123 17 Ave SW`

Results are filtered to Alberta, Canada and selected addresses store latitude/longitude for mapping.
