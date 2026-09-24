# Moulding Store Route Control — stable v11

This build intentionally uses a single self-contained `index.html`.

- No ES modules.
- No app.js/config.js dependency chain.
- No service worker is registered.
- Leaflet loads separately; the dashboard remains visible if Leaflet CDN is unavailable.
- OpenStreetMap tiles via Leaflet.
- No-key Photon/OpenStreetMap Alberta address autocomplete.
- Light luxury UI with high-contrast text.
- Drag/drop compatible jobs, optimize demo, CSV export, Google Maps route share.

## GitHub Pages
Upload `index.html` and `reset.html` to the root of the repository. Open `reset.html` once, then `index.html?v=v11`.
