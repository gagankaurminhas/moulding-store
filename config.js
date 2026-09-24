export const CONFIG={
  mapTiles:"https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",
  mapAttribution:"© OpenStreetMap contributors",
  geocoder:"https://nominatim.openstreetmap.org/search",
  geocoderCountry:"ca",
  // Alberta bounding box: west, north, east, south
  geocoderViewbox:"-120.0,60.0,-110.0,48.9",

  // For true Canada Post-style premise autocomplete, set:
  // addressProvider:"canadapost",
  // canadaPostKey:"YOUR_RESTRICTED_ADDRESSCOMPLETE_KEY"
  //
  // With no key, the app uses an Alberta-only street/premise fallback
  // and will NEVER display city-only results.
  addressProvider:"nominatim",
  canadaPostKey:"",
  apiBase:""
};
