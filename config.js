export const CONFIG={
  mapTiles:"https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",
  mapAttribution:"© OpenStreetMap contributors",
  geocoder:"https://nominatim.openstreetmap.org/search",
  geocoderCountry:"ca",
  geocoderViewbox:"-120.0,60.0,-110.0,48.9",
  // "nominatim" works without a key but has less complete address coverage.
  // Set to "canadapost" and add your restricted AddressComplete key for
  // Canada-Post-style premise/address autocomplete.
  addressProvider:"nominatim",
  canadaPostKey:"",
  apiBase:""
};
