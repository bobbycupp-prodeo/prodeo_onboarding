/**
 * Checks if an address is inside a WKT Polygon string.
 * @param {string} address The address to check.
 * @param {string} wktString The raw POLYGON string from a CSV.
 * @return {boolean} TRUE if inside, FALSE if outside.
 * @customfunction
 */
function IS_WITHIN_POLYGON(address, wktString) {
  if (!address || !wktString) return false;

  // 1. Geocode the address to get Latitude and Longitude
  var geocoder = Maps.newGeocoder();
  var response = geocoder.geocode(address);
  
  if (response.status !== 'OK') {
    return "Geocode Error: " + response.status;
  }
  
  var targetLat = response.results[0].geometry.location.lat;
  var targetLng = response.results[0].geometry.location.lng;

  // 2. Parse the WKT string into an array of [lat, lng]
  // Remove "POLYGON" and all parentheses
  var cleanString = wktString.toString().toUpperCase().replace("POLYGON", "").replace(/[()]/g, "").trim();
  
  // WKT pairs are separated by commas
  var coordPairs = cleanString.split(',');
  var polygon = [];
  
  for (var i = 0; i < coordPairs.length; i++) {
    // Inside the pair, Longitude and Latitude are separated by a space
    var pair = coordPairs[i].trim().split(/\s+/);
    if (pair.length >= 2) {
      // WKT order is [Longitude, Latitude]. We push as [Lat, Lng] for the math below.
      polygon.push([parseFloat(pair[1]), parseFloat(pair[0])]); 
    }
  }

  // 3. Ray-Casting Algorithm to check if the point is inside the polygon
  var x = targetLat, y = targetLng;
  var inside = false;
  
  for (var i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
    var xi = polygon[i][0], yi = polygon[i][1];
    var xj = polygon[j][0], yj = polygon[j][1];

    var intersect = ((yi > y) != (yj > y))
        && (x < (xj - xi) * (y - yi) / (yj - yi) + xi);
    if (intersect) inside = !inside;
  }
  
  return inside;
}