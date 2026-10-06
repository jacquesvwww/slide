/**
 * Thin wrapper around LocationIQ's Forward Geocoding and Matrix APIs —
 * powers DEV-015 (estimated travel time from a barista to a café).
 * Both endpoints are Nominatim/OSRM-compatible; see
 * https://docs.locationiq.com for the full reference.
 */

const BASE_URL = 'https://us1.locationiq.com/v1';

function getApiKey_() {
  const key = process.env.LOCATIONIQ_API_KEY;
  if (!key) throw new Error('LOCATIONIQ_API_KEY environment variable is not set.');
  return key;
}

// address -> { lat, lng } | null (null on no match — never throws on a miss)
async function geocodeAddress(address) {
  const key = getApiKey_();
  const url = `${BASE_URL}/search?key=${key}&q=${encodeURIComponent(address)}&format=json&limit=1`;
  const res = await fetch(url);
  if (res.status === 404) return null; // LocationIQ's "no match" response
  if (!res.ok) throw new Error('Geocoding request failed (' + res.status + ')');
  const data = await res.json();
  if (!Array.isArray(data) || !data.length) return null;
  return { lat: Number(data[0].lat), lng: Number(data[0].lon), displayName: data[0].display_name || '' };
}

// origin {lat,lng}, destinations [{lat,lng}, ...] -> same-length array of
// { minutes, km } | null (null means no route was found to that point).
// Matrix API only documents a 'driving' profile — there's no walking/
// cycling option to offer here.
async function getTravelTimes(origin, destinations) {
  if (!destinations.length) return [];
  const key = getApiKey_();
  const coords = [origin, ...destinations].map(p => `${p.lng},${p.lat}`).join(';');
  const destIndices = destinations.map((_, i) => i + 1).join(';');
  const url = `${BASE_URL}/matrix/driving/${coords}?sources=0&destinations=${destIndices}&annotations=distance,duration&key=${key}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error('Matrix request failed (' + res.status + ')');
  const data = await res.json();
  if (data.code !== 'Ok' || !data.durations || !data.durations[0]) {
    throw new Error('Matrix response missing durations.');
  }
  const durations = data.durations[0];
  const distances = (data.distances && data.distances[0]) || [];
  return destinations.map((_, i) => {
    const seconds = durations[i];
    const meters = distances[i];
    if (seconds == null) return null;
    return { minutes: Math.round(seconds / 60), km: meters != null ? Math.round(meters / 100) / 10 : null };
  });
}

module.exports = { geocodeAddress, getTravelTimes };
