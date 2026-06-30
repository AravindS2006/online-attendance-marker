// Sri Sairam Engineering College Coordinates
export const SAIRAM_CAMPUS_COORDS = {
  latitude: 12.9602,
  longitude: 80.0570,
  name: "Sri Sairam Engineering College Campus"
};

export interface GeoLocation {
  lat: number;
  lng: number;
}

/**
 * Calculates the distance between two coordinates in meters using the Haversine formula.
 */
export function getDistanceMeters(lat1: number, lon1: number, lat2: number, lon2: number): number {
  const R = 6371e3; // Earth radius in meters
  const phi1 = (lat1 * Math.PI) / 180;
  const phi2 = (lat2 * Math.PI) / 180;
  const deltaPhi = ((lat2 - lat1) * Math.PI) / 180;
  const deltaLambda = ((lon2 - lon1) * Math.PI) / 180;

  const a =
    Math.sin(deltaPhi / 2) * Math.sin(deltaPhi / 2) +
    Math.cos(phi1) * Math.cos(phi2) *
    Math.sin(deltaLambda / 2) * Math.sin(deltaLambda / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return R * c; // Distance in meters
}

/**
 * Checks if the student is within the campus geofence radius.
 */
export function verifyGeofence(
  studentLat: number,
  studentLng: number,
  targetLat: number = SAIRAM_CAMPUS_COORDS.latitude,
  targetLng: number = SAIRAM_CAMPUS_COORDS.longitude,
  radiusMeters: number = 500
): { inGeofence: boolean; distance: number } {
  const distance = getDistanceMeters(studentLat, studentLng, targetLat, targetLng);
  return {
    inGeofence: distance <= radiusMeters,
    distance
  };
}
