/**
 * Calculate geodesic (straight-line) distance between two coordinates
 * @param {Array} coord1 - [latitude, longitude]
 * @param {Array} coord2 - [latitude, longitude]
 * @returns {number} Distance in miles
 */
export const calculateGeodesicDistance = (coord1, coord2) => {
    const [lat1, lon1] = coord1;
    const [lat2, lon2] = coord2;
    const R = 3959; // Earth's radius in miles
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
             Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
             Math.sin(dLon/2) * Math.sin(dLon/2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
    return R * c;
};

/**
 * Calculate road-based distance using OpenRouteService API
 * Falls back to geodesic distance if API fails or no API key provided
 * @param {Array} coord1 - [latitude, longitude]
 * @param {Array} coord2 - [latitude, longitude]
 * @param {string} apiKey - OpenRouteService API key
 * @returns {Promise<number>} Distance in miles
 */
export const calculateRoadDistance = async (coord1, coord2, apiKey) => {
    if (!apiKey || apiKey.length < 10) {
        return calculateGeodesicDistance(coord1, coord2);
    }

    try {
        const [lat1, lon1] = coord1;
        const [lat2, lon2] = coord2;

        const response = await fetch('https://api.openrouteservice.org/v2/directions/driving-car', {
            method: 'POST',
            headers: {
                'Authorization': apiKey,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                coordinates: [[lon1, lat1], [lon2, lat2]],
                format: 'json',
                units: 'mi'
            })
        });

        if (response.ok) {
            const data = await response.json();
            if (data.routes && data.routes.length > 0) {
                return data.routes[0].summary.distance;
            }
        }
    } catch (error) {
        console.log('Road distance API failed, using geodesic:', error);
    }

    return calculateGeodesicDistance(coord1, coord2);
};
