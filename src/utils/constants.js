export const REGIONS = {
    'Region 1 (East/South)': {
        warehouses: ['CHICAGO', 'BOSTON', 'PHILADELPH', 'NEWYORK', 'ORLANDO', 'FTLAUDER', 'WDC', 'NASHVILLE'],
        description: 'Eastern and Southern US warehouses'
    },
    'Region 2 (West/Central)': {
        warehouses: ['ANAHEIM', 'SANFRAN', 'PHOENIX', 'LASVEGAS', 'DALLAS'],
        description: 'Western and Central US warehouses'
    }
};

export const WAREHOUSE_COORDS = {
    'ANAHEIM': [33.8366, -117.9143],
    'BOSTON': [42.3601, -71.0589],
    'CHICAGO': [41.8781, -87.6298],
    'DALLAS': [32.7767, -96.7970],
    'FTLAUDER': [26.1224, -80.1373],
    'LASVEGAS': [36.1699, -115.1398],
    'NASHVILLE': [36.1627, -86.7816],
    'NEWYORK': [40.7128, -74.0060],
    'ORLANDO': [28.5383, -81.3792],
    'PHOENIX': [33.4484, -112.0740],
    'PHILADELPH': [39.9526, -75.1652],
    'SANFRAN': [37.7749, -122.4194],
    'WDC': [38.9072, -77.0369]
};

export const OPTIMIZATION_SCENARIOS = [
    {
        id: 'minimize_distance',
        name: 'Minimize Distance',
        description: 'Prioritize closest warehouses to reduce shipping distance',
        icon: '📏'
    },
    {
        id: 'minimize_trips',
        name: 'Minimize Trips',
        description: 'Pull from fewer warehouses to reduce trip count',
        icon: '🚛'
    },
    {
        id: 'balance_inventory',
        name: 'Balance Inventory',
        description: 'Distribute pulls evenly across warehouses',
        icon: '⚖️'
    },
    {
        id: 'prefer_local',
        name: 'Prefer Local',
        description: 'Maximize pulls from destination warehouse first',
        icon: '🏠'
    },
    {
        id: 'regional_priority',
        name: 'Regional Priority',
        description: 'Prioritize warehouses within the same region',
        icon: '🌍'
    },
    {
        id: 'preferred_source',
        name: 'Preferred Source',
        description: 'Prioritize specific warehouse after destination',
        icon: '⭐',
        requiresPreferred: true
    }
];
