import { REGIONS } from './constants';

/**
 * Get the region name for a given warehouse
 * @param {string} warehouse - Warehouse name
 * @returns {string} Region name or 'Unknown'
 */
export const getWarehouseRegion = (warehouse) => {
    for (const [regionName, regionData] of Object.entries(REGIONS)) {
        if (regionData.warehouses.includes(warehouse)) {
            return regionName;
        }
    }
    return 'Unknown';
};
