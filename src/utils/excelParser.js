import { WAREHOUSE_COORDS } from './constants';

/**
 * Parse date string from Rentex format (e.g., " +1ThuJan 22 ", "Avail Today")
 * @param {string} dateStr - Date string from Rentex file
 * @returns {Date|null} Parsed date or null
 */
export const parseDateString = (dateStr) => {
    if (!dateStr || typeof dateStr !== 'string') return null;

    const trimmed = dateStr.trim();
    if (trimmed === 'Avail Today') {
        return new Date();
    }

    // Parse format like "+1ThuJan 22" or "+130SunMay 31"
    const match = trimmed.match(/\+(\d+)\w{3}(\w{3})\s+(\d+)/);
    if (!match) return null;

    const daysOffset = parseInt(match[1]);
    const monthStr = match[2];

    // Map month abbreviations
    const months = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5,
                   Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };
    const month = months[monthStr];

    if (month === undefined) return null;

    // Create date based on today + offset
    const today = new Date();
    const targetDate = new Date(today);
    targetDate.setDate(today.getDate() + daysOffset);

    return targetDate;
};

/**
 * Parse Excel inventory data (supports both Rentex and Standard formats)
 * @param {Array} excelData - Raw Excel data as 2D array
 * @param {Object} options - Parsing options
 * @param {Object} options.selectedDates - { start: string, end: string }
 * @param {Function} options.onDatesExtracted - Callback for extracted dates
 * @param {Function} options.onError - Error callback
 * @returns {Object|null} { inventory, warehouseColumns, extractedDates } or null
 */
export const parseExcelData = (excelData, options = {}) => {
    console.log('=== PARSE INVENTORY DEBUG START ===');

    if (!excelData || excelData.length < 4) {
        console.error('FAILED: No Excel data or data too short');
        console.log('excelData exists:', !!excelData);
        console.log('excelData length:', excelData ? excelData.length : 0);
        return null;
    }

    // Detect Rentex format (header at row 3, with Equipment and Location columns)
    const isRentexFormat = excelData[2] &&
        excelData[2].some(cell => cell && String(cell).toLowerCase().trim() === 'equipment') &&
        excelData[2].some(cell => cell && String(cell).toLowerCase().trim() === 'location');

    console.log('Format detected:', isRentexFormat ? 'Rentex' : 'Standard');

    if (isRentexFormat) {
        return parseRentexFormat(excelData, options);
    } else {
        return parseStandardFormat(excelData, options);
    }
};

/**
 * Parse Rentex format (Equipment/Location columns with date-based Qty columns)
 */
const parseRentexFormat = (excelData, options = {}) => {
    const { selectedDates, onError } = options;

    // Row 2 (index 1) contains dates
    // Row 3 (index 2) contains headers
    const dateRow = excelData[1] || [];
    const headers = excelData[2];
    console.log('Rentex date row:', dateRow.slice(0, 15));
    console.log('Rentex headers:', headers.slice(0, 15));

    // Find column indices
    const equipmentColIdx = headers.findIndex(h =>
        h && String(h).toLowerCase().trim() === 'equipment'
    );
    const locationColIdx = headers.findIndex(h =>
        h && String(h).toLowerCase().trim() === 'location'
    );

    // Find "Qty" columns with their corresponding dates
    const qtyColumnsWithDates = [];
    const extractedDates = [];

    headers.forEach((header, idx) => {
        if (header && String(header).toLowerCase().trim() === 'qty') {
            const dateInfo = dateRow[idx];
            const parsedDate = parseDateString(dateInfo);

            qtyColumnsWithDates.push({
                index: idx,
                dateStr: dateInfo ? String(dateInfo).trim() : '',
                date: parsedDate
            });

            if (parsedDate) {
                extractedDates.push({
                    dateStr: dateInfo ? String(dateInfo).trim() : '',
                    date: parsedDate
                });
            }
        }
    });

    console.log('Equipment column:', equipmentColIdx);
    console.log('Location column:', locationColIdx);
    console.log('Qty columns with dates:', qtyColumnsWithDates.length);
    console.log('Extracted dates:', extractedDates);

    if (equipmentColIdx === -1 || locationColIdx === -1) {
        console.error('FAILED: Required columns not found');
        if (onError) onError('❌ Could not find Equipment or Location columns');
        return null;
    }

    if (qtyColumnsWithDates.length === 0) {
        console.error('FAILED: No Qty columns found');
        if (onError) onError('❌ Could not find quantity columns');
        return null;
    }

    // Filter Qty columns by selected date range if specified
    let columnsToUse = qtyColumnsWithDates;
    if (selectedDates?.start && selectedDates?.end) {
        const startDate = new Date(selectedDates.start);
        const endDate = new Date(selectedDates.end);

        columnsToUse = qtyColumnsWithDates.filter(col => {
            if (!col.date) return false;
            return col.date >= startDate && col.date <= endDate;
        });

        console.log(`Date filter: ${selectedDates.start} to ${selectedDates.end}`);
        console.log(`Filtered to ${columnsToUse.length} columns from ${qtyColumnsWithDates.length}`);
    }

    if (columnsToUse.length === 0) {
        console.error('FAILED: No columns match date range');
        if (onError) onError('❌ No quantity columns found in selected date range');
        return null;
    }

    console.log('Using columns:', columnsToUse.map(c => c.index));

    // Group data by equipment and location
    const inventory = {};
    const warehouseColumns = {};
    let processedRows = 0;

    // Process data rows (starting from row 4, index 3)
    for (let i = 3; i < excelData.length; i++) {
        const row = excelData[i];
        if (!row || row.length === 0) continue;

        const equipment = row[equipmentColIdx];
        const location = row[locationColIdx];

        if (!equipment || !location) continue;

        // Get quantities from all selected date columns and find minimum
        const quantities = columnsToUse.map(col => {
            const qty = parseInt(row[col.index]) || 0;
            return qty;
        }).filter(q => !isNaN(q));

        // Use MINIMUM quantity across the date range (safest - guarantees availability)
        const qty = quantities.length > 0 ? Math.min(...quantities) : 0;

        const equipmentStr = String(equipment).trim();
        const locationStr = String(location).trim().toUpperCase();

        // Initialize equipment entry if needed
        if (!inventory[equipmentStr]) {
            inventory[equipmentStr] = {};
        }

        // Add/update quantity for this location
        if (!inventory[equipmentStr][locationStr]) {
            inventory[equipmentStr][locationStr] = 0;
        }
        inventory[equipmentStr][locationStr] += qty;

        // Track warehouse columns
        warehouseColumns[locationStr] = true;

        processedRows++;
    }

    console.log(`✓ Processed ${processedRows} rows`);
    console.log(`✓ Found ${Object.keys(inventory).length} unique equipment items`);
    console.log(`✓ Found ${Object.keys(warehouseColumns).length} warehouses:`, Object.keys(warehouseColumns));
    console.log('Sample equipment:', Object.keys(inventory).slice(0, 5));
    console.log('=== PARSE INVENTORY DEBUG END ===');

    return {
        inventory,
        warehouseColumns: Object.keys(warehouseColumns).reduce((acc, wh) => {
            acc[wh] = wh; // Just store warehouse name
            return acc;
        }, {}),
        extractedDates
    };
};

/**
 * Parse standard format (Equipment column + warehouse columns)
 */
const parseStandardFormat = (excelData, options = {}) => {
    const { onError } = options;

    // Standard format - first row is header
    const headers = excelData[0];
    console.log('Standard headers:', headers);

    const warehouseColumns = {};

    // Find warehouse columns
    headers.forEach((header, index) => {
        if (header && typeof header === 'string') {
            const warehouse = header.toUpperCase().trim();
            if (WAREHOUSE_COORDS[warehouse]) {
                warehouseColumns[warehouse] = index;
                console.log(`✓ Found warehouse: ${warehouse} at column ${index}`);
            }
        }
    });

    console.log('Warehouse columns found:', Object.keys(warehouseColumns).length);

    if (Object.keys(warehouseColumns).length === 0) {
        console.error('FAILED: No warehouse columns found');
        console.log('Looking for these warehouses:', Object.keys(WAREHOUSE_COORDS));
        console.log('Available headers:', headers.filter(h => h && typeof h === 'string'));
        if (onError) onError('❌ No warehouse columns found. Use Rentex format (Equipment + Location) or rename columns to: CHICAGO, BOSTON, etc.');
        return null;
    }

    // Find equipment column
    const equipmentColIndex = headers.findIndex(h =>
        h && h.toString().toLowerCase().includes('equipment')
    );

    console.log('Equipment column index:', equipmentColIndex);

    if (equipmentColIndex === -1) {
        console.error('FAILED: Equipment column not found in inventory');
        console.log('Available columns:', headers);
        if (onError) onError('❌ Equipment column not found in inventory file');
        return null;
    }

    console.log(`✓ Equipment column at index ${equipmentColIndex}`);

    // Parse inventory data
    const inventory = {};
    let equipmentCount = 0;

    for (let i = 1; i < excelData.length; i++) {
        const row = excelData[i];
        const equipmentName = row[equipmentColIndex];
        if (!equipmentName) continue;

        inventory[equipmentName] = {};
        Object.entries(warehouseColumns).forEach(([warehouse, colIndex]) => {
            const qty = parseInt(row[colIndex]) || 0;
            inventory[equipmentName][warehouse] = qty;
        });
        equipmentCount++;
    }

    console.log(`✓ Parsed ${equipmentCount} equipment items`);
    console.log('Sample equipment:', Object.keys(inventory).slice(0, 5));
    console.log('=== PARSE INVENTORY DEBUG END ===');

    return { inventory, warehouseColumns, extractedDates: [] };
};
