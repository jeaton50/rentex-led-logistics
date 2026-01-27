import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';

// Components
import Header from './components/Header';
import Toast from './components/common/Toast';
import UploadTab from './components/tabs/UploadTab';
import OptimizeTab from './components/tabs/OptimizeTab';
import ResultsTab from './components/tabs/ResultsTab';
import SettingsTab from './components/tabs/SettingsTab';

// Hooks
import { useToast } from './hooks/useToast';

// Utils
import { WAREHOUSE_COORDS, OPTIMIZATION_SCENARIOS, REGIONS } from './utils/constants';
import { parseExcelData } from './utils/excelParser';
import { calculateGeodesicDistance, calculateRoadDistance } from './utils/distanceCalculator';
import { getWarehouseRegion } from './utils/regionUtils';

const App = () => {
    // Tab management
    const [activeTab, setActiveTab] = useState('upload');

    // File handling
    const [excelFile, setExcelFile] = useState(null);
    const [excelData, setExcelData] = useState(null);
    const fileInputRef = useRef(null);

    // Form inputs
    const [destination, setDestination] = useState('');
    const [equipment, setEquipment] = useState('');
    const [quantity, setQuantity] = useState('');
    const [selectedRegion, setSelectedRegion] = useState('');
    const [preferredWarehouse, setPreferredWarehouse] = useState('');
    const [selectedDates, setSelectedDates] = useState({ start: '', end: '' });
    const [availableDates, setAvailableDates] = useState([]);

    // Optimization
    const [selectedScenarios, setSelectedScenarios] = useState(['minimize_distance']);
    const [results, setResults] = useState(null);
    const [isOptimizing, setIsOptimizing] = useState(false);

    // Settings
    const [apiKey, setApiKey] = useState(localStorage.getItem('ors_api_key') || '');
    const [useRoadDistances, setUseRoadDistances] = useState(false);
    const [equipmentWeights, setEquipmentWeights] = useState({});

    // Quote handling
    const [equipmentFromQuote, setEquipmentFromQuote] = useState({});
    const [equipmentDescriptions, setEquipmentDescriptions] = useState({});
    const quoteForOptimizerRef = useRef(null);

    // Toast
    const { toast, showToast } = useToast();

    // Load settings from localStorage on mount
    useEffect(() => {
        const savedWeights = localStorage.getItem('equipment_weights');
        if (savedWeights) {
            setEquipmentWeights(JSON.parse(savedWeights));
        }

        const savedUseRoadDistances = localStorage.getItem('use_road_distances');
        if (savedUseRoadDistances) {
            setUseRoadDistances(savedUseRoadDistances === 'true');
        }
    }, []);

    // Save API key to localStorage when it changes
    useEffect(() => {
        if (apiKey) {
            localStorage.setItem('ors_api_key', apiKey);
        }
    }, [apiKey]);

    // Save useRoadDistances to localStorage when it changes
    useEffect(() => {
        localStorage.setItem('use_road_distances', useRoadDistances.toString());
    }, [useRoadDistances]);

    // Handle file upload
    const handleFileUpload = (file) => {
        if (!file) return;

        console.log('=== FILE UPLOAD START ===');
        console.log('File name:', file.name);
        console.log('File type:', file.type);
        console.log('File size:', file.size, 'bytes');

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                console.log('File read complete, parsing...');
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                console.log('Workbook sheets:', workbook.SheetNames);

                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: null });

                console.log('Parsed rows:', jsonData.length);
                console.log('Row 1:', jsonData[0]?.slice(0, 10));
                console.log('Row 2:', jsonData[1]?.slice(0, 10));
                console.log('Row 3:', jsonData[2]?.slice(0, 10));
                console.log('Row 4:', jsonData[3]?.slice(0, 10));

                // Check if this is Rentex format (header at row 3)
                const isRentexFormat = jsonData[2] &&
                    jsonData[2].some(cell => cell && String(cell).toLowerCase().includes('equipment')) &&
                    jsonData[2].some(cell => cell && String(cell).toLowerCase().includes('location'));

                console.log('Rentex format detected:', isRentexFormat);

                setExcelData(jsonData);
                setExcelFile(file);

                console.log('✓ Excel data set in state');
                console.log('=== FILE UPLOAD END ===');

                if (isRentexFormat) {
                    showToast(`✅ Rentex inventory file uploaded: ${file.name}`, 'success');
                } else {
                    showToast(`✅ File uploaded: ${file.name}`, 'success');
                }
                setActiveTab('optimize');
            } catch (error) {
                console.error('=== FILE UPLOAD ERROR ===');
                console.error('Error:', error);
                console.error('Stack:', error.stack);
                showToast(`❌ Error reading file: ${error.message}`, 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    };

    // Handle drag and drop
    const handleDrop = (e) => {
        e.preventDefault();
        const file = e.dataTransfer.files[0];
        if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
            handleFileUpload(file);
        }
    };

    // Parse Excel data wrapper
    const parseExcelDataWrapper = () => {
        if (!excelData) return null;

        const parsedData = parseExcelData(excelData, {
            selectedDates,
            onDatesExtracted: (dates) => {
                setAvailableDates(dates);
            },
            onError: (message) => {
                showToast(message, 'error');
            }
        });

        return parsedData;
    };

    // Load quote for optimizer
    const loadQuoteForOptimizer = async (file) => {
        if (!file) return;

        try {
            showToast('📋 Loading quote file...', 'success');

            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });

            console.log('=== QUOTE LOADER DEBUG START ===');
            console.log('Available sheets:', workbook.SheetNames);

            // Look for 'Rental Items' sheet or use first sheet
            let sheetName = workbook.SheetNames[0];
            if (workbook.SheetNames.includes('Rental Items')) {
                sheetName = 'Rental Items';
            } else {
                const rentalSheet = workbook.SheetNames.find(name =>
                    name.toLowerCase().includes('rental')
                );
                if (rentalSheet) sheetName = rentalSheet;
            }

            console.log('Using sheet:', sheetName);

            const sheet = workbook.Sheets[sheetName];
            const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

            console.log('First 3 rows (raw):', rawData.slice(0, 3));

            if (!rawData || rawData.length === 0) {
                showToast('❌ File appears to be empty', 'error');
                return;
            }

            // The first row should be the header
            const headers = rawData[0];
            console.log('Headers detected:', headers);

            if (!headers || headers.length === 0) {
                showToast('❌ Could not read file headers', 'error');
                return;
            }

            // Find equipment and quantity columns with flexible matching
            let equipmentColIdx = -1;
            let quantityColIdx = -1;
            let descriptionColIdx = -1;

            headers.forEach((header, idx) => {
                if (!header) return;

                const headerStr = String(header).toLowerCase().trim();

                // Equipment column
                if (headerStr.includes('equipment') && equipmentColIdx === -1) {
                    equipmentColIdx = idx;
                    console.log(`✓ Found Equipment column at index ${idx}`);
                }

                // Description column
                if (headerStr.includes('description') && descriptionColIdx === -1) {
                    descriptionColIdx = idx;
                    console.log(`✓ Found Description column at index ${idx}`);
                }

                // Quantity column
                if ((headerStr.includes('order') ||
                     headerStr.includes('qty') ||
                     headerStr.includes('quantity') ||
                     headerStr.includes('need')) && quantityColIdx === -1) {
                    quantityColIdx = idx;
                    console.log(`✓ Found Quantity column at index ${idx}`);
                }
            });

            // Fallback to positional columns based on typical format
            if (equipmentColIdx === -1 && headers.length >= 2) {
                console.log('Using positional fallback: Equipment at column 1');
                equipmentColIdx = 1;
            }

            if (quantityColIdx === -1 && headers.length >= 4) {
                console.log('Using positional fallback: Quantity at column 3');
                quantityColIdx = 3;
            }

            if (equipmentColIdx === -1) {
                const headersList = headers.filter(h => h).join(', ');
                showToast(`❌ Equipment column not found. Available: ${headersList}`, 'error');
                return;
            }

            if (quantityColIdx === -1) {
                const headersList = headers.filter(h => h).join(', ');
                showToast(`❌ Quantity column not found. Available: ${headersList}`, 'error');
                return;
            }

            console.log(`SUCCESS: Using Equipment col ${equipmentColIdx}, Quantity col ${quantityColIdx}`);

            // Extract equipment, descriptions, and quantities
            const quoteEquipment = {};
            const quoteDescriptions = {};
            let processedRows = 0;

            for (let i = 1; i < rawData.length; i++) {
                const row = rawData[i];
                if (!row || row.length === 0 || row.every(cell => !cell)) continue;

                const equipmentName = row[equipmentColIdx];
                const qtyRaw = row[quantityColIdx];
                const description = descriptionColIdx !== -1 ? row[descriptionColIdx] : '';

                if (!equipmentName) continue;

                const equipmentStr = String(equipmentName).trim();
                const descriptionStr = description ? String(description).trim() : '';

                let qty = 0;
                if (typeof qtyRaw === 'number') {
                    qty = Math.floor(Math.abs(qtyRaw));
                } else if (qtyRaw) {
                    const numStr = String(qtyRaw).replace(/[^0-9.-]/g, '');
                    qty = Math.floor(Math.abs(parseFloat(numStr) || 0));
                }

                if (equipmentStr && qty > 0) {
                    if (quoteEquipment[equipmentStr]) {
                        quoteEquipment[equipmentStr] += qty;
                    } else {
                        quoteEquipment[equipmentStr] = qty;
                        quoteDescriptions[equipmentStr] = descriptionStr;
                    }
                    processedRows++;

                    if (processedRows <= 5) {
                        console.log(`Row ${i}: "${equipmentStr}" -> ${qty} (${descriptionStr})`);
                    }
                }
            }

            console.log(`Processed ${processedRows} rows`);
            console.log('Unique equipment items:', Object.keys(quoteEquipment).length);

            if (Object.keys(quoteEquipment).length === 0) {
                showToast('⚠️ No equipment with quantities found in quote', 'warning');
                return;
            }

            // Match against inventory
            const parsedData = parseExcelDataWrapper();
            if (!parsedData) {
                showToast('❌ Could not parse inventory data', 'error');
                return;
            }

            const availableEquipment = Object.keys(parsedData.inventory);
            console.log(`Inventory has ${availableEquipment.length} items`);

            const matched = {};
            const unmatched = [];
            let matchCount = 0;

            Object.entries(quoteEquipment).forEach(([quoteName, qty]) => {
                // Try exact match first
                if (availableEquipment.includes(quoteName)) {
                    matched[quoteName] = qty;
                    matchCount++;
                } else {
                    // Try case-insensitive match
                    const match = availableEquipment.find(inv =>
                        inv.toLowerCase() === quoteName.toLowerCase()
                    );
                    if (match) {
                        matched[match] = qty;
                        matchCount++;
                    } else {
                        unmatched.push(quoteName);
                    }
                }
            });

            console.log(`Matched ${matchCount} items`);
            if (unmatched.length > 0) {
                console.log('Unmatched items:', unmatched.slice(0, 10));
            }
            console.log('=== QUOTE LOADER DEBUG END ===');

            setEquipmentFromQuote(matched);
            setEquipmentDescriptions(quoteDescriptions);

            const message = `✅ Loaded ${processedRows} items from quote, matched ${matchCount} with inventory`;
            showToast(message, 'success');

            if (unmatched.length > 0 && unmatched.length <= 5) {
                setTimeout(() => {
                    showToast(`⚠️ ${unmatched.length} items not in inventory: ${unmatched.slice(0, 3).join(', ')}...`, 'warning');
                }, 2000);
            } else if (unmatched.length > 5) {
                setTimeout(() => {
                    showToast(`⚠️ ${unmatched.length} items from quote not found in inventory`, 'warning');
                }, 2000);
            }

        } catch (error) {
            console.error('=== QUOTE LOADER ERROR ===');
            console.error('Error:', error);
            console.error('Stack:', error.stack);
            showToast(`❌ Error loading quote: ${error.message}`, 'error');
        }
    };

    // Run optimization algorithm
    const runOptimization = (scenarioId, inventory, distances, qtyNeeded, dest, region, preferredWh = '') => {
        const pulls = [];
        let remaining = qtyNeeded;

        // Filter warehouses by region if selected
        let availableWarehouses = Object.entries(inventory)
            .filter(([wh, qty]) => qty > 0)
            .map(([wh, qty]) => ({
                warehouse: wh,
                available: qty,
                distance: distances[wh] || 0,
                region: getWarehouseRegion(wh)
            }));

        if (region) {
            availableWarehouses = availableWarehouses.filter(wh =>
                REGIONS[region].warehouses.includes(wh.warehouse)
            );
        }

        // Sort based on scenario
        switch (scenarioId) {
            case 'minimize_distance':
                availableWarehouses.sort((a, b) => a.distance - b.distance);
                break;
            case 'minimize_trips':
                availableWarehouses.sort((a, b) => b.available - a.available);
                break;
            case 'balance_inventory':
                availableWarehouses.sort((a, b) => a.available - b.available);
                break;
            case 'prefer_local':
                availableWarehouses.sort((a, b) => {
                    if (a.warehouse === dest) return -1;
                    if (b.warehouse === dest) return 1;
                    return a.distance - b.distance;
                });
                break;
            case 'regional_priority':
                const destRegion = getWarehouseRegion(dest);
                availableWarehouses.sort((a, b) => {
                    if (a.region === destRegion && b.region !== destRegion) return -1;
                    if (b.region === destRegion && a.region !== destRegion) return 1;
                    return a.distance - b.distance;
                });
                break;
            case 'preferred_source':
                // Priority: Destination → Preferred Warehouse → Others by max availability
                availableWarehouses.sort((a, b) => {
                    if (a.warehouse === dest) return -1;
                    if (b.warehouse === dest) return 1;
                    if (a.warehouse === preferredWh) return -1;
                    if (b.warehouse === preferredWh) return 1;
                    // For remaining warehouses, sort by availability then distance
                    if (a.available !== b.available) return b.available - a.available;
                    return a.distance - b.distance;
                });
                break;
        }

        // Allocate quantities
        for (const wh of availableWarehouses) {
            if (remaining <= 0) break;

            const pullQty = Math.min(remaining, wh.available);
            const weight = (equipmentWeights[equipment] || 0) * pullQty;

            pulls.push({
                warehouse: wh.warehouse,
                quantity: pullQty,
                available: wh.available,
                distance: Math.round(wh.distance),
                region: wh.region,
                weight: weight
            });

            remaining -= pullQty;
        }

        const totalDistance = pulls.reduce((sum, p) => sum + (p.distance * p.quantity), 0);
        const totalWeight = pulls.reduce((sum, p) => sum + p.weight, 0);
        const totalTrips = pulls.length;

        return {
            pulls,
            totalDistance: Math.round(totalDistance),
            avgDistance: pulls.length > 0 ? Math.round(totalDistance / qtyNeeded) : 0,
            totalWeight: Math.round(totalWeight),
            totalTrips,
            fulfilled: remaining === 0,
            shortfall: remaining > 0 ? remaining : 0
        };
    };

    // Optimize logistics (single item)
    const optimizeLogistics = async () => {
        if (!destination || !equipment || !quantity) {
            showToast('⚠️ Please fill in all fields', 'warning');
            return;
        }

        const parsedData = parseExcelDataWrapper();
        if (!parsedData) return;

        setIsOptimizing(true);
        setResults(null);

        try {
            const qty = parseInt(quantity);
            const inventory = parsedData.inventory[equipment];

            if (!inventory) {
                showToast(`❌ Equipment "${equipment}" not found in inventory`, 'error');
                setIsOptimizing(false);
                return;
            }

            const destCoords = WAREHOUSE_COORDS[destination];
            const scenarioResults = {};

            // Calculate distances to all warehouses
            const warehouseDistances = {};
            for (const warehouse of Object.keys(WAREHOUSE_COORDS)) {
                if (warehouse === destination) {
                    warehouseDistances[warehouse] = 0;
                } else {
                    const whCoords = WAREHOUSE_COORDS[warehouse];
                    if (useRoadDistances && apiKey) {
                        warehouseDistances[warehouse] = await calculateRoadDistance(destCoords, whCoords, apiKey);
                        await new Promise(resolve => setTimeout(resolve, 100)); // Rate limiting
                    } else {
                        warehouseDistances[warehouse] = calculateGeodesicDistance(destCoords, whCoords);
                    }
                }
            }

            // Run each selected scenario
            for (const scenarioId of selectedScenarios) {
                const result = runOptimization(
                    scenarioId,
                    inventory,
                    warehouseDistances,
                    qty,
                    destination,
                    selectedRegion,
                    preferredWarehouse
                );
                scenarioResults[scenarioId] = result;
            }

            setResults({
                scenarios: scenarioResults,
                equipment,
                quantity: qty,
                destination,
                distanceType: useRoadDistances ? 'Road' : 'Geodesic'
            });

            showToast('✅ Optimization complete!', 'success');
            setActiveTab('results');
        } catch (error) {
            showToast(`❌ Optimization error: ${error.message}`, 'error');
        }

        setIsOptimizing(false);
    };

    // Bulk optimize all quote items
    const bulkOptimizeAllQuoteItems = async () => {
        const allItems = Object.entries(equipmentFromQuote).map(([equipmentName, quantity]) => ({
            name: equipmentName,
            quantity: quantity
        }));

        if (allItems.length === 0) {
            showToast('⚠️ No equipment items loaded from quote', 'warning');
            return;
        }

        if (!destination) {
            showToast('⚠️ Please select a destination warehouse', 'warning');
            return;
        }

        if (selectedScenarios.length === 0) {
            showToast('⚠️ Please select at least one optimization scenario', 'warning');
            return;
        }

        const parsedData = parseExcelDataWrapper();
        if (!parsedData) return;

        setIsOptimizing(true);
        setResults(null);

        try {
            const destCoords = WAREHOUSE_COORDS[destination];
            const allResults = {};
            let processedCount = 0;
            let skippedCount = 0;

            showToast(`🔄 Bulk optimizing ${allItems.length} items from quote...`, 'success');

            // Calculate distances once (reuse for all items)
            const warehouseDistances = {};
            for (const warehouse of Object.keys(WAREHOUSE_COORDS)) {
                if (warehouse === destination) {
                    warehouseDistances[warehouse] = 0;
                } else {
                    const whCoords = WAREHOUSE_COORDS[warehouse];
                    if (useRoadDistances && apiKey) {
                        warehouseDistances[warehouse] = await calculateRoadDistance(destCoords, whCoords, apiKey);
                        await new Promise(resolve => setTimeout(resolve, 100));
                    } else {
                        warehouseDistances[warehouse] = calculateGeodesicDistance(destCoords, whCoords);
                    }
                }
            }

            // Optimize each item from quote
            for (let i = 0; i < allItems.length; i++) {
                const item = allItems[i];
                const inventory = parsedData.inventory[item.name];

                // Update progress
                if (i % 5 === 0 || i === allItems.length - 1) {
                    showToast(`🔄 Processing ${i + 1}/${allItems.length}: ${item.name.substring(0, 30)}...`, 'success');
                }

                if (!inventory) {
                    console.log(`Item "${item.name}" not found in inventory, skipping`);
                    skippedCount++;
                    continue;
                }

                const itemResults = {};
                for (const scenarioId of selectedScenarios) {
                    const result = runOptimization(
                        scenarioId,
                        inventory,
                        warehouseDistances,
                        item.quantity,
                        destination,
                        selectedRegion,
                        preferredWarehouse
                    );
                    itemResults[scenarioId] = result;
                }

                allResults[item.name] = {
                    equipment: item.name,
                    quantity: item.quantity,
                    scenarios: itemResults
                };
                processedCount++;
            }

            setResults({
                bulkMode: true,
                items: allResults,
                destination,
                distanceType: useRoadDistances ? 'Road' : 'Geodesic',
                itemCount: processedCount,
                skippedCount: skippedCount,
                totalItems: allItems.length
            });

            showToast(`✅ Bulk optimization complete! Processed ${processedCount} items${skippedCount > 0 ? `, skipped ${skippedCount}` : ''}`, 'success');
            setActiveTab('results');
        } catch (error) {
            console.error('Bulk optimization error:', error);
            showToast(`❌ Optimization error: ${error.message}`, 'error');
        }

        setIsOptimizing(false);
    };

    // Export to Excel
    const exportToExcel = () => {
        if (!results) return;

        const wb = XLSX.utils.book_new();

        // Check if this is bulk mode
        const isBulkMode = results.bulkMode || results.items;

        if (isBulkMode) {
            // BULK MODE EXPORT - One sheet per SCENARIO showing all equipment

            // Get all unique scenario IDs from first item
            const firstItem = Object.values(results.items)[0];
            const scenarioIds = Object.keys(firstItem.scenarios);

            console.log('=== BULK EXPORT DEBUG ===');
            console.log('Number of scenarios:', scenarioIds.length);
            console.log('Scenario IDs:', scenarioIds);
            console.log('Number of equipment items:', Object.keys(results.items).length);

            // Create a sheet for each scenario
            scenarioIds.forEach(scenarioId => {
                const scenarioName = OPTIMIZATION_SCENARIOS.find(s => s.id === scenarioId)?.name || scenarioId;

                console.log(`Creating sheet for scenario: ${scenarioName} (${scenarioId})`);

                const sheetData = [
                    ['Equipment', 'Description', 'Requested Qty', 'Equipment Location', 'Warehouse Region', 'Pulled Qty',
                     'Available', 'Shortfall', 'Distance (mi)', 'Weight per Unit (lbs)', 'Total Weight (lbs)', 'Destination']
                ];

                // Add data for each equipment item
                Object.entries(results.items).forEach(([equipmentName, itemResult]) => {
                    const scenarioResult = itemResult.scenarios[scenarioId];

                    if (scenarioResult && scenarioResult.pulls && scenarioResult.pulls.length > 0) {
                        // Add a row for each pull
                        scenarioResult.pulls.forEach(pull => {
                            sheetData.push([
                                equipmentName,
                                equipmentDescriptions[equipmentName] || '',
                                itemResult.quantity,
                                pull.warehouse,
                                pull.region,
                                pull.quantity,
                                pull.available,
                                scenarioResult.fulfilled ? '' : scenarioResult.shortfall,
                                pull.distance,
                                pull.quantity > 0 ? Math.round(pull.weight / pull.quantity * 100) / 100 : 0,
                                pull.weight,
                                results.destination
                            ]);
                        });

                        // If not fulfilled, add a SHORTFALL row
                        if (!scenarioResult.fulfilled && scenarioResult.shortfall > 0) {
                            sheetData.push([
                                equipmentName,
                                equipmentDescriptions[equipmentName] || '',
                                itemResult.quantity,
                                'SHORTFALL',
                                '',
                                '',
                                '',
                                `${scenarioResult.shortfall} ${equipmentName}`,
                                '',
                                0,
                                0,
                                results.destination
                            ]);
                        }
                    } else {
                        // No pulls for this item - show as complete shortfall
                        sheetData.push([
                            equipmentName,
                            equipmentDescriptions[equipmentName] || '',
                            itemResult.quantity,
                            'SHORTFALL',
                            '',
                            '',
                            '',
                            `${itemResult.quantity} ${equipmentName}`,
                            '',
                            0,
                            0,
                            results.destination
                        ]);
                    }
                });

                const sheet = XLSX.utils.aoa_to_sheet(sheetData);

                // Highlight SHORTFALL rows in red
                const range = XLSX.utils.decode_range(sheet['!ref']);
                let shortfallRowCount = 0;
                for (let R = range.s.r + 1; R <= range.e.r; R++) {
                    const equipmentLocationCell = sheet[XLSX.utils.encode_cell({r: R, c: 3})];
                    if (equipmentLocationCell && equipmentLocationCell.v === 'SHORTFALL') {
                        shortfallRowCount++;
                        // Highlight all cells in this row with red background
                        for (let C = range.s.c; C <= range.e.c; C++) {
                            const cellAddress = XLSX.utils.encode_cell({r: R, c: C});
                            if (!sheet[cellAddress]) continue;

                            sheet[cellAddress].s = {
                                fill: {
                                    patternType: "solid",
                                    fgColor: { rgb: "FFCCCC" }
                                },
                                font: {
                                    color: { rgb: "990000" },
                                    bold: true
                                },
                                alignment: {
                                    vertical: "center",
                                    horizontal: "left"
                                }
                            };
                        }
                    }
                }

                if (shortfallRowCount > 0) {
                    console.log(`  Applied red highlighting to ${shortfallRowCount} shortfall rows`);
                }

                // Auto-size columns
                const colWidths = sheetData.reduce((widths, row) => {
                    row.forEach((cell, idx) => {
                        const cellWidth = String(cell || '').length;
                        widths[idx] = Math.max(widths[idx] || 10, cellWidth + 2);
                    });
                    return widths;
                }, []);
                sheet['!cols'] = colWidths.map(w => ({ wch: w }));

                console.log(`Sheet "${scenarioName}" has ${sheetData.length - 1} data rows`);
                XLSX.utils.book_append_sheet(wb, sheet, scenarioName.substring(0, 31));
            });

            console.log('Total sheets created:', wb.SheetNames.length);
            console.log('Sheet names:', wb.SheetNames);

        } else {
            // SINGLE MODE EXPORT
            const summaryData = [
                ['Logistics Optimization Results'],
                ['Equipment:', results.equipment || 'N/A'],
                ['Quantity Needed:', results.quantity || 'N/A'],
                ['Destination:', results.destination],
                ['Distance Type:', results.distanceType],
                ['Region Filter:', selectedRegion || 'None'],
                ['Preferred Warehouse:', preferredWarehouse || 'None'],
                ['Date Range:', selectedDates.start && selectedDates.end ?
                    `${selectedDates.start} to ${selectedDates.end}` : 'All Dates'],
                [],
                ['Scenario', 'Total Distance', 'Avg Distance', 'Total Weight (lbs)', 'Trips', 'Status']
            ];

            Object.entries(results.scenarios).forEach(([scenarioId, result]) => {
                const scenarioName = OPTIMIZATION_SCENARIOS.find(s => s.id === scenarioId)?.name || scenarioId;
                summaryData.push([
                    scenarioName,
                    result.totalDistance,
                    result.avgDistance,
                    result.totalWeight,
                    result.totalTrips,
                    result.fulfilled ? 'Fulfilled' : `Short ${result.shortfall}`
                ]);
            });

            const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);

            // Auto-size columns for summary sheet
            const summaryColWidths = summaryData.reduce((widths, row) => {
                row.forEach((cell, idx) => {
                    const cellWidth = String(cell || '').length;
                    widths[idx] = Math.max(widths[idx] || 10, cellWidth + 2);
                });
                return widths;
            }, []);
            summarySheet['!cols'] = summaryColWidths.map(w => ({ wch: w }));

            XLSX.utils.book_append_sheet(wb, summarySheet, 'Summary');

            // Create detail sheets for each scenario
            Object.entries(results.scenarios).forEach(([scenarioId, result]) => {
                const scenarioName = OPTIMIZATION_SCENARIOS.find(s => s.id === scenarioId)?.name || scenarioId;
                const detailData = [
                    ['Warehouse', 'Quantity', 'Available', 'Distance (mi)', 'Weight (lbs)', 'Region'],
                    ...result.pulls.map(p => [
                        p.warehouse,
                        p.quantity,
                        p.available,
                        p.distance,
                        p.weight,
                        p.region
                    ])
                ];

                const detailSheet = XLSX.utils.aoa_to_sheet(detailData);

                // Auto-size columns for detail sheet
                const detailColWidths = detailData.reduce((widths, row) => {
                    row.forEach((cell, idx) => {
                        const cellWidth = String(cell || '').length;
                        widths[idx] = Math.max(widths[idx] || 10, cellWidth + 2);
                    });
                    return widths;
                }, []);
                detailSheet['!cols'] = detailColWidths.map(w => ({ wch: w }));

                XLSX.utils.book_append_sheet(wb, detailSheet, scenarioName.substring(0, 31));
            });
        }

        XLSX.writeFile(wb, `logistics_optimization_${Date.now()}.xlsx`, { cellStyles: true });
        showToast('📊 Excel file exported successfully!', 'success');
    };

    // Render
    return (
        <div className="app">
            <Header />

            <div className="container">
                {/* Tab Navigation */}
                <div className="tabs">
                    <button
                        className={`tab ${activeTab === 'upload' ? 'active' : ''}`}
                        onClick={() => setActiveTab('upload')}
                    >
                        📁 Upload
                    </button>
                    <button
                        className={`tab ${activeTab === 'optimize' ? 'active' : ''}`}
                        onClick={() => setActiveTab('optimize')}
                        disabled={!excelData}
                    >
                        ⚙️ Optimize
                    </button>
                    <button
                        className={`tab ${activeTab === 'results' ? 'active' : ''}`}
                        onClick={() => setActiveTab('results')}
                        disabled={!results}
                    >
                        📊 Results
                    </button>
                    <button
                        className={`tab ${activeTab === 'settings' ? 'active' : ''}`}
                        onClick={() => setActiveTab('settings')}
                    >
                        🔧 Settings
                    </button>
                </div>

                {/* Tab Content */}
                <div className="tab-content">
                    {activeTab === 'upload' && (
                        <UploadTab
                            excelFile={excelFile}
                            excelData={excelData}
                            onFileUpload={handleFileUpload}
                            onDrop={handleDrop}
                            fileInputRef={fileInputRef}
                        />
                    )}

                    {activeTab === 'optimize' && (
                        <OptimizeTab
                            excelData={excelData}
                            destination={destination}
                            setDestination={setDestination}
                            equipment={equipment}
                            setEquipment={setEquipment}
                            quantity={quantity}
                            setQuantity={setQuantity}
                            selectedRegion={selectedRegion}
                            setSelectedRegion={setSelectedRegion}
                            preferredWarehouse={preferredWarehouse}
                            setPreferredWarehouse={setPreferredWarehouse}
                            selectedDates={selectedDates}
                            setSelectedDates={setSelectedDates}
                            availableDates={availableDates}
                            selectedScenarios={selectedScenarios}
                            setSelectedScenarios={setSelectedScenarios}
                            equipmentFromQuote={equipmentFromQuote}
                            quoteForOptimizerRef={quoteForOptimizerRef}
                            loadQuoteForOptimizer={loadQuoteForOptimizer}
                            optimizeLogistics={optimizeLogistics}
                            bulkOptimizeAllQuoteItems={bulkOptimizeAllQuoteItems}
                            isOptimizing={isOptimizing}
                        />
                    )}

                    {activeTab === 'results' && (
                        <ResultsTab
                            results={results}
                            selectedScenarios={selectedScenarios}
                            selectedRegion={selectedRegion}
                            exportToExcel={exportToExcel}
                        />
                    )}

                    {activeTab === 'settings' && (
                        <SettingsTab
                            apiKey={apiKey}
                            setApiKey={setApiKey}
                            useRoadDistances={useRoadDistances}
                            setUseRoadDistances={setUseRoadDistances}
                            equipmentWeights={equipmentWeights}
                            setEquipmentWeights={setEquipmentWeights}
                            showToast={showToast}
                        />
                    )}
                </div>
            </div>

            {/* Toast Notification */}
            {toast && <Toast message={toast.message} type={toast.type} />}
        </div>
    );
};

export default App;
