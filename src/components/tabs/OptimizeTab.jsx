import React from 'react';
import Card from '../common/Card';
import Button from '../common/Button';
import { WAREHOUSE_COORDS, OPTIMIZATION_SCENARIOS, REGIONS } from '../../utils/constants';

const OptimizeTab = ({
    excelData,
    destination,
    setDestination,
    equipment,
    setEquipment,
    quantity,
    setQuantity,
    selectedRegion,
    setSelectedRegion,
    preferredWarehouse,
    setPreferredWarehouse,
    selectedDates,
    setSelectedDates,
    availableDates,
    selectedScenarios,
    setSelectedScenarios,
    equipmentFromQuote,
    quoteForOptimizerRef,
    loadQuoteForOptimizer,
    optimizeLogistics,
    bulkOptimizeAllQuoteItems,
    isOptimizing
}) => {
    const toggleScenario = (scenarioId) => {
        setSelectedScenarios(prev =>
            prev.includes(scenarioId)
                ? prev.filter(id => id !== scenarioId)
                : [...prev, scenarioId]
        );
    };

    return (
        <div>
            <Card title="⚙️ Configuration">
                {/* Quote Upload Section */}
                <div style={{
                    background: 'rgba(183, 28, 28, 0.1)',
                    padding: '1rem',
                    borderRadius: '8px',
                    marginBottom: '1.5rem',
                    border: '1px solid var(--secondary)'
                }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', flexWrap: 'wrap' }}>
                        <span style={{ fontWeight: 600 }}>📋 Quick Setup:</span>
                        <Button
                            variant="secondary"
                            onClick={() => quoteForOptimizerRef.current?.click()}
                            disabled={!excelData}
                            style={{ padding: '0.75rem 1.5rem' }}
                        >
                            Load Quote (Auto-fill Equipment)
                        </Button>
                        {Object.keys(equipmentFromQuote).length > 0 && (
                            <span style={{ color: 'var(--success)', fontWeight: 600 }}>
                                ✓ {Object.keys(equipmentFromQuote).length} items loaded from quote
                            </span>
                        )}
                    </div>
                    <input
                        ref={quoteForOptimizerRef}
                        type="file"
                        accept=".xlsx,.xls"
                        onChange={(e) => {
                            const file = e.target.files[0];
                            if (file) loadQuoteForOptimizer(file);
                        }}
                        style={{ display: 'none' }}
                    />
                    {!excelData && (
                        <p style={{
                            color: 'var(--text-secondary)',
                            fontSize: '0.9rem',
                            marginTop: '0.5rem',
                            marginBottom: 0
                        }}>
                            ⚠️ Load inventory file first before loading quote
                        </p>
                    )}
                </div>

                <div className="form-group">
                    <label className="form-label">Destination Warehouse</label>
                    <select
                        className="form-select"
                        value={destination}
                        onChange={(e) => setDestination(e.target.value)}
                    >
                        <option value="">Select destination...</option>
                        {Object.keys(WAREHOUSE_COORDS).map(wh => (
                            <option key={wh} value={wh}>{wh}</option>
                        ))}
                    </select>
                </div>

                <div className="form-group">
                    <label className="form-label">Equipment Name</label>
                    <input
                        type="text"
                        className="form-input"
                        placeholder="Enter equipment name or select from quote..."
                        value={equipment}
                        onChange={(e) => setEquipment(e.target.value)}
                        list="equipment-suggestions"
                    />
                    {Object.keys(equipmentFromQuote).length > 0 && (
                        <datalist id="equipment-suggestions">
                            {Object.keys(equipmentFromQuote).map(eq => (
                                <option key={eq} value={eq} />
                            ))}
                        </datalist>
                    )}
                </div>

                <div className="form-group">
                    <label className="form-label">Quantity Needed</label>
                    <input
                        type="number"
                        className="form-input"
                        placeholder="Enter quantity..."
                        value={quantity}
                        onChange={(e) => setQuantity(e.target.value)}
                        min="1"
                    />
                    {equipment && equipmentFromQuote[equipment] && (
                        <p style={{
                            color: 'var(--secondary)',
                            fontSize: '0.9rem',
                            marginTop: '0.5rem',
                            fontFamily: "'JetBrains Mono', monospace"
                        }}>
                            💡 Quote specifies {equipmentFromQuote[equipment]} units
                            <Button
                                variant="secondary"
                                onClick={() => setQuantity(equipmentFromQuote[equipment])}
                                style={{
                                    marginLeft: '1rem',
                                    padding: '0.25rem 0.75rem',
                                    fontSize: '0.85rem'
                                }}
                            >
                                Use Quote Qty
                            </Button>
                        </p>
                    )}
                </div>

                {Object.keys(equipmentFromQuote).length > 0 && (
                    <div style={{
                        background: 'rgba(0, 0, 0, 0.2)',
                        padding: '1rem',
                        borderRadius: '8px',
                        marginTop: '1rem'
                    }}>
                        <h4 style={{
                            color: 'var(--secondary)',
                            marginBottom: '1rem',
                            fontSize: '1rem'
                        }}>
                            📋 Equipment Loaded from Quote ({Object.keys(equipmentFromQuote).length} items):
                        </h4>
                        <div style={{
                            display: 'grid',
                            gridTemplateColumns: 'repeat(auto-fill, minmax(250px, 1fr))',
                            gap: '0.5rem',
                            maxHeight: '300px',
                            overflowY: 'auto'
                        }}>
                            {Object.entries(equipmentFromQuote).map(([eq, qty]) => (
                                <div
                                    key={eq}
                                    style={{
                                        background: 'rgba(183, 28, 28, 0.1)',
                                        padding: '0.75rem',
                                        borderRadius: '6px',
                                        fontSize: '0.85rem',
                                        fontFamily: "'JetBrains Mono', monospace",
                                        border: '1px solid rgba(183, 28, 28, 0.3)'
                                    }}
                                >
                                    <div style={{ fontWeight: 600, marginBottom: '0.25rem' }}>{eq}</div>
                                    <div style={{ color: 'var(--text-secondary)' }}>
                                        Qty: {qty} units
                                    </div>
                                </div>
                            ))}
                        </div>
                        <p style={{
                            color: 'var(--success)',
                            fontSize: '0.9rem',
                            marginTop: '1rem',
                            marginBottom: 0,
                            fontWeight: 600
                        }}>
                            ✓ Ready for bulk optimization! Select destination and scenarios, then click "Bulk Optimize All"
                        </p>
                    </div>
                )}

                <div className="form-group">
                    <label className="form-label">Regional Filter (Optional)</label>
                    <select
                        className="form-select"
                        value={selectedRegion}
                        onChange={(e) => setSelectedRegion(e.target.value)}
                    >
                        <option value="">All Regions</option>
                        {Object.entries(REGIONS).map(([name, data]) => (
                            <option key={name} value={name}>
                                {name} - {data.warehouses.length} warehouses
                            </option>
                        ))}
                    </select>
                </div>

                {selectedRegion && (
                    <div style={{
                        background: 'rgba(183, 28, 28, 0.1)',
                        padding: '1rem',
                        borderRadius: '8px',
                        marginTop: '1rem'
                    }}>
                        <strong>Selected Region:</strong> {selectedRegion}<br />
                        <strong>Warehouses:</strong> {REGIONS[selectedRegion].warehouses.join(', ')}
                    </div>
                )}

                <div className="form-group">
                    <label className="form-label">Preferred Warehouse (Optional)</label>
                    <select
                        className="form-select"
                        value={preferredWarehouse}
                        onChange={(e) => setPreferredWarehouse(e.target.value)}
                    >
                        <option value="">None - Use Standard Priority</option>
                        {Object.keys(WAREHOUSE_COORDS).map(wh => (
                            <option key={wh} value={wh}>{wh}</option>
                        ))}
                    </select>
                </div>

                {preferredWarehouse && (
                    <div style={{
                        background: 'rgba(255, 215, 0, 0.1)',
                        padding: '1rem',
                        borderRadius: '8px',
                        border: '1px solid rgba(255, 215, 0, 0.3)',
                        marginTop: '0.5rem'
                    }}>
                        <strong>⭐ Preferred Warehouse:</strong> {preferredWarehouse}<br />
                        <span style={{ fontSize: '0.85rem', color: 'var(--text-secondary)' }}>
                            Will prioritize after destination in "Preferred Source" scenario
                        </span>
                    </div>
                )}

                <div className="form-group">
                    <label className="form-label">Date Range (Optional)</label>
                    <div style={{ display: 'flex', gap: '1rem', alignItems: 'center' }}>
                        <input
                            type="date"
                            className="form-input"
                            value={selectedDates.start}
                            onChange={(e) => setSelectedDates({ ...selectedDates, start: e.target.value })}
                            style={{ flex: 1 }}
                        />
                        <span style={{ color: 'var(--text-secondary)' }}>to</span>
                        <input
                            type="date"
                            className="form-input"
                            value={selectedDates.end}
                            onChange={(e) => setSelectedDates({ ...selectedDates, end: e.target.value })}
                            style={{ flex: 1 }}
                        />
                    </div>
                    <p style={{ color: 'var(--text-secondary)', fontSize: '0.85rem', marginTop: '0.5rem' }}>
                        📅 Filter inventory availability by date range
                    </p>
                </div>

                {selectedDates.start && selectedDates.end && (
                    <div style={{
                        background: 'rgba(76, 175, 80, 0.1)',
                        padding: '1rem',
                        borderRadius: '8px',
                        border: '1px solid rgba(76, 175, 80, 0.3)',
                        marginTop: '0.5rem'
                    }}>
                        <strong>📅 Date Range Selected:</strong> {selectedDates.start} to {selectedDates.end}<br />
                        <span style={{ fontSize: '0.85rem', color: 'var(--text-secondary)' }}>
                            Only inventory available during this period will be used
                        </span>
                    </div>
                )}

                {availableDates.length > 0 && (
                    <div style={{
                        background: 'rgba(139, 28, 28, 0.1)',
                        padding: '1rem',
                        borderRadius: '8px',
                        border: '1px solid rgba(139, 28, 28, 0.3)',
                        marginTop: '0.5rem'
                    }}>
                        <strong>📆 Available Dates in Inventory:</strong><br />
                        <span style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', fontFamily: "'JetBrains Mono', monospace" }}>
                            {availableDates[0]?.dateStr} to {availableDates[availableDates.length - 1]?.dateStr}
                            {' '}({availableDates.length} dates)
                        </span>
                    </div>
                )}
            </Card>

            <Card title="🎯 Optimization Scenarios">
                <p style={{ color: 'var(--text-secondary)', marginBottom: '1.5rem' }}>
                    Select one or more scenarios to compare
                </p>

                <div className="scenarios-grid">
                    {OPTIMIZATION_SCENARIOS.map(scenario => (
                        <div
                            key={scenario.id}
                            className={`scenario-card ${selectedScenarios.includes(scenario.id) ? 'selected' : ''}`}
                            onClick={() => toggleScenario(scenario.id)}
                        >
                            <div className="scenario-title">
                                {scenario.icon} {scenario.name}
                            </div>
                            <div className="scenario-description">
                                {scenario.description}
                            </div>
                        </div>
                    ))}
                </div>
            </Card>

            <div style={{ textAlign: 'center', marginTop: '2rem', display: 'flex', gap: '1rem', justifyContent: 'center', flexWrap: 'wrap' }}>
                {Object.keys(equipmentFromQuote).length > 0 && (
                    <Button
                        onClick={bulkOptimizeAllQuoteItems}
                        disabled={!destination || selectedScenarios.length === 0 || isOptimizing}
                        style={{ fontSize: '1.2rem', padding: '1.5rem 3rem' }}
                    >
                        {isOptimizing ? (
                            <>
                                <span className="loading"></span>
                                Optimizing...
                            </>
                        ) : (
                            <>
                                🚀 Bulk Optimize All ({Object.keys(equipmentFromQuote).length} items)
                            </>
                        )}
                    </Button>
                )}

                <Button
                    onClick={optimizeLogistics}
                    disabled={!destination || !equipment || !quantity || selectedScenarios.length === 0 || isOptimizing}
                    style={{ fontSize: '1.2rem', padding: '1.5rem 3rem' }}
                >
                    {isOptimizing ? (
                        <>
                            <span className="loading"></span>
                            Optimizing...
                        </>
                    ) : (
                        <>
                            🎯 Optimize Single Item
                        </>
                    )}
                </Button>
            </div>
        </div>
    );
};

export default OptimizeTab;
