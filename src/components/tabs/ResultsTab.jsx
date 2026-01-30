import React from 'react';
import Card from '../common/Card';
import Button from '../common/Button';
import { OPTIMIZATION_SCENARIOS } from '../../utils/constants';

const ResultsTab = ({ results, selectedScenarios, selectedRegion, exportToExcel }) => {
    if (!results) {
        return (
            <Card style={{ textAlign: 'center', padding: '4rem' }}>
                <h3 style={{ color: 'var(--text-secondary)' }}>
                    No results yet. Run an optimization first.
                </h3>
            </Card>
        );
    }

    // Bulk mode results
    if (results.bulkMode) {
        return (
            <div>
                <Card>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2rem' }}>
                        <div>
                            <h2 className="card-title">📦 Bulk Optimization Results</h2>
                            <p style={{ color: 'var(--text-secondary)' }}>
                                {results.itemCount} items optimized • Destination: {results.destination}
                            </p>
                        </div>
                        <Button variant="success" onClick={exportToExcel}>
                            📥 Export to Excel
                        </Button>
                    </div>

                    <div className="stats-grid">
                        <div className="stat-card">
                            <div className="stat-label">Total Items</div>
                            <div className="stat-value">{results.itemCount}</div>
                        </div>
                        <div className="stat-card">
                            <div className="stat-label">Destination</div>
                            <div className="stat-value" style={{ fontSize: '1.2rem' }}>
                                {results.destination}
                            </div>
                        </div>
                        <div className="stat-card">
                            <div className="stat-label">Distance Method</div>
                            <div className="stat-value" style={{ fontSize: '1.2rem' }}>
                                {results.distanceType}
                            </div>
                        </div>
                        <div className="stat-card">
                            <div className="stat-label">Scenarios Compared</div>
                            <div className="stat-value">{selectedScenarios.length}</div>
                        </div>
                    </div>

                    <div style={{
                        marginTop: '2rem',
                        padding: '2rem',
                        background: 'rgba(139, 28, 28, 0.1)',
                        borderRadius: '12px',
                        border: '2px solid var(--secondary)',
                        textAlign: 'center'
                    }}>
                        <div style={{ fontSize: '3rem', marginBottom: '1rem' }}>✅</div>
                        <h3 style={{ color: 'var(--success)', marginBottom: '1rem' }}>
                            Optimization Complete!
                        </h3>
                        <p style={{ color: 'var(--text-secondary)', fontSize: '1.1rem', marginBottom: '1.5rem' }}>
                            {results.itemCount} equipment items have been optimized across {selectedScenarios.length} scenario{selectedScenarios.length > 1 ? 's' : ''}.
                        </p>
                        <p style={{ color: 'var(--text-primary)', fontSize: '1rem' }}>
                            📊 Click <strong>"Export to Excel"</strong> above to download the complete results with all equipment details, pulls, and shortfalls for each scenario.
                        </p>
                    </div>

                    {/* Show scenarios selected */}
                    <div style={{ marginTop: '2rem' }}>
                        <h4 style={{ color: 'var(--secondary)', marginBottom: '1rem' }}>
                            Scenarios Included:
                        </h4>
                        <div style={{ display: 'flex', gap: '0.5rem', flexWrap: 'wrap' }}>
                            {selectedScenarios.map(scenarioId => {
                                const scenario = OPTIMIZATION_SCENARIOS.find(s => s.id === scenarioId);
                                return (
                                    <div key={scenarioId} style={{
                                        padding: '0.75rem 1.5rem',
                                        background: 'rgba(183, 28, 28, 0.2)',
                                        border: '1px solid var(--secondary)',
                                        borderRadius: '8px',
                                        color: 'var(--secondary)',
                                        fontWeight: 600
                                    }}>
                                        {scenario?.icon} {scenario?.name}
                                    </div>
                                );
                            })}
                        </div>
                    </div>
                </Card>
            </div>
        );
    }

    // Single item results
    return (
        <div>
            <Card>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2rem' }}>
                    <div>
                        <h2 className="card-title">📊 Optimization Results</h2>
                        <p style={{ color: 'var(--text-secondary)' }}>
                            {results.equipment} • {results.quantity} units • {results.destination}
                        </p>
                    </div>
                    <Button variant="success" onClick={exportToExcel}>
                        📥 Export to Excel
                    </Button>
                </div>

                <div className="stats-grid">
                    <div className="stat-card">
                        <div className="stat-label">Scenarios Analyzed</div>
                        <div className="stat-value">{Object.keys(results.scenarios).length}</div>
                    </div>
                    <div className="stat-card">
                        <div className="stat-label">Distance Method</div>
                        <div className="stat-value" style={{ fontSize: '1.2rem' }}>
                            {results.distanceType}
                        </div>
                    </div>
                    <div className="stat-card">
                        <div className="stat-label">Regional Filter</div>
                        <div className="stat-value" style={{ fontSize: '1.2rem' }}>
                            {selectedRegion || 'None'}
                        </div>
                    </div>
                </div>
            </Card>

            {Object.entries(results.scenarios).map(([scenarioId, result]) => {
                const scenario = OPTIMIZATION_SCENARIOS.find(s => s.id === scenarioId);
                return (
                    <div key={scenarioId} className="result-card">
                        <div className="result-header">
                            <div>
                                <div className="result-title">
                                    {scenario?.icon} {scenario?.name}
                                </div>
                                <p style={{ color: 'var(--text-secondary)', marginTop: '0.5rem' }}>
                                    {scenario?.description}
                                </p>
                            </div>
                            <div style={{ textAlign: 'right' }}>
                                {result.fulfilled ? (
                                    <span className="api-status connected">✓ Fulfilled</span>
                                ) : (
                                    <span className="api-status disconnected">
                                        ⚠️ Short {result.shortfall}
                                    </span>
                                )}
                            </div>
                        </div>

                        <div className="stats-grid">
                            <div className="stat-card">
                                <div className="stat-label">Total Distance</div>
                                <div className="stat-value">
                                    {result.totalDistance.toLocaleString()}
                                    <span className="stat-unit">mi</span>
                                </div>
                            </div>
                            <div className="stat-card">
                                <div className="stat-label">Avg Distance</div>
                                <div className="stat-value">
                                    {result.avgDistance.toLocaleString()}
                                    <span className="stat-unit">mi</span>
                                </div>
                            </div>
                            <div className="stat-card">
                                <div className="stat-label">Total Weight</div>
                                <div className="stat-value">
                                    {result.totalWeight.toLocaleString()}
                                    <span className="stat-unit">lbs</span>
                                </div>
                            </div>
                            <div className="stat-card">
                                <div className="stat-label">Trips Required</div>
                                <div className="stat-value">{result.totalTrips}</div>
                            </div>
                        </div>

                        <h3 style={{ marginTop: '2rem', marginBottom: '1rem', color: 'var(--secondary)' }}>
                            Warehouse Allocation:
                        </h3>
                        <div className="warehouse-list">
                            {result.pulls.map((pull, idx) => (
                                <div key={idx} className="warehouse-item">
                                    <div className="warehouse-name">{pull.warehouse}</div>
                                    <div>
                                        <span className={`region-tag ${pull.region.includes('1') ? 'region-1' : 'region-2'}`}>
                                            {pull.region}
                                        </span>
                                    </div>
                                    <div className="warehouse-qty">
                                        <strong>{pull.quantity}</strong> / {pull.available} units
                                    </div>
                                    <div className="warehouse-distance">
                                        {pull.distance} mi
                                    </div>
                                    <div className="warehouse-weight">
                                        {pull.weight.toLocaleString()} lbs
                                    </div>
                                </div>
                            ))}
                        </div>
                    </div>
                );
            })}
        </div>
    );
};

export default ResultsTab;
