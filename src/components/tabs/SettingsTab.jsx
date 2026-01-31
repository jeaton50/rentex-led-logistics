import React, { useState } from 'react';
import Card from '../common/Card';
import Button from '../common/Button';
import { REGIONS } from '../../utils/constants';

const SettingsTab = ({
    apiKey,
    setApiKey,
    useRoadDistances,
    setUseRoadDistances,
    equipmentWeights,
    setEquipmentWeights,
    showToast
}) => {
    const [newEquipmentName, setNewEquipmentName] = useState('');
    const [newEquipmentWeight, setNewEquipmentWeight] = useState('');

    const handleSaveWeight = () => {
        const name = newEquipmentName.trim();
        const weight = parseFloat(newEquipmentWeight);

        if (name && weight > 0) {
            const newWeights = { ...equipmentWeights, [name]: weight };
            setEquipmentWeights(newWeights);
            localStorage.setItem('equipment_weights', JSON.stringify(newWeights));
            showToast(`✅ Weight saved for ${name}`, 'success');
            setNewEquipmentName('');
            setNewEquipmentWeight('');
        }
    };

    const handleRemoveWeight = (name) => {
        const newWeights = { ...equipmentWeights };
        delete newWeights[name];
        setEquipmentWeights(newWeights);
        localStorage.setItem('equipment_weights', JSON.stringify(newWeights));
        showToast(`🗑️ Weight removed for ${name}`, 'success');
    };

    return (
        <div>
            <Card title="🔧 API Configuration">
                <div className="form-group">
                    <label className="form-label">OpenRouteService API Key</label>
                    <input
                        type="text"
                        className="form-input"
                        placeholder="Get free key at openrouteservice.org"
                        value={apiKey}
                        onChange={(e) => setApiKey(e.target.value)}
                    />
                    <p style={{ color: 'var(--text-secondary)', fontSize: '0.9rem', marginTop: '0.5rem' }}>
                        Free tier: 2,000 requests/day • Required for road-based distances
                    </p>
                </div>

                <div className="checkbox-wrapper">
                    <input
                        type="checkbox"
                        checked={useRoadDistances}
                        onChange={(e) => setUseRoadDistances(e.target.checked)}
                        id="road-distances"
                    />
                    <label htmlFor="road-distances">
                        Use road-based distances (requires API key)
                    </label>
                </div>

                <div style={{ marginTop: '1rem' }}>
                    <span className={`api-status ${apiKey ? 'connected' : 'disconnected'}`}>
                        <span className={`status-dot ${apiKey ? 'active' : 'inactive'}`}></span>
                        {apiKey ? 'API Key Configured' : 'No API Key'}
                    </span>
                </div>
            </Card>

            <Card title="⚖️ Equipment Weights">
                <p style={{ color: 'var(--text-secondary)', marginBottom: '1.5rem' }}>
                    Set default weights for equipment (lbs per unit)
                </p>

                <div className="form-group">
                    <label className="form-label">Equipment Name</label>
                    <input
                        type="text"
                        className="form-input"
                        placeholder="Enter equipment name..."
                        value={newEquipmentName}
                        onChange={(e) => setNewEquipmentName(e.target.value)}
                    />
                </div>

                <div className="form-group">
                    <label className="form-label">Weight (lbs)</label>
                    <input
                        type="number"
                        className="form-input"
                        placeholder="Enter weight..."
                        value={newEquipmentWeight}
                        onChange={(e) => setNewEquipmentWeight(e.target.value)}
                        min="0"
                        step="0.1"
                    />
                </div>

                <Button variant="secondary" onClick={handleSaveWeight}>
                    💾 Save Weight
                </Button>

                {Object.keys(equipmentWeights).length > 0 && (
                    <div style={{ marginTop: '2rem' }}>
                        <h3 style={{ marginBottom: '1rem' }}>Saved Weights:</h3>
                        <div className="table-container">
                            <table className="data-table">
                                <thead>
                                    <tr>
                                        <th>Equipment</th>
                                        <th>Weight (lbs)</th>
                                        <th>Action</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {Object.entries(equipmentWeights).map(([name, weight]) => (
                                        <tr key={name}>
                                            <td>{name}</td>
                                            <td>{weight}</td>
                                            <td>
                                                <Button
                                                    variant="warning"
                                                    style={{ padding: '0.5rem 1rem', fontSize: '0.9rem' }}
                                                    onClick={() => handleRemoveWeight(name)}
                                                >
                                                    🗑️ Remove
                                                </Button>
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                )}
            </Card>

            <Card title="🌍 Regional Configuration">
                {Object.entries(REGIONS).map(([regionName, regionData]) => (
                    <div key={regionName} style={{ marginBottom: '1.5rem' }}>
                        <h3 style={{
                            color: 'var(--secondary)',
                            marginBottom: '0.5rem',
                            fontSize: '1.1rem'
                        }}>
                            {regionName}
                        </h3>
                        <p style={{ color: 'var(--text-secondary)', marginBottom: '0.5rem' }}>
                            {regionData.description}
                        </p>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.5rem' }}>
                            {regionData.warehouses.map(wh => (
                                <span key={wh} className={`badge ${regionName.includes('1') ? 'region-1' : 'region-2'}`}>
                                    {wh}
                                </span>
                            ))}
                        </div>
                    </div>
                ))}
            </Card>
        </div>
    );
};

export default SettingsTab;
