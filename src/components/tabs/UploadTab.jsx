import React from 'react';
import Card from '../common/Card';

const UploadTab = ({ excelFile, excelData, onFileUpload, onDrop, fileInputRef }) => {
    return (
        <Card title="📁 Upload Inventory Excel">
            <div
                className="file-upload-zone"
                onDrop={onDrop}
                onDragOver={(e) => e.preventDefault()}
                onClick={() => fileInputRef.current?.click()}
            >
                <div className="file-upload-icon">📊</div>
                <h3>Drop Excel file here or click to browse</h3>
                <p style={{ color: 'var(--text-secondary)', marginTop: '1rem' }}>
                    Supported: .xlsx, .xls
                </p>
                {excelFile && <div className="file-name">✓ {excelFile.name}</div>}
            </div>
            <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => onFileUpload(e.target.files[0])}
                style={{ display: 'none' }}
            />

            {excelData && (
                <div style={{ marginTop: '2rem' }}>
                    <h3 style={{ marginBottom: '1rem' }}>📋 Preview</h3>
                    <div className="table-container">
                        <table className="data-table">
                            <thead>
                                <tr>
                                    {excelData[0]?.slice(0, 10).map((header, i) => (
                                        <th key={i}>{header}</th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody>
                                {excelData.slice(1, 6).map((row, i) => (
                                    <tr key={i}>
                                        {row.slice(0, 10).map((cell, j) => (
                                            <td key={j}>{cell}</td>
                                        ))}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                    <p style={{ marginTop: '1rem', color: 'var(--text-secondary)' }}>
                        Showing first 5 rows of {excelData.length - 1} total rows
                    </p>
                </div>
            )}
        </Card>
    );
};

export default UploadTab;
