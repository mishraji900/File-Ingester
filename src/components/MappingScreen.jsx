import React, { useEffect, useState } from 'react';
import PreviewModal from './PreviewModal.jsx';

export default function MappingScreen({ payload, onBack }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [previewRow, setPreviewRow] = useState(null);

  useEffect(() => {
    async function loadSheets() {
      try {
        const planSheets = await window.api.getSheets(payload.planFile);
        const trustSheets = await window.api.getSheets(payload.trustFile);

        if (planSheets.error) throw new Error(planSheets.error);
        if (trustSheets.error) throw new Error(trustSheets.error);

        setRows([
          {
            rowId: 'plan-1',
            fileLabel: 'plan operations file',
            filePath: payload.planFile,
            sheetOptions: planSheets.sheets,
            selectedSheet: planSheets.sheets[0] || ''
          },
          {
            rowId: 'plan-2',
            fileLabel: 'plan operations file',
            filePath: payload.planFile,
            sheetOptions: planSheets.sheets,
            selectedSheet: planSheets.sheets[1] || planSheets.sheets[0] || ''
          },
          {
            rowId: 'trust-1',
            fileLabel: 'net trust assets file',
            filePath: payload.trustFile,
            sheetOptions: trustSheets.sheets,
            selectedSheet: trustSheets.sheets[0] || ''
          }
        ]);
      } catch (err) {
        setError('Error loading sheets: ' + err.message);
      } finally {
        setLoading(false);
      }
    }
    loadSheets();
  }, [payload]);

  function handleSheetChange(rowId, newSheet) {
    setRows(prev =>
      prev.map(r => (r.rowId === rowId ? { ...r, selectedSheet: newSheet } : r))
    );
  }

  function handleConfirm() {
    console.log('Confirmed mapping:', rows);
    alert('Mapping confirmed. Validation stage not yet implemented.');
  }

  if (loading) return (
    <div className="page">
      <div className="card" style={{ textAlign: 'center', padding: '48px' }}>
        <p style={{ color: 'var(--muted-foreground)' }}>Loading sheet names...</p>
      </div>
    </div>
  );

  if (error) return (
    <div className="page">
      <div className="card">
        <p className="status">{error}</p>
        <div className="btn-row" style={{ marginTop: 16 }}>
          <button className="btn btn-ghost" onClick={onBack}>← Back</button>
        </div>
      </div>
    </div>
  );

  return (
    <div className="page" style={{ justifyContent: 'flex-start', paddingTop: 40 }}>
      <div className="steps">
        <div className="step">
          <div className="step-num">1</div>
          <span>Select Files</span>
        </div>
        <div className="step-line" />
        <div className="step active">
          <div className="step-num">2</div>
          <span>Map Sheets</span>
        </div>
        <div className="step-line" />
        <div className="step">
          <div className="step-num">3</div>
          <span>Validate</span>
        </div>
      </div>

      <div className="card" style={{ maxWidth: 820 }}>
        <div className="card-icon">🗂️</div>
        <h1>Map Sheets</h1>
        <p className="hint">
          Select which sheet from each file maps to each required data slot.
          Use Preview to verify contents before confirming.
        </p>

        <table>
          <thead>
            <tr>
              <th>File</th>
              <th>Sheet</th>
              <th style={{ width: 100 }}>Preview</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.rowId}>
                <td>{row.fileLabel}</td>
                <td>
                  <select
                    value={row.selectedSheet}
                    onChange={(e) => handleSheetChange(row.rowId, e.target.value)}
                  >
                    {row.sheetOptions.map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </td>
                <td>
                  <button className="preview-btn" onClick={() => setPreviewRow(row)}>
                    Preview
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        <div className="btn-row" style={{ marginTop: 24 }}>
          <button className="btn btn-ghost" onClick={onBack}>← Back</button>
          <button className="btn btn-primary" onClick={handleConfirm}>
            Confirm & Proceed to Validation →
          </button>
        </div>
      </div>

      {previewRow && (
        <PreviewModal row={previewRow} onClose={() => setPreviewRow(null)} />
      )}
    </div>
  );
}