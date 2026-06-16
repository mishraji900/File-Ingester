import React, { useEffect, useState } from 'react';

export default function PreviewModal({ row, onClose }) {
  const [columns, setColumns] = useState([]);
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  useEffect(() => {
    async function load() {
      try {
        const result = await window.api.previewSheet(row.filePath, row.selectedSheet);
        if (result.error) throw new Error(result.error);
        setColumns(result.columns);
        setData(result.rows);
      } catch (err) {
        setError('Error loading preview: ' + err.message);
      } finally {
        setLoading(false);
      }
    }
    load();
  }, [row]);

  return (
    <div className="modal">
      <div className="modal-content">
        <span className="close" onClick={onClose}>&times;</span>
        <h3>{row.fileLabel} — {row.selectedSheet}</h3>

        {loading && <p>Loading...</p>}
        {error && <p className="status">{error}</p>}

        {!loading && !error && data.length === 0 && <p>No data found.</p>}

        {!loading && !error && data.length > 0 && (
          <table className="preview-table">
            <thead>
              <tr>
                {columns.map(c => <th key={c}>{c}</th>)}
              </tr>
            </thead>
            <tbody>
              {data.map((row, i) => (
                <tr key={i}>
                  {columns.map(c => <td key={c}>{row[c] ?? ''}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}