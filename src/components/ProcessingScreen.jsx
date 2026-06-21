// filename: src/components/ProcessingScreen.jsx
import React, { useEffect, useState } from 'react';

function guessTargetSheetName(row) {
  const label = (row.fileLabel || '').toLowerCase();
  if (label.includes('net trust assets')) return '2. Summary of Net Trust Asset';
  if (row.rowId === 'plan-1') return '3. Summary of Plan Ops';
  if (row.rowId === 'plan-2') return '4. Loan';
  return row.rowId || 'Sheet';
}

function parseLines(text) {
  return text
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean);
}

function parseCombineGroups(text) {
  // one per line: "Output Name: Input A, Input B"
  const result = {};
  for (const line of parseLines(text)) {
    const [outName, rest] = line.split(':');
    if (!outName || !rest) continue;
    result[outName.trim()] = rest.split(',').map(s => s.trim()).filter(Boolean);
  }
  return result;
}

function inputStyle(width = '100%') {
  return {
    width,
    padding: '9px 11px',
    borderRadius: 7,
    border: '1px solid var(--border)',
    background: 'var(--bg)',
    color: 'var(--fg)',
    fontSize: 13
  };
}

function Field({ label, children }) {
  return (
    <label style={{ display: 'flex', flexDirection: 'column', gap: 6, fontSize: 12, color: 'var(--fg-muted)' }}>
      <span style={{ fontWeight: 600, letterSpacing: 0.2 }}>{label}</span>
      {children}
    </label>
  );
}

const DEFAULT_GROUP_HEADERS = [
  'Contributions/Employer',
  'Contributions/Employee',
  'Adjustment (+)',
  'Adjustment (-)',
  'Administrative Fee',
  'Realized Gain/(Loss)',
  'Unrealized Gain/(Loss)',
  'Interest and Dividends',
  'Benefit Payments'
];
const DEFAULT_SECTION_GROUPS = ['Contributions/Employer', 'Contributions/Employee'];
const DEFAULT_ROLL_GROUPS = ['Contributions/Employee'];
const DEFAULT_COMBINE_GROUPS_TEXT =
  'Adjustments: Adjustment (+), Adjustment (-)\nRealized/Unrealized Gain Loss: Realized Gain/(Loss), Unrealized Gain/(Loss)';

export default function ProcessingScreen({ files, onBack }) {
  const [rows, setRows] = useState(() =>
    files.map(f => ({
      ...f,
      target_sheet_name: guessTargetSheetName(f),
      use_for_summary: false,
      index_col: f.columns?.[0]?.expected_header || '',
      sum_col: f.columns?.find(c => c.numeric)?.expected_header || ''
    }))
  );

  const [templatePath, setTemplatePath] = useState('');
  const [reportingDate, setReportingDate] = useState('');
  const [summaryEnabled, setSummaryEnabled] = useState(false);
  const [summarySheetName, setSummarySheetName] = useState('6. SOPO Summary');
  const [groupHeadersText, setGroupHeadersText] = useState(DEFAULT_GROUP_HEADERS.join('\n'));
  const [sectionGroupsText, setSectionGroupsText] = useState(DEFAULT_SECTION_GROUPS.join('\n'));
  const [rollGroupsText, setRollGroupsText] = useState(DEFAULT_ROLL_GROUPS.join('\n'));
  const [combineGroupsText, setCombineGroupsText] = useState(DEFAULT_COMBINE_GROUPS_TEXT);

  const [running, setRunning] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState(null);

  useEffect(() => {
    const today = new Date();
    setReportingDate(`${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`);
  }, []);

  function updateRow(rowId, key, value) {
    setRows(current => current.map(r => (r.rowId === rowId ? { ...r, [key]: value } : r)));
  }

  function selectSummaryRow(rowId) {
    setRows(current => current.map(r => ({ ...r, use_for_summary: r.rowId === rowId })));
  }

  async function handlePickTemplate() {
    const picked = await window.api.selectTemplateFile();
    if (picked) setTemplatePath(picked);
  }

  async function handleRun() {
    setError('');
    setResult(null);

    if (!templatePath) {
      setError('Select a template workbook first.');
      return;
    }

    const missingTarget = rows.find(r => !r.target_sheet_name.trim());
    if (missingTarget) {
      setError(`Row ${missingTarget.rowId} needs a target sheet name.`);
      return;
    }

    let summary = null;
    if (summaryEnabled) {
      const summaryRow = rows.find(r => r.use_for_summary);
      if (!summaryRow) {
        setError('Summary is enabled but no row is marked as the summary source.');
        return;
      }
      if (!summaryRow.index_col || !summaryRow.sum_col) {
        setError(`Row ${summaryRow.rowId}: pick both a description column and a numeric column for the summary.`);
        return;
      }
      summary = {
        sheet_name: summarySheetName,
        group_headers: parseLines(groupHeadersText),
        section_groups: parseLines(sectionGroupsText),
        roll_groups: parseLines(rollGroupsText),
        combine_groups: parseCombineGroups(combineGroupsText)
      };
    }

    setRunning(true);
    try {
      const payload = {
        template_path: templatePath,
        reporting_date: reportingDate,
        rows: rows.map(r => ({
          row_id: r.rowId,
          validated_path: r.validatedPath,
          target_sheet_name: r.target_sheet_name,
          use_for_summary: r.use_for_summary,
          index_col: r.index_col,
          sum_col: r.sum_col
        })),
        summary
      };

      const response = await window.api.runTrialBalance(payload);
      if (!response.success) {
        throw new Error(response.message || 'Trial balance generation failed.');
      }
      setResult(response);
    } catch (err) {
      setError(err.message || 'Trial balance generation failed.');
    } finally {
      setRunning(false);
    }
  }

  async function openOutput() {
    if (result?.outputPath) {
      await window.api.openReport(result.outputPath);
    }
  }

  return (
    <div className="page" style={{ justifyContent: 'flex-start', paddingTop: 40 }}>
      <div className="steps">
        <div className="step"><div className="step-num">1</div><span>Select Files</span></div>
        <div className="step-line" />
        <div className="step"><div className="step-num">2</div><span>Map Sheets</span></div>
        <div className="step-line" />
        <div className="step"><div className="step-num">3</div><span>Validate</span></div>
        <div className="step-line" />
        <div className="step active"><div className="step-num">4</div><span>Process</span></div>
      </div>

      <div className="container">
        <div className="card-icon">⚙️</div>
        <h1>Build Trial Balance</h1>
        <p className="hint">
          Validated sheets below were written to <code>converted/</code> after passing validation.
          Map each one to a destination tab in the template workbook, optionally pick one as the
          source for the hierarchical summary, then run.
        </p>

        <div style={{ display: 'flex', gap: 12, alignItems: 'end', flexWrap: 'wrap', marginTop: 18 }}>
          <Field label="Template workbook">
            <button type="button" className="btn btn-ghost" onClick={handlePickTemplate} style={{ minWidth: 220 }}>
              {templatePath ? 'Change Template' : 'Select Template Workbook *'}
            </button>
          </Field>

          {templatePath && (
            <span style={{ fontSize: 12, color: 'var(--primary)', maxWidth: 320, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {templatePath}
            </span>
          )}

          <Field label="Reporting date (1. Procedures!A2)">
            <input value={reportingDate} onChange={e => setReportingDate(e.target.value)} style={inputStyle(180)} />
          </Field>

          <button className="btn btn-ghost" onClick={onBack} disabled={running}>
            ← Back
          </button>

          <button className="btn btn-primary" onClick={handleRun} disabled={running}>
            {running ? 'Running...' : 'Run Trial Balance'}
          </button>

          {result?.outputPath && (
            <button className="btn btn-ghost" onClick={openOutput}>
              Open Output
            </button>
          )}
        </div>

        {/* per-row mapping */}
        {rows.map(row => (
          <div
            key={row.rowId}
            style={{ marginTop: 22, padding: 16, background: 'rgba(255,255,255,0.03)', border: '1px solid var(--border)', borderRadius: 12 }}
          >
            <h3 style={{ marginBottom: 10 }}>
              {row.rowId} — {row.fileName} — {row.sheetName}
            </h3>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12 }}>
              <Field label="Target sheet name (in template)">
                <input
                  value={row.target_sheet_name}
                  onChange={e => updateRow(row.rowId, 'target_sheet_name', e.target.value)}
                  style={inputStyle()}
                />
              </Field>

              <Field label="Description column (for summary)">
                <select
                  value={row.index_col}
                  onChange={e => updateRow(row.rowId, 'index_col', e.target.value)}
                  style={inputStyle()}
                  disabled={!summaryEnabled || !row.use_for_summary}
                >
                  <option value="">—</option>
                  {(row.columns || []).map(c => (
                    <option key={c.column} value={c.expected_header}>{c.expected_header || c.column}</option>
                  ))}
                </select>
              </Field>

              <Field label="Numeric column (for summary)">
                <select
                  value={row.sum_col}
                  onChange={e => updateRow(row.rowId, 'sum_col', e.target.value)}
                  style={inputStyle()}
                  disabled={!summaryEnabled || !row.use_for_summary}
                >
                  <option value="">—</option>
                  {(row.columns || []).filter(c => c.numeric).map(c => (
                    <option key={c.column} value={c.expected_header}>{c.expected_header || c.column}</option>
                  ))}
                </select>
              </Field>

              <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: 'var(--fg-muted)', alignSelf: 'end', paddingBottom: 8 }}>
                <input
                  type="radio"
                  name="summary-source-row"
                  checked={row.use_for_summary}
                  disabled={!summaryEnabled}
                  onChange={() => selectSummaryRow(row.rowId)}
                />
                Use as summary source
              </label>
            </div>
          </div>
        ))}

        {/* summary config */}
        <div style={{ marginTop: 24, padding: 16, background: 'rgba(255,255,255,0.03)', border: '1px solid var(--border)', borderRadius: 12 }}>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, marginBottom: 14 }}>
            <input type="checkbox" checked={summaryEnabled} onChange={e => setSummaryEnabled(e.target.checked)} />
            Generate hierarchical summary sheet
          </label>

          {summaryEnabled && (
            <>
              <Field label="Summary sheet name">
                <input value={summarySheetName} onChange={e => setSummarySheetName(e.target.value)} style={{ ...inputStyle(), marginBottom: 14 }} />
              </Field>

              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 14 }}>
                <Field label="Group headers (one per line)">
                  <textarea value={groupHeadersText} onChange={e => setGroupHeadersText(e.target.value)} rows={6} style={{ ...inputStyle(), resize: 'vertical' }} />
                </Field>
                <Field label="Section groups (one per line)">
                  <textarea value={sectionGroupsText} onChange={e => setSectionGroupsText(e.target.value)} rows={6} style={{ ...inputStyle(), resize: 'vertical' }} />
                </Field>
                <Field label="Roll groups (one per line)">
                  <textarea value={rollGroupsText} onChange={e => setRollGroupsText(e.target.value)} rows={6} style={{ ...inputStyle(), resize: 'vertical' }} />
                </Field>
                <Field label='Combine groups ("Output: A, B" per line)'>
                  <textarea value={combineGroupsText} onChange={e => setCombineGroupsText(e.target.value)} rows={6} style={{ ...inputStyle(), resize: 'vertical' }} />
                </Field>
              </div>
            </>
          )}
        </div>

        {error && <div className="status" style={{ marginTop: 20 }}>{error}</div>}

        {result && (
          <div style={{ marginTop: 24, padding: 16, background: 'rgba(255,255,255,0.03)', border: '1px solid var(--border)', borderRadius: 12 }}>
            <div className="status" style={{ marginBottom: 10 }}>
              Saved to: {result.outputPath}
            </div>
            {result.log?.map((line, idx) => (
              <div key={idx} style={{ fontSize: 12, color: 'var(--fg-muted)', marginBottom: 4 }}>{line}</div>
            ))}
            {result.summary && (
              <table style={{ marginTop: 14 }}>
                <thead>
                  <tr><th>Group</th><th>Total</th></tr>
                </thead>
                <tbody>
                  {Object.entries(result.summary).map(([k, v]) => (
                    <tr key={k}><td>{k}</td><td>{v}</td></tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        )}
      </div>
    </div>
  );
}