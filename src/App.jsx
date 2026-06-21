// filename: src/App.jsx
import React, { useState } from 'react';
import UploadScreen from './components/UploadScreen.jsx';
import MappingScreen from './components/MappingScreen.jsx';
import ValidationScreen from './components/ValidationScreen.jsx';
import ProcessingScreen from './components/ProcessingScreen.jsx';

export default function App() {
  const [screen, setScreen] = useState('upload');
  const [ingestPayload, setIngestPayload] = useState(null);
  const [mappedRows, setMappedRows] = useState([]);
  const [validatedFiles, setValidatedFiles] = useState([]);

  function goToMapping(payload) {
    setIngestPayload(payload);
    setMappedRows([]);
    setScreen('mapping');
  }

  function goToValidation(rows) {
    setMappedRows(rows);
    setScreen('validation');
  }

  function goToProcessing(files) {
    setValidatedFiles(files);
    setScreen('processing');
  }

  function goBackToUpload() {
    setScreen('upload');
    setIngestPayload(null);
    setMappedRows([]);
    setValidatedFiles([]);
  }

  function goBackToMapping() {
    setScreen('mapping');
  }

  function goBackToValidation() {
    setScreen('validation');
  }

  return (
    <>
      <div className="topbar">
        <span className="topbar-logo">Deloitte.</span>
        <div className="topbar-divider" />
        <span className="topbar-app">Sheet Ingestor</span>
      </div>

      {screen === 'processing' && validatedFiles.length > 0 ? (
        <ProcessingScreen files={validatedFiles} onBack={goBackToValidation} />
      ) : screen === 'validation' && mappedRows.length > 0 ? (
        <ValidationScreen
          rows={mappedRows}
          onBack={goBackToMapping}
          onContinue={goToProcessing}
        />
      ) : screen === 'mapping' && ingestPayload ? (
        <MappingScreen
          payload={ingestPayload}
          onBack={goBackToUpload}
          onContinue={goToValidation}
        />
      ) : (
        <UploadScreen onContinue={goToMapping} />
      )}
    </>
  );
}