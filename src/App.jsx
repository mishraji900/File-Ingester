import React, { useState } from 'react';
import UploadScreen from './components/UploadScreen.jsx';
import MappingScreen from './components/MappingScreen.jsx';

export default function App() {
  const [screen, setScreen] = useState('upload');
  const [ingestPayload, setIngestPayload] = useState(null);

  function goToMapping(payload) {
    setIngestPayload(payload);
    setScreen('mapping');
  }

  function goBackToUpload() {
    setScreen('upload');
    setIngestPayload(null);
  }

  return (
    <>
      <div className="topbar">
        <span className="topbar-logo">Deloitte.</span>
        <div className="topbar-divider"></div>
        <span className="topbar-app">Sheet Ingestor</span>
      </div>

      {screen === 'mapping' && ingestPayload
        ? <MappingScreen payload={ingestPayload} onBack={goBackToUpload} />
        : <UploadScreen onContinue={goToMapping} />
      }
    </>
  );
}