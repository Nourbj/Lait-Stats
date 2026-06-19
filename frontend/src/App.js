import { useState } from 'react';

const API = process.env.REACT_APP_API_URL || (process.env.NODE_ENV === 'development' ? 'http://localhost:5000/api' : '/api');

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} o`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} Ko`;
  return `${(bytes / 1024 / 1024).toFixed(1)} Mo`;
}

function App() {
  const [file, setFile] = useState(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [status, setStatus] = useState({
    type: '',
    message: 'Format attendu : colonne A = noms, colonnes suivantes = dates, valeurs 0 ou 1.'
  });

  const canDownload = Boolean(file) && !isDownloading;

  function selectFile(nextFile) {
    if (!nextFile) return;

    if (!nextFile.name.match(/\.xlsx?$/i)) {
      setFile(null);
      setStatus({ type: 'error', message: 'Veuillez sélectionner un fichier .xlsx ou .xls.' });
      return;
    }

    setFile(nextFile);
    setStatus({ type: 'success', message: 'Fichier chargé. Vous pouvez télécharger le rapport.' });
  }

  function clearFile() {
    setFile(null);
    setStatus({
      type: '',
      message: 'Format attendu : colonne A = noms, colonnes suivantes = dates, valeurs 0 ou 1.'
    });
  }

  async function downloadReport() {
    if (!canDownload) return;

    setIsDownloading(true);
    setStatus({ type: '', message: 'Génération du rapport en cours...' });

    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch(`${API}/download`, { method: 'POST', body: formData });
      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.error || 'Erreur lors de la génération.');
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `rapport_lait_${Date.now()}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);

      setStatus({ type: 'success', message: 'Rapport téléchargé avec succès.' });
    } catch (error) {
      setStatus({ type: 'error', message: error.message });
    } finally {
      setIsDownloading(false);
    }
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="logo" aria-label="LaitTrack">
          <div className="logo-icon" aria-hidden="true">L</div>
          Lait<span>Track</span>
        </div>
        <div className="header-badge">Import Excel</div>
      </header>

      <main className="main-content">
        <section className="workspace" aria-label="Génération de rapport lait">
          <div className="intro">
            <div className="eyebrow">Rapport automatique</div>
            <h1>Importer un fichier lait.</h1>
            <p>
              Ajoutez le fichier Excel source, puis téléchargez directement le rapport généré.
            </p>
          </div>

          <div className="panel">
            <label
              className={`drop-zone${isDragOver ? ' drag-over' : ''}`}
              onDragOver={(event) => {
                event.preventDefault();
                setIsDragOver(true);
              }}
              onDragLeave={() => setIsDragOver(false)}
              onDrop={(event) => {
                event.preventDefault();
                setIsDragOver(false);
                selectFile(event.dataTransfer.files?.[0]);
              }}
            >
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(event) => selectFile(event.target.files?.[0])}
              />
              <div className="drop-icon" aria-hidden="true">+</div>
              <div className="drop-title">Glissez votre fichier Excel ici</div>
              <div className="drop-sub">ou cliquez pour sélectionner un fichier</div>
              <div className="formats" aria-label="Formats acceptés">
                <span className="format-badge">.xlsx</span>
                <span className="format-badge">.xls</span>
              </div>
            </label>

            {file && (
              <div className="file-card">
                <div className="file-details">
                  <div className="file-name" title={file.name}>{file.name}</div>
                  <div className="file-meta">{formatSize(file.size)} - prêt pour génération</div>
                </div>
                <button className="btn btn-secondary" type="button" onClick={clearFile}>Changer</button>
              </div>
            )}

            <div className="actions">
              <button className="btn btn-primary" type="button" disabled={!canDownload} onClick={downloadReport}>
                {isDownloading && <span className="spinner" aria-hidden="true" />}
                {isDownloading ? 'Génération...' : 'Télécharger le rapport'}
              </button>
            </div>

            <div className={`status ${status.type}`.trim()}>{status.message}</div>
          </div>
        </section>
      </main>

      <footer>LaitTrack - Import et téléchargement du rapport Excel</footer>
    </div>
  );
}

export default App;
