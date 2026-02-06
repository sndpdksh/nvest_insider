import { useState, useEffect } from 'react';
import './FolderPicker.css';

function FolderPicker({ isOpen, onClose, onSelect, fetchFolders }) {
  const [folders, setFolders] = useState([]);
  const [breadcrumb, setBreadcrumb] = useState([]);
  const [loading, setLoading] = useState(false);
  const [selectedFolderId, setSelectedFolderId] = useState('root');
  const [selectedFolderName, setSelectedFolderName] = useState('My OneDrive (Root)');

  useEffect(() => {
    if (isOpen) {
      loadFolders(null);
      setBreadcrumb([]);
      setSelectedFolderId('root');
      setSelectedFolderName('My OneDrive (Root)');
    }
  }, [isOpen]);

  const loadFolders = async (folderId) => {
    setLoading(true);
    try {
      const items = await fetchFolders(folderId);
      setFolders(items);
    } catch (err) {
      console.error('Error loading folders:', err);
      setFolders([]);
    } finally {
      setLoading(false);
    }
  };

  const handleDrillDown = (folder) => {
    setBreadcrumb(prev => [...prev, { id: folder.id, name: folder.name }]);
    setSelectedFolderId(folder.id);
    setSelectedFolderName(folder.name);
    loadFolders(folder.id);
  };

  const handleBreadcrumbClick = (index) => {
    if (index === -1) {
      setBreadcrumb([]);
      setSelectedFolderId('root');
      setSelectedFolderName('My OneDrive (Root)');
      loadFolders(null);
    } else {
      const newBreadcrumb = breadcrumb.slice(0, index + 1);
      setBreadcrumb(newBreadcrumb);
      const target = newBreadcrumb[newBreadcrumb.length - 1];
      setSelectedFolderId(target.id);
      setSelectedFolderName(target.name);
      loadFolders(target.id);
    }
  };

  const handleConfirm = () => {
    onSelect({ id: selectedFolderId, name: selectedFolderName });
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div className="folder-picker-overlay" onClick={onClose}>
      <div className="folder-picker-modal" onClick={(e) => e.stopPropagation()}>
        <div className="folder-picker-header">
          <h3>Choose destination folder</h3>
          <button className="folder-picker-close" onClick={onClose}>&times;</button>
        </div>

        <div className="folder-picker-breadcrumb">
          <button onClick={() => handleBreadcrumbClick(-1)}>My OneDrive</button>
          {breadcrumb.map((crumb, idx) => (
            <span key={crumb.id}>
              <span className="breadcrumb-sep">/</span>
              <button onClick={() => handleBreadcrumbClick(idx)}>{crumb.name}</button>
            </span>
          ))}
        </div>

        <div className="folder-picker-list">
          {loading ? (
            <div className="folder-picker-loading">Loading folders...</div>
          ) : folders.length === 0 ? (
            <div className="folder-picker-empty">No subfolders here</div>
          ) : (
            folders.map((folder) => (
              <div
                key={folder.id}
                className={`folder-picker-item ${selectedFolderId === folder.id ? 'selected' : ''}`}
                onClick={() => {
                  setSelectedFolderId(folder.id);
                  setSelectedFolderName(folder.name);
                }}
                onDoubleClick={() => handleDrillDown(folder)}
              >
                <span className="folder-picker-icon">&#128193;</span>
                <span className="folder-picker-name">{folder.name}</span>
                {folder.hasChildren && (
                  <button
                    className="folder-picker-expand"
                    onClick={(e) => { e.stopPropagation(); handleDrillDown(folder); }}
                    title="Open folder"
                  >
                    &rsaquo;
                  </button>
                )}
              </div>
            ))
          )}
        </div>

        <div className="folder-picker-footer">
          <div className="folder-picker-selected">
            Upload to: <strong>{selectedFolderName}</strong>
          </div>
          <div className="folder-picker-actions">
            <button className="folder-picker-cancel" onClick={onClose}>Cancel</button>
            <button className="folder-picker-confirm" onClick={handleConfirm}>Select Folder</button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default FolderPicker;
