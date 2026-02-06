import { useState, useEffect, useCallback } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest, teamsFilesRequest, getAdminConsentUrl } from '../config/authConfig';
import {
  getMyDriveFiles,
  getSharedWithMe,
  getRecentFiles,
  searchFiles,
  formatFileSize,
  getJoinedTeams,
  getTeamDriveFiles,
  getTeamFolderFiles,
  getParentFolderUrl,
} from '../services/graphService';
import './FileSearch.css';

const getFileIcon = (type, isFolder) => {
  if (isFolder) return '📁';
  switch (type) {
    case 'document': return '📄';
    case 'spreadsheet': return '📊';
    case 'pdf': return '📕';
    case 'presentation': return '📽️';
    case 'design': return '🎨';
    case 'image': return '🖼️';
    case 'video': return '🎬';
    case 'audio': return '🎵';
    case 'archive': return '📦';
    default: return '📄';
  }
};

function FileSearch({ isDarkMode = true, setIsDarkMode }) {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [searchTerm, setSearchTerm] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  const [filterType, setFilterType] = useState('All');
  const [files, setFiles] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  // Toggle theme
  const toggleTheme = () => {
    if (setIsDarkMode) {
      setIsDarkMode(prev => !prev);
    }
  };
  const [activeTab, setActiveTab] = useState('myfiles');
  const [isSearchMode, setIsSearchMode] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showMobileSidebar, setShowMobileSidebar] = useState(false);

  // Teams state
  const [teams, setTeams] = useState([]);
  const [selectedTeam, setSelectedTeam] = useState(null);
  const [teamsLoading, setTeamsLoading] = useState(false);
  const [breadcrumb, setBreadcrumb] = useState([]); // [{id, name}]
  const [needsAdminConsent, setNeedsAdminConsent] = useState(false);

  // Fetch teams list when teams tab is selected
  const fetchTeams = useCallback(async () => {
    if (!isAuthenticated || !accounts?.length) return;
    setTeamsLoading(true);
    try {
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });
      const fetchedTeams = await getJoinedTeams(tokenResponse.accessToken);
      setTeams(fetchedTeams);
    } catch (err) {
      console.error('Error fetching teams:', err);
      if (err.name === 'InteractionRequiredAuthError' ||
          err.message?.includes('AADSTS65001') ||
          err.message?.includes('consent')) {
        await instance.acquireTokenRedirect(loginRequest);
        return;
      }
      setError(`Could not load teams: ${err.message}`);
    } finally {
      setTeamsLoading(false);
    }
  }, [instance, accounts, isAuthenticated]);

  // Acquire token with Sites.Read.All via incremental consent
  const acquireTeamsToken = useCallback(async () => {
    try {
      // First try silent with the extra scope
      const tokenResponse = await instance.acquireTokenSilent({
        scopes: [...loginRequest.scopes, ...teamsFilesRequest.scopes],
        account: accounts[0],
      });
      setNeedsAdminConsent(false);
      return tokenResponse.accessToken;
    } catch (silentErr) {
      console.log('Silent token with Sites.Read.All failed:', silentErr.message);
      // If consent is needed, try interactive
      if (silentErr.name === 'InteractionRequiredAuthError') {
        try {
          const tokenResponse = await instance.acquireTokenPopup({
            scopes: [...loginRequest.scopes, ...teamsFilesRequest.scopes],
            account: accounts[0],
          });
          setNeedsAdminConsent(false);
          return tokenResponse.accessToken;
        } catch (popupErr) {
          console.error('Popup consent failed:', popupErr);
          // Admin consent required
          if (popupErr.message?.includes('AADSTS65001') ||
              popupErr.message?.includes('AADSTS70011') ||
              popupErr.message?.includes('admin') ||
              popupErr.message?.includes('consent') ||
              popupErr.errorCode === 'consent_required') {
            setNeedsAdminConsent(true);
            return null;
          }
          throw popupErr;
        }
      }
      // Check if it's specifically an admin consent issue
      if (silentErr.message?.includes('AADSTS65001') ||
          silentErr.message?.includes('admin') ||
          silentErr.message?.includes('consent')) {
        setNeedsAdminConsent(true);
        return null;
      }
      throw silentErr;
    }
  }, [instance, accounts]);

  // Fetch files from selected team's drive (root or folder)
  const fetchTeamFiles = useCallback(async (teamId, folderId = null, folderPath = '') => {
    if (!isAuthenticated || !accounts?.length) return;
    setLoading(true);
    setError(null);
    try {
      const accessToken = await acquireTeamsToken();
      if (!accessToken) {
        setLoading(false);
        return; // needsAdminConsent is already set
      }
      let fetchedFiles;
      if (folderId) {
        fetchedFiles = await getTeamFolderFiles(accessToken, teamId, folderId, folderPath);
      } else {
        fetchedFiles = await getTeamDriveFiles(accessToken, teamId);
      }
      setFiles(fetchedFiles);
    } catch (err) {
      console.error('Error fetching team files:', err);
      if (err.statusCode === 403 || err.code === 'Authorization_RequestDenied') {
        setNeedsAdminConsent(true);
      } else {
        setError(`Could not load team files: ${err.message}`);
      }
    } finally {
      setLoading(false);
    }
  }, [isAuthenticated, accounts, acquireTeamsToken]);

  const fetchFiles = useCallback(async () => {
    console.log('fetchFiles called, isAuthenticated:', isAuthenticated);
    console.log('accounts:', accounts);

    if (!isAuthenticated) {
      console.log('Not authenticated, skipping fetch');
      return;
    }

    if (!accounts || accounts.length === 0) {
      console.log('No accounts available');
      setError('No account found. Please sign in again.');
      return;
    }

    setLoading(true);
    setError(null);

    try {
      console.log('Acquiring token for account:', accounts[0]?.username);
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });
      console.log('Token acquired successfully');

      let fetchedFiles = [];

      switch (activeTab) {
        case 'myfiles':
          fetchedFiles = await getMyDriveFiles(tokenResponse.accessToken);
          break;
        case 'shared':
          fetchedFiles = await getSharedWithMe(tokenResponse.accessToken);
          break;
        case 'recent':
          fetchedFiles = await getRecentFiles(tokenResponse.accessToken);
          break;
        case 'teams':
          // Teams tab is handled separately
          return;
        default:
          fetchedFiles = await getMyDriveFiles(tokenResponse.accessToken);
      }

      setFiles(fetchedFiles);
    } catch (err) {
      console.error('Error fetching files:', err);
      if (err.name === 'InteractionRequiredAuthError' ||
          err.message?.includes('AADSTS65001') ||
          err.message?.includes('consent')) {
        // Use redirect for consent to avoid popup blocker issues
        await instance.acquireTokenRedirect(loginRequest);
        return;
      } else {
        // Show detailed error message
        const errorMsg = err.message || err.body?.error?.message || 'Unknown error';
        const errorCode = err.statusCode || err.code || '';
        setError(`Error ${errorCode}: ${errorMsg}`);
      }
    } finally {
      setLoading(false);
    }
  }, [instance, accounts, isAuthenticated, activeTab]);

  useEffect(() => {
    if (isAuthenticated && !isSearchMode) {
      if (activeTab === 'teams') {
        fetchTeams();
      } else {
        fetchFiles();
      }
    }
  }, [isAuthenticated, fetchFiles, fetchTeams, isSearchMode, activeTab]);

  // Handle team selection
  const handleSelectTeam = (team) => {
    setSelectedTeam(team);
    setBreadcrumb([]);
    fetchTeamFiles(team.id);
  };

  // Handle folder click in teams view
  const handleTeamFolderClick = (folder) => {
    if (!folder.isFolder || !selectedTeam) return;
    const newBreadcrumb = [...breadcrumb, { id: folder.id, name: folder.name }];
    setBreadcrumb(newBreadcrumb);
    const folderPath = newBreadcrumb.map(b => b.name).join('/');
    fetchTeamFiles(selectedTeam.id, folder.id, folderPath);
  };

  // Navigate breadcrumb
  const handleBreadcrumbClick = (index) => {
    if (!selectedTeam) return;
    if (index === -1) {
      // Clicked team name -> go to root
      setBreadcrumb([]);
      fetchTeamFiles(selectedTeam.id);
    } else {
      const newBreadcrumb = breadcrumb.slice(0, index + 1);
      setBreadcrumb(newBreadcrumb);
      const target = newBreadcrumb[newBreadcrumb.length - 1];
      const folderPath = newBreadcrumb.map(b => b.name).join('/');
      fetchTeamFiles(selectedTeam.id, target.id, folderPath);
    }
  };

  // Go back to teams list
  const handleBackToTeams = () => {
    setSelectedTeam(null);
    setBreadcrumb([]);
    setFiles([]);
  };

  // Search entire drive using Graph API
  const handleSearch = async () => {
    if (!searchQuery.trim()) {
      setError('Please enter a search term');
      return;
    }

    if (!accounts || accounts.length === 0) {
      setError('Please sign in first');
      return;
    }

    setLoading(true);
    setError(null);
    setIsSearchMode(true);

    try {
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      const results = await searchFiles(tokenResponse.accessToken, searchQuery.trim());
      setFiles(results);
    } catch (err) {
      console.error('Search error:', err);
      if (err.name === 'InteractionRequiredAuthError') {
        await instance.acquireTokenRedirect(loginRequest);
        return;
      }
      setError(`Search failed: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const clearSearch = () => {
    setSearchQuery('');
    setIsSearchMode(false);
    fetchFiles();
  };

  const handleLogin = async () => {
    try {
      // Use redirect instead of popup to avoid COOP issues
      await instance.loginRedirect(loginRequest);
    } catch (err) {
      console.error('Login failed:', err);
      setError('Login failed. Please try again.');
    }
  };

  const handleLogout = () => {
    instance.logoutRedirect();
  };

  // Get unique file types for filter
  const fileTypes = ['All', 'folder', 'document', 'spreadsheet', 'pdf', 'presentation', 'image', 'video', 'audio', 'archive', 'file'];

  // Filter files based on search term (searches name, path, and sharedBy)
  const filteredFiles = files.filter(file => {
    const searchLower = searchTerm.toLowerCase();
    const matchesSearch =
      file.name?.toLowerCase().includes(searchLower) ||
      file.path?.toLowerCase().includes(searchLower) ||
      file.parentPath?.toLowerCase().includes(searchLower) ||
      file.sharedBy?.toLowerCase().includes(searchLower);

    const matchesType = filterType === 'All' || file.type === filterType;

    return matchesSearch && matchesType;
  });

  const handleOpenFile = (webUrl) => {
    if (webUrl) {
      window.open(webUrl, '_blank');
    }
  };

  const openFileLocation = async (file) => {
    if (!file) return;
    console.log('Opening file location for:', file.name);

    try {
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      const folderUrl = await getParentFolderUrl(file.id, tokenResponse.accessToken, file.parentId);
      if (folderUrl) {
        console.log('Opening folder:', folderUrl);
        window.open(folderUrl, '_blank');
        return;
      }
    } catch (err) {
      console.error('Error getting parent folder:', err);
    }

    // Fallback: Open OneDrive root
    if (file.webUrl) {
      try {
        const url = new URL(file.webUrl);
        const match = url.pathname.match(/^(\/personal\/[^/]+|\/sites\/[^/]+)/);
        if (match) {
          const siteUrl = `${url.origin}${match[1]}/_layouts/15/onedrive.aspx`;
          window.open(siteUrl, '_blank');
        }
      } catch (e) {
        console.error('URL parse error:', e);
      }
    }
  };

  if (!isAuthenticated) {
    return (
      <div className="file-search-wrapper dark-mode">
        <div className="login-screen-dark">
          <div className="login-card-dark">
            <div className="login-icon">📁</div>
            <h1>Teams Shared Files</h1>
            <p>Sign in with your Microsoft account to access shared files</p>
            <button className="login-btn-dark" onClick={handleLogin}>
              Sign in with Microsoft
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className={`file-search-wrapper ${isDarkMode ? 'dark-mode' : 'light-mode'}`}>
      {/* Mobile Sidebar Overlay */}
      {showMobileSidebar && <div className="mobile-sidebar-overlay" onClick={() => setShowMobileSidebar(false)} />}

      {/* Sidebar */}
      <aside className={`fs-sidebar ${showMobileSidebar ? 'mobile-open' : ''}`}>
        <div className="fs-sidebar-tabs">
          <button
            className={`fs-tab-btn ${activeTab === 'myfiles' ? 'active' : ''}`}
            onClick={() => { setActiveTab('myfiles'); setIsSearchMode(false); setShowMobileSidebar(false); }}
          >
            📂 My Files
          </button>
          <button
            className={`fs-tab-btn ${activeTab === 'recent' ? 'active' : ''}`}
            onClick={() => { setActiveTab('recent'); setIsSearchMode(false); setShowMobileSidebar(false); }}
          >
            🕐 Recent
          </button>
          <button
            className={`fs-tab-btn ${activeTab === 'shared' ? 'active' : ''}`}
            onClick={() => { setActiveTab('shared'); setIsSearchMode(false); setShowMobileSidebar(false); }}
          >
            👥 Shared
          </button>
          <button
            className={`fs-tab-btn ${activeTab === 'teams' ? 'active' : ''}`}
            onClick={() => { setActiveTab('teams'); setIsSearchMode(false); setSelectedTeam(null); setBreadcrumb([]); setNeedsAdminConsent(false); setShowMobileSidebar(false); }}
          >
            🏢 Teams
          </button>
          <button className="mobile-sidebar-close" onClick={() => setShowMobileSidebar(false)}>✕</button>
        </div>

        <div className="fs-sidebar-filters">
          <label>FILTER BY TYPE</label>
          <select
            value={filterType}
            onChange={(e) => setFilterType(e.target.value)}
            className="fs-filter-select"
          >
            {fileTypes.map(type => (
              <option key={type} value={type}>
                {type.charAt(0).toUpperCase() + type.slice(1)}
              </option>
            ))}
          </select>
        </div>

        <div className="fs-sidebar-footer">
          <div className="fs-settings-btn">
            <span className="fs-user-avatar">👤</span>
            <span className="fs-user-name">{accounts[0]?.name || accounts[0]?.username}</span>
          </div>
          
          <button className="fs-logout-btn" onClick={handleLogout}>
            <span>⬅️</span> Sign out
          </button>
          <button className="fs-theme-toggle" onClick={toggleTheme}>
            <span>{isDarkMode ? '☀️' : '🌙'}</span>
            {isDarkMode ? 'Light Mode' : 'Dark Mode'}
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="fs-main">
        <div className="fs-header">
          <button className="mobile-menu-btn" onClick={() => setShowMobileSidebar(true)}>
            <span>☰</span>
          </button>
          <div className="fs-header-left">
            <h2>{isSearchMode ? 'Search Results' : activeTab === 'myfiles' ? 'My Files' : activeTab === 'recent' ? 'Recent Files' : activeTab === 'teams' ? (selectedTeam ? selectedTeam.name : 'Teams') : 'Shared with Me'}</h2>
            <span className="fs-file-count">{filteredFiles.length} items</span>
          </div>
          <button className="fs-refresh-btn" onClick={fetchFiles} disabled={loading}>
            {loading ? '...' : '↻'}
          </button>
        </div>

        {/* Search Bar */}
        <div className="fs-search-section">
          <div className="fs-search-wrapper">
            <svg className="fs-search-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="11" cy="11" r="8"/>
              <path d="m21 21-4.35-4.35"/>
            </svg>
            <input
              type="text"
              placeholder="Search files and folders..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && handleSearch()}
              className="fs-search-input"
            />
            {isSearchMode ? (
              <button className="fs-clear-btn" onClick={clearSearch}>Clear</button>
            ) : (
              <button
                className="fs-search-btn"
                onClick={handleSearch}
                disabled={loading || !searchQuery.trim()}
              >
                Search
              </button>
            )}
          </div>
          {!isSearchMode && (
            <input
              type="text"
              placeholder="Filter current results..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="fs-filter-input"
            />
          )}
        </div>

        {isSearchMode && (
          <div className="fs-search-indicator">
            Showing results for "<strong>{searchQuery}</strong>"
          </div>
        )}

        {error && (
          <div className="fs-error">
            {error}
            <button onClick={() => setError(null)}>✕</button>
          </div>
        )}

        {/* Teams: Admin consent required banner */}
        {activeTab === 'teams' && needsAdminConsent && (
          <div className="fs-admin-consent-banner">
            <div className="fs-consent-icon">🔒</div>
            <div className="fs-consent-info">
              <h3>Admin Consent Required</h3>
              <p>
                Accessing Teams group documents requires the <strong>Sites.Read.All</strong> permission,
                which must be granted by your tenant admin.
              </p>
              <p className="fs-consent-help">Share this link with your admin to grant access:</p>
              <div className="fs-consent-url-row">
                <input
                  type="text"
                  readOnly
                  value={getAdminConsentUrl()}
                  className="fs-consent-url"
                  onClick={(e) => e.target.select()}
                />
                <button
                  className="fs-consent-copy-btn"
                  onClick={() => {
                    navigator.clipboard.writeText(getAdminConsentUrl());
                    setError(null);
                  }}
                >
                  Copy
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Teams: Team selector (when no team selected) */}
        {activeTab === 'teams' && !selectedTeam && !isSearchMode && !needsAdminConsent && (
          <div className="fs-teams-list">
            {teamsLoading ? (
              <div className="fs-loading">
                <div className="fs-spinner"></div>
                <p>Loading teams...</p>
              </div>
            ) : teams.length === 0 ? (
              <div className="fs-no-results">
                <span className="fs-no-results-icon">🏢</span>
                <p>No teams found</p>
                <span>You may not be a member of any teams</span>
              </div>
            ) : (
              teams.map(team => (
                <div
                  key={team.id}
                  className="fs-team-card"
                  onClick={() => handleSelectTeam(team)}
                >
                  <div className="fs-team-icon">🏢</div>
                  <div className="fs-team-info">
                    <h3 className="fs-team-name">{team.name}</h3>
                    {team.description && <p className="fs-team-desc">{team.description}</p>}
                  </div>
                  <span className="fs-team-arrow">›</span>
                </div>
              ))
            )}
          </div>
        )}

        {/* Teams: Breadcrumb navigation */}
        {activeTab === 'teams' && selectedTeam && !isSearchMode && (
          <div className="fs-breadcrumb">
            <button className="fs-breadcrumb-back" onClick={handleBackToTeams}>
              ← Teams
            </button>
            <span className="fs-breadcrumb-sep">/</span>
            <button
              className={`fs-breadcrumb-item ${breadcrumb.length === 0 ? 'active' : ''}`}
              onClick={() => handleBreadcrumbClick(-1)}
            >
              {selectedTeam.name}
            </button>
            {breadcrumb.map((crumb, idx) => (
              <span key={crumb.id}>
                <span className="fs-breadcrumb-sep">/</span>
                <button
                  className={`fs-breadcrumb-item ${idx === breadcrumb.length - 1 ? 'active' : ''}`}
                  onClick={() => handleBreadcrumbClick(idx)}
                >
                  {crumb.name}
                </button>
              </span>
            ))}
          </div>
        )}

        {/* File List (shown for all tabs, and for teams when a team is selected) */}
        {(activeTab !== 'teams' || selectedTeam) && (
        <div className="fs-file-list">
          {loading ? (
            <div className="fs-loading">
              <div className="fs-spinner"></div>
              <p>Loading files...</p>
            </div>
          ) : filteredFiles.length === 0 ? (
            <div className="fs-no-results">
              <span className="fs-no-results-icon">📭</span>
              <p>No files found</p>
              <span>Try a different search term or filter</span>
            </div>
          ) : (
            filteredFiles.map(file => (
              <div
                key={file.id}
                className={`fs-file-card ${file.isFolder ? 'fs-folder-card' : ''}`}
                onClick={() => {
                  if (file.isFolder && activeTab === 'teams') {
                    handleTeamFolderClick(file);
                  }
                }}
                style={file.isFolder && activeTab === 'teams' ? { cursor: 'pointer' } : {}}
              >
                <div className="fs-file-icon">{getFileIcon(file.type, file.isFolder)}</div>
                <div className="fs-file-details">
                  <h3 className="fs-file-name">{file.name}</h3>
                  <div className="fs-file-path">
                    <span>📂 {file.path || file.name}</span>
                  </div>
                  <div className="fs-file-meta">
                    <span className="fs-meta-item">
                      <span className="fs-meta-label">Modified:</span> {file.date || 'N/A'}
                    </span>
                    <span className="fs-meta-item">
                      <span className="fs-meta-label">Owner:</span> {file.sharedBy || 'Unknown'}
                    </span>
                    {file.size && (
                      <span className="fs-meta-item">
                        <span className="fs-meta-label">Size:</span> {formatFileSize(file.size)}
                      </span>
                    )}
                    <span className="fs-type-badge">{file.isFolder ? 'Folder' : file.type}</span>
                  </div>
                </div>
                <div className="fs-file-actions">
                  <button
                    className="fs-open-btn"
                    onClick={(e) => { e.stopPropagation(); handleOpenFile(file.webUrl); }}
                  >
                    Open
                  </button>
                  <button
                    className="fs-location-btn"
                    onClick={(e) => { e.stopPropagation(); openFileLocation(file); }}
                    title="Open file location"
                  >
                    📁
                  </button>
                </div>
              </div>
            ))
          )}
        </div>
        )}
      </main>

      {/* Settings Modal */}
      {showSettings && (
        <div className="fs-settings-overlay" onClick={() => setShowSettings(false)}>
          <div className="fs-settings-modal" onClick={(e) => e.stopPropagation()}>
            <div className="fs-settings-header">
              <h2>Settings</h2>
              <button className="fs-settings-close" onClick={() => setShowSettings(false)}>
                ✕
              </button>
            </div>
            <div className="fs-settings-content">
              <div className="fs-settings-section">
                <h3>Appearance</h3>
                <p className="fs-settings-description">Customize the look and feel</p>
                <button className="fs-settings-theme-btn" onClick={toggleTheme}>
                  <span>{isDarkMode ? '☀️' : '🌙'}</span>
                  {isDarkMode ? 'Switch to Light Mode' : 'Switch to Dark Mode'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default FileSearch;
