// Microsoft Authentication Library (MSAL) Configuration

export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_CLIENT_ID || 'YOUR_CLIENT_ID_HERE',
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_TENANT_ID || 'common'}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

// Scopes that DO NOT require admin consent (used at login)
export const loginRequest = {
  scopes: [
    'openid',
    'profile',
    'User.Read',
    'Files.Read',           // User's own OneDrive files
    'Files.ReadWrite',      // Read/write user's files
    'Team.ReadBasic.All',   // List joined teams
  ],
};

// Additional scopes for Teams group document access (requires admin consent)
// Requested incrementally only when user clicks the Teams tab
export const teamsFilesRequest = {
  scopes: [
    'Sites.Read.All',       // Read group/team SharePoint document libraries
  ],
};

// Build admin consent URL for tenant admin to grant Sites.Read.All
export function getAdminConsentUrl() {
  const clientId = import.meta.env.VITE_AZURE_CLIENT_ID || '';
  const tenantId = import.meta.env.VITE_AZURE_TENANT_ID || 'common';
  return `https://login.microsoftonline.com/${tenantId}/adminconsent?client_id=${clientId}&redirect_uri=${encodeURIComponent(window.location.origin)}`;
}

// Graph API endpoints
export const graphConfig = {
  graphMeEndpoint: 'https://graph.microsoft.com/v1.0/me',
  graphDriveEndpoint: 'https://graph.microsoft.com/v1.0/me/drive',
};
