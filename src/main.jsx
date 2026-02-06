import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { PublicClientApplication, EventType } from '@azure/msal-browser'
import { MsalProvider } from '@azure/msal-react'
import { msalConfig } from './config/authConfig'
import './index.css'
import App from './App.jsx'

// Initialize MSAL instance
const msalInstance = new PublicClientApplication(msalConfig)

// Handle redirect response
msalInstance.initialize().then(() => {
  // Handle redirect promise
  msalInstance.handleRedirectPromise().then((response) => {
    if (response) {
      console.log('Login successful:', response.account?.username);
      msalInstance.setActiveAccount(response.account);
    }
  }).catch((error) => {
    console.error('Redirect error:', error);
  });

  // Set active account if available
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
  }

  // Listen for login events
  msalInstance.addEventCallback((event) => {
    if (event.eventType === EventType.LOGIN_SUCCESS && event.payload?.account) {
      msalInstance.setActiveAccount(event.payload.account);
    }
  });

  // Render app after MSAL is initialized
  createRoot(document.getElementById('root')).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </StrictMode>,
  )
});
