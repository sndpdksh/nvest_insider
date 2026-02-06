import { useState, useEffect } from 'react'
import FileSearch from './components/FileSearch'
import ChatBot from './components/ChatBot'
import nvestLogo from './assets/nvest-logo.svg'
import './App.css'

function App() {
  const [activeView, setActiveView] = useState(() => {
    return localStorage.getItem('nvest-active-tab') || 'search';
  })
  const [isDarkMode, setIsDarkMode] = useState(() => {
    const saved = localStorage.getItem('nvest-theme');
    return saved !== null ? saved === 'dark' : true;
  })

  // Persist theme to localStorage
  useEffect(() => {
    localStorage.setItem('nvest-theme', isDarkMode ? 'dark' : 'light');
  }, [isDarkMode])

  // Persist active tab to localStorage
  useEffect(() => {
    localStorage.setItem('nvest-active-tab', activeView);
  }, [activeView])

  // Dynamic theme class based on isDarkMode state
  const getAppClass = () => {
    return isDarkMode ? 'dark-mode-app' : 'light-mode-app'
  }

  return (
    <div className={`app ${getAppClass()}`}>
      <nav className="app-nav">
        <div className="nav-brand">
          <img src={nvestLogo} alt="Nvest" className="nav-logo-image" />
          <span className="nav-app-name">Insider</span>
        </div>
        <div className="nav-tabs">
          <button
            className={`nav-tab ${activeView === 'search' ? 'active' : ''}`}
            onClick={() => setActiveView('search')}
          >
            📁 File Search
          </button>
          <button
            className={`nav-tab ${activeView === 'chat' ? 'active' : ''}`}
            onClick={() => setActiveView('chat')}
          >
            🤖 Chat Assistant
          </button>
        </div>
      </nav>

      <main className="app-content">
        <div style={{ display: activeView === 'search' ? 'contents' : 'none' }}>
          <FileSearch isDarkMode={isDarkMode} setIsDarkMode={setIsDarkMode} />
        </div>
        <div style={{ display: activeView === 'chat' ? 'contents' : 'none' }}>
          <ChatBot isDarkMode={isDarkMode} setIsDarkMode={setIsDarkMode} />
        </div>
      </main>
    </div>
  )
}

export default App
