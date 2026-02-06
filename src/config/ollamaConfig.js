// Ollama Local LLM Configuration

export const ollamaConfig = {
  // Ollama server URL (default: http://localhost:11434)
  baseUrl: import.meta.env.VITE_OLLAMA_BASE_URL || 'http://localhost:11434',

  // Model to use (e.g., llama3, mistral, phi, gemma)
  model: import.meta.env.VITE_OLLAMA_MODEL || 'llama3',

  // Request timeout in milliseconds (3 minutes for first load)
  timeout: parseInt(import.meta.env.VITE_OLLAMA_TIMEOUT) || 180000,
};

// Check if Ollama is configured (always true since it uses defaults)
export function isOllamaConfigured() {
  return true;
}

// Check if Ollama server is running
export async function checkOllamaStatus() {
  try {
    const response = await fetch(`${ollamaConfig.baseUrl}/api/tags`, {
      method: 'GET',
      signal: AbortSignal.timeout(5000),
    });

    if (response.ok) {
      const data = await response.json();
      return {
        running: true,
        models: data.models || [],
      };
    }
    return { running: false, models: [] };
  } catch (error) {
    console.log('Ollama not running:', error.message);
    return { running: false, models: [] };
  }
}
