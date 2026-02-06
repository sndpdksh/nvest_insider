// Google Gemini API Configuration

export const geminiConfig = {
  // API Key from Google AI Studio (https://aistudio.google.com/app/apikey)
  apiKey: import.meta.env.VITE_GEMINI_API_KEY || '',

  // Model to use - use gemini-2.0-flash or gemini-1.5-flash-latest
  model: import.meta.env.VITE_GEMINI_MODEL || 'gemini-2.0-flash',

  // API endpoint - v1beta supports latest models
  baseUrl: 'https://generativelanguage.googleapis.com/v1beta',
};

// Check if Gemini is configured
export function isGeminiConfigured() {
  return !!geminiConfig.apiKey;
}
