// Groq API Configuration - Fast LLM Inference

export const groqConfig = {
  // API Key from Groq Console (https://console.groq.com/keys)
  apiKey: import.meta.env.VITE_GROQ_API_KEY || '',

  // Model to use - Smaller, stable model less likely to be deprecated
  // llama-3.1-8b-instant is fast and reliable
  model: import.meta.env.VITE_GROQ_MODEL || 'llama-3.1-8b-instant',

  // API endpoint
  baseUrl: 'https://api.groq.com/openai/v1',
};

  // API endpoint
export function isGroqConfigured() {
  return !!groqConfig.apiKey;
}
