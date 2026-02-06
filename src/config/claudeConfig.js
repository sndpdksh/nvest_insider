// Claude Configuration

export const claudeConfig = {
  // Anthropic API Key
  apiKey: import.meta.env.VITE_CLAUDE_API_KEY || '',

  // Model name (e.g., claude-3-5-sonnet-20241022)
  model: import.meta.env.VITE_CLAUDE_MODEL || 'claude-3-5-sonnet-20241022',

  // API endpoint
  apiEndpoint: 'https://api.anthropic.com/v1/messages',
};

// Check if Claude is configured
export function isClaudeConfigured() {
  return !!(claudeConfig.apiKey && claudeConfig.model);
}
