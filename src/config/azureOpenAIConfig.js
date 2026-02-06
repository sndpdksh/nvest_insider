// Azure OpenAI Service Configuration

export const azureOpenAIConfig = {
  // Azure OpenAI endpoint (e.g., https://your-resource-name.openai.azure.com)
  endpoint: import.meta.env.VITE_AZURE_OPENAI_ENDPOINT || '',

  // API Key for Azure OpenAI
  apiKey: import.meta.env.VITE_AZURE_OPENAI_API_KEY || '',

  // Deployment name for your model (e.g., gpt-4, gpt-35-turbo)
  deploymentName: import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT || 'gpt-35-turbo',

  // API Version
  apiVersion: import.meta.env.VITE_AZURE_OPENAI_API_VERSION || '2024-02-15-preview',
};

// Check if Azure OpenAI is configured
export function isAzureOpenAIConfigured() {
  return !!(azureOpenAIConfig.endpoint && azureOpenAIConfig.apiKey && azureOpenAIConfig.deploymentName);
}
