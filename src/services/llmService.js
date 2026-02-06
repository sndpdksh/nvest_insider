// Unified LLM Service - Supports Azure OpenAI, Groq, Google Gemini, Claude, and Ollama
import { isAzureOpenAIConfigured } from '../config/azureOpenAIConfig';
import { isGroqConfigured } from '../config/groqConfig';
import { isGeminiConfigured } from '../config/geminiConfig';
import { isClaudeConfigured } from '../config/claudeConfig';
import { checkOllamaStatus } from '../config/ollamaConfig';
import {
  generateDocumentSummary as azureSummary,
  answerDocumentQuestion as azureAnswer,
  getChatResponse as azureChat,
} from './azureOpenAIService';
import {
  generateDocumentSummaryGemini,
  answerDocumentQuestionGemini,
  getChatResponseGemini,
} from './geminiService';
import {
  generateDocumentSummaryGroq,
  answerDocumentQuestionGroq,
  getChatResponseGroq,
} from './groqService';
import {
  generateDocumentSummaryClaude,
  answerDocumentQuestionClaude,
  getChatResponseClaude,
} from './claudeService';
import {
  generateDocumentSummaryOllama,
  answerDocumentQuestionOllama,
  getChatResponseOllama,
} from './ollamaService';

// // LLM Provider types
// exCLAUDE: 'claude',
//   port const LLM_PROVIDERS = {
//   NONE: 'none',
//   AZURE_OPENAI: 'azure-openai',
//   GROQ: 'groq',
//   GEMINI: 'gemini',
//   OLLAMA: 'ollama',
// };
// LLM Provider types
export const LLM_PROVIDERS = {
  NONE: 'none',
  CLAUDE: 'claude',  
  AZURE_OPENAI: 'azure-openai',
  GROQ: 'groq',
  GEMINI: 'gemini',
  OLLAMA: 'ollama',
};

// Current active provider (will be set during initialization)
let activeProvider = LLM_PROVIDERS.NONE;
let ollamaStatus = { running: false, models: [] };
let availableProviders = []; // List of configured providers

/**
 * Initialize and detect available LLM providers
 * @returns {Promise<Object>} - Active provider and list of available providers
 */
export async function initializeLLM() {
  // Reset available providers list
  const providers = [];

  // Check Claude (highest priority - though has CORS issues from browser)
  if (isClaudeConfigured()) {
    providers.push({
      id: LLM_PROVIDERS.CLAUDE,
      name: 'Claude 3.5 Sonnet',
      icon: '🤖'
    });
  }

  // Check Gemini (stable, reliable, no deprecation issues)
  if (isGeminiConfigured()) {
    providers.push({
      id: LLM_PROVIDERS.GEMINI,
      name: 'Google Gemini',
      icon: '✨'
    });
  }

  // Check Azure OpenAI
  if (isAzureOpenAIConfigured()) {
    providers.push({
      id: LLM_PROVIDERS.AZURE_OPENAI,
      name: 'Azure OpenAI',
      icon: '☁️'
    });
  }

  // Check Groq (fallback - smaller model to avoid deprecation)
  if (isGroqConfigured()) {
    providers.push({
      id: LLM_PROVIDERS.GROQ,
      name: 'Groq (Llama 3.1 8B)',
      icon: '⚡'
    });
  }

  // Check Ollama
  ollamaStatus = await checkOllamaStatus();
  if (ollamaStatus.running) {
    providers.push({
      id: LLM_PROVIDERS.OLLAMA,
      name: 'Ollama (Local)',
      icon: '🦙'
    });
  }

  // Update the module-level variable
  availableProviders = providers;

  // Set default active provider (first available) only if not already set
  if (activeProvider === LLM_PROVIDERS.NONE && providers.length > 0) {
    activeProvider = providers[0].id;
  }

  console.log('Available LLM Providers:', availableProviders.map(p => p.name));
  console.log('Active LLM Provider:', activeProvider);

  return { activeProvider, availableProviders: [...providers] };
}

/**
 * Get list of available providers
 * @returns {Array}
 */
export function getAvailableProviders() {
  return availableProviders;
}

/**
 * Manually set the active LLM provider
 * @param {string} providerId - Provider ID to set
 * @returns {boolean} - Success status
 */
export function setLLMProvider(providerId) {
  const provider = availableProviders.find(p => p.id === providerId);
  if (provider) {
    activeProvider = providerId;
    console.log('LLM Provider changed to:', provider.name);
    return true;
  }
  console.warn('Provider not available:', providerId);
  return false;
}

/**
 * Get current LLM provider
 * @returns {string} - Current provider
 */
export function getLLMProvider() {
  return activeProvider;
}

/**
 * Check if any LLM is available
 * @returns {boolean}
 */
export function isLLMAvailable() {
  return activeProvider !== LLM_PROVIDERS.NONE;
}

/**
 * Get provider display name
 * @returns {string}
 */
export function getProviderName() {
  switch (activeProvider) {
    case LLM_PROVIDERS.AZURE_OPENAI:
      return 'Azure OpenAI';
    case LLM_PROVIDERS.GROQ:
      return 'Groq (Llama 3.1)';
    case LLM_PROVIDERS.GEMINI:
      return 'Google Gemini';
    case LLM_PROVIDERS.OLLAMA:
      return 'Ollama (Local)';
    case LLM_PROVIDERS.CLAUDE:
      return 'Claude 3.5 Sonnet';
    default:
      return 'None';
  }
}

/**
 * Get available Ollama models
 * @returns {Array}
 */
export function getOllamaModels() {
  return ollamaStatus.models;
}

/**
 * Generate document summary using available LLM
 * @param {string} content - Document content
 * @param {string} fileName - File name
 * @returns {Promise<string|null>}
 */
export async function generateSummary(content, fileName) {
  if (!content || content.trim().length === 0) {
    return null;
  }

  try {
    switch (activeProvider) {
      case LLM_PROVIDERS.CLAUDE:
        return await generateDocumentSummaryClaude(content, fileName);
      case LLM_PROVIDERS.AZURE_OPENAI:
        return await azureSummary(content, fileName);
      case LLM_PROVIDERS.GROQ:
        return await generateDocumentSummaryGroq(content, fileName);
      case LLM_PROVIDERS.GEMINI:
        return await generateDocumentSummaryGemini(content, fileName);
      case LLM_PROVIDERS.OLLAMA:
        return await generateDocumentSummaryOllama(content, fileName);
      default:
        return null;
    }
  } catch (error) {
    console.error('LLM summary failed:', error);
    return null;
  }
}

/**
 * Answer question about document using available LLM
 * @param {string} content - Document content
 * @param {string} question - User question
 * @param {string} fileName - File name
 * @returns {Promise<string>}
 */
export async function answerQuestion(content, question, fileName) {
  try {
    switch (activeProvider) {
      case LLM_PROVIDERS.CLAUDE:
        return await answerDocumentQuestionClaude(content, question, fileName);
      case LLM_PROVIDERS.AZURE_OPENAI:
        return await azureAnswer(content, question, fileName);
      case LLM_PROVIDERS.GROQ:
        return await answerDocumentQuestionGroq(content, question, fileName);
      case LLM_PROVIDERS.GEMINI:
        return await answerDocumentQuestionGemini(content, question, fileName);
      case LLM_PROVIDERS.OLLAMA:
        return await answerDocumentQuestionOllama(content, question, fileName);
      default:
        return "No AI service is configured. Please set up Claude, Groq, Gemini, or Ollama.";
    }
  } catch (error) {
    console.error('LLM answer failed:', error);
    return "I encountered an error. Please try again.";
  }
}

/**
 * Get chat response using available LLM
 * @param {string} message - User message
 * @param {Array} history - Conversation history
 * @returns {Promise<string>}
 */
export async function getChatResponse(message, history = []) {
  try {
    switch (activeProvider) {
      case LLM_PROVIDERS.CLAUDE:
        return await getChatResponseClaude(message, history);
      case LLM_PROVIDERS.AZURE_OPENAI:
        return await azureChat(message, history);
      case LLM_PROVIDERS.GROQ:
        return await getChatResponseGroq(message, history);
      case LLM_PROVIDERS.GEMINI:
        return await getChatResponseGemini(message, history);
      case LLM_PROVIDERS.OLLAMA:
        return await getChatResponseOllama(message, history);
      default:
        return "No AI service is configured. I can still help you search for files.";
    }
  } catch (error) {
    console.error('LLM chat failed:', error);
    return "I'm having trouble processing your request.";
  }
}

/**
 * Re-check LLM availability (useful after config changes)
 * @returns {Promise<string>}
 */
export async function recheckLLM() {
  return await initializeLLM();
}
