// Ollama Local LLM Service for document processing
import { ollamaConfig } from '../config/ollamaConfig';

/**
 * Call Ollama API for chat completion
 * @param {string} prompt - The prompt to send
 * @param {string} systemPrompt - System/context prompt
 * @param {Object} options - Optional parameters
 * @returns {Promise<string>} - The AI response text
 */
export async function callOllama(prompt, systemPrompt = '', options = {}) {
  const { baseUrl, model } = ollamaConfig;
  const url = `${baseUrl}/api/generate`;

  const fullPrompt = systemPrompt
    ? `${systemPrompt}\n\nUser: ${prompt}\n\nAssistant:`
    : prompt;

  const requestBody = {
    model: options.model || model,
    prompt: fullPrompt,
    stream: true, // Use streaming to avoid timeout
    options: {
      temperature: options.temperature || 0.7,
      num_predict: options.maxTokens || 500, // Reduced for faster response
      top_p: options.topP || 0.9,
    },
  };

  try {
    console.log('Ollama: Calling', url, 'with model:', requestBody.model);

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Ollama API error: ${response.status} - ${errorText}`);
    }

    // Handle streaming response
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let fullResponse = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      const chunk = decoder.decode(value);
      const lines = chunk.split('\n').filter(line => line.trim());

      for (const line of lines) {
        try {
          const json = JSON.parse(line);
          if (json.response) {
            fullResponse += json.response;
          }
          if (json.done) {
            console.log('Ollama: Response complete, length:', fullResponse.length);
            return fullResponse;
          }
        } catch (e) {
          // Skip invalid JSON lines
        }
      }
    }

    return fullResponse;
  } catch (error) {
    console.error('Ollama API call failed:', error);
    throw error;
  }
}

/**
 * Call Ollama Chat API (for models that support chat format)
 * @param {Array} messages - Array of message objects with role and content
 * @param {Object} options - Optional parameters
 * @returns {Promise<string>} - The AI response text
 */
export async function callOllamaChat(messages, options = {}) {
  const { baseUrl, model, timeout } = ollamaConfig;
  const url = `${baseUrl}/api/chat`;

  const requestBody = {
    model: options.model || model,
    messages: messages,
    stream: false,
    options: {
      temperature: options.temperature || 0.7,
      num_predict: options.maxTokens || 1000,
    },
  };

  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeout);

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Ollama Chat API error: ${response.status} - ${errorText}`);
    }

    const data = await response.json();
    return data.message?.content || '';
  } catch (error) {
    if (error.name === 'AbortError') {
      throw new Error('Ollama request timed out.');
    }
    console.error('Ollama Chat API call failed:', error);
    throw error;
  }
}

/**
 * Generate an intelligent summary of document content
 * @param {string} content - The document content to summarize
 * @param {string} fileName - The name of the document
 * @returns {Promise<string>} - The AI-generated summary
 */
export async function generateDocumentSummaryOllama(content, fileName) {
  if (!content || content.trim().length === 0) {
    return null;
  }

  // Truncate content if too long (Ollama models have context limits)
  const maxContentLength = 4000;
  const truncatedContent = content.length > maxContentLength
    ? content.substring(0, maxContentLength) + '...[truncated]'
    : content;

  const systemPrompt = `You are a helpful document assistant. Provide clear, concise summaries of documents.
Focus on: main topics, key points, important facts/dates, and document purpose.
Keep summaries to 3-5 sentences.`;

  const prompt = `Please summarize this document "${fileName}":\n\n${truncatedContent}`;

  try {
    const summary = await callOllama(prompt, systemPrompt, {
      maxTokens: 500,
      temperature: 0.5,
    });
    return summary.trim();
  } catch (error) {
    console.error('Failed to generate document summary with Ollama:', error);
    return null;
  }
}

/**
 * Answer a question about document content
 * @param {string} content - The document content
 * @param {string} question - The user's question
 * @param {string} fileName - The name of the document
 * @returns {Promise<string>} - The AI-generated answer
 */
export async function answerDocumentQuestionOllama(content, question, fileName) {
  if (!content || content.trim().length === 0) {
    return "I couldn't read the document content to answer your question.";
  }

  const maxContentLength = 4000;
  const truncatedContent = content.length > maxContentLength
    ? content.substring(0, maxContentLength) + '...[truncated]'
    : content;

  const systemPrompt = `You are a helpful document assistant. Answer questions based on the provided document content.
If the answer cannot be found in the document, say so clearly.
Be accurate and cite specific information from the document.`;

  const prompt = `Document: "${fileName}"\n\nContent:\n${truncatedContent}\n\nQuestion: ${question}`;

  try {
    const answer = await callOllama(prompt, systemPrompt, {
      maxTokens: 800,
      temperature: 0.3,
    });
    return answer.trim();
  } catch (error) {
    console.error('Failed to answer document question with Ollama:', error);
    return "I encountered an error processing your question. Please try again.";
  }
}

/**
 * General chat response
 * @param {string} userMessage - The user's message
 * @param {Array} conversationHistory - Previous messages
 * @returns {Promise<string>} - The AI response
 */
export async function getChatResponseOllama(userMessage, conversationHistory = []) {
  const systemPrompt = `You are a helpful document assistant for a corporate environment. You help users:
- Find and understand documents
- Answer questions about company policies
- Provide information about processes and procedures

Be professional, helpful, and concise.`;

  // Build context from history
  let contextPrompt = systemPrompt;
  if (conversationHistory.length > 0) {
    const recentHistory = conversationHistory.slice(-6);
    const historyText = recentHistory.map(msg =>
      `${msg.role === 'user' ? 'User' : 'Assistant'}: ${msg.content}`
    ).join('\n');
    contextPrompt += `\n\nRecent conversation:\n${historyText}`;
  }

  try {
    const response = await callOllama(userMessage, contextPrompt, {
      maxTokens: 800,
      temperature: 0.7,
    });
    return response.trim();
  } catch (error) {
    console.error('Failed to get chat response from Ollama:', error);
    return "I'm having trouble connecting to the AI. Please check if Ollama is running.";
  }
}

/**
 * List available models in Ollama
 * @returns {Promise<Array>} - List of model names
 */
export async function listOllamaModels() {
  const { baseUrl } = ollamaConfig;

  try {
    const response = await fetch(`${baseUrl}/api/tags`);
    if (response.ok) {
      const data = await response.json();
      return data.models?.map(m => m.name) || [];
    }
    return [];
  } catch (error) {
    console.error('Failed to list Ollama models:', error);
    return [];
  }
}

/**
 * Pull/download a model in Ollama
 * @param {string} modelName - Name of the model to pull
 * @returns {Promise<boolean>} - Success status
 */
export async function pullOllamaModel(modelName) {
  const { baseUrl } = ollamaConfig;

  try {
    const response = await fetch(`${baseUrl}/api/pull`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: modelName, stream: false }),
    });
    return response.ok;
  } catch (error) {
    console.error('Failed to pull Ollama model:', error);
    return false;
  }
}
