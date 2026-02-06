// Azure OpenAI Service for intelligent document processing
import { azureOpenAIConfig, isAzureOpenAIConfigured } from '../config/azureOpenAIConfig';

/**
 * Call Azure OpenAI Chat Completion API
 * @param {Array} messages - Array of message objects with role and content
 * @param {Object} options - Optional parameters
 * @returns {Promise<string>} - The AI response text
 */
export async function callAzureOpenAI(messages, options = {}) {
  if (!isAzureOpenAIConfigured()) {
    throw new Error('Azure OpenAI is not configured. Please set environment variables.');
  }

  const { endpoint, apiKey, deploymentName, apiVersion } = azureOpenAIConfig;
  const url = `${endpoint}/openai/deployments/${deploymentName}/chat/completions?api-version=${apiVersion}`;

  const requestBody = {
    messages: messages,
    max_tokens: options.maxTokens || 1000,
    temperature: options.temperature || 0.7,
    top_p: options.topP || 0.95,
    frequency_penalty: options.frequencyPenalty || 0,
    presence_penalty: options.presencePenalty || 0,
    stop: options.stop || null,
  };

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey,
      },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(`Azure OpenAI API error: ${response.status} - ${errorData.error?.message || response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0]?.message?.content || '';
  } catch (error) {
    console.error('Azure OpenAI API call failed:', error);
    throw error;
  }
}

/**
 * Generate an intelligent summary of document content
 * @param {string} content - The document content to summarize
 * @param {string} fileName - The name of the document
 * @returns {Promise<string>} - The AI-generated summary
 */
export async function generateDocumentSummary(content, fileName) {
  if (!content || content.trim().length === 0) {
    return null;
  }

  // Truncate content if too long (keep within token limits)
  const maxContentLength = 8000;
  const truncatedContent = content.length > maxContentLength
    ? content.substring(0, maxContentLength) + '...[truncated]'
    : content;

  const messages = [
    {
      role: 'system',
      content: `You are a helpful document assistant. Your task is to provide clear, concise summaries of documents.
Focus on:
- Main topics and key points
- Important facts, figures, or dates
- The purpose or intent of the document
Keep summaries informative but concise (3-5 sentences).`
    },
    {
      role: 'user',
      content: `Please summarize the following document "${fileName}":\n\n${truncatedContent}`
    }
  ];

  try {
    const summary = await callAzureOpenAI(messages, {
      maxTokens: 500,
      temperature: 0.5,
    });
    return summary;
  } catch (error) {
    console.error('Failed to generate document summary:', error);
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
export async function answerDocumentQuestion(content, question, fileName) {
  if (!content || content.trim().length === 0) {
    return "I couldn't read the document content to answer your question.";
  }

  // Truncate content if too long
  const maxContentLength = 8000;
  const truncatedContent = content.length > maxContentLength
    ? content.substring(0, maxContentLength) + '...[truncated]'
    : content;

  const messages = [
    {
      role: 'system',
      content: `You are a helpful document assistant. Answer questions based on the provided document content.
If the answer cannot be found in the document, say so clearly.
Be accurate and cite specific information from the document when possible.`
    },
    {
      role: 'user',
      content: `Document: "${fileName}"\n\nDocument Content:\n${truncatedContent}\n\nQuestion: ${question}`
    }
  ];

  try {
    const answer = await callAzureOpenAI(messages, {
      maxTokens: 800,
      temperature: 0.3,
    });
    return answer;
  } catch (error) {
    console.error('Failed to answer document question:', error);
    return "I encountered an error while processing your question. Please try again.";
  }
}

/**
 * General chat response (for conversational queries)
 * @param {string} userMessage - The user's message
 * @param {Array} conversationHistory - Previous messages in the conversation
 * @returns {Promise<string>} - The AI response
 */
export async function getChatResponse(userMessage, conversationHistory = []) {
  const systemMessage = {
    role: 'system',
    content: `You are a helpful document assistant for a corporate environment. You help users:
- Find and understand documents
- Answer questions about company policies
- Provide information about processes and procedures

Be professional, helpful, and concise. If you don't have specific information, offer to search for relevant documents.`
  };

  // Build messages array with conversation history
  const messages = [
    systemMessage,
    ...conversationHistory.slice(-10), // Keep last 10 messages for context
    { role: 'user', content: userMessage }
  ];

  try {
    const response = await callAzureOpenAI(messages, {
      maxTokens: 800,
      temperature: 0.7,
    });
    return response;
  } catch (error) {
    console.error('Failed to get chat response:', error);
    return "I'm having trouble processing your request right now. Please try again.";
  }
}

/**
 * Extract key information from document
 * @param {string} content - The document content
 * @param {string} fileName - The name of the document
 * @returns {Promise<Object>} - Extracted key information
 */
export async function extractKeyInfo(content, fileName) {
  if (!content || content.trim().length === 0) {
    return null;
  }

  // Truncate content if too long
  const maxContentLength = 6000;
  const truncatedContent = content.length > maxContentLength
    ? content.substring(0, maxContentLength) + '...[truncated]'
    : content;

  const messages = [
    {
      role: 'system',
      content: `You are a document analysis assistant. Extract key information from documents in a structured format.
Return the information as a JSON object with these fields (use null for missing fields):
- title: Document title if mentioned
- type: Document type (policy, procedure, report, etc.)
- date: Any relevant dates
- keyPoints: Array of 3-5 key points
- topics: Array of main topics covered`
    },
    {
      role: 'user',
      content: `Extract key information from this document "${fileName}":\n\n${truncatedContent}`
    }
  ];

  try {
    const response = await callAzureOpenAI(messages, {
      maxTokens: 600,
      temperature: 0.3,
    });

    // Try to parse as JSON
    try {
      // Extract JSON from response (it might have extra text)
      const jsonMatch = response.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        return JSON.parse(jsonMatch[0]);
      }
    } catch (parseError) {
      console.log('Could not parse key info as JSON, returning as text');
    }

    return { rawResponse: response };
  } catch (error) {
    console.error('Failed to extract key info:', error);
    return null;
  }
}

/**
 * Compare multiple documents
 * @param {Array} documents - Array of {content, fileName} objects
 * @returns {Promise<string>} - Comparison analysis
 */
export async function compareDocuments(documents) {
  if (!documents || documents.length < 2) {
    return "Need at least 2 documents to compare.";
  }

  const docsText = documents.map((doc, i) => {
    const truncated = doc.content.length > 3000
      ? doc.content.substring(0, 3000) + '...[truncated]'
      : doc.content;
    return `Document ${i + 1}: "${doc.fileName}"\n${truncated}`;
  }).join('\n\n---\n\n');

  const messages = [
    {
      role: 'system',
      content: `You are a document analysis assistant. Compare the provided documents and identify:
- Key similarities
- Key differences
- Unique information in each document
Be concise and focus on the most important points.`
    },
    {
      role: 'user',
      content: `Compare these documents:\n\n${docsText}`
    }
  ];

  try {
    const comparison = await callAzureOpenAI(messages, {
      maxTokens: 1000,
      temperature: 0.5,
    });
    return comparison;
  } catch (error) {
    console.error('Failed to compare documents:', error);
    return "I encountered an error while comparing the documents.";
  }
}
