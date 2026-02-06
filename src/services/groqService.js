// Groq API Service - Fast LLM Inference with Llama 3.1
import { groqConfig } from '../config/groqConfig';

/**
 * Call Groq API (OpenAI-compatible format)
 * @param {Array} messages - Array of message objects with role and content
 * @param {Object} options - Optional parameters
 * @returns {Promise<string>} - The AI response text
 */
export async function callGroq(messages, options = {}) {
  const { apiKey, model, baseUrl } = groqConfig;
  const url = `${baseUrl}/chat/completions`;

  const requestBody = {
    model: options.model || model,
    messages: messages,
    max_tokens: options.maxTokens || 1000,
    temperature: options.temperature || 0.7,
    top_p: options.topP || 0.95,
  };

  try {
    console.log('Groq: Calling', model);

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify(requestBody),
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('Groq API error:', data);
      throw new Error(`Groq API error: ${response.status} - ${data.error?.message || response.statusText}`);
    }

    const text = data.choices?.[0]?.message?.content || '';
    console.log('Groq: Response length:', text.length);

    return text;
  } catch (error) {
    console.error('Groq API call failed:', error);
    throw error;
  }
}

/**
 * Generate document summary using Groq
 * @param {string} content - Document content
 * @param {string} fileName - File name
 * @returns {Promise<string|null>} - AI-generated summary
 */
export async function generateDocumentSummaryGroq(content, fileName) {
  if (!content || content.trim().length === 0) {
    return null;
  }

  // Clean and truncate content
  let cleanContent = content
    .split('\n')
    .filter((line, index, arr) => arr.indexOf(line) === index)
    .join('\n')
    .replace(/\s+/g, ' ')
    .trim();

  const maxContentLength = 12000; // Llama 3.1 70B has large context
  const truncatedContent = cleanContent.length > maxContentLength
    ? cleanContent.substring(0, maxContentLength) + '...'
    : cleanContent;

  const messages = [
    {
      role: 'system',
      content: 'You are a helpful document assistant. Provide clear, concise summaries focusing on main topics, key points, and important details. Keep summaries to 3-5 sentences.'
    },
    {
      role: 'user',
      content: `Summarize this document "${fileName}":\n\n${truncatedContent}`
    }
  ];

  try {
    const summary = await callGroq(messages, {
      maxTokens: 500,
      temperature: 0.5,
    });
    return summary.trim();
  } catch (error) {
    console.error('Groq summary failed:', error);
    return null;
  }
}

/**
 * Answer a question about document content using Groq
 * @param {string} content - Document content
 * @param {string} question - User's question
 * @param {string} fileName - File name
 * @returns {Promise<string>} - AI-generated answer
 */
export async function answerDocumentQuestionGroq(content, question, fileName) {
  if (!content || content.trim().length === 0) {
    return "I couldn't read the document content to answer your question.";
  }

  const maxContentLength = 12000;
  const truncatedContent = content.length > maxContentLength
    ? content.substring(0, maxContentLength) + '...'
    : content;

  const messages = [
    {
      role: 'system',
      content: 'You are a helpful document assistant. Answer questions based only on the provided document content. If the answer is not in the document, say so clearly. Be accurate and cite specific information when possible.'
    },
    {
      role: 'user',
      content: `Document: "${fileName}"\n\nContent:\n${truncatedContent}\n\nQuestion: ${question}`
    }
  ];

  try {
    const answer = await callGroq(messages, {
      maxTokens: 800,
      temperature: 0.3,
    });
    return answer.trim();
  } catch (error) {
    console.error('Groq answer failed:', error);
    return "I encountered an error processing your question. Please try again.";
  }
}

/**
 * Get chat response using Groq
 * @param {string} userMessage - User's message
 * @param {Array} conversationHistory - Previous messages
 * @returns {Promise<string>} - AI response
 */
export async function getChatResponseGroq(userMessage, conversationHistory = []) {
  const messages = [
    {
      role: 'system',
      content: `You are a helpful document assistant for a corporate environment. You help users:
- Find and understand documents
- Answer questions about company policies
- Provide information about processes and procedures

Be professional, helpful, and concise.`
    },
    ...conversationHistory.slice(-8).map(msg => ({
      role: msg.role,
      content: msg.content
    })),
    {
      role: 'user',
      content: userMessage
    }
  ];

  try {
    const response = await callGroq(messages, {
      maxTokens: 800,
      temperature: 0.7,
    });
    return response.trim();
  } catch (error) {
    console.error('Groq chat failed:', error);
    return "I'm having trouble processing your request right now. Please try again.";
  }
}
