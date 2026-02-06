// Google Gemini API Service for document processing
import { geminiConfig } from '../config/geminiConfig';

/**
 * Call Google Gemini API
 * @param {string} prompt - The prompt to send
 * @param {Object} options - Optional parameters
 * @returns {Promise<string>} - The AI response text
 */
export async function callGemini(prompt, options = {}) {
  const { apiKey, model, baseUrl } = geminiConfig;
  const url = `${baseUrl}/models/${options.model || model}:generateContent?key=${apiKey}`;

  const requestBody = {
    contents: [
      {
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      temperature: options.temperature || 0.7,
      maxOutputTokens: options.maxTokens || 1000,
      topP: options.topP || 0.95,
    },
    safetySettings: [
      { category: 'HARM_CATEGORY_HARASSMENT', threshold: 'BLOCK_ONLY_HIGH' },
      { category: 'HARM_CATEGORY_HATE_SPEECH', threshold: 'BLOCK_ONLY_HIGH' },
      { category: 'HARM_CATEGORY_SEXUALLY_EXPLICIT', threshold: 'BLOCK_ONLY_HIGH' },
      { category: 'HARM_CATEGORY_DANGEROUS_CONTENT', threshold: 'BLOCK_ONLY_HIGH' },
    ],
  };

  try {
    console.log('Gemini API: Calling model:', options.model || model);

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('Gemini API error response:', data);
      throw new Error(`Gemini API error: ${response.status} - ${data.error?.message || response.statusText}`);
    }

    // Check for blocked content
    if (data.candidates?.[0]?.finishReason === 'SAFETY') {
      console.warn('Gemini: Content blocked by safety filters');
      return null;
    }

    // Extract text from response
    const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
    console.log('Gemini API: Response length:', text.length);

    if (!text) {
      console.warn('Gemini API: Empty response', data);
    }

    return text;
  } catch (error) {
    console.error('Gemini API call failed:', error);
    throw error;
  }
}

/**
 * Call Gemini with chat history (multi-turn conversation)
 * @param {Array} messages - Array of {role, content} objects
 * @param {Object} options - Optional parameters
 * @returns {Promise<string>} - The AI response text
 */
export async function callGeminiChat(messages, options = {}) {
  const { apiKey, model, baseUrl } = geminiConfig;
  const url = `${baseUrl}/models/${options.model || model}:generateContent?key=${apiKey}`;

  // Convert messages to Gemini format
  const contents = messages.map(msg => ({
    role: msg.role === 'assistant' ? 'model' : 'user',
    parts: [{ text: msg.content }]
  }));

  const requestBody = {
    contents: contents,
    generationConfig: {
      temperature: options.temperature || 0.7,
      maxOutputTokens: options.maxTokens || 1000,
      topP: options.topP || 0.95,
    },
  };

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(`Gemini Chat API error: ${response.status} - ${errorData.error?.message || response.statusText}`);
    }

    const data = await response.json();
    return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
  } catch (error) {
    console.error('Gemini Chat API call failed:', error);
    throw error;
  }
}

/**
 * Generate document summary using Gemini
 * @param {string} content - Document content
 * @param {string} fileName - File name
 * @returns {Promise<string|null>} - AI-generated summary
 */
export async function generateDocumentSummaryGemini(content, fileName) {
  if (!content || content.trim().length === 0) {
    console.log('Gemini summary: No content provided');
    return null;
  }

  // Clean content - remove duplicate lines and excessive whitespace
  let cleanContent = content
    .split('\n')
    .filter((line, index, arr) => arr.indexOf(line) === index) // Remove duplicate lines
    .join('\n')
    .replace(/\s+/g, ' ')
    .trim();

  // Gemini has large context window, but let's be reasonable
  const maxContentLength = 8000;
  const truncatedContent = cleanContent.length > maxContentLength
    ? cleanContent.substring(0, maxContentLength) + '...'
    : cleanContent;

  console.log('Gemini summary: Content length:', truncatedContent.length);

  const prompt = `Summarize this document in 3-5 clear sentences. Focus on the main purpose, key points, and any important details like dates, names, or numbers.

Document name: "${fileName}"

Document content:
${truncatedContent}

Provide a professional summary:`;

  try {
    console.log('Gemini: Generating summary for', fileName);
    const summary = await callGemini(prompt, {
      maxTokens: 500,
      temperature: 0.5,
    });

    if (summary && summary.trim().length > 0) {
      console.log('Gemini: Summary generated successfully');
      return summary.trim();
    }

    console.log('Gemini: No summary returned');
    return null;
  } catch (error) {
    console.error('Gemini summary error:', error.message);
    return null;
  }
}

/**
 * Answer a question about document content
 * @param {string} content - Document content
 * @param {string} question - User's question
 * @param {string} fileName - File name
 * @returns {Promise<string>} - AI-generated answer
 */
export async function answerDocumentQuestionGemini(content, question, fileName) {
  if (!content || content.trim().length === 0) {
    return "I couldn't read the document content to answer your question.";
  }

  const maxContentLength = 10000;
  const truncatedContent = content.length > maxContentLength
    ? content.substring(0, maxContentLength) + '...[truncated]'
    : content;

  const prompt = `You are a helpful document assistant. Answer the question based on the provided document content.

Document: "${fileName}"

Document Content:
${truncatedContent}

Question: ${question}

Instructions:
- Answer based only on information in the document
- If the answer is not in the document, say so clearly
- Be accurate and cite specific information when possible

Answer:`;

  try {
    const answer = await callGemini(prompt, {
      maxTokens: 800,
      temperature: 0.3,
    });
    return answer.trim();
  } catch (error) {
    console.error('Failed to answer question with Gemini:', error);
    return "I encountered an error processing your question. Please try again.";
  }
}

/**
 * Get chat response using Gemini
 * @param {string} userMessage - User's message
 * @param {Array} conversationHistory - Previous messages
 * @returns {Promise<string>} - AI response
 */
export async function getChatResponseGemini(userMessage, conversationHistory = []) {
  const systemContext = `You are a helpful document assistant for a corporate environment. You help users:
- Find and understand documents
- Answer questions about company policies
- Provide information about processes and procedures

Be professional, helpful, and concise. If you don't have specific information, offer to search for relevant documents.`;

  // Build prompt with context
  let prompt = systemContext + '\n\n';

  if (conversationHistory.length > 0) {
    const recentHistory = conversationHistory.slice(-6);
    prompt += 'Recent conversation:\n';
    recentHistory.forEach(msg => {
      prompt += `${msg.role === 'user' ? 'User' : 'Assistant'}: ${msg.content}\n`;
    });
    prompt += '\n';
  }

  prompt += `User: ${userMessage}\n\nAssistant:`;

  try {
    const response = await callGemini(prompt, {
      maxTokens: 800,
      temperature: 0.7,
    });
    return response.trim();
  } catch (error) {
    console.error('Failed to get chat response from Gemini:', error);
    return "I'm having trouble processing your request right now. Please try again.";
  }
}

/**
 * Analyze document and extract key information
 * @param {string} content - Document content
 * @param {string} fileName - File name
 * @returns {Promise<Object|null>} - Extracted information
 */
export async function analyzeDocumentGemini(content, fileName) {
  if (!content || content.trim().length === 0) {
    return null;
  }

  const maxContentLength = 8000;
  const truncatedContent = content.length > maxContentLength
    ? content.substring(0, maxContentLength) + '...[truncated]'
    : content;

  const prompt = `Analyze this document and extract key information in JSON format.

Document: "${fileName}"

Content:
${truncatedContent}

Return a JSON object with these fields (use null if not found):
{
  "title": "document title if mentioned",
  "type": "document type (policy, procedure, report, etc.)",
  "date": "any relevant dates mentioned",
  "keyPoints": ["array of 3-5 key points"],
  "topics": ["array of main topics covered"],
  "summary": "2-3 sentence summary"
}

JSON:`;

  try {
    const response = await callGemini(prompt, {
      maxTokens: 800,
      temperature: 0.3,
    });

    // Try to parse JSON from response
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
    return { rawResponse: response };
  } catch (error) {
    console.error('Failed to analyze document with Gemini:', error);
    return null;
  }
}
