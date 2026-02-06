// Claude Service for intelligent document processing using Anthropic's Claude API
// NOTE: Claude API has CORS restrictions when called from browser
// This service includes fallback to Groq for browser compatibility
import { claudeConfig, isClaudeConfigured } from '../config/claudeConfig';
import { generateDocumentSummaryGroq, answerDocumentQuestionGroq, getChatResponseGroq } from './groqService';

/**
 * Call Claude API with fallback to Groq
 * @param {Array} messages - Array of message objects with role and content
 * @param {Object} options - Optional parameters
 * @returns {Promise<string>} - The Claude response text
 */
export async function callClaude(messages, options = {}) {
  if (!isClaudeConfigured()) {
    throw new Error('Claude is not configured. Please set environment variables.');
  }

  const requestBody = {
    model: claudeConfig.model,
    max_tokens: options.maxTokens || 1000,
    messages: messages,
    temperature: options.temperature !== undefined ? options.temperature : 0.7,
    top_p: options.topP || 1,
  };

  try {
    console.log('Claude: Attempting to call', claudeConfig.model, 'with', messages.length, 'messages');
    
    const response = await fetch(claudeConfig.apiEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': claudeConfig.apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(requestBody),
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('Claude API error response:', data);
      throw new Error(
        `Claude API error: ${response.status} - ${
          data.error?.message || response.statusText
        }`
      );
    }

    const text = data.content?.[0]?.text || '';
    console.log('Claude: Response received, length:', text.length);
    return text;
  } catch (error) {
    console.error('Claude API call failed:', error.message);
    console.warn('Claude failed with:', error.message, '- This is likely a CORS issue. Claude API cannot be called directly from browser.');
    throw error;
  }
}

/**
 * Generate an intelligent summary of document content using Claude (with Groq fallback)
 * @param {string} content - The document content to summarize
 * @param {string} fileName - The name of the document
 * @returns {Promise<string>} - The Claude-generated summary
 */
export async function generateDocumentSummaryClaude(content, fileName) {
  if (!content || content.trim().length === 0) {
    return null;
  }

  // Truncate content if too long (Claude has good context window but still needs limits)
  const maxContentLength = 16000;
  const truncatedContent =
    content.length > maxContentLength
      ? content.substring(0, maxContentLength) + '\n...[content truncated]'
      : content;

  const systemPrompt = `You are a professional document summarizer for a corporate healthcare environment. Your task is to provide clear, concise, and actionable summaries.

Guidelines:
- Extract the most important information first
- Focus on key decisions, dates, and metrics
- Highlight any action items or requirements
- Use bullet points for clarity
- Keep language professional and precise
- Summaries should be 3-5 sentences or 5-7 bullet points`;

  const userPrompt = `Please provide a comprehensive summary of the following document titled "${fileName}":

${truncatedContent}

Summary:`;

  const messages = [
    {
      role: 'user',
      content: userPrompt,
    },
  ];

  try {
    console.log('Claude Summary: Starting for', fileName, 'content length:', truncatedContent.length);
    const summary = await callClaude(messages, {
      maxTokens: 800,
      temperature: 0.3,
    });
    console.log('Claude Summary: Generated', summary.length, 'characters');
    return summary;
  } catch (error) {
    console.error('Claude summary failed:', error.message);
    console.warn('Claude not available, falling back to Groq...');
    
    // Fallback to Groq
    try {
      return await generateDocumentSummaryGroq(content, fileName);
    } catch (groqError) {
      console.error('Groq fallback also failed:', groqError.message);
      return null;
    }
  }
}

/**
 * Answer a question about document content using Claude (with Groq fallback)
 * @param {string} content - The document content
 * @param {string} question - The user's question
 * @param {string} fileName - The name of the document
 * @returns {Promise<string>} - The Claude-generated answer
 */
export async function answerDocumentQuestionClaude(
  content,
  question,
  fileName
) {
  if (!content || content.trim().length === 0) {
    return "I couldn't read the document content to answer your question.";
  }

  // Truncate content if too long
  const maxContentLength = 16000;
  const truncatedContent =
    content.length > maxContentLength
      ? content.substring(0, maxContentLength) + '\n...[content truncated]'
      : content;

  const messages = [
    {
      role: 'user',
      content: `Answer the following question based on the provided document content.

Document: "${fileName}"

Document Content:
${truncatedContent}

Question: ${question}

If the answer cannot be found in the document, say so clearly.
Be accurate and cite specific information from the document when possible.`,
    },
  ];

  try {
    console.log('Claude Q&A: Starting for question:', question);
    const answer = await callClaude(messages, {
      maxTokens: 800,
      temperature: 0.3,
    });
    console.log('Claude Q&A: Response length:', answer.length);
    return answer;
  } catch (error) {
    console.error('Claude Q&A failed:', error.message);
    console.warn('Claude not available, falling back to Groq...');
    
    // Fallback to Groq
    try {
      return await answerDocumentQuestionGroq(content, question, fileName);
    } catch (groqError) {
      console.error('Groq fallback also failed:', groqError.message);
      return "I encountered an error while processing your question. Please try again.";
    }
  }
}

/**
 * Get chat response using Claude (with Groq fallback) for conversational queries
 * @param {string} userMessage - The user's message
 * @param {Array} conversationHistory - Previous messages in the conversation
 * @returns {Promise<string>} - The Claude response
 */
export async function getChatResponseClaude(
  userMessage,
  conversationHistory = []
) {
  // Build messages array with conversation history
  const messages = [
    ...conversationHistory.slice(-10), // Keep last 10 messages for context
    { role: 'user', content: userMessage },
  ];

  try {
    console.log('Claude Chat: Processing message from user');
    const response = await callClaude(messages, {
      maxTokens: 800,
      temperature: 0.7,
    });
    console.log('Claude Chat: Response length:', response.length);
    return response;
  } catch (error) {
    console.error('Claude chat failed:', error.message);
    console.warn('Claude not available, falling back to Groq...');
    
    // Fallback to Groq
    try {
      return await getChatResponseGroq(userMessage, conversationHistory);
    } catch (groqError) {
      console.error('Groq fallback also failed:', groqError.message);
      return "I'm having trouble processing your request right now. Please try again.";
    }
  }
}
