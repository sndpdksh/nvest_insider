import { useState, useRef, useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest } from '../config/authConfig';
import { searchFiles, searchMedia, getDocumentContent, getRecentFiles, getFoldersOnly, uploadFileToOneDrive, getFileById, getFileType, SUPPORTED_UPLOAD_EXTENSIONS, MAX_UPLOAD_SIZE } from '../services/graphService';
import FolderPicker from './FolderPicker';
import PMDocumentForm from './PMDocumentForm';
import { generatePMDocument } from '../services/docxGenerator';
import { saveAs } from 'file-saver';
import {
  initializeLLM,
  isLLMAvailable,
  getProviderName,
  getAvailableProviders,
  setLLMProvider,
  getLLMProvider,
  generateSummary as llmGenerateSummary,
  getChatResponse as llmGetChatResponse,
  answerQuestion as llmAnswerQuestion,
} from '../services/llmService';
import './ChatBot.css';

// Knowledge base with Q&A patterns
const knowledgeBase = [
  {
    keywords: ['approval', 'project', 'change', 'process'],
    topic: 'project_change',
    initialResponse: `Project change approval involves submission of a change request, impact analysis, and formal approval before implementation.

Could you confirm:
• Which department is this for?
• Is this for scope, timeline, or cost change?`,
    followUp: {
      'pmo': {
        'scope': {
          response: 'Scope change approval requires a formal change request, impact assessment on cost and timeline, and sign-off from the client sponsor before implementation.',
          sourceFile: 'Project_Change_Process'
        },
        'timeline': {
          response: 'Timeline change approval requires justification, updated schedule, and approval from project sponsor and stakeholders.',
          sourceFile: 'Project_Change_Process'
        },
        'cost': {
          response: 'Cost change approval requires budget revision request, financial impact analysis, and CFO approval for changes exceeding 10%.',
          sourceFile: 'Budget_Change_Policy'
        }
      },
      'finance': {
        'scope': {
          response: 'Finance scope changes require budget committee review and CFO approval.',
          sourceFile: 'Finance_Change_Policy'
        }
      }
    }
  },
  {
    keywords: ['leave', 'policy', 'vacation', 'pto'],
    topic: 'leave_policy',
    initialResponse: `Our leave policy covers various types of time off including vacation, sick leave, and personal days.

What would you like to know:
• Annual leave entitlement?
• Sick leave policy?
• Leave application process?`,
    followUp: {
      'annual': {
        response: 'Annual leave entitlement is 20 days per year for regular employees, accrued monthly. Unused leave can be carried forward up to 5 days.',
        sourceFile: 'Leave_Policy'
      },
      'sick': {
        response: 'Sick leave is 10 days per year. Medical certificate required for absences exceeding 3 consecutive days.',
        sourceFile: 'Leave_Policy'
      },
      'application': {
        response: 'Leave applications should be submitted via HR portal at least 5 business days in advance. Manager approval is required.',
        sourceFile: 'Leave_Policy'
      }
    }
  },
  {
    keywords: ['expense', 'reimbursement', 'claim'],
    topic: 'expense',
    initialResponse: `Expense reimbursement follows our corporate policy for business-related expenses.

Please specify:
• Travel expenses?
• Office supplies?
• Client entertainment?`,
    followUp: {
      'travel': {
        response: 'Travel expenses require pre-approval for trips over $500. Submit receipts within 30 days via expense portal. Per diem rates apply for meals.',
        sourceFile: 'Travel_Expense_Policy'
      },
      'supplies': {
        response: 'Office supplies under $100 can be purchased directly. Amounts over $100 require manager approval.',
        sourceFile: 'Procurement_Policy'
      },
      'entertainment': {
        response: 'Client entertainment requires pre-approval and itemized receipts. Maximum $75 per person for meals.',
        sourceFile: 'Entertainment_Policy'
      }
    }
  },
  {
    keywords: ['onboarding', 'new', 'employee', 'hire', 'joining'],
    topic: 'onboarding',
    initialResponse: `Welcome! Onboarding process includes IT setup, HR documentation, and department orientation.

What do you need help with:
• IT access and equipment?
• HR documentation?
• Training schedule?`,
    followUp: {
      'it': {
        response: 'IT setup includes laptop provisioning, email activation, and system access. Ticket is auto-created on hire. Equipment delivered within 2 business days.',
        sourceFile: 'IT_Onboarding_Checklist'
      },
      'hr': {
        response: 'HR documentation includes I-9 verification, tax forms, benefits enrollment, and emergency contacts. Complete within first 3 days.',
        sourceFile: 'HR_Onboarding_Guide'
      },
      'training': {
        response: 'Mandatory training includes compliance (Day 1), security awareness (Week 1), and role-specific training (Week 2-4).',
        sourceFile: 'Training_Schedule'
      }
    }
  }
];

function ChatBot({ isDarkMode, setIsDarkMode }) {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [aiEnabled, setAiEnabled] = useState(false);
  const [aiProvider, setAiProvider] = useState('');
  const [availableLLMs, setAvailableLLMs] = useState([]);
  const [selectedLLM, setSelectedLLM] = useState('');
  const [conversationHistory, setConversationHistory] = useState([]);
  const [showSettings, setShowSettings] = useState(false);

  // Toggle theme
  const toggleTheme = () => {
    setIsDarkMode(prev => !prev);
  };

  // Handle logout
  const handleLogout = () => {
    instance.logoutPopup();
  };

  // Initialize LLM on mount
  useEffect(() => {
    let isMounted = true;
    const init = async () => {
      const result = await initializeLLM();
      if (isMounted) {
        setAiEnabled(isLLMAvailable());
        setAiProvider(getProviderName());
        setAvailableLLMs(result.availableProviders);
        setSelectedLLM(result.activeProvider);
      }
    };
    init();
    return () => { isMounted = false; };
  }, []);

  // Handle LLM provider change
  const handleLLMChange = (e) => {
    const newProvider = e.target.value;
    if (setLLMProvider(newProvider)) {
      setSelectedLLM(newProvider);
      setAiProvider(getProviderName());
      console.log('Switched to:', getProviderName());
    }
  };

  const [messages, setMessages] = useState([
    {
      type: 'bot',
      text: aiEnabled
        ? `Hello! I'm your AI-powered document assistant. I can help you:

• Find and summarize documents
• Answer questions about your files
• Explain document contents

Try asking me things like:
• "Summarize the Arogya document"
• "What is in the health policy?"
• "Tell me about leave policy"

How can I help you today?`
        : `Hello! I’m Nvest Insider, your Knowledge centre. How can I help you today?`
    }
  ]);
  const [input, setInput] = useState('');
  const [isTyping, setIsTyping] = useState(false);
  const [currentContext, setCurrentContext] = useState(null);
  const [sourceDocuments, setSourceDocuments] = useState([]);
  const [recentDocuments, setRecentDocuments] = useState([]); // Accumulated recent docs for sidebar
  const [activeDocument, setActiveDocument] = useState(null); // Track current document for Q&A
  const lastFileListRef = useRef([]); // Store last displayed file list for number selection
  const lastSearchResultsRef = useRef([]); // Store full search results for "show all files"
  const messagesEndRef = useRef(null);
  const fileInputRef = useRef(null);
  const recentUploadsRef = useRef([]); // Track recently uploaded items for search
  const lastSuggestedQuestionsRef = useRef([]); // Store numbered suggested questions for shortcut
  const [showFolderPicker, setShowFolderPicker] = useState(false);
  const [pendingUploadFile, setPendingUploadFile] = useState(null);
  const [uploadProgress, setUploadProgress] = useState(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const [showPMDocForm, setShowPMDocForm] = useState(false);
  const [pmDocFormData, setPmDocFormData] = useState(null);
  const [pendingPMDocBlob, setPendingPMDocBlob] = useState(null); // For upload after download
  const [showMobileSidebar, setShowMobileSidebar] = useState(false);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  // Accumulate documents into recentDocuments whenever sourceDocuments changes
  useEffect(() => {
    if (sourceDocuments.length === 0) return;
    setRecentDocuments(prev => {
      const existingIds = new Set(prev.map(d => d.id || d.name));
      const newDocs = sourceDocuments.filter(d => !existingIds.has(d.id || d.name));
      if (newDocs.length === 0) return prev;
      return [...newDocs, ...prev].slice(0, 20);
    });
  }, [sourceDocuments]);


  // Search for source documents
  const findSourceDocument = async (searchTerm) => {
    if (!isAuthenticated || !accounts || accounts.length === 0) {
      return null;
    }

    try {
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      const results = await searchFiles(tokenResponse.accessToken, searchTerm);
      return results.length > 0 ? results[0] : null;
    } catch (error) {
      console.error('Error searching for source:', error);
      return null;
    }
  };

  // Find matching knowledge base entry
  const findKnowledgeMatch = (text) => {
    const lowerText = text.toLowerCase();

    for (const entry of knowledgeBase) {
      const matchCount = entry.keywords.filter(kw => lowerText.includes(kw)).length;
      if (matchCount >= 2) {
        return entry;
      }
    }
    return null;
  };

  // Process follow-up response
  const processFollowUp = (text, context) => {
    const lowerText = text.toLowerCase();

    if (context.topic === 'project_change') {
      // Check for department
      const departments = ['pmo', 'finance', 'hr', 'it', 'operations'];
      const changeTypes = ['scope', 'timeline', 'cost', 'budget'];

      let dept = departments.find(d => lowerText.includes(d));
      let changeType = changeTypes.find(t => lowerText.includes(t));

      if (dept && changeType) {
        const followUp = context.followUp[dept]?.[changeType] || context.followUp['pmo']?.[changeType];
        if (followUp) {
          return {
            response: followUp.response,
            sourceFile: followUp.sourceFile,
            resolved: true
          };
        }
      }
    } else if (context.followUp) {
      // Check for matching follow-up keywords
      for (const [key, value] of Object.entries(context.followUp)) {
        if (lowerText.includes(key)) {
          return {
            response: value.response,
            sourceFile: value.sourceFile,
            resolved: true
          };
        }
      }
    }

    return null;
  };

  // Handle source request
  const handleSourceRequest = async () => {
    if (sourceDocuments.length > 0) {
      const sources = sourceDocuments.map(doc =>
        `📄 ${doc.name}\n   📂 ${doc.path}`
      ).join('\n\n');

      return `**Source Documents:**\n\n${sources}`;
    }
    return "I don't have a specific source document for the previous response. Would you like me to search for related documents?";
  };

  // Search for multiple documents
  const searchDocuments = async (searchTerm) => {
    if (!isAuthenticated || !accounts || accounts.length === 0) {
      return [];
    }

    try {
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      const results = await searchFiles(tokenResponse.accessToken, searchTerm);

      // Include recently uploaded files that match the search term
      // (Graph search index may take time to index new uploads)
      const searchLower = searchTerm.toLowerCase();
      const matchingUploads = recentUploadsRef.current.filter(
        u => u.name?.toLowerCase().includes(searchLower) &&
             !results.some(r => r.id === u.id)
      );

      for (const upload of matchingUploads) {
        const fileItem = await getFileById(tokenResponse.accessToken, upload.id);
        if (fileItem) {
          results.unshift(fileItem);
        }
      }

      return results;
    } catch (error) {
      console.error('Error searching documents:', error);
      return [];
    }
  };

  // Check if message is asking to read/summarize a document
  const isReadSummaryRequest = (message) => {
    const lowerMsg = message.toLowerCase();
    const readKeywords = ['read', 'summarize', 'summary', 'content', 'what does', 'what is in', 'tell me about', 'explain', 'transcript', 'meeting', 'recording', 'call'];
    return readKeywords.some(kw => lowerMsg.includes(kw));
  };

  // Check if file is a video
  const isVideoFile = (fileName) => {
    const ext = fileName?.split('.').pop()?.toLowerCase();
    return ['mp4', 'mov', 'avi', 'mkv', 'webm'].includes(ext);
  };

  // Read video transcript and generate summary
  const readVideoTranscript = async (videoItem) => {
    if (!isAuthenticated || !accounts || accounts.length === 0) {
      return null;
    }
    try {
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });
      const { getVideoTranscript } = await import('../services/graphService');
      return await getVideoTranscript(tokenResponse.accessToken, videoItem);
    } catch (error) {
      console.error('Error reading video transcript:', error);
      return null;
    }
  };

  // Check if message is a follow-up question about the active document
  const isDocumentQuestion = (message) => {
    if (!activeDocument) return false;
    const lowerMsg = message.toLowerCase();
    // Direct action phrases for CR documents
    const actionPhrases = ['simplify', 'layman', 'simple terms', 'break down', 'task breakdown', 'dev assignment', 'assign to dev', 'development tasks'];
    if (actionPhrases.some(p => lowerMsg.includes(p))) return true;
    // Question indicators
    const questionWords = ['what', 'who', 'when', 'where', 'why', 'how', 'which', 'is', 'are', 'does', 'do', 'can', 'tell', 'explain', 'describe', 'list', 'show'];
    const hasQuestionWord = questionWords.some(qw => lowerMsg.startsWith(qw) || lowerMsg.includes(' ' + qw + ' '));
    const hasQuestionMark = message.includes('?');
    // Check if it's likely a question about the document
    return hasQuestionWord || hasQuestionMark;
  };

  // Answer question about the active document
  const answerDocumentQuestion = async (question) => {
    if (!activeDocument || !activeDocument.content) {
      return null;
    }
    try {
      console.log('Answering question about:', activeDocument.name);
      const answer = await llmAnswerQuestion(activeDocument.content, question, activeDocument.name);
      return answer;
    } catch (error) {
      console.error('Error answering document question:', error);
      return null;
    }
  };

  // Read document and get content
  const readDocument = async (itemId, fileName) => {
    if (!isAuthenticated || !accounts || accounts.length === 0) {
      return null;
    }

    try {
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      const docContent = await getDocumentContent(tokenResponse.accessToken, itemId, fileName);
      return docContent;
    } catch (error) {
      console.error('Error reading document:', error);
      return null;
    }
  };

  // Generate summary from content (uses AI if configured)
  const generateSummary = async (content, fileName, maxLength = 500) => {
    if (!content || content.trim().length === 0) {
      return null;
    }

    // Try AI-powered summary first
    if (aiEnabled) {
      try {
        console.log('Generating AI summary for:', fileName, 'using:', aiProvider);
        const aiSummary = await llmGenerateSummary(content, fileName);
        if (aiSummary) {
          return aiSummary;
        }
      } catch (error) {
        console.log('AI summary failed, falling back to simple summary:', error);
      }
    }

    // Fallback: Simple extraction-based summary
    console.log('Using fallback summary (AI not available or failed)');

    // Clean the content first
    let text = content
      .split('\n')
      .filter((line, i, arr) => line.trim() && arr.indexOf(line) === i) // Remove empty and duplicate lines
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();

    // Extract meaningful sentences
    const sentences = text.split(/[.!?]+/)
      .map(s => s.trim())
      .filter(s => s.length > 20 && !s.match(/^[\[\]\(\)\{\}]/)); // Filter short and bracket-only content

    if (sentences.length > 0) {
      let summary = '';
      for (const sentence of sentences.slice(0, 5)) { // Max 5 sentences
        if (summary.length + sentence.length > maxLength) break;
        summary += sentence + '. ';
      }
      return summary.trim();
    }

    // Last resort: just first N characters
    return text.length > maxLength ? text.substring(0, maxLength) + '...' : text;
  };

  // Check if message is a file search request
  const isFileSearchRequest = (message) => {
    const lowerMsg = message.toLowerCase();
    // Action keywords
    const actionKeywords = ['find', 'search', 'show', 'get', 'list', 'check', 'want', 'need', 'look', 'related', 'about', 'share', 'give'];
    // File type keywords
    const fileKeywords = ['file', 'files', 'document', 'documents', 'sheet', 'sheets', 'video', 'recording', 'recordings', 'image', 'images'];
    const hasSearchIntent = actionKeywords.some(kw => lowerMsg.includes(kw));
    const hasFileKeyword = fileKeywords.some(kw => lowerMsg.includes(kw));

    // Also check for product names or specific terms
    const hasSpecificTerm = lowerMsg.includes('arogya') ||
                           lowerMsg.includes('sanjeevani') ||
                           lowerMsg.includes('top up') ||
                           lowerMsg.includes('topup') ||
                           lowerMsg.includes('product') ||
                           lowerMsg.includes('policy') ||
                           lowerMsg.includes('health');

    return hasSearchIntent || hasFileKeyword || hasSpecificTerm;
  };

  // Check if user wants multiple files or just one
  const wantsMultipleFiles = (message) => {
    const lowerMsg = message.toLowerCase();
    const multipleKeywords = ['files', 'all', 'list', 'multiple', 'documents', 'every', 'related'];
    const hasKeyword = multipleKeywords.some(kw => lowerMsg.includes(kw));
    const hasCount = getRequestedCount(message) !== null;
    return hasKeyword || hasCount;
  };

  // Extract requested file count from message (e.g. "top 10", "show 20 files")
  const getRequestedCount = (message) => {
    const lowerMsg = message.toLowerCase();
    // Match patterns: "top 10", "top10", "show 20", "first 15", "10 files", "10 documents"
    const patterns = [
      /top\s*(\d+)/i,
      /first\s*(\d+)/i,
      /show\s*(\d+)/i,
      /list\s*(\d+)/i,
      /get\s*(\d+)/i,
      /(\d+)\s*(?:files?|documents?|results?|items?)/i,
    ];
    for (const pattern of patterns) {
      const match = lowerMsg.match(pattern);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > 0 && num <= 50) return num;
      }
    }
    return null;
  };

  // Check if message is just a number (for selecting from file list)
  const isNumberSelection = (message) => {
    return /^\s*\d+\s*$/.test(message.trim());
  };

  // Check if user specifically asked for folders
  const wantsFolders = (message) => {
    const lowerMsg = message.toLowerCase();
    return lowerMsg.includes('folder') || lowerMsg.includes('directory') || lowerMsg.includes('directories');
  };

  // Check if user specifically asked for documents (show only doc/docx/pdf)
  const wantsDocumentsOnly = (message) => {
    const lowerMsg = message.toLowerCase();
    return lowerMsg.includes('document') || lowerMsg.includes('doc ') || lowerMsg.includes('docs') || lowerMsg.endsWith('doc');
  };

  const isDocOrPdf = (fileName) => {
    const ext = fileName?.split('.').pop()?.toLowerCase();
    return ['doc', 'docx', 'pdf'].includes(ext);
  };

  // Apply file type filters based on user request
  const applyFileFilters = (files, message) => {
    let filtered = files;
    if (!wantsFolders(message)) {
      filtered = filtered.filter(f => !f.isFolder);
    }
    if (wantsDocumentsOnly(message)) {
      filtered = filtered.filter(f => f.isFolder || isDocOrPdf(f.name));
    }
    return filtered;
  };

  // Check if message is asking for images or videos
  const isMediaRequest = (message) => {
    const lowerMsg = message.toLowerCase();
    const mediaKeywords = ['image', 'images', 'photo', 'photos', 'picture', 'pictures', 'video', 'videos', 'media', 'png', 'jpg', 'jpeg', 'gif', 'mp4', 'mov'];
    return mediaKeywords.some(kw => lowerMsg.includes(kw));
  };

  // Determine media type requested
  const getMediaType = (message) => {
    const lowerMsg = message.toLowerCase();
    const imageKeywords = ['image', 'images', 'photo', 'photos', 'picture', 'pictures', 'png', 'jpg', 'jpeg', 'gif'];
    const videoKeywords = ['video', 'videos', 'mp4', 'mov', 'clip', 'clips'];

    const hasImage = imageKeywords.some(kw => lowerMsg.includes(kw));
    const hasVideo = videoKeywords.some(kw => lowerMsg.includes(kw));

    if (hasImage && !hasVideo) return 'image';
    if (hasVideo && !hasImage) return 'video';
    return 'all';
  };

  // Search for media files
  const searchMediaFiles = async (searchTerm, mediaType = 'all') => {
    if (!isAuthenticated || !accounts || accounts.length === 0) {
      return [];
    }

    try {
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      const results = await searchMedia(tokenResponse.accessToken, searchTerm, mediaType);
      return results;
    } catch (error) {
      console.error('Error searching media:', error);
      return [];
    }
  };

  // Extract search terms from message
  const extractSearchTerms = (message) => {
    const lowerMsg = message.toLowerCase();
    const productTerms = [];

    // Check for CR/Change Request numbers (e.g., CR 19637, CR-19637, CR19637)
    const crMatch = message.match(/CR[\s\-_]*(\d+)/i);
    if (crMatch) {
      // Search with most specific terms first
      productTerms.push(`CR${crMatch[1]}`);    // CR20049 (exact, no space)
      productTerms.push(`CR_${crMatch[1]}`);   // CR_20049 (underscore variant)
      productTerms.push(`CR ${crMatch[1]}`);   // CR 20049
      productTerms.push(`CR-${crMatch[1]}`);   // CR-20049
      productTerms.push(crMatch[1]);            // 20049 (bare number last - too broad)
      return productTerms;
    }

    // Check for any standalone numbers (could be document IDs)
    const numberMatch = message.match(/\b(\d{4,})\b/);
    if (numberMatch) {
      productTerms.push(numberMatch[1]);
    }

    // Common product/file terms to search for
    if (lowerMsg.includes('arogya') && lowerMsg.includes('sanjeevani')) {
      productTerms.push('Arogya Sanjeevani');
    } else if (lowerMsg.includes('arogya')) {
      productTerms.push('Arogya');
    }

    if (lowerMsg.includes('top up') || lowerMsg.includes('topup')) {
      productTerms.push('Top up');
      productTerms.push('Topup');
    }

    if (lowerMsg.includes('sanjeevani') && !lowerMsg.includes('arogya')) {
      productTerms.push('Sanjeevani');
    }

    if (lowerMsg.includes('cross sell') || lowerMsg.includes('crosssell')) {
      productTerms.push('Cross Sell');
    }

    if (lowerMsg.includes('insuremo')) {
      productTerms.push('Insuremo');
    }

    // If we found specific terms, return them
    if (productTerms.length > 0) {
      return productTerms;
    }

    // Otherwise, extract key words from message
    const stopWords = [
      // Action keywords
      'find', 'search', 'show', 'get', 'list', 'check', 'want', 'need', 'look',
      'related', 'about', 'from', 'my', 'all', 'the', 'in', 'for', 'to', 'a',
      'an', 'of', 'every', 'multiple', 'share', 'give', 'me', 'please', 'can',
      'you', 'could', 'would', 'should', 'this', 'that', 'these', 'those',
      // Read/summary keywords
      'read', 'summarize', 'summary', 'content', 'what', 'does', 'tell', 'explain',
      'is', 'are', 'was', 'were', 'has', 'have', 'had', 'do', 'did',
      // File type keywords
      'file', 'files', 'document', 'documents', 'sheet', 'sheets',
      'recording', 'recordings',
      // Media keywords
      'image', 'images', 'photo', 'photos', 'picture', 'pictures',
      'video', 'videos', 'media', 'clip', 'clips',
      // Count/number keywords
      'top', 'first', 'last',
      // Common filler words
      'add', 'respective', 'fields', 'tags'
    ];

    const words = message.split(/\s+/).filter(w =>
      !stopWords.includes(w.toLowerCase()) &&
      w.length > 2 &&
      !/^[\-\_\.\,\;\:]+$/.test(w) // Skip punctuation-only
    );

    // Take only first 3-4 meaningful words for search
    if (words.length > 0) {
      productTerms.push(words.slice(0, 4).join(' '));
    }

    return productTerms;
  };

  const handleSend = async (directMessage = null) => {
    const messageToSend = (typeof directMessage === 'string' ? directMessage : null) || input.trim();
    if (!messageToSend) return;

    const userMessage = messageToSend;
    setInput('');

    // Add user message
    setMessages(prev => [...prev, { type: 'user', text: userMessage }]);
    setIsTyping(true);

    // Process the message
    let botResponse = '';
    let newSourceDocs = [];

    const lowerMessage = userMessage.toLowerCase();

    // Check if user wants to see all files from last search
    if ((lowerMessage.includes('show all files') || lowerMessage.includes('show all') || lowerMessage.includes('see more')) && lastSearchResultsRef.current.length > 1) {
      const allResults = lastSearchResultsRef.current;
      botResponse = `📂 **All Search Results (${allResults.length}):**\n\n`;
      allResults.forEach((file, idx) => {
        botResponse += `${idx + 1}. **${file.name}**\n`;
        botResponse += `   📁 ${file.path || 'Root'} • ${file.date || 'N/A'}\n\n`;
      });
      botResponse += `\n💡 **Type a number (e.g. "3") to summarize that document.**`;
      newSourceDocs = allResults;
      setSourceDocuments(allResults);
      lastFileListRef.current = allResults;

      setIsTyping(false);
      setMessages(prev => [...prev, { type: 'bot', text: botResponse, sources: newSourceDocs }]);
      return;
    }

    // Check if user wants to see recent/all files
    if (lowerMessage.includes('recent files') || lowerMessage.includes('my recent') || lowerMessage.includes('show recent') || lowerMessage.includes('list recent') || lowerMessage.includes('all files') || lowerMessage.includes('list files') || lowerMessage.includes('my files')) {
      try {
        const tokenResponse = await instance.acquireTokenSilent({
          ...loginRequest,
          account: accounts[0],
        });

        const recentFiles = await getRecentFiles(tokenResponse.accessToken);

        if (recentFiles.length > 0) {
          const filteredFiles = applyFileFilters(recentFiles, userMessage);
          const requestedCount = getRequestedCount(userMessage);
          const showCount = requestedCount || 10;
          const displayFiles = filteredFiles.slice(0, showCount);

          const label = wantsFolders(userMessage) ? 'Files & Folders' : wantsDocumentsOnly(userMessage) ? 'Documents' : 'Files';
          botResponse = `📂 **Your Recent ${label} (${filteredFiles.length}):**\n\n`;
          displayFiles.forEach((file, idx) => {
            botResponse += `${idx + 1}. **${file.name}**\n`;
            botResponse += `   📁 ${file.path || 'Root'} • ${file.date || 'N/A'}\n\n`;
          });
          if (filteredFiles.length > showCount) {
            botResponse += `_...and ${filteredFiles.length - showCount} more files_\n`;
          }
          botResponse += `\n💡 **Type a number (e.g. "3") to summarize that document.**`;
          newSourceDocs = displayFiles;
          setSourceDocuments(newSourceDocs);
          lastFileListRef.current = displayFiles;
        } else {
          botResponse = `No recent files found. Try uploading a file to OneDrive first.`;
        }
      } catch (error) {
        console.error('Error getting recent files:', error);
        botResponse = `Couldn't fetch recent files. Please make sure you're signed in.`;
      }

      setIsTyping(false);
      setMessages(prev => [...prev, { type: 'bot', text: botResponse, sources: newSourceDocs }]);
      return;
    }

    // Check if user typed a number to pick a suggested question
    if (isNumberSelection(userMessage) && lastSuggestedQuestionsRef.current.length > 0) {
      const qIdx = parseInt(userMessage.trim(), 10) - 1;
      if (qIdx >= 0 && qIdx < lastSuggestedQuestionsRef.current.length) {
        const selectedQuestion = lastSuggestedQuestionsRef.current[qIdx];
        lastSuggestedQuestionsRef.current = [];
        setIsTyping(false);
        // Re-send as if user typed the full question
        handleSend(selectedQuestion);
        return;
      }
    }

    // Check if user typed a number to select a file from the last list
    const fileList = lastFileListRef.current;
    if (isNumberSelection(userMessage) && fileList.length > 0) {
      const selectedIdx = parseInt(userMessage.trim(), 10) - 1;
      if (selectedIdx >= 0 && selectedIdx < fileList.length) {
        const doc = fileList[selectedIdx];
        console.log('User selected file #', selectedIdx + 1, ':', doc.name, 'id:', doc.id);

        // Check if it's a video file
        if (isVideoFile(doc.name)) {
          const transcriptData = await readVideoTranscript(doc);
          newSourceDocs = [{ ...doc, isVideo: true }];
          setSourceDocuments(newSourceDocs);

          botResponse = `🎬 **${doc.name}**\n\n`;
          if (transcriptData && transcriptData.hasTranscript && transcriptData.content) {
            setActiveDocument({ id: doc.id, name: doc.name, content: transcriptData.content, path: doc.path, isVideo: true });
            const summary = await generateSummary(transcriptData.content, doc.name);
            if (summary) {
              botResponse += `**📝 ${aiEnabled ? 'AI Summary of Transcript' : 'Transcript Preview'}:**\n${summary}\n`;
            }
            botResponse += `\n**📂 Path:** ${doc.path}`;
            if (aiEnabled) botResponse += `\n\n💡 **Ask me questions about this recording!**`;
          } else {
            botResponse += `No transcript available for this video.\nClick "Open" below to watch in OneDrive.`;
          }
        } else {
          // Regular document - read and summarize
          const docContent = await readDocument(doc.id, doc.name);
          newSourceDocs = [{ ...doc, webUrl: docContent?.webUrl || doc.webUrl }];
          setSourceDocuments(newSourceDocs);

          botResponse = `📄 **#${selectedIdx + 1}: ${doc.name}**\n\n`;

          if (docContent && docContent.content && docContent.content.length > 50) {
            setActiveDocument({ id: doc.id, name: doc.name, content: docContent.content, path: docContent.path });
            const summary = await generateSummary(docContent.content, doc.name);
            if (summary) {
              botResponse += `**📝 ${aiEnabled ? 'Document Summary' : 'Content Preview'}:**\n${summary}\n`;
            }

            // For CR documents, extract impacted areas
            const isCRDoc = /CR[\s\-_]*\d+/i.test(doc.name);
            if (isCRDoc && aiEnabled) {
              try {
                const impactedAreas = await llmAnswerQuestion(
                  docContent.content,
                  'List all impacted areas, affected modules, systems, screens, APIs, and components mentioned in this document. Format as a concise bullet list. If specific module names, screen names, or API names are mentioned, include them.',
                  doc.name
                );
                if (impactedAreas && !impactedAreas.toLowerCase().includes('not found') && !impactedAreas.toLowerCase().includes('not mentioned')) {
                  botResponse += `\n**🎯 Impacted Areas:**\n${impactedAreas}\n`;
                }
              } catch (impactErr) {
                console.log('Could not extract impacted areas:', impactErr);
              }
            }

            botResponse += `\n**📂 Path:** ${docContent.path || doc.path}`;
            if (aiEnabled) botResponse += `\n\n💡 **Ask me questions about this document!**`;
          } else {
            botResponse += `Could not extract content for summarization.\n`;
            botResponse += `**📂 Path:** ${doc.path}\n`;
            botResponse += `Click "Open" below to view in OneDrive.`;
          }
        }

        setIsTyping(false);
        setMessages(prev => [...prev, { type: 'bot', text: botResponse, sources: newSourceDocs }]);
        return;
      } else {
        botResponse = `Invalid selection. Please enter a number between 1 and ${fileList.length}.`;
        setIsTyping(false);
        setMessages(prev => [...prev, { type: 'bot', text: botResponse }]);
        return;
      }
    }

    // Check if user wants to clear current document and start fresh
    if (lowerMessage.includes('new document') || lowerMessage.includes('different document') || lowerMessage.includes('another document') || lowerMessage.includes('clear document')) {
      setActiveDocument(null);
      botResponse = `Document cleared. What document would you like me to find?\n\nTry:\n• "Read Arogya document"\n• "Summarize health policy"\n• "Show me product files"`;
      setIsTyping(false);
      setMessages(prev => [...prev, { type: 'bot', text: botResponse, sources: [] }]);
      return;
    }

    // Check if asking to read/summarize a document
    if (isReadSummaryRequest(userMessage)) {
      const searchTerms = extractSearchTerms(userMessage);
      console.log('Read/Summary request - Search terms:', searchTerms);

      let results = [];
      let usedSearchTerm = '';
      const readableExts = ['docx', 'doc', 'pdf', 'txt', 'md', 'csv', 'xlsx', 'xls', 'pptx', 'ppt'];
      const userWantsVideo = /\b(video|recording|meeting|call)\b/i.test(userMessage);
      let fallbackResults = [];
      let fallbackTerm = '';

      // Try each search term — prefer readable docs with name match
      for (const term of searchTerms) {
        if (!term) continue;
        console.log('Searching for:', term);
        const allResults = await searchDocuments(term);
        if (allResults.length === 0) continue;

        // Filter to readable file types only (unless user wants video)
        let filtered = allResults;
        if (!userWantsVideo) {
          const readable = allResults.filter(r => {
            const ext = r.name?.split('.').pop()?.toLowerCase();
            return readableExts.includes(ext);
          });
          if (readable.length > 0) filtered = readable;
        }

        // Check if any readable result has the search term in filename
        const termLower = term.toLowerCase();
        const hasNameMatch = filtered.some(r => r.name?.toLowerCase().includes(termLower));

        if (hasNameMatch) {
          results = filtered;
          usedSearchTerm = term;
          break;
        }

        // Keep first readable results as fallback
        if (fallbackResults.length === 0 && filtered.length > 0) {
          fallbackResults = filtered;
          fallbackTerm = term;
        }
      }

      // Use fallback if no name-matched readable results found
      if (results.length === 0 && fallbackResults.length > 0) {
        results = fallbackResults;
        usedSearchTerm = fallbackTerm;
      }

      console.log('Found documents:', results.length, 'using term:', usedSearchTerm);

      // Sort and prioritize results
      if (results.length > 0) {
        // Sort: name-match first
        results.sort((a, b) => {
          const aNameMatch = a.name?.toLowerCase().includes(usedSearchTerm.toLowerCase()) ? 1 : 0;
          const bNameMatch = b.name?.toLowerCase().includes(usedSearchTerm.toLowerCase()) ? 1 : 0;
          return bNameMatch - aNameMatch;
        });

        // If user mentioned "video"/"recording", prefer video results
        if (userWantsVideo) {
          const videoResult = results.find(r => isVideoFile(r.name));
          if (videoResult) {
            results = [videoResult, ...results.filter(r => r.id !== videoResult.id)];
            console.log('Prioritized video result:', videoResult.name);
          }
        }

        console.log('Best match:', results[0]?.name);
      }

      if (results.length > 0) {
          const doc = results[0];
          console.log('Reading document:', doc.name, doc.id);

          // Check if the found file is a video - use transcript flow instead
          if (isVideoFile(doc.name)) {
            console.log('Video file detected, using transcript flow:', doc.name);
            const transcriptData = await readVideoTranscript(doc);
            newSourceDocs = [{ ...doc, isVideo: true }];
            setSourceDocuments(newSourceDocs);

            botResponse = `🎬 **${doc.name}**\n\n`;
            if (transcriptData && transcriptData.hasTranscript && transcriptData.content) {
              setActiveDocument({ id: doc.id, name: doc.name, content: transcriptData.content, path: doc.path, isVideo: true });
              const summary = await generateSummary(transcriptData.content, doc.name);
              if (summary) {
                botResponse += `**📝 ${aiEnabled ? 'AI Summary of Transcript' : 'Transcript Preview'}:**\n${summary}\n`;
              }
              botResponse += `\n**📂 Path:** ${doc.path}`;
              if (aiEnabled) botResponse += `\n\n💡 **Ask me questions about this recording!**`;
            } else {
              botResponse += `No transcript available for this video.\n`;
              botResponse += `**📂 Path:** ${doc.path}\n`;
              botResponse += `Click "Open" below to watch in OneDrive.`;
            }
            if (results.length > 1) {
              lastSearchResultsRef.current = results;
              botResponse += `\n\n_${results.length - 1} more file(s) found. Ask "show all files" to see more._`;
            }
          } else {
          // Regular document flow
          const docContent = await readDocument(doc.id, doc.name);
          console.log('Document content:', docContent?.contentType, 'Length:', docContent?.content?.length);

          if (docContent) {
            newSourceDocs = [{ ...doc, webUrl: docContent.webUrl }];
            setSourceDocuments(newSourceDocs);

            // Save document for Q&A follow-up questions
            if (docContent.content && docContent.content.length > 50) {
              setActiveDocument({
                id: doc.id,
                name: doc.name,
                content: docContent.content,
                path: docContent.path,
              });
              console.log('Active document set for Q&A:', doc.name);
            }

            const ext = doc.name.split('.').pop()?.toLowerCase();
            const getFileType = (name) => {
              const e = name.split('.').pop()?.toLowerCase();
              const types = { 'docx': 'Word Document', 'xlsx': 'Excel Spreadsheet', 'pptx': 'PowerPoint', 'pdf': 'PDF', 'txt': 'Text File' };
              return types[e] || 'File';
            };

            botResponse = `📄 **${doc.name}**\n\n`;

            const isCRDoc = /CR[\s\-_]*\d+/i.test(doc.name) || /CR[\s\-_]*\d+/i.test(userMessage);

            // Show content summary if we have content
            if (docContent.content && docContent.content.length > 50) {
              const summary = await generateSummary(docContent.content, doc.name);
              if (summary) {
                botResponse += `\n**📝 ${aiEnabled ? 'Document Summary' : 'Content Preview'}:**\n${summary}\n`;
              }

              // For CR documents, automatically extract impacted areas
              if (isCRDoc && aiEnabled) {
                try {
                  const impactedAreas = await llmAnswerQuestion(
                    docContent.content,
                    'List all impacted areas, affected modules, systems, screens, APIs, and components mentioned in this document. Format as a concise bullet list. If specific module names, screen names, or API names are mentioned, include them.',
                    doc.name
                  );
                  if (impactedAreas && !impactedAreas.toLowerCase().includes('not found') && !impactedAreas.toLowerCase().includes('not mentioned')) {
                    botResponse += `\n**🎯 Impacted Areas:**\n${impactedAreas}\n`;
                  }
                } catch (impactErr) {
                  console.log('Could not extract impacted areas:', impactErr);
                }
              }
            } else {
              // No content extracted, show metadata only
              botResponse += `**📋 Document Summary:**\n`;
              botResponse += `• Type: ${getFileType(doc.name)}\n`;
              if (docContent.size) {
                const kb = docContent.size / 1024;
                const size = kb > 1024 ? (kb / 1024).toFixed(1) + ' MB' : Math.round(kb) + ' KB';
                botResponse += `• Size: ${size}\n`;
              }
              botResponse += `• Modified: ${docContent.lastModified || 'N/A'}\n`;
              if (docContent.lastModifiedBy) {
                botResponse += `• Last edited by: ${docContent.lastModifiedBy}\n`;
              }

              // For Office documents, suggest opening to view content
              if (['docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt', 'pdf'].includes(ext)) {
                botResponse += `\n_Note: Click "Open" below to view full content in OneDrive._\n`;
              }
            }

            botResponse += `\n**📂 Path:**\n${docContent.path}`;

            // Add numbered Q&A suggestions
            if (aiEnabled && docContent.content && docContent.content.length > 50) {
              const numberedQuestions = [];
              try {
                const suggestedQs = await llmAnswerQuestion(
                  docContent.content,
                  'Based on this document, suggest exactly 3 short questions a user might ask. Return ONLY the 3 questions, one per line, without numbering or bullet points.',
                  doc.name
                );
                if (suggestedQs) {
                  const questions = suggestedQs.split('\n').map(q => q.trim()).filter(q => q.length > 0).slice(0, 3);
                  numberedQuestions.push(...questions);
                }
              } catch (err) {
                console.log('Could not generate suggested questions:', err);
              }

              // CR documents also get layman and task breakdown suggestions
              if (isCRDoc) {
                numberedQuestions.push('Simplify this CR in simple words.');
                numberedQuestions.push('Break down tasks for dev assignment');
              }

              if (numberedQuestions.length > 0) {
                lastSuggestedQuestionsRef.current = numberedQuestions;
                lastFileListRef.current = [];
                botResponse += `\n\n💡 **Ask me questions about this document!**`;
                botResponse += `\n_Enter a number to ask:_`;
                numberedQuestions.forEach((q, i) => {
                  botResponse += `\n_${i + 1}. ${q}_`;
                });
              }
            }

            if (results.length > 1) {
              lastSearchResultsRef.current = results;
              botResponse += `\n\n_${results.length - 1} more file(s) found. Ask "show all files" to see more._`;
            }
          } else {
            // Could not read content, show basic info
            botResponse = `📄 **${doc.name}**\n\n`;
            botResponse += `I found the document but couldn't read its content.\n\n`;
            botResponse += `**Path:** ${doc.path}\n`;
            botResponse += `Click "Open" below to view in OneDrive.`;
            newSourceDocs = [doc];
            setSourceDocuments([doc]);
          }
          }
      } else {
        botResponse = `I couldn't find a document matching "${searchTerms.join(', ')}".\n\nTry:\n• Different keywords\n• Just the CR number (e.g., "19637")\n• "Read Arogya document"`;
      }

      if (searchTerms.length === 0) {
        botResponse = `Please specify which document you want me to read.\n\nExamples:\n• "Read CR 19637"\n• "Summarize health policy file"\n• "What is in the product sheet?"`;
      }

      setCurrentContext(null);
    }
    // Check if user wants to generate/create a PM document (independent of active document)
    else if (['generate pm', 'create pm', 'pm document', 'impact analysis document'].some(p => lowerMessage.includes(p))) {
      setIsTyping(false);
      handleGeneratePMDoc();
      return;
    }
    // Check if it's a follow-up question about the active document
    else if (aiEnabled && activeDocument && isDocumentQuestion(userMessage)) {
      console.log('Document Q&A detected for:', activeDocument.name);

      const lowerMsg = userMessage.toLowerCase();
      const isLaymanRequest = ['simplify', 'layman', 'simple terms', 'plain language', 'non-technical', 'easy to understand'].some(p => lowerMsg.includes(p));
      const isTaskBreakdown = ['break down', 'task breakdown', 'dev assignment', 'assign to dev', 'development tasks', 'dev tasks'].some(p => lowerMsg.includes(p));

      let answer;
      if (isLaymanRequest) {
        answer = await llmAnswerQuestion(
          activeDocument.content,
          'Explain this change request in simple, non-technical language that anyone can understand. Describe what is changing, why it matters, and what the end result will be for the users or the business. Keep it to 4-6 sentences. Avoid all technical jargon.',
          activeDocument.name
        );
        if (answer) {
          botResponse = `**💬 In Simple Terms — "${activeDocument.name}":**\n\n${answer}`;
        }
      } else if (isTaskBreakdown) {
        answer = await llmAnswerQuestion(
          activeDocument.content,
          'Break down this change request into specific development tasks that can be assigned to developers. For each task include: task title, brief description of what needs to be done, and the module/area it belongs to. Format as a numbered list. Be specific and actionable.',
          activeDocument.name
        );
        if (answer) {
          botResponse = `**📋 Task Breakdown for Dev Assignment — "${activeDocument.name}":**\n\n${answer}`;
        }
      } else {
        answer = await answerDocumentQuestion(userMessage);
      }

      if (answer) {
        if (!botResponse) {
          botResponse = `**📄 Answer from "${activeDocument.name}":**\n\n${answer}`;
        }
        botResponse += `\n\n💡 _Ask more questions or say "new document" to search for another file._`;

        // Keep the same source documents
        newSourceDocs = sourceDocuments;
      } else {
        botResponse = `I couldn't find an answer to that question in the document. Try rephrasing or ask a different question.`;
      }

      // Update conversation history for context
      setConversationHistory(prev => [
        ...prev.slice(-6),
        { role: 'user', content: userMessage },
        { role: 'assistant', content: botResponse }
      ]);
    }
    // Check if asking for images or videos
    else if (isMediaRequest(userMessage)) {
      const mediaType = getMediaType(userMessage);
      const searchTerms = extractSearchTerms(userMessage);
      const searchTerm = searchTerms.length > 0 ? searchTerms[0] : '';

      console.log('Media request detected:', { mediaType, searchTerms, searchTerm });

      // Use a generic term if no specific search term found
      const queryTerm = searchTerm || (mediaType === 'video' ? 'video' : 'image');
      const results = await searchMediaFiles(queryTerm, mediaType);

      console.log('Media search results:', results.length);

      if (results.length > 0) {
        // Show up to 5 only if using plural/multiple keywords (images, videos, all, list)
        const lowerMsg = userMessage.toLowerCase();
        const multipleKeywords = ['images', 'videos', 'photos', 'pictures', 'all', 'list', 'every'];
        const showMultiple = multipleKeywords.some(kw => lowerMsg.includes(kw));
        const requestedMediaCount = getRequestedCount(userMessage);
        const maxFiles = requestedMediaCount || (showMultiple ? 5 : 1);

        console.log('Multiple check:', { lowerMsg, showMultiple, maxFiles });

        const displayResults = results.slice(0, maxFiles);

        newSourceDocs = displayResults;
        setSourceDocuments(displayResults);
        lastFileListRef.current = displayResults;

        const mediaLabel = mediaType === 'video' ? 'video(s)' : mediaType === 'image' ? 'image(s)' : 'media file(s)';

        if (showMultiple || requestedMediaCount) {
          botResponse = `I found **${results.length} ${mediaLabel}**${searchTerm ? ` related to "${searchTerm}"` : ''}${requestedMediaCount ? ` (showing ${displayResults.length})` : ''}:\n\n`;
        } else {
          botResponse = `Here's ${mediaType === 'video' ? 'a video' : 'an image'}${searchTerm ? ` related to "${searchTerm}"` : ''}:\n\n`;
        }

        // Media will be rendered separately in the message
        displayResults.forEach((doc, idx) => {
          if (showMultiple || requestedMediaCount) {
            botResponse += `**${idx + 1}. ${doc.name}**\n`;
          } else {
            botResponse += `**${doc.name}**\n`;
          }
          botResponse += `📂 ${doc.path}\n\n`;
        });

        if (results.length > maxFiles) {
          botResponse += `_...and ${results.length - maxFiles} more ${mediaLabel} available._\n`;
        }
        if (displayResults.length > 1) {
          botResponse += `\n💡 **Type a number (e.g. "1") to summarize that ${mediaType === 'video' ? 'video' : 'file'}.**`;
        }
      } else {
        botResponse = `I couldn't find any ${mediaType === 'video' ? 'videos' : mediaType === 'image' ? 'images' : 'media files'}${searchTerm ? ` matching "${searchTerm}"` : ''} in your OneDrive.\n\nTry:\n• Different keywords\n• "Show my images" or "Show my videos"`;
      }

      setCurrentContext(null);
    }
    // Check if asking for source
    else if (lowerMessage.includes('source') || (lowerMessage.includes('share') && !lowerMessage.includes('sharepoint'))) {
      if (sourceDocuments.length > 0) {
        botResponse = await handleSourceRequest();
      } else if (currentContext?.lastSourceFile) {
        // Search for the source file
        const doc = await findSourceDocument(currentContext.lastSourceFile);
        if (doc) {
          newSourceDocs = [doc];
          setSourceDocuments([doc]);
          botResponse = `**Source:**\n📄 ${doc.name}\n📂 Location: ${doc.path}`;
        } else {
          botResponse = `Source: ${currentContext.lastSourceFile}.docx – Check the relevant department folder.`;
        }
      } else {
        botResponse = "I don't have a specific source for the previous response. Could you ask a question first?";
      }
    }
    // Check if it's a file search request
    else if (isFileSearchRequest(userMessage)) {
      const searchTerms = extractSearchTerms(userMessage);

      if (searchTerms.length > 0) {
        let allResults = [];

        // Search for each term
        for (const term of searchTerms) {
          const results = await searchDocuments(term);
          allResults = [...allResults, ...results];
        }

        // Remove duplicates by id
        const uniqueResults = allResults.filter((file, index, self) =>
          index === self.findIndex(f => f.id === file.id)
        );

        // Filter based on user request (folders, documents only, etc.)
        const filteredResults = applyFileFilters(uniqueResults, userMessage);

        if (filteredResults.length > 0) {
          // Determine how many files to show based on user request
          const showMultiple = wantsMultipleFiles(userMessage);
          const requestedCount = getRequestedCount(userMessage);
          const maxFiles = requestedCount || (showMultiple ? 5 : 1);
          const displayResults = filteredResults.slice(0, maxFiles);

          newSourceDocs = displayResults;
          setSourceDocuments(displayResults);

          // Store the displayed list for number selection
          lastFileListRef.current = displayResults;

          // Helper to get file type description
          const getFileTypeDesc = (name) => {
            const ext = name.split('.').pop()?.toLowerCase();
            const typeMap = {
              'docx': 'Word Document', 'doc': 'Word Document',
              'xlsx': 'Excel Spreadsheet', 'xls': 'Excel Spreadsheet',
              'pptx': 'PowerPoint Presentation', 'ppt': 'PowerPoint Presentation',
              'pdf': 'PDF Document', 'txt': 'Text File',
              'png': 'Image', 'jpg': 'Image', 'jpeg': 'Image', 'gif': 'Image',
              'mp4': 'Video', 'mov': 'Video', 'avi': 'Video',
              'mp3': 'Audio', 'wav': 'Audio',
              'zip': 'Archive', 'rar': 'Archive'
            };
            return typeMap[ext] || 'File';
          };

          // Helper to format size
          const formatSize = (bytes) => {
            if (!bytes) return '';
            const kb = bytes / 1024;
            return kb > 1024 ? (kb / 1024).toFixed(1) + ' MB' : Math.round(kb) + ' KB';
          };

          if (showMultiple || requestedCount) {
            botResponse = `I found **${filteredResults.length} file(s)** related to "${searchTerms.join(', ')}"${requestedCount ? ` (showing ${displayResults.length})` : ''}:\n\n`;
            displayResults.forEach((doc, idx) => {
              const fileType = getFileTypeDesc(doc.name);
              const size = formatSize(doc.size);
              botResponse += `**${idx + 1}. ${doc.name}**\n`;
              botResponse += `   📄 ${fileType}${size ? ` • ${size}` : ''}\n`;
              botResponse += `   📂 ${doc.path}\n`;
              botResponse += `   📅 Modified: ${doc.date || 'N/A'}${doc.sharedBy ? ` • By: ${doc.sharedBy}` : ''}\n\n`;
            });
            if (filteredResults.length > maxFiles) {
              botResponse += `_...and ${filteredResults.length - maxFiles} more files._\n`;
            }
            botResponse += `\n💡 **Type a number (e.g. "3") to summarize that document.**`;
          } else {
            // Single file response with summary
            const doc = displayResults[0];
            const fileType = getFileTypeDesc(doc.name);
            const size = formatSize(doc.size);

            botResponse = `Here's the file for "${searchTerms.join(', ')}":\n\n`;
            botResponse += `📄 **${doc.name}**\n\n`;
            botResponse += `**Summary:**\n`;
            botResponse += `• Type: ${fileType}\n`;
            if (size) botResponse += `• Size: ${size}\n`;
            botResponse += `• Modified: ${doc.date || 'N/A'}\n`;
            if (doc.sharedBy) botResponse += `• Owner: ${doc.sharedBy}\n`;
            botResponse += `\n**Path:**\n📂 ${doc.path}\n`;

            if (filteredResults.length > 1) {
              botResponse += `\n_${filteredResults.length - 1} more file(s) available. Ask "show all files" or "list files" to see more._`;
            }
          }
        } else {
          botResponse = `I couldn't find any files matching "${searchTerms.join(', ')}" in your OneDrive.\n\nPlease check:\n• The file name spelling\n• If the file exists in your accessible folders\n• Try different search terms`;
        }
      } else {
        botResponse = "Please specify what files you're looking for. For example:\n• 'Show Arogya Sanjeevani files'\n• 'Find Top up product documents'\n• 'Search health policy'";
      }

      setCurrentContext(null);
    }
    // Check if we have an active context for follow-up
    else if (currentContext) {
      const followUpResult = processFollowUp(userMessage, currentContext);

      if (followUpResult) {
        botResponse = followUpResult.response;

        // Search for source document
        if (followUpResult.sourceFile) {
          const doc = await findSourceDocument(followUpResult.sourceFile);
          if (doc) {
            newSourceDocs = [doc];
            setSourceDocuments([doc]);
          }
          setCurrentContext({
            ...currentContext,
            lastSourceFile: followUpResult.sourceFile,
            resolved: true
          });
        }
      } else {
        // Try finding new topic
        const match = findKnowledgeMatch(userMessage);
        if (match) {
          botResponse = match.initialResponse;
          setCurrentContext({ ...match, lastSourceFile: null });
          setSourceDocuments([]);
        } else {
          // Try file search as fallback
          const searchTerms = extractSearchTerms(userMessage);
          if (searchTerms.length > 0) {
            const results = await searchDocuments(searchTerms[0]);
            const filteredResultsCtx = applyFileFilters(results, userMessage);
            if (filteredResultsCtx.length > 0) {
              const showMultiple = wantsMultipleFiles(userMessage);
              const requestedCount = getRequestedCount(userMessage);
              const maxFiles = requestedCount || (showMultiple ? 5 : 1);
              const displayResults = filteredResultsCtx.slice(0, maxFiles);
              newSourceDocs = displayResults;
              setSourceDocuments(displayResults);
              lastFileListRef.current = displayResults;

              const getFileType = (name) => {
                const ext = name.split('.').pop()?.toLowerCase();
                const types = { 'docx': 'Word', 'xlsx': 'Excel', 'pptx': 'PowerPoint', 'pdf': 'PDF', 'txt': 'Text' };
                return types[ext] || 'File';
              };

              if (showMultiple || requestedCount) {
                botResponse = `I found **${filteredResultsCtx.length} file(s)** that might be relevant:\n\n`;
                displayResults.forEach((doc, idx) => {
                  botResponse += `**${idx + 1}. ${doc.name}**\n`;
                  botResponse += `   📄 ${getFileType(doc.name)} • 📂 ${doc.path}\n\n`;
                });
                if (filteredResultsCtx.length > maxFiles) {
                  botResponse += `_...and ${filteredResultsCtx.length - maxFiles} more files._\n`;
                }
                botResponse += `\n💡 **Type a number (e.g. "3") to summarize that document.**`;
              } else {
                const doc = displayResults[0];
                botResponse = `Here's a relevant file:\n\n📄 **${doc.name}**\n\n`;
                botResponse += `**Summary:**\n• Type: ${getFileType(doc.name)}\n• Path: ${doc.path}`;
                if (filteredResultsCtx.length > 1) {
                  botResponse += `\n\n_${filteredResultsCtx.length - 1} more file(s) available. Ask "show all files" to see more._`;
                }
              }
            } else {
              botResponse = "I'm not sure I understand. Could you please rephrase your question or try searching for specific file names?";
            }
          } else {
            botResponse = "I'm not sure I understand. Could you please rephrase your question or provide more details?";
          }
        }
      }
    }
    // Check for new topic
    else {
      const match = findKnowledgeMatch(userMessage);

      if (match) {
        botResponse = match.initialResponse;
        setCurrentContext({ ...match, lastSourceFile: null });
        setSourceDocuments([]);
      } else {
        // Try file search
        const searchTerms = extractSearchTerms(userMessage);
        if (searchTerms.length > 0) {
          const results = await searchDocuments(searchTerms[0]);
          const filteredResultsElse = applyFileFilters(results, userMessage);
          if (filteredResultsElse.length > 0) {
            const showMultiple = wantsMultipleFiles(userMessage);
            const requestedCount = getRequestedCount(userMessage);
            const maxFiles = requestedCount || (showMultiple ? 5 : 1);
            const displayResults = filteredResultsElse.slice(0, maxFiles);
            newSourceDocs = displayResults;
            setSourceDocuments(displayResults);
            lastFileListRef.current = displayResults;

            const getFileType = (name) => {
              const ext = name.split('.').pop()?.toLowerCase();
              const types = { 'docx': 'Word Document', 'xlsx': 'Excel Spreadsheet', 'pptx': 'PowerPoint', 'pdf': 'PDF', 'txt': 'Text File' };
              return types[ext] || 'File';
            };
            const formatSize = (bytes) => {
              if (!bytes) return '';
              const kb = bytes / 1024;
              return kb > 1024 ? (kb / 1024).toFixed(1) + ' MB' : Math.round(kb) + ' KB';
            };

            if (showMultiple || requestedCount) {
              botResponse = `I found **${filteredResultsElse.length} file(s)** matching your query:\n\n`;
              displayResults.forEach((doc, idx) => {
                const size = formatSize(doc.size);
                botResponse += `**${idx + 1}. ${doc.name}**\n`;
                botResponse += `   📄 ${getFileType(doc.name)}${size ? ` • ${size}` : ''}\n`;
                botResponse += `   📂 ${doc.path}\n`;
                botResponse += `   📅 ${doc.date || 'N/A'}${doc.sharedBy ? ` • ${doc.sharedBy}` : ''}\n\n`;
              });
              if (filteredResultsElse.length > maxFiles) {
                botResponse += `_...and ${filteredResultsElse.length - maxFiles} more files._\n`;
              }
              botResponse += `\n💡 **Type a number (e.g. "3") to summarize that document.**`;
            } else {
              const doc = displayResults[0];
              const size = formatSize(doc.size);
              botResponse = `Here's the file matching your query:\n\n`;
              botResponse += `📄 **${doc.name}**\n\n`;
              botResponse += `**Summary:**\n`;
              botResponse += `• Type: ${getFileType(doc.name)}\n`;
              if (size) botResponse += `• Size: ${size}\n`;
              botResponse += `• Modified: ${doc.date || 'N/A'}\n`;
              if (doc.sharedBy) botResponse += `• Owner: ${doc.sharedBy}\n`;
              botResponse += `\n**Path:**\n📂 ${doc.path}`;
              if (filteredResultsElse.length > 1) {
                botResponse += `\n\n_${filteredResultsElse.length - 1} more file(s) available. Ask "show all files" or "list files" to see more._`;
              }
            }
          } else {
            botResponse = `I couldn't find files matching "${userMessage}".\n\nTry:\n• Different keywords\n• Product names like "Arogya", "Sanjeevani", "Top up"\n• Or ask about company policies`;
          }
        } else {
          // If AI is enabled, try getting a conversational response
          if (aiEnabled) {
            try {
              console.log('Getting AI chat response for:', userMessage, 'using:', aiProvider);
              botResponse = await llmGetChatResponse(userMessage, conversationHistory);
              // Update conversation history
              setConversationHistory(prev => [
                ...prev.slice(-8), // Keep last 8 messages
                { role: 'user', content: userMessage },
                { role: 'assistant', content: botResponse }
              ]);
            } catch (error) {
              console.log('AI chat failed:', error);
              botResponse = `I'm sorry, I couldn't process your request. Try asking:
• "Show Arogya Sanjeevani file"
• "Find Top up product"
• "Summarize health policy"`;
            }
          } else {
            botResponse = `I'm sorry, I couldn't find specific information about "${userMessage}".

Try asking:
• "Show Arogya Sanjeevani file"
• "Find Top up product"
• "List all health policy files"
• Or ask about project approval, leave policy, expenses`;
          }
        }
      }
    }

    // Simulate typing delay
    await new Promise(resolve => setTimeout(resolve, 500 + Math.random() * 500));

    setIsTyping(false);
    setMessages(prev => [...prev, {
      type: 'bot',
      text: botResponse,
      sources: newSourceDocs
    }]);
  };

  const handleKeyPress = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const openDocument = (url) => {
    if (url) {
      window.open(url, '_blank');
    }
  };

  const openFileLocation = async (doc) => {
    if (!doc) return;
    console.log('Opening file location for:', doc.name, 'id:', doc.id);

    try {
      // Get fresh access token using same scopes as rest of app
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      // Use Graph API to get the parent folder's webUrl
      const { getParentFolderUrl } = await import('../services/graphService.js');
      const folderUrl = await getParentFolderUrl(doc.id, tokenResponse.accessToken, doc.parentId);
      if (folderUrl) {
        console.log('Opening folder:', folderUrl);
        window.open(folderUrl, '_blank');
        return;
      }
    } catch (err) {
      console.error('Error getting parent folder:', err);
    }

    // Fallback: Open the user's OneDrive root
    if (doc.webUrl) {
      try {
        const url = new URL(doc.webUrl);
        // Extract personal/site path
        const match = url.pathname.match(/^(\/personal\/[^/]+|\/sites\/[^/]+)/);
        if (match) {
          const siteUrl = `${url.origin}${match[1]}/_layouts/15/onedrive.aspx`;
          console.log('Opening OneDrive root:', siteUrl);
          window.open(siteUrl, '_blank');
          return;
        }
      } catch (e) {
        console.error('URL parse error:', e);
      }
    }
  };

  // ===== UPLOAD HANDLERS =====

  const validateFile = (file) => {
    const ext = file.name.split('.').pop()?.toLowerCase();
    if (!SUPPORTED_UPLOAD_EXTENSIONS.includes(ext)) {
      return `Unsupported file type ".${ext}". Supported: ${SUPPORTED_UPLOAD_EXTENSIONS.join(', ')}`;
    }
    if (file.size > MAX_UPLOAD_SIZE) {
      return `File too large (${(file.size / 1024 / 1024).toFixed(1)}MB). Maximum: ${MAX_UPLOAD_SIZE / 1024 / 1024}MB`;
    }
    return null;
  };

  const fetchFoldersForPicker = async (folderId) => {
    const tokenResponse = await instance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0],
    });
    return await getFoldersOnly(tokenResponse.accessToken, folderId);
  };

  const handleFileSelected = (file) => {
    const error = validateFile(file);
    if (error) {
      setMessages(prev => [...prev, { type: 'bot', text: `Upload failed: ${error}` }]);
      return;
    }
    setPendingUploadFile(file);
    setShowFolderPicker(true);
  };

  const handleFolderSelected = async (folder) => {
    if (!pendingUploadFile) return;

    const file = pendingUploadFile;
    setPendingUploadFile(null);

    setMessages(prev => [...prev, { type: 'user', text: `Upload "${file.name}" to ${folder.name}` }]);
    setIsTyping(true);
    setUploadProgress(0);

    try {
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      const uploadedItem = await uploadFileToOneDrive(
        tokenResponse.accessToken,
        folder.id,
        file,
        (progress) => setUploadProgress(progress),
      );

      setUploadProgress(null);

      // Track recently uploaded file so search can find it immediately
      recentUploadsRef.current = [
        { id: uploadedItem.id, name: file.name, webUrl: uploadedItem.webUrl, parentFolderId: folder.id, parentFolderName: folder.name },
        ...recentUploadsRef.current,
      ].slice(0, 20);

      let botResponse = `File **"${file.name}"** uploaded successfully to **${folder.name}**.\n\n`;
      let newSourceDocs = [];

      try {
        const docContent = await getDocumentContent(
          tokenResponse.accessToken,
          uploadedItem.id,
          file.name
        );

        if (docContent) {
          newSourceDocs = [{
            id: uploadedItem.id,
            name: file.name,
            path: docContent.path || file.name,
            webUrl: uploadedItem.webUrl || docContent.webUrl,
            date: new Date().toISOString().split('T')[0],
            type: getFileType(file.name),
            isFolder: false,
          }];
          setSourceDocuments(newSourceDocs);

          if (docContent.content && docContent.content.length > 50) {
            setActiveDocument({
              id: uploadedItem.id,
              name: file.name,
              content: docContent.content,
              path: docContent.path,
            });

            botResponse += `📄 **${file.name}**\n\n`;

            // AI summary
            const summary = await generateSummary(docContent.content, file.name);
            if (summary) {
              botResponse += `**📝 ${aiEnabled ? 'Document Summary' : 'Content Preview'}:**\n${summary}\n`;
            }

            // For CR documents, extract impacted areas
            const isCRDoc = /CR[\s\-_]*\d+/i.test(file.name);
            if (isCRDoc && aiEnabled) {
              try {
                const impactedAreas = await llmAnswerQuestion(
                  docContent.content,
                  'List all impacted areas, affected modules, systems, screens, APIs, and components mentioned in this document. Format as a concise bullet list. If specific module names, screen names, or API names are mentioned, include them.',
                  file.name
                );
                if (impactedAreas && !impactedAreas.toLowerCase().includes('not found') && !impactedAreas.toLowerCase().includes('not mentioned')) {
                  botResponse += `\n**🎯 Impacted Areas:**\n${impactedAreas}\n`;
                }
              } catch (impactErr) {
                console.log('Could not extract impacted areas:', impactErr);
              }
            }

            botResponse += `\n**📂 Path:**\n${docContent.path}`;

            // Numbered Q&A suggestions
            if (aiEnabled) {
              const numberedQuestions = [];
              try {
                const suggestedQs = await llmAnswerQuestion(
                  docContent.content,
                  'Based on this document, suggest exactly 3 short questions a user might ask. Return ONLY the 3 questions, one per line, without numbering or bullet points.',
                  file.name
                );
                if (suggestedQs) {
                  const questions = suggestedQs.split('\n').map(q => q.trim()).filter(q => q.length > 0).slice(0, 3);
                  numberedQuestions.push(...questions);
                }
              } catch (err) {
                console.log('Could not generate suggested questions:', err);
              }

              if (isCRDoc) {
                numberedQuestions.push('Simplify this CR in layman terms');
                numberedQuestions.push('Break down tasks for dev assignment');
              }

              if (numberedQuestions.length > 0) {
                lastSuggestedQuestionsRef.current = numberedQuestions;
                lastFileListRef.current = [];
                botResponse += `\n\n💡 **Ask me questions about this document!**`;
                botResponse += `\n_Enter a number to ask:_`;
                numberedQuestions.forEach((q, i) => {
                  botResponse += `\n_${i + 1}. ${q}_`;
                });
              }
            }
          } else {
            // No content extracted, show metadata
            const ext = file.name.split('.').pop()?.toLowerCase();
            const fileTypeLabel = { 'docx': 'Word Document', 'xlsx': 'Excel Spreadsheet', 'pptx': 'PowerPoint', 'pdf': 'PDF', 'txt': 'Text File' }[ext] || 'File';
            botResponse += `**📋 Document Summary:**\n`;
            botResponse += `• Type: ${fileTypeLabel}\n`;
            if (docContent.size) {
              const kb = docContent.size / 1024;
              const size = kb > 1024 ? (kb / 1024).toFixed(1) + ' MB' : Math.round(kb) + ' KB';
              botResponse += `• Size: ${size}\n`;
            }
            botResponse += `\n**📂 Path:**\n${docContent.path}`;
            botResponse += `\n_Click "Open" below to view full content in OneDrive._`;
          }
        } else {
          botResponse += `Content could not be extracted for summarization. You can view the file in OneDrive.`;
          newSourceDocs = [{
            id: uploadedItem.id,
            name: file.name,
            webUrl: uploadedItem.webUrl,
            date: new Date().toISOString().split('T')[0],
          }];
          setSourceDocuments(newSourceDocs);
        }
      } catch (contentErr) {
        console.error('Post-upload content extraction failed:', contentErr);
        botResponse += `File uploaded but content extraction failed. Open in OneDrive to view.`;
        newSourceDocs = [{
          id: uploadedItem.id,
          name: file.name,
          webUrl: uploadedItem.webUrl,
          date: new Date().toISOString().split('T')[0],
        }];
        setSourceDocuments(newSourceDocs);
      }

      setIsTyping(false);
      setMessages(prev => [...prev, { type: 'bot', text: botResponse, sources: newSourceDocs }]);
    } catch (uploadErr) {
      console.error('Upload error:', uploadErr);
      setUploadProgress(null);
      setIsTyping(false);
      setMessages(prev => [...prev, {
        type: 'bot',
        text: `Upload failed: ${uploadErr.message}\n\nPlease try again or check your permissions.`,
      }]);
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
    const files = e.dataTransfer?.files;
    if (files && files.length > 0) {
      handleFileSelected(files[0]);
    }
  };

  const handleAttachClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileInputChange = (e) => {
    const file = e.target.files?.[0];
    if (file) {
      handleFileSelected(file);
    }
    e.target.value = '';
  };

  // ===== PM DOCUMENT GENERATION =====

  const handleGeneratePMDoc = async () => {
    // If an active document exists, use AI to pre-fill form fields
    if (activeDocument && activeDocument.content) {
      setMessages(prev => [...prev, { type: 'bot', text: 'Extracting details from the document for PM Impact Analysis form...' }]);
      setIsTyping(true);

      try {
        const extractPrompt = `Extract the following from this document and return ONLY valid JSON (no markdown, no code blocks):
{
  "pmNumber": "the PM number (digits only, e.g. 13366) - this is different from the CR number",
  "crNumber": "the CR number (digits only, e.g. 19078)",
  "issueDescription": "brief description of the issue or change",
  "systemImpacts": [{"application": "app name", "components": "file/component names", "remarks": "what is impacted and why"}],
  "risks": [{"assumptions": "key assumptions", "risks": "identified risks", "otherImpacts": "other impacts", "remarks": "additional notes"}]
}`;

        const aiResult = await llmAnswerQuestion(activeDocument.content, extractPrompt, activeDocument.name);

        let prefill = {};
        if (aiResult) {
          try {
            const cleaned = aiResult.replace(/```json?\s*/g, '').replace(/```\s*/g, '').trim();
            prefill = JSON.parse(cleaned);
          } catch (parseErr) {
            console.log('Could not parse AI prefill JSON:', parseErr);
          }
        }

        setPmDocFormData(prefill);
        setShowPMDocForm(true);
      } catch (err) {
        console.error('PM doc prefill error:', err);
        setPmDocFormData({});
        setShowPMDocForm(true);
      }
      setIsTyping(false);
    } else {
      // No active document — open blank form
      setPmDocFormData({});
      setShowPMDocForm(true);
    }
  };

  const handlePMDocGenerate = async (formData, mode) => {
    setShowPMDocForm(false);

    try {
      const { blob, fileName } = await generatePMDocument(formData);

      // Always download
      saveAs(blob, fileName);

      setMessages(prev => [...prev, {
        type: 'bot',
        text: `**${fileName}** has been downloaded.\n\n${mode === 'upload' ? 'Select a folder to upload...' : ''}`,
      }]);

      // If upload requested, open folder picker
      if (mode === 'upload') {
        const file = new File([blob], fileName, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        setPendingUploadFile(file);
        setShowFolderPicker(true);
      }
    } catch (err) {
      console.error('PM document generation error:', err);
      setMessages(prev => [...prev, {
        type: 'bot',
        text: `Failed to generate the document: ${err.message}`,
      }]);
    }
  };

  // Suggested questions for quick access
  const suggestedQuestions = [
    "Tell me about Arogya Sanjeevani Product",
    "Summarize Production Deployment SOP",
    "Read CR20049 document",
    "Summarize CR20049 video"
  ];

  // Get current time for message timestamps
  const formatTime = (date = new Date()) => {
    return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: true });
  };

  return (
    <div className={`chatbot-wrapper ${isDarkMode ? 'dark-mode' : 'light-mode'}`}>
      {/* Mobile Sidebar Overlay */}
      {showMobileSidebar && <div className="mobile-sidebar-overlay" onClick={() => setShowMobileSidebar(false)} />}

      {/* Left Sidebar */}
      <aside className={`chat-sidebar ${showMobileSidebar ? 'mobile-open' : ''}`}>
        <div className="sidebar-tabs">
          <button className="tab-btn active">Recent</button>
          <button className="mobile-sidebar-close" onClick={() => setShowMobileSidebar(false)}>✕</button>
        </div>

        <div className="sidebar-documents">
          {recentDocuments.length > 0 ? (
            recentDocuments.slice(0, 8).map((doc, idx) => (
              <div
                key={idx}
                className="sidebar-doc-item"
                onClick={() => openDocument(doc.webUrl)}
              >
                <span className="doc-icon">📄</span>
                <div className="doc-info">
                  <span className="doc-name">{doc.name}</span>
                  <br/>
                  <span className="doc-date">📅{doc.date || 'Recently accessed'}</span>
                </div>
              </div>
            ))
          ) : (
            <div className="sidebar-empty">
              <p>No recent documents</p>
              <span>Your recently accessed files will appear here</span>
            </div>
          )}
        </div>

        <div className="fs-sidebar-footer">
          
          <div className="fs-settings-btn">
            <span className="fs-user-avatar">👤</span>
            <span className="fs-user-name">{accounts[0]?.name || accounts[0]?.username}</span>
          </div>
          {/* <button className="sidebar-action logout-btn" onClick={handleLogout}>
            Sign out
          </button> */}
           <button className="fs-logout-btn" onClick={handleLogout}>
            <span>⬅️</span> Sign out
          </button>
          <button className="sidebar-action settings-btn" onClick={() => setShowSettings(true)}>
            <span>⚙️</span> Settings
          </button>
        </div>
      </aside>

      {/* Main Chat Area */}
      <main
        className={`chat-main ${isDragOver ? 'drag-over' : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        {isDragOver && (
          <div className="drag-overlay">
            <div className="drag-overlay-content">
              <span className="drag-overlay-icon">&#128206;</span>
              <p>Drop file to upload to OneDrive</p>
            </div>
          </div>
        )}
        <div className="chat-header-dark">
          <button className="mobile-menu-btn" onClick={() => setShowMobileSidebar(true)}>
            <span>☰</span>
          </button>
          <div className="header-left">
            <h2>Knowledge Centre</h2>
            <span className="connection-status">
              {isAuthenticated ? '● Connected' : '○ Not signed in'}
            </span>
          </div>
        </div>

        <div className="chat-messages-dark">
          {messages.map((msg, idx) => (
            <div key={idx} className={`message-dark ${msg.type}`}>
              {msg.type === 'bot' && (
                <div className="message-avatar">
                  <span>🤖</span>
                </div>
              )}
              <div className="message-bubble">
                <div className="message-text-dark">
                  {(msg.text || '').split('\n').map((line, i) => (
                    <p key={i}>{line}</p>
                  ))}
                </div>
                {/* Display media (images and videos) */}
                {msg.sources && msg.sources.length > 0 && msg.sources.some(doc => doc.isImage || doc.isVideo) && (
                  <div className="message-media-dark">
                    {msg.sources.filter(doc => doc.isImage || doc.isVideo).map((doc, i) => (
                      <div key={i} className="media-item-dark">
                        {doc.isImage && (doc.thumbnailUrl || doc.downloadUrl) && (
                          <div className="image-preview-dark">
                            <img
                              src={doc.downloadUrl || doc.thumbnailUrl}
                              alt={doc.name}
                              onClick={() => openDocument(doc.webUrl)}
                              onError={(e) => {
                                e.target.style.display = 'none';
                                e.target.nextSibling.style.display = 'block';
                              }}
                            />
                            <div className="image-fallback-dark" style={{ display: 'none' }}>
                              <span>🖼️</span>
                              <p>{doc.name}</p>
                            </div>
                          </div>
                        )}
                        {doc.isVideo && (
                          <div className="video-preview-dark">
                            {doc.embedUrl ? (
                              <div className="video-embed-dark">
                                <iframe
                                  src={doc.embedUrl}
                                  title={doc.name}
                                  frameBorder="0"
                                  allowFullScreen
                                  allow="autoplay; encrypted-media"
                                />
                              </div>
                            ) : doc.downloadUrl ? (
                              <>
                                <video
                                  controls
                                  controlsList="nodownload"
                                  poster={doc.thumbnailUrl}
                                  preload="auto"
                                  playsInline
                                  onError={(e) => {
                                    e.target.parentElement.querySelector('.video-fallback-dark').style.display = 'flex';
                                    e.target.style.display = 'none';
                                  }}
                                >
                                  <source src={doc.downloadUrl} type={`video/${doc.name.split('.').pop()?.toLowerCase() || 'mp4'}`} />
                                </video>
                                <div className="video-fallback-dark" style={{ display: 'none' }} onClick={() => openDocument(doc.webUrl)}>
                                  <div className="play-btn">▶</div>
                                  <p>Click to play in OneDrive</p>
                                </div>
                              </>
                            ) : (
                              <div className="video-thumb-dark" onClick={() => openDocument(doc.webUrl)}>
                                {doc.thumbnailUrl ? (
                                  <img src={doc.thumbnailUrl} alt={doc.name} />
                                ) : (
                                  <div className="video-placeholder-dark">🎬</div>
                                )}
                                <div className="play-overlay-dark">▶</div>
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                )}
                {/* Display file search results in card layout */}
                {msg.sources && msg.sources.length > 0 && !msg.sources.some(doc => doc.isImage || doc.isVideo) && (
                  <div className="message-file-results">
                    {msg.sources.map((doc, i) => {
                      const getFileIcon = (name, isFolder) => {
                        if (isFolder) return '📁';
                        if (!name) return '📄';
                        const ext = name.split('.').pop()?.toLowerCase();
                        const icons = {
                          'docx': '📄', 'doc': '📄',
                          'xlsx': '📊', 'xls': '📊',
                          'pdf': '📕',
                          'pptx': '📽️', 'ppt': '📽️',
                          'png': '🖼️', 'jpg': '🖼️', 'jpeg': '🖼️', 'gif': '🖼️',
                          'mp4': '🎬', 'mov': '🎬', 'avi': '🎬',
                          'mp3': '🎵', 'wav': '🎵',
                          'zip': '📦', 'rar': '📦',
                        };
                        return icons[ext] || '📄';
                      };
                      const formatSize = (bytes) => {
                        if (!bytes) return '';
                        const kb = bytes / 1024;
                        return kb > 1024 ? (kb / 1024).toFixed(1) + ' MB' : Math.round(kb) + ' KB';
                      };
                      return (
                        <div key={i} className="file-result-card">
                          <div className="file-icon">{getFileIcon(doc.name, doc.isFolder)}</div>
                          <div className="file-info">
                            <h4 className="file-name">{doc.name}</h4>
                            <div className="file-path">📂 {doc.path || doc.name}</div>
                            <div className="file-metadata">
                              {doc.date && <span className="meta-item">📅 {doc.date}</span>}
                              {doc.sharedBy && <span className="meta-item">👤 {doc.sharedBy}</span>}
                              {doc.size && <span className="meta-item">💾 {formatSize(doc.size)}</span>}
                              <span className="meta-badge">{doc.isFolder ? 'Folder' : doc.type || 'File'}</span>
                            </div>
                          </div>
                          <div className="file-actions">
                            <button
                              className="file-open-btn"
                              onClick={() => openDocument(doc.webUrl)}
                              title="Open file"
                            >
                              Open
                            </button>
                            <button
                              className="file-location-btn"
                              onClick={() => openFileLocation(doc)}
                              title="Open file location"
                            >
                              📁
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
                <span className="message-time">{formatTime()}</span>
              </div>
              {msg.type === 'user' && (
                <div className="message-avatar user-avatar-dark">
                  <span>👤</span>
                </div>
              )}
            </div>
          ))}
          {isTyping && (
            <div className="message-dark bot">
              <div className="message-avatar">
                <span>🤖</span>
              </div>
              <div className="message-bubble">
                <div className="typing-dots">
                  <span></span>
                  <span></span>
                  <span></span>
                </div>
              </div>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>

        {/* Suggested Questions */}
        {messages.length <= 2 && (
          <div className="suggested-questions">
            <h3>Suggested questions:</h3>
            <div className="question-cards">
              {suggestedQuestions.map((question, idx) => (
                <button
                  key={idx}
                  className="question-card"
                  onClick={() => handleSend(question)}
                >
                  {question}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Active document indicator for Q&A */}
        {activeDocument && aiEnabled && (
          <div className="active-doc-bar">
            <span>📄 Asking about: <strong>{activeDocument.name}</strong></span>
            <button onClick={() => setActiveDocument(null)}>✕</button>
          </div>
        )}

        {uploadProgress !== null && (
          <div className="upload-progress-bar">
            <div className="upload-progress-fill" style={{ width: `${uploadProgress}%` }} />
            <span className="upload-progress-text">Uploading... {uploadProgress}%</span>
          </div>
        )}

        <div className="chat-input-section">
          <input
            type="file"
            ref={fileInputRef}
            style={{ display: 'none' }}
            accept={SUPPORTED_UPLOAD_EXTENSIONS.map(ext => `.${ext}`).join(',')}
            onChange={handleFileInputChange}
          />
          <button
            className="chat-attach-btn"
            onClick={handleAttachClick}
            disabled={isTyping || uploadProgress !== null}
            title="Upload a document"
          >
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M21.44 11.05l-9.19 9.19a6 6 0 01-8.49-8.49l9.19-9.19a4 4 0 015.66 5.66l-9.2 9.19a2 2 0 01-2.83-2.83l8.49-8.48"/>
            </svg>
          </button>
          <div className="chat-input-container">
            <input
              type="text"
              value={input}
              onChange={(e) => setInput(e.target.value)}
              onKeyPress={handleKeyPress}
              placeholder={activeDocument ? `Ask about ${activeDocument.name}...` : "Type your question here..."}
              className="chat-input-field"
            />
          </div>
          <button
            className="chat-send-btn"
            onClick={handleSend}
            disabled={!input.trim() || isTyping}
          >
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M22 2L11 13"/>
              <path d="M22 2l-7 20-4-9-9-4 20-7z"/>
            </svg>
          </button>
        </div>
      </main>

      {/* Folder Picker Modal */}
      <FolderPicker
        isOpen={showFolderPicker}
        onClose={() => { setShowFolderPicker(false); setPendingUploadFile(null); }}
        onSelect={handleFolderSelected}
        fetchFolders={fetchFoldersForPicker}
      />

      {/* PM Document Form Modal */}
      <PMDocumentForm
        isOpen={showPMDocForm}
        onClose={() => setShowPMDocForm(false)}
        onGenerate={handlePMDocGenerate}
        initialData={pmDocFormData}
      />

      {/* Settings Modal */}
      {showSettings && (
        <div className="settings-overlay" onClick={() => setShowSettings(false)}>
          <div className="settings-modal" onClick={(e) => e.stopPropagation()}>
            <div className="settings-header">
              <h2>Settings</h2>
              <button className="settings-close" onClick={() => setShowSettings(false)}>
                ✕
              </button>
            </div>
            <div className="settings-content">
              <div className="settings-section">
                <h3>AI Provider</h3>
                <p className="settings-description">Select the AI model to use for chat responses</p>
                {availableLLMs.length > 0 ? (
                  <select
                    value={selectedLLM}
                    onChange={handleLLMChange}
                    className="settings-select"
                  >
                    {availableLLMs.map((llm) => (
                      <option key={llm.id} value={llm.id}>
                        {llm.icon} {llm.name}
                      </option>
                    ))}
                  </select>
                ) : (
                  <p className="settings-no-llm">No AI providers available</p>
                )}
              </div>
              <div className="settings-section">
                <h3>Appearance</h3>
                <p className="settings-description">Customize the look and feel</p>
                <button className="settings-theme-btn" onClick={toggleTheme}>
                  <span>{isDarkMode ? '☀️' : '🌙'}</span>
                  {isDarkMode ? 'Switch to Light Mode' : 'Switch to Dark Mode'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default ChatBot;
