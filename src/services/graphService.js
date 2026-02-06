import { Client } from '@microsoft/microsoft-graph-client';

// PDF.js will be loaded dynamically to avoid worker issues
let pdfjsLib = null;

// Initialize Graph client with access token
function getGraphClient(accessToken) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
}

// Get user profile
export async function getUserProfile(accessToken) {
  const client = getGraphClient(accessToken);
  return await client.api('/me').get();
}

// Get parent folder's webUrl for a file
export async function getParentFolderUrl(fileId, accessToken, parentId = null) {
  if (!accessToken) return null;

  try {
    const client = getGraphClient(accessToken);

    // If parentId is provided, use it directly
    let parentFolderId = parentId;

    // Otherwise get it from the file
    if (!parentFolderId && fileId) {
      const fileDetails = await client
        .api(`/me/drive/items/${fileId}`)
        .select('id,name,parentReference')
        .get();
      parentFolderId = fileDetails.parentReference?.id;
    }

    if (parentFolderId) {
      // Get parent folder details
      const parentFolder = await client
        .api(`/me/drive/items/${parentFolderId}`)
        .select('id,name,webUrl')
        .get();

      return parentFolder.webUrl;
    }
  } catch (err) {
    console.error('Error getting parent folder URL:', err);
  }
  return null;
}

// Recursively get all files from a folder
async function getFilesRecursive(client, driveId, folderId, parentPath = '') {
  const allFiles = [];

  try {
    const response = await client
      .api(`/me/drive/items/${folderId}/children`)
      .select('id,name,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,file,folder,size,parentReference')
      .top(200)
      .get();

    for (const item of response.value) {
      const currentPath = parentPath ? `${parentPath}/${item.name}` : item.name;

      if (item.folder) {
        // It's a folder - add it to the list and recurse into it
        allFiles.push({
          id: item.id,
          name: item.name,
          path: currentPath,
          parentPath: parentPath || '/',
          team: 'My OneDrive',
          sharedBy: item.createdBy?.user?.displayName || 'Me',
          lastModifiedBy: item.lastModifiedBy?.user?.displayName || 'Me',
          date: item.lastModifiedDateTime?.split('T')[0] || '',
          webUrl: item.webUrl,
          size: null,
          type: 'folder',
          isFolder: true,
        });

        // Recursively get files from this folder
        const subFiles = await getFilesRecursive(client, driveId, item.id, currentPath);
        allFiles.push(...subFiles);
      } else {
        // It's a file
        allFiles.push({
          id: item.id,
          name: item.name,
          path: currentPath,
          parentPath: parentPath || '/',
          team: 'My OneDrive',
          sharedBy: item.createdBy?.user?.displayName || 'Me',
          lastModifiedBy: item.lastModifiedBy?.user?.displayName || 'Me',
          date: item.lastModifiedDateTime?.split('T')[0] || '',
          webUrl: item.webUrl,
          size: item.size,
          type: getFileType(item.name),
          isFolder: false,
        });
      }
    }
  } catch (error) {
    console.error('Error fetching folder contents:', error);
  }

  return allFiles;
}

// Get files from user's OneDrive (root level only for fast loading)
export async function getMyDriveFiles(accessToken) {
  console.log('getMyDriveFiles called');
  const client = getGraphClient(accessToken);

  try {
    const me = await client.api('/me').get();
    console.log('User:', me.displayName);

    // Get root children only (fast)
    const response = await client
      .api('/me/drive/root/children')
      .select('id,name,file,folder,webUrl,size,lastModifiedDateTime,createdBy,parentReference')
      .top(100)
      .get();

    console.log('Found', response.value?.length, 'items in root');

    return (response.value || []).map(item => ({
      id: item.id,
      name: item.name,
      path: item.name,
      parentPath: 'Root',
      team: 'My OneDrive',
      sharedBy: item.createdBy?.user?.displayName || 'Me',
      date: item.lastModifiedDateTime?.split('T')[0] || '',
      webUrl: item.webUrl,
      size: item.size,
      type: item.folder ? 'folder' : getFileType(item.name),
      isFolder: !!item.folder,
    }));
  } catch (error) {
    console.error('Error fetching drive files:', error);
    throw error;
  }
}

// Get files from a specific folder (for folder navigation)
export async function getFolderFiles(accessToken, folderId, folderPath = '') {
  console.log('getFolderFiles called for:', folderId);
  const client = getGraphClient(accessToken);

  try {
    const response = await client
      .api(`/me/drive/items/${folderId}/children`)
      .select('id,name,file,folder,webUrl,size,lastModifiedDateTime,createdBy,parentReference')
      .top(100)
      .get();

    return (response.value || []).map(item => ({
      id: item.id,
      name: item.name,
      path: folderPath ? `${folderPath}/${item.name}` : item.name,
      parentPath: folderPath || 'Root',
      team: 'My OneDrive',
      sharedBy: item.createdBy?.user?.displayName || 'Me',
      date: item.lastModifiedDateTime?.split('T')[0] || '',
      webUrl: item.webUrl,
      size: item.size,
      type: item.folder ? 'folder' : getFileType(item.name),
      isFolder: !!item.folder,
    }));
  } catch (error) {
    console.error('Error fetching folder files:', error);
    throw error;
  }
}

// Get files shared with the current user
export async function getSharedWithMe(accessToken) {
  const client = getGraphClient(accessToken);
  try {
    console.log('Fetching shared files...');
    const response = await client
      .api('/me/drive/sharedWithMe')
      .select('id,name,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,file,folder,size,remoteItem,parentReference')
      .get();

    console.log('Shared files response:', response);

    if (!response.value || response.value.length === 0) {
      console.log('No shared files found');
      return [];
    }

    return response.value.map(item => {
      const remotePath = item.remoteItem?.parentReference?.path || '';
      const pathParts = remotePath.split('/').filter(p => p && !p.includes(':'));
      const parentPath = pathParts.join('/') || '/';

      return {
        id: item.id,
        name: item.name || item.remoteItem?.name,
        path: parentPath ? `${parentPath}/${item.name || item.remoteItem?.name}` : (item.name || item.remoteItem?.name),
        parentPath: parentPath,
        team: 'Shared with me',
        sharedBy: item.remoteItem?.shared?.sharedBy?.user?.displayName ||
                  item.createdBy?.user?.displayName ||
                  item.remoteItem?.createdBy?.user?.displayName ||
                  'Unknown',
        lastModifiedBy: item.lastModifiedBy?.user?.displayName || 'Unknown',
        date: item.lastModifiedDateTime?.split('T')[0] || item.remoteItem?.lastModifiedDateTime?.split('T')[0] || '',
        webUrl: item.webUrl || item.remoteItem?.webUrl,
        size: item.size || item.remoteItem?.size,
        type: getFileType(item.name || item.remoteItem?.name || ''),
        isFolder: !!(item.folder || item.remoteItem?.folder),
      };
    });
  } catch (error) {
    console.error('Error fetching shared files:', error);
    if (error.statusCode === 403 || error.code === 'Authorization_RequestDenied') {
      throw new Error('Access denied. "Shared with Me" requires admin consent for Files.Read.All permission. Please use "My Files" or "Recent" tabs instead.');
    }
    throw error;
  }
}

// Get recent files
export async function getRecentFiles(accessToken) {
  const client = getGraphClient(accessToken);
  try {
    const response = await client
      .api('/me/drive/recent')
      .select('id,name,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,file,folder,size,parentReference')
      .top(50)
      .get();

    return response.value.map(item => {
      const pathFromRef = item.parentReference?.path || '';
      const pathParts = pathFromRef.split('/').filter(p => p && !p.includes(':'));
      const parentPath = pathParts.join('/') || '/';

      return {
        id: item.id,
        name: item.name,
        path: parentPath !== '/' ? `${parentPath}/${item.name}` : item.name,
        parentPath: parentPath,
        team: item.parentReference?.name || 'Recent',
        sharedBy: item.createdBy?.user?.displayName || 'Unknown',
        lastModifiedBy: item.lastModifiedBy?.user?.displayName || 'Unknown',
        date: item.lastModifiedDateTime?.split('T')[0] || '',
        webUrl: item.webUrl,
        size: item.size,
        type: getFileType(item.name),
        isFolder: !!item.folder,
      };
    });
  } catch (error) {
    console.error('Error fetching recent files:', error);
    return [];
  }
}

// Get teams the user has joined
export async function getJoinedTeams(accessToken) {
  const client = getGraphClient(accessToken);
  try {
    const response = await client
      .api('/me/joinedTeams')
      .select('id,displayName,description')
      .get();

    return (response.value || []).map(team => ({
      id: team.id,
      name: team.displayName,
      description: team.description || '',
    }));
  } catch (error) {
    console.error('Error fetching joined teams:', error);
    throw error;
  }
}

// Get channels for a team
export async function getTeamChannels(accessToken, teamId) {
  const client = getGraphClient(accessToken);
  try {
    const response = await client
      .api(`/teams/${teamId}/channels`)
      .select('id,displayName,description')
      .get();

    return (response.value || []).map(channel => ({
      id: channel.id,
      name: channel.displayName,
      description: channel.description || '',
    }));
  } catch (error) {
    console.error('Error fetching team channels:', error);
    throw error;
  }
}

// Get files from a team's (group's) document library root
export async function getTeamDriveFiles(accessToken, teamId) {
  const client = getGraphClient(accessToken);
  try {
    const response = await client
      .api(`/groups/${teamId}/drive/root/children`)
      .select('id,name,file,folder,webUrl,size,lastModifiedDateTime,lastModifiedBy,createdBy,parentReference')
      .top(200)
      .get();

    return (response.value || []).map(item => ({
      id: item.id,
      name: item.name,
      path: item.name,
      parentPath: 'Root',
      team: 'Team Documents',
      sharedBy: item.createdBy?.user?.displayName || 'Unknown',
      lastModifiedBy: item.lastModifiedBy?.user?.displayName || 'Unknown',
      date: item.lastModifiedDateTime?.split('T')[0] || '',
      webUrl: item.webUrl,
      size: item.size,
      type: item.folder ? 'folder' : getFileType(item.name),
      isFolder: !!item.folder,
      driveType: 'group',
      groupId: teamId,
    }));
  } catch (error) {
    console.error('Error fetching team drive files:', error);
    throw error;
  }
}

// Get files from a specific folder in a team's drive
export async function getTeamFolderFiles(accessToken, teamId, folderId, folderPath = '') {
  const client = getGraphClient(accessToken);
  try {
    const response = await client
      .api(`/groups/${teamId}/drive/items/${folderId}/children`)
      .select('id,name,file,folder,webUrl,size,lastModifiedDateTime,lastModifiedBy,createdBy,parentReference')
      .top(200)
      .get();

    return (response.value || []).map(item => ({
      id: item.id,
      name: item.name,
      path: folderPath ? `${folderPath}/${item.name}` : item.name,
      parentPath: folderPath || 'Root',
      team: 'Team Documents',
      sharedBy: item.createdBy?.user?.displayName || 'Unknown',
      lastModifiedBy: item.lastModifiedBy?.user?.displayName || 'Unknown',
      date: item.lastModifiedDateTime?.split('T')[0] || '',
      webUrl: item.webUrl,
      size: item.size,
      type: item.folder ? 'folder' : getFileType(item.name),
      isFolder: !!item.folder,
      driveType: 'group',
      groupId: teamId,
    }));
  } catch (error) {
    console.error('Error fetching team folder files:', error);
    throw error;
  }
}

// Search files in user's entire OneDrive (all folders)
// Also searches recent files and Recordings folder for videos
export async function searchFiles(accessToken, searchQuery) {
  console.log('Searching for:', searchQuery);
  const client = getGraphClient(accessToken);
  const searchLower = searchQuery.toLowerCase();

  // Helper to recursively search a folder
  async function searchInFolder(folderPath, depth = 0) {
    if (depth > 2) return [];
    const items = [];
    try {
      const response = await client
        .api(`/me/drive/root:/${folderPath}:/children`)
        .select('id,name,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,file,folder,size,parentReference')
        .top(200)
        .get();

      for (const item of response.value || []) {
        // Check if item name matches search
        const decodedName = decodeURIComponent(item.name || '');
        if (decodedName.toLowerCase().includes(searchLower) ||
            item.name?.toLowerCase().includes(searchLower)) {
          items.push(item);
        }
        // Recursively search subfolders
        if (item.folder) {
          const subItems = await searchInFolder(`${folderPath}/${item.name}`, depth + 1);
          items.push(...subItems);
        }
      }
    } catch (err) {
      console.log('Could not search folder:', folderPath, err.message);
    }
    return items;
  }

  try {
    // Search using Graph API search
    const searchPromise = client
      .api(`/me/drive/root/search(q='${searchQuery}')`)
      .select('id,name,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,file,folder,size,parentReference')
      .top(200)
      .get();

    // Also get recent files (catches newly uploaded files faster)
    const recentPromise = client
      .api('/me/drive/recent')
      .select('id,name,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,file,folder,size,parentReference')
      .top(50)
      .get();

    // Also get root children for very new files
    const rootPromise = client
      .api('/me/drive/root/children')
      .select('id,name,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,file,folder,size,parentReference')
      .top(100)
      .get();

    // Search in Recordings folder specifically (Teams meeting recordings)
    const recordingsPromise = searchInFolder('Recordings');

    const [searchResponse, recentResponse, rootResponse, recordingsItems] = await Promise.all([
      searchPromise.catch(e => ({ value: [] })),
      recentPromise.catch(e => ({ value: [] })),
      rootPromise.catch(e => ({ value: [] })),
      recordingsPromise.catch(e => [])
    ]);

    // Combine and filter results
    const allItems = [
      ...(searchResponse.value || []),
      ...(recentResponse.value || []).filter(item =>
        item.name?.toLowerCase().includes(searchLower)
      ),
      ...(rootResponse.value || []).filter(item =>
        item.name?.toLowerCase().includes(searchLower)
      ),
      ...recordingsItems
    ];

    // Remove duplicates by ID
    const uniqueItems = allItems.filter((item, index, self) =>
      index === self.findIndex(t => t.id === item.id)
    );

    console.log('Search found', uniqueItems.length, 'results (search:', searchResponse.value?.length, '+ recent + root)');

    // For each item, get its full path by making additional API call
    const filesWithPaths = await Promise.all(
      uniqueItems.map(async (item) => {
        let folderPath = '';
        let parentFolderName = '';

        try {
          // Get the item details including path
          const itemDetails = await client
            .api(`/me/drive/items/${item.id}`)
            .select('id,name,parentReference')
            .get();

          console.log('Item details:', item.name, JSON.stringify(itemDetails.parentReference));

          // Extract path from parentReference.path
          const rawPath = itemDetails.parentReference?.path || '';
          parentFolderName = itemDetails.parentReference?.name || '';

          // Remove the "/drive/root:" or "/drives/{id}/root:" prefix
          folderPath = rawPath
            .replace(/^\/drives\/[^/]+\/root:?\/?/, '')
            .replace(/^\/drive\/root:?\/?/, '');

          // Decode URL encoded characters
          try {
            folderPath = decodeURIComponent(folderPath);
          } catch (e) {}
        } catch (err) {
          console.log('Could not get path for:', item.name);
          // Try to extract from webUrl as fallback
          if (item.webUrl) {
            try {
              const url = new URL(item.webUrl);
              const pathParts = url.pathname.split('/');
              // Remove filename and common prefixes
              const idx = pathParts.findIndex(p => p === 'Documents' || p === 'Shared%20Documents');
              if (idx !== -1) {
                folderPath = pathParts.slice(idx, -1).map(p => decodeURIComponent(p)).join('/');
              }
            } catch (e) {}
          }
        }

        // Decode name if URL encoded
        let decodedName = item.name;
        try {
          decodedName = decodeURIComponent(item.name);
        } catch (e) {}

        // Build full path
        const fullPath = folderPath ? `${folderPath}/${decodedName}` : decodedName;
        const location = folderPath || parentFolderName || 'Root';

        return {
          id: item.id,
          name: decodedName,
          path: fullPath,
          parentPath: location,
          parentFolderName: parentFolderName,
          parentId: item.parentReference?.id,
          team: 'Search Results',
          sharedBy: item.createdBy?.user?.displayName || 'Unknown',
          date: item.lastModifiedDateTime?.split('T')[0] || '',
          webUrl: item.webUrl,
          size: item.size,
          type: item.folder ? 'folder' : getFileType(decodedName),
          isFolder: !!item.folder,
        };
      })
    );

    return filesWithPaths;
  } catch (error) {
    console.error('Error searching files:', error);
    throw error;
  }
}

// Supported file types for upload + content extraction
export const SUPPORTED_UPLOAD_EXTENSIONS = [
  'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt',
  'pdf', 'txt', 'md', 'csv', 'json', 'xml', 'html', 'htm',
];

export const MAX_UPLOAD_SIZE = 150 * 1024 * 1024; // 150MB

// Helper function to determine file type based on extension
export function getFileType(filename) {
  if (!filename) return 'file';
  const ext = filename.split('.').pop()?.toLowerCase();
  const typeMap = {
    'docx': 'document',
    'doc': 'document',
    'txt': 'document',
    'md': 'document',
    'xlsx': 'spreadsheet',
    'xls': 'spreadsheet',
    'csv': 'spreadsheet',
    'pdf': 'pdf',
    'pptx': 'presentation',
    'ppt': 'presentation',
    'png': 'image',
    'jpg': 'image',
    'jpeg': 'image',
    'gif': 'image',
    'svg': 'image',
    'fig': 'design',
    'sketch': 'design',
    'mp4': 'video',
    'mov': 'video',
    'avi': 'video',
    'mp3': 'audio',
    'wav': 'audio',
    'zip': 'archive',
    'rar': 'archive',
  };
  return typeMap[ext] || 'file';
}

// Format file size
export function formatFileSize(bytes) {
  if (!bytes) return '';
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${sizes[i]}`;
}

// Get document content/preview for summary
export async function getDocumentContent(accessToken, itemId, fileName) {
  const client = getGraphClient(accessToken);
  const ext = fileName.split('.').pop()?.toLowerCase();

  try {
    // Get file details first
    const itemDetails = await client
      .api(`/me/drive/items/${itemId}`)
      .get();

    const downloadUrl = itemDetails['@microsoft.graph.downloadUrl'];
    let content = '';
    let contentType = 'unknown';

    // For text files, fetch content directly
    if (['txt', 'md', 'csv', 'json', 'xml', 'html', 'htm'].includes(ext)) {
      contentType = 'text';
      if (downloadUrl) {
        try {
          const response = await fetch(downloadUrl);
          content = await response.text();
          // Limit content length
          if (content.length > 5000) {
            content = content.substring(0, 5000) + '...';
          }
        } catch (e) {
          console.log('Could not fetch text content:', e);
        }
      }
    }
    // For Word documents, try to extract content directly from .docx
    else if (['docx', 'doc'].includes(ext)) {
      contentType = 'word';

      // Try downloading and parsing .docx first (most reliable)
      if (downloadUrl && ext === 'docx') {
        try {
          console.log('Fetching Word document:', fileName);
          const docResponse = await fetch(downloadUrl);
          const docBlob = await docResponse.blob();
          const textFromDocx = await extractTextFromDocx(docBlob);
          if (textFromDocx && textFromDocx.length > 20) {
            content = textFromDocx;
            contentType = 'word-extracted';
            console.log('Extracted Word content:', content.substring(0, 100));
          }
        } catch (docxErr) {
          console.log('Could not extract docx content:', docxErr);
        }
      }

    }
    // For Excel documents - extract from downloaded file
    else if (['xlsx', 'xls'].includes(ext)) {
      contentType = 'excel';
      if (downloadUrl && ext === 'xlsx') {
        try {
          console.log('Fetching Excel document:', fileName);
          const xlsResponse = await fetch(downloadUrl);
          if (xlsResponse.ok) {
            const xlsBlob = await xlsResponse.blob();
            const textFromXlsx = await extractTextFromOfficeXml(xlsBlob, 'xlsx');
            if (textFromXlsx && textFromXlsx.length > 20) {
              content = textFromXlsx;
              contentType = 'excel-extracted';
            }
          }
        } catch (xlsErr) {
          console.log('Could not extract xlsx content:', xlsErr);
        }
      }
    }
    // For PowerPoint - extract from downloaded file
    else if (['pptx', 'ppt'].includes(ext)) {
      contentType = 'powerpoint';
      if (downloadUrl && ext === 'pptx') {
        try {
          console.log('Fetching PowerPoint document:', fileName);
          const pptResponse = await fetch(downloadUrl);
          if (pptResponse.ok) {
            const pptBlob = await pptResponse.blob();
            const textFromPptx = await extractTextFromOfficeXml(pptBlob, 'pptx');
            if (textFromPptx && textFromPptx.length > 20) {
              content = textFromPptx;
              contentType = 'ppt-extracted';
            }
          }
        } catch (pptErr) {
          console.log('Could not extract pptx content:', pptErr);
        }
      }
    }
    // For PDFs - extract text using pdf.js
    else if (ext === 'pdf') {
      contentType = 'pdf';
      if (downloadUrl) {
        try {
          console.log('Fetching PDF for text extraction...');
          const pdfResponse = await fetch(downloadUrl);
          const pdfBlob = await pdfResponse.blob();
          content = await extractTextFromPdf(pdfBlob);
          console.log('PDF text extracted, length:', content.length);
        } catch (pdfErr) {
          console.log('Could not extract PDF content:', pdfErr);
        }
      }
    }

    // Extract path
    const rawPath = itemDetails.parentReference?.path || '';
    let folderPath = rawPath
      .replace(/^\/drives\/[^/]+\/root:?\/?/, '')
      .replace(/^\/drive\/root:?\/?/, '');
    try { folderPath = decodeURIComponent(folderPath); } catch (e) {}

    return {
      id: itemId,
      name: fileName,
      content: content,
      contentType: contentType,
      path: folderPath ? `${folderPath}/${fileName}` : fileName,
      size: itemDetails.size,
      lastModified: itemDetails.lastModifiedDateTime?.split('T')[0],
      lastModifiedBy: itemDetails.lastModifiedBy?.user?.displayName,
      webUrl: itemDetails.webUrl,
      downloadUrl: downloadUrl
    };
  } catch (error) {
    console.error('Error getting document content:', error);
    throw error;
  }
}

// Helper to extract text from HTML content
function extractTextFromHtml(html) {
  try {
    // Create a temporary DOM element to parse HTML
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    // Remove script and style elements
    const scripts = doc.querySelectorAll('script, style, noscript');
    scripts.forEach(s => s.remove());

    // Get text content
    let text = doc.body?.textContent || doc.documentElement?.textContent || '';

    // Clean up the text
    text = text
      .replace(/\s+/g, ' ')  // Replace multiple spaces with single space
      .replace(/\n\s*\n/g, '\n')  // Replace multiple newlines
      .trim();

    // Limit length
    if (text.length > 3000) {
      text = text.substring(0, 3000) + '...';
    }

    return text;
  } catch (e) {
    console.log('Error extracting text from HTML:', e);
    return '';
  }
}

// Helper to extract text from .docx files (with ZIP decompression)
async function extractTextFromDocx(blob) {
  try {
    console.log('extractTextFromDocx - Blob size:', blob.size);

    const arrayBuffer = await blob.arrayBuffer();
    const uint8Array = new Uint8Array(arrayBuffer);

    // Check if it's a valid ZIP file (starts with PK)
    if (uint8Array[0] !== 0x50 || uint8Array[1] !== 0x4B) {
      console.log('Not a valid ZIP/DOCX file');
      return '';
    }

    console.log('Valid DOCX file, extracting document.xml...');

    // Extract and decompress document.xml from the ZIP
    const documentXml = await extractFileFromZip(uint8Array, 'word/document.xml');

    if (!documentXml) {
      console.log('Could not extract document.xml');
      return '';
    }

    console.log('document.xml extracted, length:', documentXml.length);

    // Extract text from the XML
    let extractedText = '';

    // Look for <w:t> tags (Word text content)
    const wtMatches = documentXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g);
    console.log('Found w:t matches:', wtMatches ? wtMatches.length : 0);

    if (wtMatches && wtMatches.length > 0) {
      extractedText = wtMatches
        .map(match => match.replace(/<[^>]+>/g, ''))
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();
    }

    console.log('Extracted text length:', extractedText.length);

    if (extractedText.length > 3000) {
      return extractedText.substring(0, 3000) + '...';
    }
    return extractedText;
  } catch (e) {
    console.log('Error extracting text from docx:', e);
    return '';
  }
}

// Helper to extract text from PDF files using pdf.js
async function extractTextFromPdf(blob) {
  try {
    console.log('extractTextFromPdf - Blob size:', blob.size);

    // Dynamically import pdfjs-dist
    if (!pdfjsLib) {
      const pdfjs = await import('pdfjs-dist');
      pdfjsLib = pdfjs;
      // Set worker from jsdelivr CDN (reliable and has all versions)
      const version = pdfjsLib.version;
      pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${version}/build/pdf.worker.min.mjs`;
      console.log('PDF.js version:', version);
    }

    const arrayBuffer = await blob.arrayBuffer();

    // Load the PDF document
    const loadingTask = pdfjsLib.getDocument({
      data: arrayBuffer,
    });
    const pdf = await loadingTask.promise;

    console.log('PDF loaded, pages:', pdf.numPages);

    let fullText = '';
    const maxPages = Math.min(pdf.numPages, 20); // Limit to first 20 pages

    // Extract text from each page
    for (let pageNum = 1; pageNum <= maxPages; pageNum++) {
      try {
        const page = await pdf.getPage(pageNum);
        const textContent = await page.getTextContent();

        // Concatenate text items
        const pageText = textContent.items
          .map(item => item.str)
          .join(' ');

        fullText += pageText + '\n\n';
      } catch (pageErr) {
        console.log(`Error extracting page ${pageNum}:`, pageErr);
      }
    }

    // Clean up the text
    fullText = fullText
      .replace(/\s+/g, ' ')  // Replace multiple spaces
      .replace(/\n\s*\n/g, '\n\n')  // Clean up newlines
      .trim();

    console.log('PDF text extracted, length:', fullText.length);

    // Limit text length for LLM processing
    if (fullText.length > 15000) {
      fullText = fullText.substring(0, 15000) + '...\n\n[Content truncated - PDF has ' + pdf.numPages + ' pages]';
    }

    if (pdf.numPages > maxPages) {
      fullText += `\n\n[Showing first ${maxPages} of ${pdf.numPages} pages]`;
    }

    return fullText;
  } catch (e) {
    console.log('Error extracting text from PDF:', e);
    return '';
  }
}

// Extract a file from ZIP archive
async function extractFileFromZip(zipData, targetFileName) {
  try {
    let offset = 0;

    while (offset < zipData.length - 4) {
      // Check for local file header signature (PK\x03\x04)
      if (zipData[offset] !== 0x50 || zipData[offset + 1] !== 0x4B ||
          zipData[offset + 2] !== 0x03 || zipData[offset + 3] !== 0x04) {
        break;
      }

      // Parse local file header
      const compressionMethod = zipData[offset + 8] | (zipData[offset + 9] << 8);
      const compressedSize = zipData[offset + 18] | (zipData[offset + 19] << 8) |
                            (zipData[offset + 20] << 16) | (zipData[offset + 21] << 24);
      const uncompressedSize = zipData[offset + 22] | (zipData[offset + 23] << 8) |
                               (zipData[offset + 24] << 16) | (zipData[offset + 25] << 24);
      const fileNameLength = zipData[offset + 26] | (zipData[offset + 27] << 8);
      const extraFieldLength = zipData[offset + 28] | (zipData[offset + 29] << 8);

      // Get file name
      const fileNameBytes = zipData.slice(offset + 30, offset + 30 + fileNameLength);
      const fileName = new TextDecoder().decode(fileNameBytes);

      const dataOffset = offset + 30 + fileNameLength + extraFieldLength;

      console.log('ZIP entry:', fileName, 'compression:', compressionMethod, 'size:', compressedSize);

      if (fileName === targetFileName) {
        const compressedData = zipData.slice(dataOffset, dataOffset + compressedSize);

        if (compressionMethod === 0) {
          // No compression - stored
          return new TextDecoder().decode(compressedData);
        } else if (compressionMethod === 8) {
          // Deflate compression - use DecompressionStream
          try {
            const decompressed = await decompressDeflate(compressedData);
            return new TextDecoder().decode(decompressed);
          } catch (decompressErr) {
            console.log('Decompression failed:', decompressErr);
            return '';
          }
        }
      }

      // Move to next file
      offset = dataOffset + compressedSize;
    }

    return '';
  } catch (e) {
    console.log('Error parsing ZIP:', e);
    return '';
  }
}

// Decompress deflate data using browser's DecompressionStream
async function decompressDeflate(compressedData) {
  try {
    // Add zlib header for DecompressionStream (it expects raw deflate without header)
    // DecompressionStream 'deflate-raw' handles raw deflate
    const stream = new DecompressionStream('deflate-raw');
    const writer = stream.writable.getWriter();
    const reader = stream.readable.getReader();

    // Write compressed data
    writer.write(compressedData);
    writer.close();

    // Read decompressed data
    const chunks = [];
    let result;
    while (!(result = await reader.read()).done) {
      chunks.push(result.value);
    }

    // Combine chunks
    const totalLength = chunks.reduce((acc, chunk) => acc + chunk.length, 0);
    const decompressed = new Uint8Array(totalLength);
    let position = 0;
    for (const chunk of chunks) {
      decompressed.set(chunk, position);
      position += chunk.length;
    }

    return decompressed;
  } catch (e) {
    console.log('DecompressionStream error:', e);
    throw e;
  }
}

// Extract text from XML content (Word document format)
function extractTextFromXmlContent(content) {
  try {
    // Method 1: Extract text between <w:t> tags (Word text elements)
    const textMatches = content.match(/<w:t[^>]*>([^<]*)<\/w:t>/g);

    if (textMatches && textMatches.length > 0) {
      const extractedText = textMatches
        .map(match => {
          // Remove XML tags and decode
          const text = match.replace(/<[^>]+>/g, '');
          return text;
        })
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();

      console.log('Extracted text length from w:t tags:', extractedText.length);

      if (extractedText.length > 10) {
        if (extractedText.length > 3000) {
          return extractedText.substring(0, 3000) + '...';
        }
        return extractedText;
      }
    }

    // Method 2: Try to find any readable text patterns
    const readableText = content.match(/[A-Za-z][A-Za-z\s,.'"-]{20,}/g);
    if (readableText && readableText.length > 0) {
      const combinedText = readableText
        .filter(t => !t.includes('xml') && !t.includes('schema') && !t.includes('http'))
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();

      if (combinedText.length > 50) {
        console.log('Extracted readable text:', combinedText.substring(0, 100));
        if (combinedText.length > 3000) {
          return combinedText.substring(0, 3000) + '...';
        }
        return combinedText;
      }
    }

    return '';
  } catch (e) {
    console.log('Error in extractTextFromXmlContent:', e);
    return '';
  }
}

// Extract text from Office XML files (xlsx, pptx)
async function extractTextFromOfficeXml(blob, fileType) {
  try {
    const arrayBuffer = await blob.arrayBuffer();
    const uint8Array = new Uint8Array(arrayBuffer);

    // Check if it's a valid ZIP file
    if (uint8Array[0] !== 0x50 || uint8Array[1] !== 0x4B) {
      console.log('Not a valid ZIP/Office file');
      return '';
    }

    // Convert to string
    let binaryString = '';
    for (let i = 0; i < uint8Array.length; i++) {
      binaryString += String.fromCharCode(uint8Array[i]);
    }

    let textContent = '';

    if (fileType === 'xlsx') {
      // Excel: Look for shared strings in xl/sharedStrings.xml
      const matches = binaryString.match(/<t[^>]*>([^<]+)<\/t>/g);
      if (matches) {
        textContent = matches
          .map(m => m.replace(/<[^>]+>/g, ''))
          .join(' ');
      }
    } else if (fileType === 'pptx') {
      // PowerPoint: Look for text in <a:t> tags
      const matches = binaryString.match(/<a:t>([^<]+)<\/a:t>/g);
      if (matches) {
        textContent = matches
          .map(m => m.replace(/<[^>]+>/g, ''))
          .join(' ');
      }
    }

    textContent = textContent.replace(/\s+/g, ' ').trim();

    if (textContent.length > 3000) {
      return textContent.substring(0, 3000) + '...';
    }
    return textContent;
  } catch (e) {
    console.log('Error extracting Office XML content:', e);
    return '';
  }
}

// Get thumbnail/preview URL for an item
export async function getItemThumbnail(accessToken, itemId) {
  const client = getGraphClient(accessToken);
  try {
    const thumbnails = await client
      .api(`/me/drive/items/${itemId}/thumbnails`)
      .get();

    if (thumbnails.value && thumbnails.value.length > 0) {
      // Return large thumbnail if available, otherwise medium or small
      const thumb = thumbnails.value[0];
      return thumb.large?.url || thumb.medium?.url || thumb.small?.url || null;
    }
    return null;
  } catch (error) {
    console.error('Error getting thumbnail:', error);
    return null;
  }
}

// Get download URL for a file (for video streaming)
export async function getDownloadUrl(accessToken, itemId) {
  const client = getGraphClient(accessToken);
  try {
    const item = await client
      .api(`/me/drive/items/${itemId}`)
      .select('@microsoft.graph.downloadUrl')
      .get();

    return item['@microsoft.graph.downloadUrl'] || null;
  } catch (error) {
    console.error('Error getting download URL:', error);
    return null;
  }
}

// Get only folders from a specific location (for folder picker)
export async function getFoldersOnly(accessToken, folderId = null) {
  const client = getGraphClient(accessToken);
  try {
    const apiPath = folderId
      ? `/me/drive/items/${folderId}/children`
      : '/me/drive/root/children';

    const response = await client
      .api(apiPath)
      .select('id,name,folder,parentReference')
      .top(100)
      .get();

    return (response.value || [])
      .filter(item => item.folder)
      .map(item => ({
        id: item.id,
        name: item.name,
        hasChildren: (item.folder?.childCount || 0) > 0,
      }));
  } catch (error) {
    console.error('Error fetching folders:', error);
    return [];
  }
}

// Upload a small file (<4MB) to OneDrive
async function uploadSmallFile(accessToken, parentFolderId, fileName, fileContent) {
  const apiPath = parentFolderId === 'root'
    ? `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/content`
    : `https://graph.microsoft.com/v1.0/me/drive/items/${parentFolderId}:/${encodeURIComponent(fileName)}:/content`;

  const response = await fetch(apiPath, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/octet-stream',
    },
    body: fileContent,
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData?.error?.message || `Upload failed with status ${response.status}`);
  }

  return await response.json();
}

// Upload a large file (>4MB) using upload session with progress
async function uploadLargeFile(accessToken, parentFolderId, fileName, file, onProgress) {
  const sessionUrl = parentFolderId === 'root'
    ? `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/createUploadSession`
    : `https://graph.microsoft.com/v1.0/me/drive/items/${parentFolderId}:/${encodeURIComponent(fileName)}:/createUploadSession`;

  const sessionResponse = await fetch(sessionUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      item: { '@microsoft.graph.conflictBehavior': 'rename' },
    }),
  });

  if (!sessionResponse.ok) {
    throw new Error('Failed to create upload session');
  }

  const session = await sessionResponse.json();
  const uploadUrl = session.uploadUrl;

  const CHUNK_SIZE = 5 * 1024 * 1024; // 5MB chunks
  const totalSize = file.size;
  let offset = 0;
  let result = null;

  while (offset < totalSize) {
    const end = Math.min(offset + CHUNK_SIZE, totalSize);
    const chunk = file.slice(offset, end);

    const chunkResponse = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Length': `${end - offset}`,
        'Content-Range': `bytes ${offset}-${end - 1}/${totalSize}`,
      },
      body: chunk,
    });

    if (!chunkResponse.ok && chunkResponse.status !== 202) {
      throw new Error(`Upload failed at ${Math.round((offset / totalSize) * 100)}%`);
    }

    result = await chunkResponse.json();
    offset = end;

    if (onProgress) {
      onProgress(Math.round((offset / totalSize) * 100));
    }
  }

  return result;
}

// Upload file to OneDrive (auto-selects small vs large)
const SMALL_FILE_LIMIT = 4 * 1024 * 1024; // 4MB

export async function uploadFileToOneDrive(accessToken, parentFolderId, file, onProgress) {
  if (file.size <= SMALL_FILE_LIMIT) {
    if (onProgress) onProgress(50);
    const result = await uploadSmallFile(accessToken, parentFolderId, file.name, file);
    if (onProgress) onProgress(100);
    return result;
  } else {
    return await uploadLargeFile(accessToken, parentFolderId, file.name, file, onProgress);
  }
}

// Get a file item by ID (for recently uploaded files that may not be indexed yet)
export async function getFileById(accessToken, itemId) {
  const client = getGraphClient(accessToken);
  try {
    const item = await client
      .api(`/me/drive/items/${itemId}`)
      .select('id,name,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,file,folder,size,parentReference')
      .get();

    const rawPath = item.parentReference?.path || '';
    let folderPath = rawPath
      .replace(/^\/drives\/[^/]+\/root:?\/?/, '')
      .replace(/^\/drive\/root:?\/?/, '');
    try { folderPath = decodeURIComponent(folderPath); } catch (e) {}

    return {
      id: item.id,
      name: item.name,
      path: folderPath ? `${folderPath}/${item.name}` : item.name,
      parentPath: folderPath || '/',
      team: item.parentReference?.name || 'OneDrive',
      sharedBy: item.createdBy?.user?.displayName || 'Unknown',
      lastModifiedBy: item.lastModifiedBy?.user?.displayName || 'Unknown',
      date: item.lastModifiedDateTime?.split('T')[0] || '',
      webUrl: item.webUrl,
      size: item.size,
      type: getFileType(item.name),
      isFolder: !!item.folder,
    };
  } catch (error) {
    console.error('Error fetching file by ID:', error);
    return null;
  }
}

// Get video transcript content for summarization
export async function getVideoTranscript(accessToken, videoItem) {
  console.log('Getting transcript for video:', videoItem.name);
  const client = getGraphClient(accessToken);

  try {
    // Extract base name without extension
    const videoName = videoItem.name;
    const baseName = videoName.replace(/\.(mp4|mov|avi|mkv|webm)$/i, '');

    // Get the parent folder path
    let parentPath = videoItem.parentPath || videoItem.path?.replace(/\/[^/]+$/, '') || '';
    if (parentPath === videoItem.name) parentPath = '';

    console.log('Looking for transcript, video base name:', baseName);

    let transcriptContent = '';
    let transcriptFile = null;

    // Method 1: Check same folder as video for transcript with same base name
    if (videoItem.id) {
      try {
        // Get the video item's parent folder
        const videoDetails = await client
          .api(`/me/drive/items/${videoItem.id}`)
          .select('id,name,parentReference')
          .get();

        if (videoDetails.parentReference?.id) {
          const siblingFiles = await client
            .api(`/me/drive/items/${videoDetails.parentReference.id}/children`)
            .select('id,name,file,webUrl')
            .top(100)
            .get();

          console.log('Files in same folder:', siblingFiles.value?.map(f => f.name));

          // Look for transcript files
          for (const item of siblingFiles.value || []) {
            const name = item.name?.toLowerCase() || '';
            const ext = name.split('.').pop();

            // Check if it's a transcript file type
            if (ext === 'vtt' || ext === 'srt' ||
                (name.includes('transcript') && (ext === 'docx' || ext === 'txt'))) {

              // Check if it matches the video name
              const itemBase = item.name.replace(/\.(vtt|srt|docx|txt)$/i, '').toLowerCase();
              const videoBase = baseName.toLowerCase();

              if (itemBase === videoBase ||
                  itemBase.includes(videoBase.substring(0, 30)) ||
                  videoBase.includes(itemBase.substring(0, 30))) {
                transcriptFile = item;
                console.log('Found matching transcript in same folder:', item.name);
                break;
              }
            }
          }

          // If no exact match, take any .vtt file in the folder
          if (!transcriptFile) {
            for (const item of siblingFiles.value || []) {
              const ext = item.name?.split('.').pop()?.toLowerCase();
              if (ext === 'vtt' || ext === 'srt') {
                transcriptFile = item;
                console.log('Using first transcript file found:', item.name);
                break;
              }
            }
          }
        }
      } catch (err) {
        console.log('Could not check parent folder:', err.message);
      }
    }

    // Method 2: Search by video name patterns
    if (!transcriptFile) {
      const searchPatterns = [
        baseName,
        baseName.replace(/-Meeting Recording$/i, ''),
        baseName.replace(/Meeting Recording$/i, '').trim(),
        baseName.split('-')[0],
        'transcript',
      ];

      for (const pattern of searchPatterns) {
        if (!pattern || pattern.length < 3) continue;

        try {
          const searchResponse = await client
            .api(`/me/drive/root/search(q='${encodeURIComponent(pattern)}')`)
            .select('id,name,file,webUrl,parentReference')
            .top(50)
            .get();

          for (const item of searchResponse.value || []) {
            const name = item.name?.toLowerCase() || '';
            const ext = name.split('.').pop();

            if (ext === 'vtt' || ext === 'srt' ||
                (name.includes('transcript') && (ext === 'docx' || ext === 'txt'))) {
              transcriptFile = item;
              console.log('Found transcript via search:', item.name);
              break;
            }
          }

          if (transcriptFile) break;
        } catch (err) {
          console.log('Search error for pattern:', pattern, err.message);
        }
      }
    }

    // Method 3: Check Recordings folder directly
    if (!transcriptFile) {
      try {
        const recordingsResponse = await client
          .api('/me/drive/root:/Recordings:/children')
          .select('id,name,file,webUrl')
          .top(200)
          .get();

        console.log('Files in Recordings folder:', recordingsResponse.value?.map(f => f.name).filter(n => n.endsWith('.vtt') || n.includes('transcript')));

        for (const item of recordingsResponse.value || []) {
          const name = item.name?.toLowerCase() || '';
          const ext = name.split('.').pop();

          if (ext === 'vtt' || ext === 'srt' ||
              (name.includes('transcript') && (ext === 'docx' || ext === 'txt'))) {

            // Check for name match
            const itemBase = item.name.replace(/\.(vtt|srt|docx|txt)$/i, '').toLowerCase();
            const videoBase = baseName.toLowerCase();

            if (itemBase.includes(videoBase.substring(0, 20)) ||
                videoBase.includes(itemBase.substring(0, 20))) {
              transcriptFile = item;
              console.log('Found transcript in Recordings:', item.name);
              break;
            }
          }
        }

        // If still not found, list all .vtt files for debugging
        if (!transcriptFile) {
          const vttFiles = (recordingsResponse.value || []).filter(item =>
            item.name?.toLowerCase().endsWith('.vtt')
          );
          console.log('All .vtt files in Recordings:', vttFiles.map(f => f.name));
        }
      } catch (err) {
        console.log('Could not search Recordings folder:', err.message);
      }
    }

    // Method 4: Try Microsoft Stream transcripts location
    if (!transcriptFile) {
      try {
        // Teams recordings may store transcripts in a subfolder
        const paths = [
          '/me/drive/root:/Recordings/Transcripts:/children',
          '/me/drive/root:/Microsoft Teams Chat Files:/children',
        ];

        for (const apiPath of paths) {
          try {
            const response = await client
              .api(apiPath)
              .select('id,name,file,webUrl')
              .top(100)
              .get();

            for (const item of response.value || []) {
              const ext = item.name?.split('.').pop()?.toLowerCase();
              if (ext === 'vtt' || ext === 'srt' || ext === 'docx') {
                const itemBase = item.name.replace(/\.(vtt|srt|docx)$/i, '').toLowerCase();
                if (baseName.toLowerCase().includes(itemBase.substring(0, 15))) {
                  transcriptFile = item;
                  console.log('Found transcript in:', apiPath, item.name);
                  break;
                }
              }
            }
            if (transcriptFile) break;
          } catch (pathErr) {
            // Path doesn't exist, continue
          }
        }
      } catch (err) {
        console.log('Could not search alternate locations:', err.message);
      }
    }

    if (transcriptFile) {
      console.log('Reading transcript file:', transcriptFile.name);

      const itemDetails = await client
        .api(`/me/drive/items/${transcriptFile.id}`)
        .get();

      const downloadUrl = itemDetails['@microsoft.graph.downloadUrl'];

      if (downloadUrl) {
        const response = await fetch(downloadUrl);
        let content = await response.text();

        const ext = transcriptFile.name.split('.').pop()?.toLowerCase();
        if (ext === 'vtt' || ext === 'srt') {
          content = parseSubtitleFile(content);
        } else if (ext === 'docx') {
          // Try to extract text from docx
          try {
            const blob = await (await fetch(downloadUrl)).blob();
            content = await extractTextFromDocx(blob);
          } catch (docxErr) {
            console.log('Could not parse docx transcript:', docxErr);
          }
        }

        transcriptContent = content;
        console.log('Transcript content length:', transcriptContent.length);
      }

      return {
        hasTranscript: true,
        transcriptFile: transcriptFile.name,
        content: transcriptContent,
        videoName: videoItem.name,
        videoPath: videoItem.path,
        videoWebUrl: videoItem.webUrl,
      };
    }

    // No transcript found
    console.log('No transcript file found for:', videoItem.name);
    return {
      hasTranscript: false,
      videoName: videoItem.name,
      videoPath: videoItem.path,
      videoWebUrl: videoItem.webUrl,
      message: 'No transcript file found for this video.',
    };
  } catch (error) {
    console.error('Error getting video transcript:', error);
    return {
      hasTranscript: false,
      videoName: videoItem.name,
      error: error.message,
    };
  }
}

// Parse VTT/SRT subtitle file to extract text
function parseSubtitleFile(content) {
  // Remove WEBVTT header and metadata
  let text = content
    .replace(/^WEBVTT.*$/gm, '')
    .replace(/^NOTE.*$/gm, '')
    .replace(/^Kind:.*$/gm, '')
    .replace(/^Language:.*$/gm, '');

  // Remove timestamps (00:00:00.000 --> 00:00:05.000 format)
  text = text.replace(/\d{2}:\d{2}:\d{2}[.,]\d{3}\s*-->\s*\d{2}:\d{2}:\d{2}[.,]\d{3}/g, '');

  // Remove SRT sequence numbers
  text = text.replace(/^\d+\s*$/gm, '');

  // Remove positioning tags like <v Speaker Name>
  text = text.replace(/<v\s+[^>]+>/g, '');
  text = text.replace(/<\/v>/g, '');

  // Remove other tags
  text = text.replace(/<[^>]+>/g, '');

  // Clean up whitespace
  text = text
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.length > 0)
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();

  // Limit length for LLM processing
  if (text.length > 15000) {
    text = text.substring(0, 15000) + '...\n\n[Transcript truncated]';
  }

  return text;
}

// Search for images and videos specifically
export async function searchMedia(accessToken, searchQuery, mediaType = 'all') {
  console.log('Searching media for:', searchQuery, 'type:', mediaType);
  const client = getGraphClient(accessToken);

  try {
    const response = await client
      .api(`/me/drive/root/search(q='${searchQuery}')`)
      .select('id,name,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,file,folder,size,parentReference')
      .top(200)
      .get();

    // Filter for images and videos
    const imageExtensions = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'];
    const videoExtensions = ['mp4', 'mov', 'avi', 'wmv', 'mkv', 'webm'];

    let mediaFiles = response.value.filter(item => {
      if (item.folder) return false;
      const ext = item.name.split('.').pop()?.toLowerCase();

      if (mediaType === 'image') {
        return imageExtensions.includes(ext);
      } else if (mediaType === 'video') {
        return videoExtensions.includes(ext);
      } else {
        return imageExtensions.includes(ext) || videoExtensions.includes(ext);
      }
    });

    // Get thumbnails and download URLs for each media file
    const mediaWithPreviews = await Promise.all(
      mediaFiles.map(async (item) => {
        const ext = item.name.split('.').pop()?.toLowerCase();
        const isVideo = videoExtensions.includes(ext);
        const isImage = imageExtensions.includes(ext);

        let thumbnailUrl = null;
        let downloadUrl = null;
        let folderPath = '';

        let embedUrl = null;

        try {
          // Get item details including download URL (don't use select for @microsoft.graph.downloadUrl)
          const itemDetails = await client
            .api(`/me/drive/items/${item.id}`)
            .get();

          downloadUrl = itemDetails['@microsoft.graph.downloadUrl'];
          console.log('Download URL for', item.name, ':', downloadUrl ? 'Found' : 'Not found');

          // Extract path
          const rawPath = itemDetails.parentReference?.path || '';
          folderPath = rawPath
            .replace(/^\/drives\/[^/]+\/root:?\/?/, '')
            .replace(/^\/drive\/root:?\/?/, '');
          try { folderPath = decodeURIComponent(folderPath); } catch (e) {}

          // For videos, get the embed URL for iframe playback
          if (isVideo) {
            try {
              const preview = await client
                .api(`/me/drive/items/${item.id}/preview`)
                .post({});
              embedUrl = preview.getUrl;
              console.log('Embed URL for', item.name, ':', embedUrl ? 'Found' : 'Not found');
            } catch (previewErr) {
              console.log('Could not get embed URL for:', item.name);
            }
          }

          // Get thumbnail
          try {
            const thumbnails = await client
              .api(`/me/drive/items/${item.id}/thumbnails`)
              .get();

            if (thumbnails.value && thumbnails.value.length > 0) {
              const thumb = thumbnails.value[0];
              thumbnailUrl = thumb.large?.url || thumb.medium?.url || thumb.small?.url;
            }
          } catch (thumbErr) {
            console.log('Could not get thumbnail for:', item.name);
          }
        } catch (err) {
          console.log('Could not get details for:', item.name, err);
        }

        const fullPath = folderPath ? `${folderPath}/${item.name}` : item.name;

        return {
          id: item.id,
          name: item.name,
          path: fullPath,
          parentPath: folderPath || 'Root',
          team: 'Search Results',
          sharedBy: item.createdBy?.user?.displayName || 'Unknown',
          date: item.lastModifiedDateTime?.split('T')[0] || '',
          webUrl: item.webUrl,
          size: item.size,
          type: isVideo ? 'video' : 'image',
          isFolder: false,
          thumbnailUrl: thumbnailUrl,
          downloadUrl: downloadUrl,
          embedUrl: embedUrl,
          isVideo: isVideo,
          isImage: isImage,
        };
      })
    );

    return mediaWithPreviews;
  } catch (error) {
    console.error('Error searching media:', error);
    throw error;
  }
}
