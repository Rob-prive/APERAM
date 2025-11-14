// ===== CONFIGURATION =====
const CACHE_DURATION = 300; // 5 minutes for sheet data
const CACHE_DURATION_INSTALLATIONS = 600; // 10 minutes for installations
const SPREADSHEET_ID = '1dRfxF-38TmpVgZYReE6mbn3ekSoLpehzCPzelMdtYo4';

// ===== FIREBASE CONFIGURATION =====
const FIREBASE_URL = 'https://aanvragen-a3d22-default-rtdb.europe-west1.firebasedatabase.app';
const FIREBASE_SECRET = 'lUqckdwDhaz6NNjjzTllOPv2q5LZ0Tg5tsfxhdMn';

// Firebase for Aanvragen (separate database)
const FIREBASE_AANVRAGEN_URL = 'https://aanvragen-groepen-default-rtdb.europe-west1.firebasedatabase.app';
const FIREBASE_AANVRAGEN_SECRET = 'Cv5y8Dk4hPEGbjAiilKeReRvijqECnUu89qSY2TZ';

// ===== FIREBASE DATA FETCHING =====

/**
 * Fetch data from Firebase Realtime Database
 * @param {string} path - Firebase path (e.g., '/uploadData')
 * @returns {Object|Array} Firebase data
 */
function getFirebaseData(path) {
  try {
    const url = `${FIREBASE_URL}${path}.json?auth=${FIREBASE_SECRET}`;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      console.error('Firebase error:', responseCode, response.getContentText());
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    return data;
  } catch (error) {
    console.error('Error fetching Firebase data:', error);
    return null;
  }
}

/**
 * Get all vergrendelpunten from Firebase
 * @returns {Array} Array of vergrendelpunten objects
 */
function getAllVergrendelpunten() {
  console.log('Fetching Firebase data...');
  const data = getFirebaseData('/uploadData');
  
  if (!data) {
    console.error('No data returned from Firebase');
    return [];
  }
  
  // Convert to array if it's an object
  let dataArray = Array.isArray(data) ? data : Object.values(data);
  
  // Filter out any null/undefined entries
  dataArray = dataArray.filter(item => item && typeof item === 'object');
  
  console.log(`✓ Loaded ${dataArray.length} items from Firebase`);
  return dataArray;
}

/**
 * Get filtered vergrendelpunten data for a specific installation
 * FILTERING LOGIC: Code field starts with installation name
 * Example: Code "APL3-0155" matches installation "APL3"
 *
 * @param {string|null} installation - Installation name to filter by, or null for all data
 * @returns {Object} {headers: Array, data: Array[]}
 */
function getFilteredFirebaseData(installation) {
  try {
    // If installation is null, undefined, or empty string, return ALL data
    if (!installation || installation === null || installation === undefined || installation === '') {
      console.log('No installation filter - returning ALL data');
      return getAllFirebaseDataFormatted();
    }

    console.log(`Filtering Firebase data for installation: ${installation}`);

    // Get all data
    const allData = getAllVergrendelpunten();

    if (!allData || allData.length === 0) {
      console.warn('No data available from Firebase');
      return { headers: [], data: [] };
    }

    // Filter by Code field - must START with installation name
    const filtered = allData.filter(item => {
      if (!item) return false;

      const code = item['Code'] || item['code'] || '';

      // Check if code starts with installation name
      // Case-insensitive comparison
      return String(code).toUpperCase().startsWith(String(installation).toUpperCase());
    });

    console.log(`✓ Filtered to ${filtered.length} items for ${installation} (matched by Code prefix)`);

    // Determine headers from first item (or use default structure)
    let headers = [];
    if (filtered.length > 0) {
      headers = Object.keys(filtered[0]);
    } else if (allData.length > 0) {
      // Use headers from all data even if no matches
      headers = Object.keys(allData[0]);
    } else {
      // Fallback headers based on Firebase structure
      headers = ['Code', 'Geografische locatie', 'Naam', 'Naam geografische locatie', 'Sleutelbox', 'Type'];
    }

    // Convert objects to 2D array format (like Sheets)
    const data = filtered.map(item =>
      headers.map(header => {
        const value = item[header];
        // Convert null/undefined to empty string
        return value !== null && value !== undefined ? String(value) : '';
      })
    );

    return {
      headers: headers,
      data: data
    };

  } catch (error) {
    console.error('Error in getFilteredFirebaseData:', error);
    return { headers: [], data: [] };
  }
}

/**
 * Get ALL vergrendelpunten data without filtering
 * Returns formatted data structure matching filtered version
 * @returns {Object} {headers: Array, data: Array[]}
 */
function getAllFirebaseDataFormatted() {
  try {
    console.log('Getting ALL Firebase data (no filter)');

    const allData = getAllVergrendelpunten();

    if (!allData || allData.length === 0) {
      console.warn('No data available from Firebase');
      return { headers: [], data: [] };
    }

    console.log(`✓ Loaded ${allData.length} total vergrendelpunten (unfiltered)`);

    // Determine headers from first item
    let headers = [];
    if (allData.length > 0) {
      headers = Object.keys(allData[0]);
    } else {
      // Fallback headers
      headers = ['Code', 'Geografische locatie', 'Naam', 'Naam geografische locatie', 'Sleutelbox', 'Type'];
    }

    // Convert objects to 2D array format (like Sheets)
    const data = allData.map(item =>
      headers.map(header => {
        const value = item[header];
        // Convert null/undefined to empty string
        return value !== null && value !== undefined ? String(value) : '';
      })
    );

    return {
      headers: headers,
      data: data
    };

  } catch (error) {
    console.error('Error in getAllFirebaseDataFormatted:', error);
    return { headers: [], data: [] };
  }
}

/**
 * Main function to get filtered sheet data
 * NOW USES FIREBASE instead of Google Sheets!
 *
 * @param {string|null} installation - Installation name to filter by, or null/empty for ALL data
 * @returns {Object} {headers: Array, data: Array[]}
 *
 * Usage:
 *   getFilteredSheetBData('APL3')  -> Returns only APL3 data
 *   getFilteredSheetBData(null)    -> Returns ALL data (no filter)
 *   getFilteredSheetBData('')      -> Returns ALL data (no filter)
 */
function getFilteredSheetBData(installation) {
  return getFilteredFirebaseData(installation);
}

/**
 * Prefetch Firebase data (no-op since we don't cache Firebase data due to size)
 */
function prefetchFirebaseData() {
  // Firebase data is fetched on-demand without caching
  // This function kept for backwards compatibility
  return { success: true };
}

/**
 * Get all essential data at once
 * Fetches installations from Sheets
 * Firebase data is fetched on-demand per installation
 */
function getAllData() {
  try {
    console.log('Fetching installations from Sheets...');
    
    // Fetch installations from Sheets
    const installations = getInstallations();
    
    return {
      installations: installations
    };
  } catch (error) {
    console.error('Error in getAllData:', error);
    return {
      installations: []
    };
  }
}

/**
 * Test function - check Firebase connection and filtering
 */
function testFirebaseConnection() {
  console.log('Testing Firebase connection...');
  
  const data = getAllVergrendelpunten();
  console.log(`Total items: ${data.length}`);
  
  if (data.length > 0) {
    console.log('First item:', JSON.stringify(data[0], null, 2));
    console.log('Headers:', Object.keys(data[0]));
    
    // Test filtering with a sample installation
    const sampleCode = data[0]['Code'] || data[0]['code'] || '';
    const testInstallation = sampleCode.split('-')[0]; // Get prefix before first dash
    
    console.log(`\nTesting filter for installation: ${testInstallation}`);
    const filtered = getFilteredFirebaseData(testInstallation);
    console.log(`Filtered results: ${filtered.data.length} items`);
    
    if (filtered.data.length > 0) {
      console.log('Sample filtered codes:', 
        filtered.data.slice(0, 5).map(row => row[0]).join(', ')
      );
    }
  }
  
  return {
    success: data.length > 0,
    totalItems: data.length,
    sampleItem: data[0] || null
  };
}

// ===== AANVRAGEN SUBMISSION =====

/**
 * Generate unique Request ID
 * Format: 5a2e681f-b775-4f99-8d8d-4123442dc8cd
 */
function generateRequestId() {
  return Utilities.getUuid();
}

/**
 * Generate Approval Token (for future approval flow)
 * 32 character hex string
 */
function generateApprovalToken() {
  const chars = '0123456789abcdef';
  let token = '';
  for (let i = 0; i < 32; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

/**
 * Write data to Firebase Aanvragen database
 * @param {Object} data - Data object to write
 * @returns {Object} {success: boolean, requestId: string, error?: string}
 */
function writeToFirebaseAanvragen(data) {
  try {
    const url = `${FIREBASE_AANVRAGEN_URL}/uploadData.json?auth=${FIREBASE_AANVRAGEN_SECRET}`;
    
    const payload = JSON.stringify(data);
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: payload,
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const result = JSON.parse(response.getContentText());
      console.log('✓ Successfully wrote to Firebase Aanvragen:', result);
      return {
        success: true,
        firebaseKey: result.name // Firebase auto-generated key
      };
    } else {
      console.error('Firebase write error:', responseCode, response.getContentText());
      return {
        success: false,
        error: `HTTP ${responseCode}: ${response.getContentText()}`
      };
    }
  } catch (error) {
    console.error('Error writing to Firebase Aanvragen:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Submit aanvraag - saves to Firebase Aanvragen database
 * @param {Object} formData - Form data from modal
 * @param {Array} selectedRows - Array of selected row objects
 * @param {string} userName - Name of requester
 * @returns {Object} {success: boolean, requestId: string, aanvraagNr?: string, error?: string}
 */
function submitAanvraag(formData, selectedRows, userName) {
  try {
    console.log('Submitting aanvraag...');
    console.log('Form data:', JSON.stringify(formData, null, 2));
    console.log('Selected rows:', selectedRows.length);
    
    // Generate unique IDs
    const requestId = generateRequestId();
    const approvalToken = generateApprovalToken();
    const timestamp = new Date().toISOString();
    
    // Get requester email
    const requesterEmail = getUserEmail();
    
    // Prepare Form Data object (matches screenshot structure)
    const formDataObject = {
      installatie: formData.installatie,
      naam: formData.naam,
      interneOpmerking: formData.interneOpmerking || '',
      externeOpmerking: formData.externeOpmerking || '',
      opmerkingen: formData.opmerkingen || ''
    };
    
    // Prepare selected items array (verzamelmandje data)
    const verzamelmandjeData = selectedRows.map(row => {
      // row.data is the array of cell values
      // We want to create a readable string or object
      const rowData = row.data;
      
      // Assuming headers match: Code, Geografische locatie, Naam, etc.
      return rowData.join(','); // Simple comma-separated for now
    });
    
    // Create the aanvraag object matching screenshot structure
    const aanvraagData = {
      "Request ID": requestId,
      "Approval Token": approvalToken,
      "Approval Link": "", // Empty for now, will be generated later
      "Requester Email": requesterEmail,
      "Approver Email": formData.autorisator,
      "Status": "Pending",
      "Timestamp": timestamp,
      "Decision Timestamp": "", // Empty until approved/rejected
      "Authorizer Remark": "", // Empty until decision made
      "Form Data": JSON.stringify(formDataObject),
      "Verzamelmandje Data": verzamelmandjeData
    };
    
    console.log('Aanvraag data prepared:', JSON.stringify(aanvraagData, null, 2));
    
    // Write to Firebase
    const result = writeToFirebaseAanvragen(aanvraagData);
    
    if (result.success) {
      console.log('✓ Aanvraag submitted successfully');
      console.log('Request ID:', requestId);
      console.log('Firebase Key:', result.firebaseKey);
      
      return {
        success: true,
        requestId: requestId,
        aanvraagNr: requestId.split('-')[0], // First part of UUID as display number
        firebaseKey: result.firebaseKey
      };
    } else {
      console.error('Failed to submit aanvraag:', result.error);
      return {
        success: false,
        error: result.error
      };
    }
    
  } catch (error) {
    console.error('Error in submitAanvraag:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get all aanvragen from Firebase Aanvragen database
 * @returns {Array} Array of aanvraag objects
 */
function getAanvragen() {
  try {
    console.log('Fetching aanvragen from Firebase...');
    
    const url = `${FIREBASE_AANVRAGEN_URL}/uploadData.json?auth=${FIREBASE_AANVRAGEN_SECRET}`;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      console.error('Firebase error:', responseCode, response.getContentText());
      return [];
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data) {
      console.log('No aanvragen found');
      return [];
    }
    
    // Convert Firebase object to array with keys
    const aanvragen = Object.keys(data).map(key => ({
      firebaseKey: key,
      ...data[key]
    }));
    
    console.log(`✓ Loaded ${aanvragen.length} aanvragen from Firebase`);
    
    return aanvragen;
    
  } catch (error) {
    console.error('Error fetching aanvragen:', error);
    return [];
  }
}

/**
 * Delete aanvraag from Firebase Aanvragen database
 * @param {string} firebaseKey - Firebase key of the aanvraag to delete
 * @returns {Object} {success: boolean, error?: string}
 */
function deleteAanvraag(firebaseKey) {
  try {
    console.log('Deleting aanvraag:', firebaseKey);
    
    if (!firebaseKey) {
      return {
        success: false,
        error: 'Geen Firebase key opgegeven'
      };
    }
    
    const url = `${FIREBASE_AANVRAGEN_URL}/uploadData/${firebaseKey}.json?auth=${FIREBASE_AANVRAGEN_SECRET}`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'delete',
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      console.log('✓ Successfully deleted aanvraag:', firebaseKey);
      return {
        success: true
      };
    } else {
      console.error('Firebase delete error:', responseCode, response.getContentText());
      return {
        success: false,
        error: `HTTP ${responseCode}: ${response.getContentText()}`
      };
    }
    
  } catch (error) {
    console.error('Error deleting aanvraag:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Test function - test reading aanvragen from Firebase
 */
function testGetAanvragen() {
  console.log('Testing getAanvragen...');
  
  const aanvragen = getAanvragen();
  console.log(`Total aanvragen: ${aanvragen.length}`);
  
  if (aanvragen.length > 0) {
    console.log('First aanvraag:', JSON.stringify(aanvragen[0], null, 2));
  }
  
  return {
    success: true,
    totalAanvragen: aanvragen.length,
    sampleAanvraag: aanvragen[0] || null
  };
}

/**
 * Test function - test writing to Firebase Aanvragen database
 */
function testAanvraagSubmission() {
  console.log('Testing aanvraag submission...');
  
  // Mock data
  const testFormData = {
    naam: 'Test Aanvraag',
    installatie: 'BAL1',
    interneOpmerking: 'Dit is een test',
    externeOpmerking: 'Test externe opmerking',
    opmerkingen: 'Test opmerkingen',
    autorisator: 'test@example.com'
  };
  
  const testSelectedRows = [
    {
      id: 'test-row-1',
      data: ['BAL1-001', 'Test Type', 'KG_zone1', 'Test Naam', 'Test Locatie', 'Mechanisch'],
      originalIndex: 0
    }
  ];
  
  const result = submitAanvraag(testFormData, testSelectedRows, 'Test User');
  
  console.log('Test result:', JSON.stringify(result, null, 2));
  
  return result;
}

// ===== HTML SERVING =====

/**
 * Serves the HTML page
 */
function doGet(e) {
  // Check if debug page is requested
  const page = e && e.parameter && e.parameter.page;

  if (page === 'debug') {
    // Serve debug page
    return HtmlService.createHtmlOutputFromFile('debug')
        .setTitle('Debug & Test - Aanvragen Rimses')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Default: serve main application
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Vergrendelgroepaanvraag')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Includes HTML file content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== USER AUTHENTICATION & INFO =====

/**
 * Gets the current user's name from their email
 */
function getUserName() {
  const email = Session.getActiveUser().getEmail();
  const parts = email.split('@')[0].split('.');
  
  if (parts.length >= 2) {
    const surname = parts[0].charAt(0).toUpperCase() + parts[0].slice(1);
    const name = parts[1].charAt(0).toUpperCase() + parts[1].slice(1);
    return surname + ' ' + name;
  }
  
  return email;
}

/**
 * Gets the current user's email
 */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Sends a test email (for debug purposes)
 * @returns {Object} Result of the email sending operation
 */
function sendTestEmail() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date().toLocaleString('nl-NL');
    const fromEmail = 'reliabilitycmms@gmail.com';
    const fromName = 'Reliability CMMS';

    const subject = '🔧 Test Email from Aanvragen Rimses Debug Panel';
    const body = `
Hallo,

Dit is een test email verzonden vanuit het Debug & Test panel van de Aanvragen Rimses applicatie.

Details:
- Verzonden door gebruiker: ${userEmail}
- Verzonden vanuit account: ${fromEmail}
- Tijdstip: ${timestamp}
- Applicatie: Aanvragen Rimses v2.8.9-EMAIL-ALIAS
- Environment: Google Apps Script

Als je deze email ontvangt, betekent dit dat de email functionaliteit correct werkt!

Met vriendelijke groet,
Reliability CMMS
Aanvragen Rimses Systeem
    `;

    // Send via GmailApp with alias
    GmailApp.sendEmail(
      'rob.oversteyns@gmail.com',
      subject,
      body,
      {
        from: fromEmail,
        name: fromName
      }
    );

    return {
      success: true,
      message: 'Test email succesvol verzonden naar rob.oversteyns@gmail.com vanuit ' + fromEmail,
      timestamp: timestamp,
      from: fromEmail
    };

  } catch (error) {
    console.error('Error sending test email:', error);
    return {
      success: false,
      message: 'Fout bij verzenden email: ' + error.message,
      error: error.toString()
    };
  }
}

/**
 * Validates user credentials against DATA sheet
 */
function validateUser(username, password) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === username && data[i][2] === password) {
      return true;
    }
  }
  
  return false;
}

// ===== GOOGLE SHEETS DATA (Installations & Autorisators) =====

/**
 * Gets list of installations from DATA sheet, column A (with caching)
 */
function getInstallations() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('installations');
  
  if (cachedData != null) {
    return JSON.parse(cachedData);
  }
  
  return fetchAndCacheInstallations();
}

/**
 * Fetches and caches installations
 */
function fetchAndCacheInstallations() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('DATA');
    
    if (!sheet) {
      throw new Error('Sheet DATA not found');
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }
    
    // Get column A values, starting from row 2 (skip header)
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    // Filter out empty values and flatten array
    const installations = data
      .map(row => row[0])
      .filter(value => value !== null && value !== undefined && value !== '');
    
    // Cache the result
    const cache = CacheService.getScriptCache();
    cache.put('installations', JSON.stringify(installations), CACHE_DURATION_INSTALLATIONS);
    
    return installations;
  } catch (error) {
    console.error('Error fetching installations:', error);
    return [];
  }
}

/**
 * Prefetch installations (called during login typing)
 */
function prefetchInstallations() {
  try {
    fetchAndCacheInstallations();
    return { success: true };
  } catch (error) {
    console.error('Error prefetching installations:', error);
    return { success: false };
  }
}

/**
 * Gets list of autorisators from DATA sheet, column B
 */
function getAutorisators() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('DATA');
    
    if (!sheet) {
      throw new Error('Sheet DATA not found');
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }
    
    // Get column B values, starting from row 2 (skip header)
    const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    
    // Filter out empty values and flatten array
    const autorisators = data
      .map(row => row[0])
      .filter(value => value !== null && value !== undefined && value !== '');
    
    return autorisators;
  } catch (error) {
    console.error('Error fetching autorisators:', error);
    return [];
  }
}

// ===== LEGACY / BACKWARDS COMPATIBILITY =====

/**
 * Prefetch all data at once for faster performance
 */
function prefetchAllData() {
  try {
    fetchAndCacheInstallations();
    // Firebase data fetched on-demand (no prefetch due to size)
    return { success: true };
  } catch (error) {
    console.error('Error prefetching all data:', error);
    return { success: false };
  }
}

/**
 * Refreshes all caches
 */
function refreshAllCaches() {
  fetchAndCacheInstallations();
  return { success: true };
}