// ===== Google Apps Script Backend =====
// Version: 2.40.1-LOGIN-VERSION-DISPLAY
// Last Updated: November 2025

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

// Firebase for Users database
const FIREBASE_USERS_URL = 'https://users-6e913.firebaseio.com';
const FIREBASE_USERS_SECRET = 'BBcmkVVW6jrsfA3GSJI0f7NovJNJ8yN8lKOQRzrK';

// Firebase for Criticaliteit database
const FIREBASE_CRITICALITY_URL = 'https://criticality-32c16-default-rtdb.europe-west1.firebasedatabase.app';
const FIREBASE_CRITICALITY_SECRET = 'uU8khW9s4Zg4mAcPSjwuXhNfBNAtxLrV4tBHtftP';

// Firebase for LOG IN LOG database
const FIREBASE_LOGIN_LOG_URL = 'https://userlogin-f7339-default-rtdb.europe-west1.firebasedatabase.app/';
const FIREBASE_LOGIN_LOG_SECRET = 'eStXHU0LSnsvPime0pgGNQOq78VQSgOYiHumh7hn';


// ===== HELPER FUNCTIONS =====

/**
 * Encode email address for Firebase-safe key
 * Firebase doesn't allow . $ # [ ] / in keys, so we replace them
 * @param {string} email - Email address to encode
 * @returns {string} Firebase-safe key
 */
function encodeEmailForFirebase(email) {
  if (!email) return '';
  return email
    .replace(/\./g, '_DOT_')
    .replace(/@/g, '_AT_')
    .replace(/\$/g, '_DOLLAR_')
    .replace(/#/g, '_HASH_')
    .replace(/\[/g, '_LBRACKET_')
    .replace(/\]/g, '_RBRACKET_')
    .replace(/\//g, '_SLASH_');
}

/**
 * Decode Firebase key back to email address
 * @param {string} key - Firebase key to decode
 * @returns {string} Original email address
 */
function decodeEmailFromFirebase(key) {
  if (!key) return '';
  return key
    .replace(/_DOT_/g, '.')
    .replace(/_AT_/g, '@')
    .replace(/_DOLLAR_/g, '$')
    .replace(/_HASH_/g, '#')
    .replace(/_LBRACKET_/g, '[')
    .replace(/_RBRACKET_/g, ']')
    .replace(/_SLASH_/g, '/');
}

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
function submitAanvraag(formData, selectedRows, userName, requesterEmail) {
  try {
    console.log('Submitting aanvraag...');
    console.log('Form data:', JSON.stringify(formData, null, 2));
    console.log('Selected rows:', selectedRows.length);
    console.log('Requester email (from login):', requesterEmail);

    // Generate unique IDs
    const requestId = generateRequestId();
    const approvalToken = generateApprovalToken();
    const timestamp = new Date().toISOString();

    // Requester email now comes from login username parameter
    // This works correctly with "Execute as: Me" deployment
    
    // Prepare Form Data object (matches screenshot structure)
    const formDataObject = {
      installatie: formData.installatie,
      naam: formData.naam,
      interneOpmerking: formData.interneOpmerking || '',
      externeOpmerking: formData.externeOpmerking || '',
      opmerkingen: formData.opmerkingen || ''
    };
    
    // Prepare selected items array (vergrendelpunten with structured data)
    const lockedPoints = selectedRows.map(row => {
      const rowData = row.data;
      // rowData structure: [Code, Type, Sleutelbox, Naam, Naam Geografisch, ...]
      return {
        code: rowData[0] || '',
        type: rowData[1] || '',
        sleutelbox: rowData[2] || '',
        naam: rowData[3] || '',
        geografisch: rowData[4] || ''
      };
    });

    // Also keep verzamelmandje data for backwards compatibility
    const verzamelmandjeData = selectedRows.map(row => row.data.join(','));
    
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
      "Locked Points": JSON.stringify(lockedPoints),
      "Verzamelmandje Data": verzamelmandjeData
    };
    
    console.log('Aanvraag data prepared:', JSON.stringify(aanvraagData, null, 2));
    
    // Write to Firebase
    const result = writeToFirebaseAanvragen(aanvraagData);
    
    if (result.success) {
      console.log('✓ Aanvraag submitted successfully');
      console.log('Request ID:', requestId);
      console.log('Firebase Key:', result.firebaseKey);

      // Send notification email to authorizer
      console.log('Sending notification email to authorizer:', formData.autorisator);
      const emailResult = sendAuthorizerNotification(aanvraagData, formData, selectedRows, userName);

      if (emailResult.success) {
        console.log('✓ Notification email sent successfully to authorizer');
      } else {
        console.warn('⚠ Failed to send notification email to authorizer:', emailResult.message);
        // Continue anyway - the aanvraag was saved successfully
      }

      // Send confirmation email to requester
      console.log('Sending confirmation email to requester:', requesterEmail);
      const requesterEmailResult = sendRequesterConfirmation(aanvraagData, formData, selectedRows, userName);

      if (requesterEmailResult.success) {
        console.log('✓ Confirmation email sent successfully to requester');
      } else {
        console.warn('⚠ Failed to send confirmation email to requester:', requesterEmailResult.message);
        // Continue anyway - the aanvraag was saved successfully
      }

      return {
        success: true,
        requestId: requestId,
        aanvraagNr: requestId.split('-')[0], // First part of UUID as display number
        firebaseKey: result.firebaseKey,
        emailSent: emailResult.success && requesterEmailResult.success,
        emailMessage: `Authorizer: ${emailResult.message}, Requester: ${requesterEmailResult.message}`
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
 * Get all users from Firebase Users database
 * @returns {Array} Array of user objects
 */
function getUsers() {
  try {
    console.log('Fetching users from Firebase...');

    const url = `${FIREBASE_USERS_URL}/fire-data.json?auth=${FIREBASE_USERS_SECRET}`;
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
      console.log('No users found');
      return [];
    }

    // Convert Firebase object to array with keys
    const users = Object.keys(data).map(key => ({
      firebaseKey: key,
      ...data[key]
    }));

    console.log(`✓ Loaded ${users.length} users from Firebase`);

    return users;

  } catch (error) {
    console.error('Error fetching users:', error);
    return [];
  }
}

/**
 * Update user in Firebase Users database
 * @param {Object} userData - User data to update
 * @returns {Object} {success: boolean, message: string}
 */
function updateUser(userData) {
  try {
    console.log('Updating user:', userData.firebaseKey);

    // Prepare update data (exclude firebaseKey and newEmail from update object)
    const updateData = {
      ROLE: userData.ROLE,
      DEPARTEMENT: userData.DEPARTEMENT || '',
      TITEL: userData.TITEL || '',
      ACTIVE: userData.ACTIVE,
      USER_NAME: userData.USER_NAME
    };

    // Only include password if it's being updated
    if (userData.PASSWORD && userData.PASSWORD.trim() !== '') {
      updateData.PASSWORD = userData.PASSWORD;
    }

    // If email is changing, we need to delete old entry and create new one
    if (userData.newEmail && userData.newEmail !== userData.firebaseKey) {
      console.log('Email is changing from', userData.firebaseKey, 'to', userData.newEmail);

      // First, get the current user data to preserve PASSWORD if not updating
      const getCurrentUrl = `${FIREBASE_USERS_URL}/fire-data/${userData.firebaseKey}.json?auth=${FIREBASE_USERS_SECRET}`;
      const currentResponse = UrlFetchApp.fetch(getCurrentUrl, {
        method: 'get',
        muteHttpExceptions: true
      });

      const currentData = JSON.parse(currentResponse.getContentText());

      // Preserve password if not updating
      if (!updateData.PASSWORD && currentData.PASSWORD) {
        updateData.PASSWORD = currentData.PASSWORD;
      }

      // Create new entry with new email as key
      const createUrl = `${FIREBASE_USERS_URL}/fire-data/${userData.newEmail}.json?auth=${FIREBASE_USERS_SECRET}`;
      UrlFetchApp.fetch(createUrl, {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify(updateData),
        muteHttpExceptions: true
      });

      // Delete old entry
      const deleteUrl = `${FIREBASE_USERS_URL}/fire-data/${userData.firebaseKey}.json?auth=${FIREBASE_USERS_SECRET}`;
      UrlFetchApp.fetch(deleteUrl, {
        method: 'delete',
        muteHttpExceptions: true
      });

      console.log('✓ User email updated successfully');
      return {
        success: true,
        message: 'Gebruiker succesvol bijgewerkt (email gewijzigd)'
      };
    } else {
      // Email not changing, just update existing entry
      // First get current password if we're not updating it
      if (!updateData.PASSWORD) {
        const getCurrentUrl = `${FIREBASE_USERS_URL}/fire-data/${userData.firebaseKey}.json?auth=${FIREBASE_USERS_SECRET}`;
        const currentResponse = UrlFetchApp.fetch(getCurrentUrl, {
          method: 'get',
          muteHttpExceptions: true
        });
        const currentData = JSON.parse(currentResponse.getContentText());
        if (currentData.PASSWORD) {
          updateData.PASSWORD = currentData.PASSWORD;
        }
      }

      const url = `${FIREBASE_USERS_URL}/fire-data/${userData.firebaseKey}.json?auth=${FIREBASE_USERS_SECRET}`;
      const response = UrlFetchApp.fetch(url, {
        method: 'patch',
        contentType: 'application/json',
        payload: JSON.stringify(updateData),
        muteHttpExceptions: true
      });

      const responseCode = response.getResponseCode();
      if (responseCode !== 200) {
        console.error('Firebase error:', responseCode, response.getContentText());
        return {
          success: false,
          message: 'Fout bij bijwerken gebruiker in database'
        };
      }

      console.log('✓ User updated successfully');
      return {
        success: true,
        message: 'Gebruiker succesvol bijgewerkt'
      };
    }

  } catch (error) {
    console.error('Error updating user:', error);
    return {
      success: false,
      message: 'Fout bij bijwerken: ' + error.toString()
    };
  }
}

/**
 * Delete user from Firebase Users database
 * @param {string} userEmail - Email (firebaseKey) of the user to delete
 * @returns {Object} {success: boolean, message: string}
 */
function deleteUser(userEmail) {
  try {
    console.log('Deleting user:', userEmail);

    if (!userEmail) {
      return {
        success: false,
        message: 'Geen email opgegeven'
      };
    }

    // Encode email for Firebase-safe key
    const encodedEmail = encodeEmailForFirebase(userEmail);
    const url = `${FIREBASE_USERS_URL}/fire-data/${encodedEmail}.json?auth=${FIREBASE_USERS_SECRET}`;

    // Delete from Firebase
    const response = UrlFetchApp.fetch(url, {
      method: 'delete',
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      console.log('✓ Successfully deleted user:', userEmail);
      return {
        success: true,
        message: 'Gebruiker succesvol verwijderd'
      };
    } else {
      console.error('Firebase delete error:', responseCode, response.getContentText());
      return {
        success: false,
        message: `Fout bij verwijderen: HTTP ${responseCode}`
      };
    }

  } catch (error) {
    console.error('Error deleting user:', error);
    return {
      success: false,
      message: 'Fout bij verwijderen: ' + error.toString()
    };
  }
}

/**
 * Create new user in Firebase Users database
 * @param {Object} userData - User data for new user
 * @returns {Object} {success: boolean, message: string}
 */
function createUser(userData) {
  try {
    console.log('Creating user:', userData.email);

    if (!userData.email) {
      return {
        success: false,
        message: 'Email is verplicht'
      };
    }

    if (!userData.PASSWORD) {
      return {
        success: false,
        message: 'Wachtwoord is verplicht'
      };
    }

    // Check if user already exists by getting all users and checking USER_NAME
    console.log('🔍 Checking for duplicate USER_NAME:', userData.email);

    const checkAllUrl = `${FIREBASE_USERS_URL}/fire-data.json?auth=${FIREBASE_USERS_SECRET}`;
    const checkAllResponse = UrlFetchApp.fetch(checkAllUrl, {
      method: 'get',
      muteHttpExceptions: true
    });

    const checkAllData = JSON.parse(checkAllResponse.getContentText());

    if (checkAllData) {
      // Check if any existing user has the same USER_NAME (email)
      const existingUser = Object.keys(checkAllData).find(key => {
        const user = checkAllData[key];
        return user && user.USER_NAME && user.USER_NAME.toLowerCase() === userData.email.toLowerCase();
      });

      if (existingUser) {
        console.log('❌ User already exists with USER_NAME:', userData.email);
        return {
          success: false,
          message: 'Gebruiker met dit email adres bestaat al'
        };
      }
    }

    console.log('✓ User does not exist, proceeding with creation');

    // Reuse the data we already fetched for finding the next numeric key
    const allUsersData = checkAllData;

    // Find the highest numeric key
    let nextKey = 0;
    if (allUsersData) {
      Object.keys(allUsersData).forEach(key => {
        const numKey = parseInt(key);
        if (!isNaN(numKey) && numKey > nextKey) {
          nextKey = numKey;
        }
      });
    }
    nextKey = nextKey + 1; // Increment for new user

    console.log('✓ Next user key will be:', nextKey);

    // Prepare user data
    const newUserData = {
      ROLE: userData.ROLE || 'A',
      DEPARTEMENT: userData.DEPARTEMENT || '',
      TITEL: userData.TITEL || '',
      ACTIVE: userData.ACTIVE !== undefined ? userData.ACTIVE : 1,
      USER_NAME: userData.USER_NAME || userData.email,
      PASSWORD: userData.PASSWORD
    };

    // Create new user in Firebase with numeric key
    const createUrl = `${FIREBASE_USERS_URL}/fire-data/${nextKey}.json?auth=${FIREBASE_USERS_SECRET}`;

    const createResponse = UrlFetchApp.fetch(createUrl, {
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify(newUserData),
      muteHttpExceptions: true
    });

    const createResponseCode = createResponse.getResponseCode();

    if (createResponseCode === 200) {
      console.log('✓ Successfully created user:', userData.email);
      return {
        success: true,
        message: 'Gebruiker succesvol aangemaakt'
      };
    } else {
      console.error('Firebase create error:', createResponseCode, createResponse.getContentText());
      return {
        success: false,
        message: `Fout bij aanmaken: HTTP ${createResponseCode}`
      };
    }

  } catch (error) {
    console.error('Error creating user:', error);
    return {
      success: false,
      message: 'Fout bij aanmaken: ' + error.toString()
    };
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
 * Update aanvraag decision (approve/reject) in Firebase Aanvragen database
 * @param {Object} decisionData - Object containing firebaseKey, decision, authorizerRemark, decisionTimestamp
 * @returns {Object} {success: boolean, message?: string, error?: string}
 */
function updateAanvraagDecision(decisionData) {
  try {
    console.log('Updating aanvraag decision:', decisionData);

    if (!decisionData || !decisionData.firebaseKey || !decisionData.decision) {
      return {
        success: false,
        error: 'Ongeldige decision data'
      };
    }

    const { firebaseKey, decision, authorizerRemark, decisionTimestamp } = decisionData;

    // Prepare update payload
    const updatePayload = {
      'Status': decision,
      'Authorizer Remark': authorizerRemark || '',
      'Decision Timestamp': decisionTimestamp
    };

    const url = `${FIREBASE_AANVRAGEN_URL}/uploadData/${firebaseKey}.json?auth=${FIREBASE_AANVRAGEN_SECRET}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'patch',
      contentType: 'application/json',
      payload: JSON.stringify(updatePayload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      console.log('✓ Successfully updated aanvraag decision:', firebaseKey);

      // Fetch the complete aanvraag data to send decision email
      const getUrl = `${FIREBASE_AANVRAGEN_URL}/uploadData/${firebaseKey}.json?auth=${FIREBASE_AANVRAGEN_SECRET}`;
      const getResponse = UrlFetchApp.fetch(getUrl, {
        method: 'get',
        muteHttpExceptions: true
      });

      if (getResponse.getResponseCode() === 200) {
        const aanvraagData = JSON.parse(getResponse.getContentText());

        // Send decision notification email to requester
        console.log('Sending decision email to requester:', aanvraagData['Requester Email']);
        const emailResult = sendDecisionNotification(aanvraagData, decision, authorizerRemark, decisionTimestamp);

        if (emailResult.success) {
          console.log('✓ Decision email sent successfully');
        } else {
          console.warn('⚠ Failed to send decision email:', emailResult.message);
          // Continue anyway - the decision was saved successfully
        }
      } else {
        console.warn('⚠ Could not fetch aanvraag data for email notification');
      }

      return {
        success: true,
        message: `Aanvraag ${decision === 'Approved' ? 'goedgekeurd' : 'afgekeurd'}`
      };
    } else {
      console.error('Firebase update error:', responseCode, response.getContentText());
      return {
        success: false,
        error: `HTTP ${responseCode}: ${response.getContentText()}`
      };
    }

  } catch (error) {
    console.error('Error updating aanvraag decision:', error);
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
 * VERWIJDERD: getUserName() - niet meer nodig
 * Username wordt nu direct uit login form gebruikt (werkt ongeacht deployment setting)
 * Dit voorkomt problemen met "Execute as: Me" deployment waarbij Session.getActiveUser()
 * altijd de publisher's email zou teruggeven i.p.v. de ingelogde gebruiker
 */

/**
 * Gets the current user's email
 * LET OP: Bij "Execute as: Me" deployment geeft dit de PUBLISHER's email,
 * niet de daadwerkelijke gebruiker die de app gebruikt!
 */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Sends notification email to authorizer when a new request is submitted
 * @param {Object} aanvraagData - The request data
 * @param {Object} formData - The form data
 * @param {Array} selectedRows - Selected vergrendelpunten
 * @param {string} userName - Name of the requester
 * @returns {Object} Result of email sending
 */
function sendAuthorizerNotification(aanvraagData, formData, selectedRows, userName) {
  try {
    const authorizerEmail = formData.autorisator;
    const requestId = aanvraagData['Request ID'];
    const timestamp = new Date(aanvraagData.Timestamp).toLocaleString('nl-NL');
    const requesterEmail = aanvraagData['Requester Email'];

    // Build HTML table rows for selected vergrendelpunten
    let vergrendelpuntenRows = '';
    selectedRows.forEach((row, index) => {
      const rowData = row.data;
      // Assuming structure: Code, Type, Sleutelbox, Naam, Naam Geografisch
      vergrendelpuntenRows += `
        <tr style="border-bottom: 1px solid #e0e0e0;">
          <td style="padding: 8px; font-size: 11px; color: #666; white-space: nowrap;">${index + 1}</td>
          <td style="padding: 8px; font-size: 11px; font-weight: 600; font-family: monospace; white-space: nowrap;">${rowData[0] || 'N/A'}</td>
          <td style="padding: 8px; font-size: 11px; white-space: nowrap;">${rowData[1] || 'N/A'}</td>
          <td style="padding: 8px; font-size: 11px; white-space: nowrap;">${rowData[2] || 'N/A'}</td>
          <td style="padding: 8px; font-size: 11px;">${rowData[3] || 'N/A'}</td>
          <td style="padding: 8px; font-size: 11px;">${rowData[4] || 'N/A'}</td>
        </tr>
      `;
    });

    const subject = `Nieuwe Aanvraag Vergrendelgroep - ${requestId}`;

    // Plain text fallback
    const plainTextBody = `
Beste,

Er is een nieuwe aanvraag voor een vergrendelgroep ingediend die uw goedkeuring vereist.

Aanvraag Nummer: ${requestId}
Ingediend op: ${timestamp}
Aanvrager: ${requesterEmail}

Installatie: ${formData.installatie || 'N/A'}
Naam: ${formData.naam || 'N/A'}

Deze aanvraag wacht op uw goedkeuring. Log in op de Aanvragen Rimses applicatie om de aanvraag te beoordelen.

Met vriendelijke groet,
Reliability CMMS
    `;

    // HTML email body
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Arial, sans-serif; background-color: #f5f5f5;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f5f5f5; padding: 20px;">
    <tr>
      <td align="center">
        <!-- Main container -->
        <table width="750" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">

          <!-- Header -->
          <tr>
            <td style="background: linear-gradient(135deg, #1a237e 0%, #4a148c 100%); padding: 30px; text-align: center;">
              <h1 style="margin: 0; color: #ffffff; font-size: 24px; font-weight: 700;">
                Nieuwe Aanvraag Vergrendelgroep
              </h1>
              <p style="margin: 10px 0 0 0; color: #e0e0e0; font-size: 14px;">
                Reliability CMMS - Aanvragen Rimses
              </p>
            </td>
          </tr>

          <!-- Content -->
          <tr>
            <td style="padding: 30px;">

              <!-- Introduction -->
              <p style="margin: 0 0 20px 0; color: #333; font-size: 15px; line-height: 1.6;">
                Beste,
              </p>
              <p style="margin: 0 0 25px 0; color: #666; font-size: 14px; line-height: 1.6;">
                Er is een nieuwe aanvraag voor een vergrendelgroep ingediend die uw goedkeuring vereist.
              </p>

              <!-- Aanvraag Details -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #e8f5e9; border-left: 4px solid #4caf50; border-radius: 4px; margin-bottom: 25px;">
                <tr>
                  <td style="padding: 15px;">
                    <h3 style="margin: 0 0 12px 0; color: #2e7d32; font-size: 16px; font-weight: 600;">
                      Aanvraag Details
                    </h3>
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600; width: 140px;">Aanvraag Nummer:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px; font-family: monospace;">${requestId}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Ingediend op:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${timestamp}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Aanvrager:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${requesterEmail}</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Formulier Gegevens -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #e3f2fd; border-left: 4px solid #2196f3; border-radius: 4px; margin-bottom: 25px;">
                <tr>
                  <td style="padding: 15px;">
                    <h3 style="margin: 0 0 12px 0; color: #1565c0; font-size: 16px; font-weight: 600;">
                      Formulier Gegevens
                    </h3>
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600; width: 140px;">Installatie:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.installatie || 'N/A'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Naam:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.naam || 'N/A'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Interne Opmerking:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.interneOpmerking || 'Geen'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Externe Opmerking:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.externeOpmerking || 'Geen'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Opmerkingen:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.opmerkingen || 'Geen'}</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Vergrendelpunten Table -->
              <div style="margin-bottom: 25px;">
                <h3 style="margin: 0 0 12px 0; color: #333; font-size: 16px; font-weight: 600;">
                  Geselecteerde Vergrendelpunten (${selectedRows.length})
                </h3>
                <table width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;">
                  <thead>
                    <tr style="background-color: #f5f5f5;">
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">#</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Code</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Type</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Sleutelbox</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Naam</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Locatie</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${vergrendelpuntenRows}
                  </tbody>
                </table>
              </div>

              <!-- Call to Action -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #fff3e0; border-left: 4px solid #ff9800; border-radius: 4px; margin-bottom: 20px;">
                <tr>
                  <td style="padding: 15px;">
                    <h3 style="margin: 0 0 8px 0; color: #e65100; font-size: 15px; font-weight: 600;">
                      Actie Vereist
                    </h3>
                    <p style="margin: 0; color: #666; font-size: 13px; line-height: 1.6;">
                      Deze aanvraag wacht op uw goedkeuring. Log in op de Aanvragen Rimses applicatie om de aanvraag te beoordelen en goed te keuren of af te wijzen.
                    </p>
                  </td>
                </tr>
              </table>

              <!-- Closing -->
              <p style="margin: 0; color: #666; font-size: 13px; line-height: 1.6;">
                Met vriendelijke groet,<br>
                <strong>Reliability CMMS</strong><br>
                Aanvragen Rimses Systeem
              </p>

            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background-color: #f5f5f5; padding: 15px; text-align: center; border-top: 1px solid #e0e0e0;">
              <p style="margin: 0; color: #999; font-size: 11px;">
                Dit is een automatisch gegenereerde email. Gelieve niet te antwoorden op dit bericht.
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>
    `;

    // Send email via GmailApp with alias
    GmailApp.sendEmail(
      authorizerEmail,
      subject,
      plainTextBody,
      {
        from: 'reliabilitycmms@gmail.com',
        name: 'Reliability CMMS',
        htmlBody: htmlBody
      }
    );

    console.log('✓ Notification email sent to:', authorizerEmail);

    return {
      success: true,
      message: 'Notification email sent to authorizer: ' + authorizerEmail
    };

  } catch (error) {
    console.error('Error sending authorizer notification:', error);
    return {
      success: false,
      message: 'Failed to send notification email: ' + error.message,
      error: error.toString()
    };
  }
}

/**
 * Send confirmation email to requester when new aanvraag is submitted
 * @param {Object} aanvraagData - The aanvraag data
 * @param {Object} formData - The form data
 * @param {Array} selectedRows - Selected vergrendelpunten
 * @param {string} userName - Name of requester
 * @returns {Object} Result object with success status
 */
function sendRequesterConfirmation(aanvraagData, formData, selectedRows, userName) {
  try {
    const requesterEmail = aanvraagData['Requester Email'];
    const authorizerEmail = formData.autorisator;
    const requestId = aanvraagData['Request ID'];
    const timestamp = new Date(aanvraagData.Timestamp).toLocaleString('nl-NL');

    // Build HTML table rows for selected vergrendelpunten
    let vergrendelpuntenRows = '';
    selectedRows.forEach((row, index) => {
      const rowData = row.data;
      vergrendelpuntenRows += `
        <tr style="border-bottom: 1px solid #e0e0e0;">
          <td style="padding: 8px; font-size: 11px; color: #666; white-space: nowrap;">${index + 1}</td>
          <td style="padding: 8px; font-size: 11px; font-weight: 600; font-family: monospace; white-space: nowrap;">${rowData[0] || 'N/A'}</td>
          <td style="padding: 8px; font-size: 11px; white-space: nowrap;">${rowData[1] || 'N/A'}</td>
          <td style="padding: 8px; font-size: 11px; white-space: nowrap;">${rowData[2] || 'N/A'}</td>
          <td style="padding: 8px; font-size: 11px;">${rowData[3] || 'N/A'}</td>
          <td style="padding: 8px; font-size: 11px;">${rowData[4] || 'N/A'}</td>
        </tr>
      `;
    });

    const subject = `Bevestiging Aanvraag Vergrendelgroep - ${requestId}`;

    // Plain text fallback
    const plainTextBody = `
Beste ${userName || requesterEmail},

Uw aanvraag voor een vergrendelgroep is succesvol ingediend.

Aanvraag Nummer: ${requestId}
Ingediend op: ${timestamp}
Autorisator: ${authorizerEmail}

Installatie: ${formData.installatie || 'N/A'}
Naam: ${formData.naam || 'N/A'}

Uw aanvraag wordt beoordeeld door de autorisator. U ontvangt een email wanneer de aanvraag is goedgekeurd of afgewezen.

Met vriendelijke groet,
Reliability CMMS
    `;

    // HTML email body
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Arial, sans-serif; background-color: #f5f5f5;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f5f5f5; padding: 20px;">
    <tr>
      <td align="center">
        <!-- Main container -->
        <table width="750" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">

          <!-- Header -->
          <tr>
            <td style="background: linear-gradient(135deg, #1976d2 0%, #42a5f5 100%); padding: 30px; text-align: center;">
              <h1 style="margin: 0; color: #ffffff; font-size: 24px; font-weight: 700;">
                ✓ Aanvraag Bevestiging
              </h1>
              <p style="margin: 10px 0 0 0; color: #e0e0e0; font-size: 14px;">
                Reliability CMMS - Aanvragen Rimses
              </p>
            </td>
          </tr>

          <!-- Content -->
          <tr>
            <td style="padding: 30px;">

              <!-- Introduction -->
              <p style="margin: 0 0 20px 0; color: #333; font-size: 15px; line-height: 1.6;">
                Beste ${userName || requesterEmail},
              </p>
              <p style="margin: 0 0 25px 0; color: #666; font-size: 14px; line-height: 1.6;">
                Uw aanvraag voor een vergrendelgroep is succesvol ingediend en wacht op goedkeuring van de autorisator.
              </p>

              <!-- Aanvraag Details -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #e8f5e9; border-left: 4px solid #4caf50; border-radius: 4px; margin-bottom: 25px;">
                <tr>
                  <td style="padding: 15px;">
                    <h3 style="margin: 0 0 12px 0; color: #2e7d32; font-size: 16px; font-weight: 600;">
                      Aanvraag Details
                    </h3>
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600; width: 140px;">Aanvraag Nummer:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px; font-family: monospace;">${requestId}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Ingediend op:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${timestamp}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Autorisator:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${authorizerEmail}</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Formulier Gegevens -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #e3f2fd; border-left: 4px solid #2196f3; border-radius: 4px; margin-bottom: 25px;">
                <tr>
                  <td style="padding: 15px;">
                    <h3 style="margin: 0 0 12px 0; color: #1565c0; font-size: 16px; font-weight: 600;">
                      Formulier Gegevens
                    </h3>
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600; width: 140px;">Installatie:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.installatie || 'N/A'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Naam:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.naam || 'N/A'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Interne Opmerking:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.interneOpmerking || 'Geen'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Externe Opmerking:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.externeOpmerking || 'Geen'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Opmerkingen:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.opmerkingen || 'Geen'}</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Vergrendelpunten Table -->
              <div style="margin-bottom: 25px;">
                <h3 style="margin: 0 0 12px 0; color: #333; font-size: 16px; font-weight: 600;">
                  Geselecteerde Vergrendelpunten (${selectedRows.length})
                </h3>
                <table width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;">
                  <thead>
                    <tr style="background-color: #f5f5f5;">
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">#</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Code</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Type</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Sleutelbox</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Naam</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Locatie</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${vergrendelpuntenRows}
                  </tbody>
                </table>
              </div>

              <!-- Info Box -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #fff3e0; border-left: 4px solid #ff9800; border-radius: 4px; margin-bottom: 20px;">
                <tr>
                  <td style="padding: 15px;">
                    <h3 style="margin: 0 0 8px 0; color: #e65100; font-size: 15px; font-weight: 600;">
                      Volgende Stap
                    </h3>
                    <p style="margin: 0; color: #666; font-size: 13px; line-height: 1.6;">
                      Uw aanvraag wordt nu beoordeeld door de autorisator. U ontvangt een nieuwe email wanneer de aanvraag is goedgekeurd of afgewezen.
                    </p>
                  </td>
                </tr>
              </table>

              <!-- Closing -->
              <p style="margin: 0; color: #666; font-size: 13px; line-height: 1.6;">
                Met vriendelijke groet,<br>
                <strong>Reliability CMMS</strong><br>
                Aanvragen Rimses Systeem
              </p>

            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background-color: #f5f5f5; padding: 15px; text-align: center; border-top: 1px solid #e0e0e0;">
              <p style="margin: 0; color: #999; font-size: 11px;">
                Dit is een automatisch gegenereerde email. Gelieve niet te antwoorden op dit bericht.
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>
    `;

    // Send email via GmailApp with alias
    GmailApp.sendEmail(
      requesterEmail,
      subject,
      plainTextBody,
      {
        from: 'reliabilitycmms@gmail.com',
        name: 'Reliability CMMS',
        htmlBody: htmlBody
      }
    );

    console.log('✓ Confirmation email sent to requester:', requesterEmail);

    return {
      success: true,
      message: 'Confirmation email sent to requester: ' + requesterEmail
    };

  } catch (error) {
    console.error('Error sending requester confirmation:', error);
    return {
      success: false,
      message: 'Failed to send confirmation email: ' + error.message,
      error: error.toString()
    };
  }
}

/**
 * Send decision notification email to requester when aanvraag is approved/rejected
 * @param {Object} aanvraagData - The complete aanvraag data from Firebase
 * @param {string} decision - 'Approved' or 'Rejected'
 * @param {string} authorizerRemark - Remark from authorizer
 * @param {string} decisionTimestamp - When the decision was made
 * @returns {Object} Result object with success status
 */
function sendDecisionNotification(aanvraagData, decision, authorizerRemark, decisionTimestamp) {
  try {
    const requesterEmail = aanvraagData['Requester Email'];
    const requesterName = aanvraagData['Requester Name'] || requesterEmail;
    const authorizerEmail = aanvraagData['Approver Email'];
    const requestId = aanvraagData['Request ID'];
    const timestamp = new Date(aanvraagData.Timestamp).toLocaleString('nl-NL');
    const decisionDate = new Date(decisionTimestamp).toLocaleString('nl-NL');

    const isApproved = decision === 'Approved';
    const statusText = isApproved ? 'Goedgekeurd' : 'Afgekeurd';
    const statusColor = isApproved ? '#4caf50' : '#f44336';
    const statusBgColor = isApproved ? '#e8f5e9' : '#ffebee';
    const statusIcon = isApproved ? '✓' : '✗';

    // Parse Form Data (JSON string) to get formulier gegevens
    let formData = {};
    try {
      formData = JSON.parse(aanvraagData['Form Data'] || '{}');
      console.log('Parsed Form Data:', formData);
    } catch (e) {
      console.error('Error parsing Form Data:', e);
      formData = {
        installatie: 'N/A',
        naam: 'N/A',
        interneOpmerking: 'N/A',
        externeOpmerking: 'N/A',
        opmerkingen: 'N/A'
      };
    }

    // Parse vergrendelpunten from 'Locked Points' field (JSON string)
    let vergrendelpuntenRows = '';
    try {
      console.log('Raw Locked Points data:', aanvraagData['Locked Points']);
      const lockedPoints = JSON.parse(aanvraagData['Locked Points'] || '[]');
      console.log('Parsed Locked Points:', lockedPoints);
      lockedPoints.forEach((point, index) => {
        vergrendelpuntenRows += `
          <tr style="border-bottom: 1px solid #e0e0e0;">
            <td style="padding: 8px; font-size: 11px; color: #666; white-space: nowrap;">${index + 1}</td>
            <td style="padding: 8px; font-size: 11px; font-weight: 600; font-family: monospace; white-space: nowrap;">${point.code || 'N/A'}</td>
            <td style="padding: 8px; font-size: 11px; white-space: nowrap;">${point.type || 'N/A'}</td>
            <td style="padding: 8px; font-size: 11px; white-space: nowrap;">${point.sleutelbox || 'N/A'}</td>
            <td style="padding: 8px; font-size: 11px;">${point.naam || 'N/A'}</td>
            <td style="padding: 8px; font-size: 11px;">${point.geografisch || 'N/A'}</td>
          </tr>
        `;
      });
    } catch (e) {
      console.error('Error parsing locked points:', e);
      vergrendelpuntenRows = '<tr><td colspan="6" style="padding: 8px; text-align: center; color: #999;">Geen vergrendelpunten beschikbaar</td></tr>';
    }

    const subject = `${statusIcon} Aanvraag ${statusText} - ${requestId}`;

    // Plain text fallback
    const plainTextBody = `
Beste ${requesterName},

Uw aanvraag voor een vergrendelgroep is ${statusText.toLowerCase()}.

Aanvraag Nummer: ${requestId}
Ingediend op: ${timestamp}
Beslissing: ${statusText}
Beslissing datum: ${decisionDate}
Autorisator: ${authorizerEmail}
${authorizerRemark ? `Opmerking: ${authorizerRemark}` : ''}

Installatie: ${formData.installatie || 'N/A'}
Naam: ${formData.naam || 'N/A'}

Met vriendelijke groet,
Reliability CMMS
    `;

    // HTML email body
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Arial, sans-serif; background-color: #f5f5f5;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f5f5f5; padding: 20px;">
    <tr>
      <td align="center">
        <!-- Main container -->
        <table width="750" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">

          <!-- Header -->
          <tr>
            <td style="background: linear-gradient(135deg, ${statusColor} 0%, ${statusColor}dd 100%); padding: 30px; text-align: center;">
              <h1 style="margin: 0; color: #ffffff; font-size: 24px; font-weight: 700;">
                ${statusIcon} Aanvraag ${statusText}
              </h1>
              <p style="margin: 10px 0 0 0; color: #ffffff; font-size: 14px;">
                Reliability CMMS - Aanvragen Rimses
              </p>
            </td>
          </tr>

          <!-- Content -->
          <tr>
            <td style="padding: 30px;">

              <!-- Introduction -->
              <p style="margin: 0 0 20px 0; color: #333; font-size: 15px; line-height: 1.6;">
                Beste ${requesterName},
              </p>
              <p style="margin: 0 0 25px 0; color: #666; font-size: 14px; line-height: 1.6;">
                Uw aanvraag voor een vergrendelgroep is ${isApproved ? '<strong style="color: #4caf50;">goedgekeurd</strong>' : '<strong style="color: #f44336;">afgekeurd</strong>'}.
              </p>

              <!-- Decision Status -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: ${statusBgColor}; border-left: 4px solid ${statusColor}; border-radius: 4px; margin-bottom: 25px;">
                <tr>
                  <td style="padding: 15px;">
                    <h3 style="margin: 0 0 12px 0; color: ${statusColor}; font-size: 16px; font-weight: 600;">
                      Beslissing: ${statusText}
                    </h3>
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600; width: 140px;">Beslissing datum:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${decisionDate}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Autorisator:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${authorizerEmail}</td>
                      </tr>
                      ${authorizerRemark ? `
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600; vertical-align: top;">Opmerking:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${authorizerRemark}</td>
                      </tr>
                      ` : ''}
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Aanvraag Details -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #e8f5e9; border-left: 4px solid #4caf50; border-radius: 4px; margin-bottom: 25px;">
                <tr>
                  <td style="padding: 15px;">
                    <h3 style="margin: 0 0 12px 0; color: #2e7d32; font-size: 16px; font-weight: 600;">
                      Aanvraag Details
                    </h3>
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600; width: 140px;">Aanvraag Nummer:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px; font-family: monospace;">${requestId}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Ingediend op:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${timestamp}</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Formulier Gegevens -->
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #e3f2fd; border-left: 4px solid #2196f3; border-radius: 4px; margin-bottom: 25px;">
                <tr>
                  <td style="padding: 15px;">
                    <h3 style="margin: 0 0 12px 0; color: #1565c0; font-size: 16px; font-weight: 600;">
                      Formulier Gegevens
                    </h3>
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600; width: 140px;">Installatie:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.installatie || 'N/A'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Naam:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.naam || 'N/A'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Interne Opmerking:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.interneOpmerking || 'Geen'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Externe Opmerking:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.externeOpmerking || 'Geen'}</td>
                      </tr>
                      <tr>
                        <td style="padding: 4px 0; color: #666; font-size: 13px; font-weight: 600;">Opmerkingen:</td>
                        <td style="padding: 4px 0; color: #333; font-size: 13px;">${formData.opmerkingen || 'Geen'}</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Vergrendelpunten Table -->
              <div style="margin-bottom: 25px;">
                <h3 style="margin: 0 0 12px 0; color: #333; font-size: 16px; font-weight: 600;">
                  Geselecteerde Vergrendelpunten
                </h3>
                <table width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;">
                  <thead>
                    <tr style="background-color: #f5f5f5;">
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">#</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Code</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Type</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Sleutelbox</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Naam</th>
                      <th style="padding: 10px 8px; text-align: left; font-size: 11px; font-weight: 600; color: #666; border-bottom: 2px solid #e0e0e0; white-space: nowrap;">Locatie</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${vergrendelpuntenRows}
                  </tbody>
                </table>
              </div>

              <!-- Closing -->
              <p style="margin: 0; color: #666; font-size: 13px; line-height: 1.6;">
                Met vriendelijke groet,<br>
                <strong>Reliability CMMS</strong><br>
                Aanvragen Rimses Systeem
              </p>

            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background-color: #f5f5f5; padding: 15px; text-align: center; border-top: 1px solid #e0e0e0;">
              <p style="margin: 0; color: #999; font-size: 11px;">
                Dit is een automatisch gegenereerde email. Gelieve niet te antwoorden op dit bericht.
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>
    `;

    // Send email via GmailApp with alias
    GmailApp.sendEmail(
      requesterEmail,
      subject,
      plainTextBody,
      {
        from: 'reliabilitycmms@gmail.com',
        name: 'Reliability CMMS',
        htmlBody: htmlBody
      }
    );

    console.log(`✓ Decision notification email (${statusText}) sent to requester:`, requesterEmail);

    return {
      success: true,
      message: `Decision email sent to requester: ${requesterEmail}`
    };

  } catch (error) {
    console.error('Error sending decision notification:', error);
    return {
      success: false,
      message: 'Failed to send decision email: ' + error.message,
      error: error.toString()
    };
  }
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
 * Test aanvraag notification email with dummy data
 */
function testAanvraagEmail() {
  try {
    const userEmail = Session.getActiveUser().getEmail();

    // Create dummy aanvraag data
    const dummyAanvraagData = {
      "Request ID": "TEST-" + new Date().getTime(),
      "Approval Token": "dummy-token-123",
      "Requester Email": userEmail,
      "Approver Email": "rob.oversteyns@gmail.com",
      "Status": "Pending",
      "Timestamp": new Date().toISOString()
    };

    const dummyFormData = {
      installatie: "G4",
      naam: "Test Vergrendelgroep",
      interneOpmerking: "Dit is een test interne opmerking",
      externeOpmerking: "Dit is een test externe opmerking",
      opmerkingen: "Test algemene opmerkingen",
      autorisator: "rob.oversteyns@gmail.com"
    };

    const dummySelectedRows = [
      { data: ["VP001", "Type A", "Box 1", "Test Punt 1", "Locatie A"] },
      { data: ["VP002", "Type B", "Box 2", "Test Punt 2", "Locatie B"] },
      { data: ["VP003", "Type C", "Box 3", "Test Punt 3", "Locatie C"] }
    ];

    const dummyUserName = "Test User (Debug)";

    console.log('Testing aanvraag email with dummy data...');
    console.log('Sending to:', dummyFormData.autorisator);

    // Call the actual sendAuthorizerNotification function
    const result = sendAuthorizerNotification(
      dummyAanvraagData,
      dummyFormData,
      dummySelectedRows,
      dummyUserName
    );

    return {
      success: result.success,
      message: result.message,
      testData: {
        requestId: dummyAanvraagData["Request ID"],
        sentTo: dummyFormData.autorisator,
        timestamp: new Date().toLocaleString('nl-NL')
      },
      error: result.error || null
    };

  } catch (error) {
    console.error('Error in testAanvraagEmail:', error);
    return {
      success: false,
      message: 'Fout bij testen aanvraag email: ' + error.message,
      error: error.toString()
    };
  }
}

/**
 * Generates a random password of specified length
 * @param {number} length - Password length (default: 8)
 * @returns {string} Random password
 */
function generateRandomPassword(length = 8) {
  const charset = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%';
  let password = '';
  for (let i = 0; i < length; i++) {
    const randomIndex = Math.floor(Math.random() * charset.length);
    password += charset[randomIndex];
  }
  return password;
}

/**
 * Updates user password in Firebase users database
 * @param {string} firebaseKey - Firebase key of user record
 * @param {string} newPassword - New password to set
 * @returns {Object} { success: boolean, message: string }
 */
function updateUserPassword(firebaseKey, newPassword) {
  try {
    const url = `${FIREBASE_USERS_URL}/fire-data/${firebaseKey}.json?auth=${FIREBASE_USERS_SECRET}`;

    const payload = {
      PASSWORD: newPassword
    };

    const options = {
      method: 'patch',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      console.log('Password updated successfully for key:', firebaseKey);
      return {
        success: true,
        message: 'Wachtwoord succesvol bijgewerkt'
      };
    } else {
      console.error('Failed to update password:', responseCode, response.getContentText());
      return {
        success: false,
        message: 'Fout bij bijwerken wachtwoord'
      };
    }
  } catch (error) {
    console.error('Error in updateUserPassword:', error);
    return {
      success: false,
      message: 'Fout: ' + error.message
    };
  }
}

/**
 * Sends welcome email with generated password to new user
 * @param {string} userEmail - User email address
 * @param {string} userName - User display name
 * @param {string} generatedPassword - The generated password
 * @returns {Object} { success: boolean, message: string }
 */
function sendWelcomePasswordEmail(userEmail, userName, generatedPassword) {
  try {
    const subject = 'Welkom bij Aanvragen Rimses - Uw Wachtwoord';

    // HTML email body
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; background-color: #ffffff; }
    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; text-align: center; color: white; }
    .header h1 { margin: 0; font-size: 28px; font-weight: 700; }
    .content { padding: 30px; }
    .password-box { background-color: #f8f9fa; border-left: 4px solid #667eea; padding: 20px; margin: 20px 0; border-radius: 4px; }
    .password { font-size: 24px; font-weight: 700; color: #667eea; letter-spacing: 2px; font-family: 'Courier New', monospace; }
    .info-box { background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; border-radius: 4px; }
    .footer { background-color: #f8f9fa; padding: 20px; text-align: center; font-size: 12px; color: #6c757d; }
    .button { display: inline-block; padding: 12px 30px; background-color: #667eea; color: white; text-decoration: none; border-radius: 6px; font-weight: 600; margin: 20px 0; }
    ul { padding-left: 20px; }
    li { margin: 10px 0; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Welkom bij Aanvragen Rimses</h1>
    </div>

    <div class="content">
      <h2>Hallo ${userName || userEmail},</h2>

      <p>Welkom bij de Aanvragen Rimses applicatie! Dit is uw eerste login en we hebben automatisch een wachtwoord voor u aangemaakt.</p>

      <div class="password-box">
        <p style="margin: 0 0 10px 0; font-weight: 600;">Uw tijdelijke wachtwoord:</p>
        <div class="password">${generatedPassword}</div>
      </div>

      <div class="info-box">
        <strong>BELANGRIJK:</strong>
        <ul style="margin: 10px 0;">
          <li>Bewaar dit wachtwoord op een veilige plek</li>
          <li>Log in met dit wachtwoord om toegang te krijgen tot de applicatie</li>
          <li>U kunt dit wachtwoord blijven gebruiken voor toekomstige logins</li>
        </ul>
      </div>

      <h3>Login Gegevens:</h3>
      <p>
        <strong>Gebruikersnaam:</strong> ${userEmail}<br>
        <strong>Wachtwoord:</strong> <code>${generatedPassword}</code>
      </p>

      <h3>Hoe in te loggen:</h3>
      <ol>
        <li>Ga naar de Aanvragen Rimses applicatie</li>
        <li>Voer uw gebruikersnaam (email) in</li>
        <li>Voer het bovenstaande wachtwoord in</li>
        <li>Klik op "Inloggen"</li>
      </ol>

      <p style="margin-top: 30px;">Heeft u vragen? Neem contact op met de beheerder.</p>
    </div>

    <div class="footer">
      <p>Dit is een automatisch gegenereerde email van Aanvragen Rimses.</p>
      <p style="margin: 5px 0;">Reageer niet op deze email.</p>
      <p style="margin: 10px 0; color: #999;">Generated with Reliability CMMS</p>
    </div>
  </div>
</body>
</html>
    `;

    // Plain text fallback
    const plainTextBody = `
Welkom bij Aanvragen Rimses

Hallo ${userName || userEmail},

Dit is uw eerste login en we hebben automatisch een wachtwoord voor u aangemaakt.

UW TIJDELIJKE WACHTWOORD: ${generatedPassword}

Login Gegevens:
- Gebruikersnaam: ${userEmail}
- Wachtwoord: ${generatedPassword}

BELANGRIJK:
- Bewaar dit wachtwoord op een veilige plek
- Log in met dit wachtwoord om toegang te krijgen tot de applicatie
- U kunt dit wachtwoord blijven gebruiken voor toekomstige logins

Hoe in te loggen:
1. Ga naar de Aanvragen Rimses applicatie
2. Voer uw gebruikersnaam (email) in
3. Voer het bovenstaande wachtwoord in
4. Klik op "Inloggen"

Heeft u vragen? Neem contact op met de beheerder.

---
Dit is een automatisch gegenereerde email van Aanvragen Rimses.
Reageer niet op deze email.
    `;

    // Send email via GmailApp
    GmailApp.sendEmail(
      userEmail,
      subject,
      plainTextBody,
      {
        from: 'reliabilitycmms@gmail.com',
        name: 'Reliability CMMS - Aanvragen Rimses',
        htmlBody: htmlBody
      }
    );

    console.log('✓ Welcome password email sent to:', userEmail);

    return {
      success: true,
      message: 'Welkom email succesvol verzonden naar ' + userEmail
    };

  } catch (error) {
    console.error('Error sending welcome password email:', error);
    return {
      success: false,
      message: 'Fout bij verzenden email: ' + error.message
    };
  }
}

/**
 * Validates user credentials against Firebase users database
 * NEW: Automatically generates and emails password on first login
 *
 * Flow:
 * 1. User tries to login (password input can be empty)
 * 2. Check if user exists in database
 * 3. Check if database password is empty
 *    - If DB password is empty → Generate password, save to DB, email user, show modal
 *    - If DB password is set AND input is empty → Show "please enter password" error
 *    - If DB password is set AND input is wrong → Show "incorrect password" error
 *    - If DB password is set AND input matches → Login success
 *
 * @param {string} username - Username (email address)
 * @param {string} password - Password to validate (can be empty)
 * @returns {Object} { success: boolean, emptyPassword: boolean, firstLogin: boolean, message: string }
 */
function validateUser(username, password) {
  try {
    // Fetch users from Firebase users database (fire-data path)
    const url = `${FIREBASE_USERS_URL}/fire-data.json?auth=${FIREBASE_USERS_SECRET}`;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      console.error('Firebase users error:', responseCode, response.getContentText());
      return {
        success: false,
        emptyPassword: false,
        firstLogin: false,
        message: 'Database connectie fout'
      };
    }

    const data = JSON.parse(response.getContentText());

    if (!data || typeof data !== 'object') {
      console.error('Invalid Firebase users data structure');
      return {
        success: false,
        emptyPassword: false,
        firstLogin: false,
        message: 'Database data fout'
      };
    }

    // Search for user in Firebase data
    // Try multiple strategies to find the user:
    // 1. Direct lookup with encoded email (new users)
    // 2. Direct lookup with plain email (some legacy users)
    // 3. Search through all users matching USER_NAME field (legacy users)

    let user = null;
    let userKey = null;

    // Strategy 1: Try Firebase-safe encoded email as key
    const encodedUsername = encodeEmailForFirebase(username);
    if (data[encodedUsername]) {
      user = data[encodedUsername];
      userKey = encodedUsername;
      console.log('✓ Found user via encoded email key');
    }

    // Strategy 2: Try plain email as key (some legacy data)
    if (!user && data[username]) {
      user = data[username];
      userKey = username;
      console.log('✓ Found user via plain email key');
    }

    // Strategy 3: Search by USER_NAME field (legacy users where key != email)
    if (!user) {
      for (const key in data) {
        const currentUser = data[key];
        // Check if currentUser is valid object and has USER_NAME field
        if (currentUser && typeof currentUser === 'object' && currentUser.USER_NAME === username) {
          user = currentUser;
          userKey = key;
          console.log('✓ Found user via USER_NAME field search, key:', key);
          break;
        }
      }
    }

    if (user) {
        // Get database password (can be empty)
        const dbPassword = user.PASSWORD || '';
        const dbPasswordIsEmpty = !dbPassword || dbPassword.trim() === '';

        // SCENARIO 1: Database password is empty - FIRST LOGIN
        // → Generate password, save to DB, email user, show modal
        if (dbPasswordIsEmpty) {
          console.log('🔐 First login detected for user:', username);

          // Generate random password
          const generatedPassword = generateRandomPassword(8);
          console.log('Generated password for', username);

          // Update password in Firebase (use the actual key we found)
          const updateResult = updateUserPassword(userKey, generatedPassword);

          if (!updateResult.success) {
            console.error('Failed to update password in Firebase');
            return {
              success: false,
              emptyPassword: true,
              firstLogin: true,
              message: 'Fout bij aanmaken wachtwoord. Probeer opnieuw of neem contact op met de beheerder.'
            };
          }

          console.log('✓ Password updated in Firebase');

          // Send welcome email with password
          const emailResult = sendWelcomePasswordEmail(username, user.DISPLAY_NAME || username, generatedPassword);

          if (!emailResult.success) {
            console.warn('⚠ Failed to send welcome email, but password was saved:', emailResult.message);
            // Continue anyway - password is saved
          } else {
            console.log('✓ Welcome email sent successfully');
          }

          // Return first login response
          return {
            success: false,
            emptyPassword: true,
            firstLogin: true,
            emailSent: emailResult.success,
            message: 'Dit is uw eerste login. Een wachtwoord is aangemaakt en gemaild naar ' + username
          };
        }

        // SCENARIO 2: Database password exists, but user didn't enter anything
        // → Show error asking to enter password
        const inputPassword = password || '';
        if (inputPassword.trim() === '') {
          return {
            success: false,
            emptyPassword: false,
            firstLogin: false,
            message: 'Voer uw wachtwoord in'
          };
        }

        // SCENARIO 3: Database password exists, user entered something
        // → Validate the password
        if (dbPassword === inputPassword) {
          // Get user role (A = user, B = authorizer, C = admin)
          const userRole = user.ROLE || 'A'; // Default to 'A' if not set

          return {
            success: true,
            emptyPassword: false,
            firstLogin: false,
            role: userRole,
            message: 'Login succesvol'
          };
        } else {
          return {
            success: false,
            emptyPassword: false,
            firstLogin: false,
            message: 'Onjuist wachtwoord'
          };
        }
    }

    // User not found with any strategy
    console.error('❌ User not found. Searched for:', username);
    console.error('Available keys:', Object.keys(data).join(', '));
    return {
      success: false,
      emptyPassword: false,
      firstLogin: false,
      message: 'Gebruiker niet gevonden'
    };

  } catch (error) {
    console.error('Error in validateUser:', error);
    return {
      success: false,
      emptyPassword: false,
      firstLogin: false,
      message: 'Fout bij validatie: ' + error.message
    };
  }
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
    // Fetch users from Firebase users database (fire-data path)
    const url = `${FIREBASE_USERS_URL}/fire-data.json?auth=${FIREBASE_USERS_SECRET}`;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      console.error('Firebase users error:', responseCode, response.getContentText());
      return [];
    }

    const data = JSON.parse(response.getContentText());

    if (!data) {
      console.log('No users data found in Firebase');
      return [];
    }

    // Convert Firebase object to array (same as getUsers function)
    const users = Object.keys(data).map(key => ({
      firebaseKey: key,
      ...data[key]
    }));

    console.log(`Total users in Firebase: ${users.length}`);

    // Extract USER_NAME from users where ROLE is 'B' (Autorisator) or 'C' (Admin)
    const autorisators = [];

    users.forEach((user, index) => {
      console.log(`User ${index}:`, {
        USER_NAME: user?.USER_NAME,
        ROLE: user?.ROLE,
        ACTIVE: user?.ACTIVE
      });

      if (user && user.USER_NAME && user.ACTIVE === 1 && (user.ROLE === 'B' || user.ROLE === 'C')) {
        autorisators.push(user.USER_NAME);
      }
    });

    // Sort alphabetically and remove duplicates
    const uniqueAutorisators = [...new Set(autorisators)].sort();

    console.log('Autorisators fetched from Firebase:', uniqueAutorisators.length);
    console.log('Autorisators list:', uniqueAutorisators);
    return uniqueAutorisators;
  } catch (error) {
    console.error('Error fetching autorisators from Firebase:', error);
    return [];
  }
}

/**
 * Debug function to test Firebase users connection and return detailed info
 */
function debugLoadUsersFirebase() {
  const debugInfo = {
    timestamp: new Date().toLocaleString('nl-NL'),
    config: {
      url: FIREBASE_USERS_URL,
      hasSecret: FIREBASE_USERS_SECRET ? 'Yes' : 'No',
      secretLength: FIREBASE_USERS_SECRET ? FIREBASE_USERS_SECRET.length : 0
    },
    attempts: []
  };

  try {
    // Attempt 1: Check root of database
    const rootUrl = `${FIREBASE_USERS_URL}/.json?auth=${FIREBASE_USERS_SECRET}`;
    const rootAttempt = {
      path: '/ (root)',
      url: rootUrl.replace(FIREBASE_USERS_SECRET, '***SECRET***'),
      method: 'GET'
    };

    const rootResponse = UrlFetchApp.fetch(rootUrl, {
      method: 'get',
      muteHttpExceptions: true
    });

    rootAttempt.responseCode = rootResponse.getResponseCode();
    const rootContentText = rootResponse.getContentText();
    rootAttempt.contentLength = rootContentText.length;
    rootAttempt.rawContent = rootContentText.substring(0, 500); // First 500 chars

    if (rootAttempt.responseCode === 200) {
      const rootData = JSON.parse(rootContentText);
      rootAttempt.dataType = typeof rootData;
      rootAttempt.isNull = rootData === null;

      if (rootData && typeof rootData === 'object') {
        rootAttempt.keys = Object.keys(rootData);
        rootAttempt.keyCount = rootAttempt.keys.length;
      }
    }
    debugInfo.attempts.push(rootAttempt);

    // Attempt 2: Check /users path
    const usersUrl = `${FIREBASE_USERS_URL}/users.json?auth=${FIREBASE_USERS_SECRET}`;
    const usersAttempt = {
      path: '/users',
      url: usersUrl.replace(FIREBASE_USERS_SECRET, '***SECRET***'),
      method: 'GET'
    };

    const usersResponse = UrlFetchApp.fetch(usersUrl, {
      method: 'get',
      muteHttpExceptions: true
    });

    usersAttempt.responseCode = usersResponse.getResponseCode();
    const usersContentText = usersResponse.getContentText();
    usersAttempt.contentLength = usersContentText.length;
    usersAttempt.rawContent = usersContentText.substring(0, 500);

    if (usersAttempt.responseCode === 200) {
      const usersData = JSON.parse(usersContentText);
      usersAttempt.dataType = typeof usersData;
      usersAttempt.isNull = usersData === null;

      if (usersData && typeof usersData === 'object') {
        usersAttempt.keys = Object.keys(usersData);
        usersAttempt.keyCount = usersAttempt.keys.length;

        // Extract users with USER_NAME
        const users = [];
        for (const key in usersData) {
          if (usersData.hasOwnProperty(key)) {
            const user = usersData[key];
            users.push({
              key: key,
              hasUserName: !!user.USER_NAME,
              userName: user.USER_NAME || null,
              allFields: Object.keys(user)
            });
          }
        }

        usersAttempt.users = users.slice(0, 10);
        usersAttempt.totalUsers = users.length;
        usersAttempt.usersWithUserName = users.filter(u => u.hasUserName).length;

        const userNames = users
          .filter(u => u.hasUserName)
          .map(u => u.userName);

        usersAttempt.userNames = [...new Set(userNames)].sort();
      }
    }
    debugInfo.attempts.push(usersAttempt);

    // Summary
    debugInfo.success = true;
    debugInfo.message = 'Debug complete - check attempts array for details';

  } catch (error) {
    debugInfo.success = false;
    debugInfo.error = error.toString();
    debugInfo.message = 'Exception occurred: ' + error.message;
  }

  return debugInfo;
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

// ===== CRITICALITY DATA FUNCTIONS =====

/**
 * Fetch all criticality data from Firebase
 * @returns {Array} Array of criticality objects
 */
function getCriticalityData() {
  try {
    console.log('🔍 [DEBUG] getCriticalityData: Starting fetch...');
    console.log('🔍 [DEBUG] Firebase URL:', FIREBASE_CRITICALITY_URL);

    // Try root path to see all available data
    const url = `${FIREBASE_CRITICALITY_URL}/.json?auth=${FIREBASE_CRITICALITY_SECRET}`;
    console.log('🔍 [DEBUG] Full URL (without secret):', url.replace(FIREBASE_CRITICALITY_SECRET, '***SECRET***'));
    console.log('🔍 [DEBUG] Trying root path to discover data structure...');

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    console.log('🔍 [DEBUG] Response code:', responseCode);

    if (responseCode !== 200) {
      console.error('❌ Firebase Criticality error:', responseCode, response.getContentText());
      return [];
    }

    const responseText = response.getContentText();
    console.log('🔍 [DEBUG] Response length:', responseText.length, 'characters');
    console.log('🔍 [DEBUG] Response preview (first 200 chars):', responseText.substring(0, 200));

    const data = JSON.parse(responseText);
    console.log('🔍 [DEBUG] Parsed data type:', typeof data);
    console.log('🔍 [DEBUG] Is data null?:', data === null);
    console.log('🔍 [DEBUG] Is data object?:', typeof data === 'object');

    if (!data) {
      console.log('⚠️ No criticality data found in Firebase (data is null/undefined)');
      return [];
    }

    // Log the root keys to see the structure
    const rootKeys = Object.keys(data);
    console.log('🔍 [DEBUG] Root keys in Firebase:', rootKeys);

    // Data is under 'SpareParts' key
    let criticalityData = data;

    // Check if data has 'SpareParts' key
    if (data.SpareParts) {
      console.log('✓ Found SpareParts key, using that data');
      criticalityData = data.SpareParts;
    } else if (rootKeys.length > 0) {
      // Fallback: use first key if it contains objects
      console.log('⚠️ No SpareParts key, checking first key:', rootKeys[0]);
      if (typeof data[rootKeys[0]] === 'object' && data[rootKeys[0]] !== null) {
        const firstKeySubKeys = Object.keys(data[rootKeys[0]]);
        console.log('🔍 [DEBUG] First key has', firstKeySubKeys.length, 'sub-keys');
        if (firstKeySubKeys.length > 0) {
          criticalityData = data[rootKeys[0]];
        }
      }
    }

    const keys = Object.keys(criticalityData);
    console.log('🔍 [DEBUG] Number of items in criticalityData:', keys.length);
    console.log('🔍 [DEBUG] First 5 firebase keys:', keys.slice(0, 5));

    // Convert Firebase object to array with firebaseKey
    const dataArray = keys.map(key => ({
      firebaseKey: key,
      ...criticalityData[key]
    }));

    console.log('🔍 [DEBUG] DataArray length:', dataArray.length);
    if (dataArray.length > 0) {
      console.log('🔍 [DEBUG] First item keys:', Object.keys(dataArray[0]));
      console.log('🔍 [DEBUG] First item sample:', JSON.stringify(dataArray[0]).substring(0, 200));
    }

    console.log(`✓ Loaded ${dataArray.length} criticality items from Firebase`);
    return dataArray;

  } catch (error) {
    console.error('❌ Error fetching criticality data:', error);
    console.error('❌ Error stack:', error.stack);
    return [];
  }
}

/**
 * Save a new criticality part to Firebase
 */
function saveCriticalityPart(data) {
  try {
    console.log('💾 [saveCriticalityPart] Starting save operation...');
    console.log('📦 Data to save:', JSON.stringify(data));

    // Generate unique ID using timestamp
    const timestamp = new Date().getTime();
    const id = 'CRIT-' + timestamp;

    // Prepare data for Firebase
    const firebaseData = {
      ID: id,
      Datum: data.Datum || new Date().toISOString(),
      Naam: data.Naam || '',
      RemaxNummer: data.RemaxNummer || '',
      Kosten: data.Kosten || '',
      Levertijd: data.Levertijd || '',
      FalenFreq: data.FalenFreq || '',
      FalenSeverity: data.FalenSeverity || '',
      MachineCrit: data.MachineCrit || 'C',
      AankoopFactor: data.AankoopFactor || 0,
      Risico: data.Risico || '',
      FunctieCrit: data.FunctieCrit || 0,
      Categorie: data.Categorie || '',
      User: data.User || '',
      Status: data.Status || 'In behandeling'
    };

    console.log('🔧 Prepared Firebase data:', JSON.stringify(firebaseData));

    // Save to Firebase under SpareParts path
    const url = FIREBASE_CRITICALITY_URL + '/SpareParts.json?auth=' + FIREBASE_CRITICALITY_SECRET;

    const payload = JSON.stringify(firebaseData);
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: payload,
      muteHttpExceptions: true
    };

    console.log('📡 Sending POST request to Firebase...');
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log('📨 Response code:', responseCode);
    console.log('📨 Response text:', responseText);

    if (responseCode === 200) {
      const result = JSON.parse(responseText);
      console.log('✅ Criticality part saved successfully! Firebase key:', result.name);
      return {
        success: true,
        firebaseKey: result.name,
        id: id,
        message: 'Onderdeel opgeslagen als ' + firebaseData.Categorie
      };
    } else {
      console.error('❌ Firebase error:', responseCode, responseText);
      throw new Error('Firebase returned status ' + responseCode + ': ' + responseText);
    }

  } catch (error) {
    console.error('❌ Error in saveCriticalityPart:', error);
    console.error('❌ Error stack:', error.stack);
    throw new Error('Fout bij opslaan: ' + error.message);
  }
}

/**
 * Log user login to Firebase
 * @param {string} username - Username that logged in
 */
function logUserLogin(username) {
  try {
    console.log('🔐 [logUserLogin] Logging login for user:', username);

    // Generate timestamp and ID
    const now = new Date();
    const timestamp = now.getTime();

    // Format LoginTime as "YYYY-MM-DD HH:MM:SS"
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    const loginTime = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;

    // Prepare login data
    const loginData = {
      ID: timestamp,
      LoginTime: loginTime,
      Username: username
    };

    console.log('📦 Login data:', JSON.stringify(loginData));

    // Save to Firebase under /logins path
    const url = FIREBASE_LOGIN_LOG_URL + 'logins.json?auth=' + FIREBASE_LOGIN_LOG_SECRET;

    const payload = JSON.stringify(loginData);
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: payload,
      muteHttpExceptions: true
    };

    console.log('📡 Sending login log to Firebase...');
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log('📨 Response code:', responseCode);
    console.log('📨 Response text:', responseText);

    if (responseCode === 200) {
      const result = JSON.parse(responseText);
      console.log('✅ Login logged successfully! Firebase key:', result.name);
      return {
        success: true,
        firebaseKey: result.name,
        loginTime: loginTime
      };
    } else {
      console.error('❌ Firebase login log error:', responseCode, responseText);
      // Don't throw - we don't want to block login if logging fails
      return {
        success: false,
        error: 'Firebase returned status ' + responseCode
      };
    }

  } catch (error) {
    console.error('❌ Error in logUserLogin:', error);
    console.error('❌ Error stack:', error.stack);
    // Don't throw - we don't want to block login if logging fails
    return {
      success: false,
      error: error.message
    };
  }
}