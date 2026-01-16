// Google Apps Script for Herbal Inventory Status Tracking
// Updated version with query improvements and CORS support

/**
 * Handle HTTP POST requests from the web app
 */
function doPost(e) {
  try {
    // Better error handling for missing data
    if (!e) {
      return createResponse(false, 'No event object received');
    }

    if (!e.postData) {
      return createResponse(false, 'No postData in request');
    }

    if (!e.postData.contents) {
      return createResponse(false, 'No contents in postData');
    }

    const data = JSON.parse(e.postData.contents);
    const { action, herbs, statusType, queryType } = data;

    // Log for debugging
    Logger.log('Received action: ' + action);
    Logger.log('Herbs: ' + JSON.stringify(herbs));
    Logger.log('Status type: ' + statusType);
    Logger.log('Query type: ' + queryType);

    if (!action) {
      return createResponse(false, 'Action required');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');

    if (!sheet) {
      return createResponse(false, 'Sheet named "Inventory" not found');
    }

    let result;

    if (action === 'mark_status') {
      result = markHerbStatus(sheet, herbs, statusType);
    } else if (action === 'clear_status') {
      result = clearHerbStatus(sheet, herbs, statusType);
    } else if (action === 'query') {
      result = queryHerbStatus(sheet, herbs, queryType);
    } else {
      return createResponse(false, 'Unknown action: ' + action);
    }

    return createResponse(true, result.message, result.data);

  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return createResponse(false, 'Server error: ' + error.message);
  }
}

/**
 * Handle HTTP GET requests (for testing)
 */
function doGet(e) {
  return ContentService.createTextOutput(
    'Herbal Inventory Status Tracker API is running.\n' +
    'Send POST requests with JSON body containing:\n' +
    '- action: "mark_status", "clear_status", or "query"\n' +
    '- herbs: array of herb names\n' +
    '- statusType: "Low", "Out", "In Tincture", or "Clinic Backstock"\n' +
    '- queryType: (optional) specific status to check'
  );
}

/**
 * Mark herbs with a status (check the box)
 */
function markHerbStatus(sheet, herbs, statusType) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find column indices
  const herbNameCol = headers.indexOf('Herb Name');
  const statusCol = headers.indexOf(statusType);
  const dateStartedCol = headers.indexOf('Date Started');

  if (herbNameCol === -1) {
    throw new Error('Herb Name column not found');
  }

  if (statusCol === -1) {
    throw new Error(statusType + ' column not found. Available columns: ' + headers.join(', '));
  }

  const results = [];
  const herbsArray = Array.isArray(herbs) ? herbs : [herbs];

  for (let herbName of herbsArray) {
    herbName = herbName.trim();

    // Find herb row (case-insensitive)
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][herbNameCol] &&
          data[i][herbNameCol].toString().toLowerCase() === herbName.toLowerCase()) {
        rowIndex = i;
        break;
      }
    }

    // If herb not found, create new row
    if (rowIndex === -1) {
      rowIndex = data.length;
      sheet.getRange(rowIndex + 1, herbNameCol + 1).setValue(herbName);

      // Initialize all checkbox columns to false
      const checkboxColumns = ['Low', 'Out', 'In Tincture', 'Clinic Backstock'];
      for (let col of checkboxColumns) {
        const colIndex = headers.indexOf(col);
        if (colIndex !== -1 && col !== statusType) {
          sheet.getRange(rowIndex + 1, colIndex + 1).setValue(false);
        }
      }

      results.push('Created new herb: ' + herbName);
    }

    // Mark the status (checkbox = TRUE)
    sheet.getRange(rowIndex + 1, statusCol + 1).setValue(true);

    // If marking "In Tincture", set Date Started to today
    if (statusType === 'In Tincture' && dateStartedCol !== -1) {
      const today = new Date();
      sheet.getRange(rowIndex + 1, dateStartedCol + 1).setValue(today);
    }

    results.push('Marked ' + herbName + ' as ' + statusType);
  }

  return {
    message: results.join('; '),
    data: { herbs: herbsArray, status: statusType, action: 'marked' }
  };
}

/**
 * Clear a status from herbs (uncheck the box)
 */
function clearHerbStatus(sheet, herbs, statusType) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const herbNameCol = headers.indexOf('Herb Name');
  const statusCol = headers.indexOf(statusType);
  const dateStartedCol = headers.indexOf('Date Started');

  if (herbNameCol === -1 || statusCol === -1) {
    throw new Error('Required columns not found');
  }

  const results = [];
  const herbsArray = Array.isArray(herbs) ? herbs : [herbs];

  for (let herbName of herbsArray) {
    herbName = herbName.trim();

    // Find herb row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][herbNameCol] &&
          data[i][herbNameCol].toString().toLowerCase() === herbName.toLowerCase()) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      results.push(herbName + ' not found');
      continue;
    }

    // Clear the status (checkbox = FALSE)
    sheet.getRange(rowIndex + 1, statusCol + 1).setValue(false);

    // If clearing "In Tincture", clear Date Started too
    if (statusType === 'In Tincture' && dateStartedCol !== -1) {
      sheet.getRange(rowIndex + 1, dateStartedCol + 1).setValue('');
    }

    results.push('Cleared ' + statusType + ' from ' + herbName);
  }

  return {
    message: results.join('; '),
    data: { herbs: herbsArray, status: statusType, action: 'cleared' }
  };
}

/**
 * Query the status of herbs
 * NEW: Now supports queryType parameter to check specific status
 */
function queryHerbStatus(sheet, herbs, queryType) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const herbNameCol = headers.indexOf('Herb Name');

  if (herbNameCol === -1) {
    throw new Error('Herb Name column not found');
  }

  const results = [];
  const herbsArray = Array.isArray(herbs) ? herbs : [herbs];
  const herbStatuses = {}; // For structured data response

  for (let herbName of herbsArray) {
    herbName = herbName.trim();

    // Find herb row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][herbNameCol] &&
          data[i][herbNameCol].toString().toLowerCase() === herbName.toLowerCase()) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      results.push(herbName + ': Not found in inventory');
      herbStatuses[herbName] = { found: false, hasStatus: false };
      continue;
    }

    // If queryType is specified, only check that specific status
    if (queryType) {
      const colIndex = headers.indexOf(queryType);
      if (colIndex === -1) {
        throw new Error(queryType + ' column not found');
      }

      const hasStatus = data[rowIndex][colIndex] === true;
      herbStatuses[herbName] = {
        found: true,
        hasStatus: hasStatus,
        statusType: queryType
      };

      results.push(herbName + ': ' + (hasStatus ? 'YES' : 'NO'));

    } else {
      // Check all status columns (original behavior)
      const statuses = [];
      const statusColumns = ['Low', 'Out', 'In Tincture', 'Clinic Backstock'];

      for (let statusCol of statusColumns) {
        const colIndex = headers.indexOf(statusCol);
        if (colIndex !== -1 && data[rowIndex][colIndex] === true) {
          statuses.push(statusCol);
        }
      }

      // Check Date Ready if in tincture
      const dateReadyCol = headers.indexOf('Date Ready');
      let dateReady = '';
      if (dateReadyCol !== -1 && data[rowIndex][dateReadyCol]) {
        const date = new Date(data[rowIndex][dateReadyCol]);
        if (!isNaN(date.getTime())) {
          dateReady = ' (ready ' + formatDate(date) + ')';
        }
      }

      herbStatuses[herbName] = {
        found: true,
        statuses: statuses,
        dateReady: dateReady
      };

      if (statuses.length === 0) {
        results.push(herbName + ': No status marked');
      } else {
        results.push(herbName + ': ' + statuses.join(', ') + dateReady);
      }
    }
  }

  return {
    message: results.join('\n'),
    data: {
      herbs: herbsArray,
      queryType: queryType,
      results: herbStatuses,
      action: 'queried'
    }
  };
}

/**
 * Format date as readable string
 */
function formatDate(date) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return months[date.getMonth()] + ' ' + date.getDate();
}

/**
 * Create standardized response with CORS headers
 */
function createResponse(success, message, data) {
  const response = {
    success: success,
    message: message,
    timestamp: new Date().toISOString()
  };

  if (data) {
    response.data = data;
  }

  // IMPORTANT: Add CORS headers so web app can read the response
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Test function - Mark status
 */
function testMarkStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');

  if (!sheet) {
    Logger.log('ERROR: Sheet named "Inventory" not found!');
    return;
  }

  Logger.log('Test 1: Mark single herb as Low');
  const result1 = markHerbStatus(sheet, ['Yarrow'], 'Low');
  Logger.log(result1.message);

  Logger.log('Test 2: Mark multiple herbs as In Tincture');
  const result2 = markHerbStatus(sheet, ['Chamomile', 'Motherwort'], 'In Tincture');
  Logger.log(result2.message);

  Logger.log('Test 3: Mark as Clinic Backstock');
  const result3 = markHerbStatus(sheet, ['Yarrow'], 'Clinic Backstock');
  Logger.log(result3.message);

  Logger.log('All tests complete! Check your sheet.');
}

/**
 * Test function - Query status
 */
function testQuery() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');

  Logger.log('Query Test 1: Check all statuses for Yarrow');
  const result1 = queryHerbStatus(sheet, ['Yarrow']);
  Logger.log(result1.message);
  Logger.log('Data: ' + JSON.stringify(result1.data, null, 2));

  Logger.log('\nQuery Test 2: Check Clinic Backstock for multiple herbs');
  const result2 = queryHerbStatus(sheet, ['Yarrow', 'Angelica', 'Baptisia'], 'Clinic Backstock');
  Logger.log(result2.message);
  Logger.log('Data: ' + JSON.stringify(result2.data, null, 2));
}

/**
 * Test function - Clear status
 */
function testClear() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');

  Logger.log('Clear Test: Remove Low status from Yarrow');
  const result = clearHerbStatus(sheet, ['Yarrow'], 'Low');
  Logger.log(result.message);
}

/**
 * Manual test of doPost - simulates a web app call
 */
function testDoPost() {
  // Test 1: Mark status
  Logger.log('=== Test 1: Mark Status ===');
  const testEvent1 = {
    postData: {
      contents: JSON.stringify({
        action: 'mark_status',
        herbs: ['Test Herb'],
        statusType: 'Low'
      })
    }
  };
  const response1 = doPost(testEvent1);
  Logger.log('Response: ' + response1.getContent());

  // Test 2: Query specific status
  Logger.log('\n=== Test 2: Query Specific Status ===');
  const testEvent2 = {
    postData: {
      contents: JSON.stringify({
        action: 'query',
        herbs: ['Yarrow', 'Angelica', 'Baptisia'],
        queryType: 'Clinic Backstock'
      })
    }
  };
  const response2 = doPost(testEvent2);
  Logger.log('Response: ' + response2.getContent());

  // Test 3: Query all statuses
  Logger.log('\n=== Test 3: Query All Statuses ===');
  const testEvent3 = {
    postData: {
      contents: JSON.stringify({
        action: 'query',
        herbs: ['Yarrow']
      })
    }
  };
  const response3 = doPost(testEvent3);
  Logger.log('Response: ' + response3.getContent());
}
