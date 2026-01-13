// ============================================
// SPRINKLER PERMIT FORM HANDLER
// Version 1.0 - Complete Solution
// ============================================

// CONFIGURATION - CHANGE THESE VALUES
const CONFIG = {
  SENDER_EMAIL: 'vaidyaajith12@gmail.com',   // Your Gmail that sends emails
  RECIPIENT_EMAIL: 'facilities@team4lifespaces.com',  // Where to send
  SHEET_NAME: 'Sprinkler_Permit_Log',        // Google Sheet name
  EMAIL_SUBJECT_PREFIX: '🚒 Sprinkler Permit - '
};

// Main function that receives form data
function doPost(e) {
  try {
    // Parse the incoming data
    const data = JSON.parse(e.postData.contents);
    const timestamp = new Date();
    
    // 1. LOG TO GOOGLE SHEET
    const logResult = logToGoogleSheet(data, timestamp);
    
    // 2. SEND EMAIL WITH PDF
    const emailResult = sendNotificationEmail(data, timestamp);
    
    // Return success response
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Form submitted successfully',
      logId: logResult.logId,
      emailSent: emailResult.success,
      timestamp: timestamp.toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // Return error response
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error processing form: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Function to log data to Google Sheet
function logToGoogleSheet(data, timestamp) {
  try {
    // Get or create the spreadsheet
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = spreadsheet.getSheetByName('Submissions') || spreadsheet.insertSheet('Submissions');
    
    // Set headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      const headers = [
        'Timestamp', 'Form ID', 'Scroll No', 'Date', 'Tower', 'Flat No',
        'Owner Name', 'Contact', 'Owner Email', 'Extension Locations',
        'Advisory Accepted', 'Signature Saved', 'PDF Generated', 'Email Status'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
    }
    
    // Generate unique Form ID
    const formId = 'SPR-' + timestamp.getTime().toString(36).toUpperCase();
    
    // Prepare row data
    const rowData = [
      timestamp.toLocaleString(),
      formId,
      data.scroll_no || '',
      data.date || '',
      data.tower || '',
      data.flat_no || '',
      data.owner_name || '',
      data.contact_number || '',
      data.email || '',
      data.extension_locations || 'None',
      data.advisory_acknowledged ? 'YES' : 'NO',
      data.signature ? 'YES' : 'NO',
      data.pdf_base64 ? 'YES' : 'NO',
      'PENDING'
    ];
    
    // Append row to sheet
    sheet.appendRow(rowData);
    
    // Format the new row
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, rowData.length).setBorder(true, true, true, true, true, true);
    
    return {
      success: true,
      logId: formId,
      rowNumber: lastRow
    };
    
  } catch (error) {
    console.error('Error logging to sheet:', error);
    throw error;
  }
}

// Function to send email with PDF attachment
function sendNotificationEmail(data, timestamp) {
  try {
    // Generate Form ID
    const formId = 'SPR-' + timestamp.getTime().toString(36).toUpperCase();
    
    // Create email content
    const subject = CONFIG.EMAIL_SUBJECT_PREFIX + data.tower + ' - Flat ' + data.flat_no;
    
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
    .header { background-color: #d32f2f; color: white; padding: 20px; text-align: center; }
    .content { padding: 20px; border: 1px solid #ddd; margin: 20px; }
    .field { margin-bottom: 15px; }
    .label { font-weight: bold; color: #555; display: inline-block; width: 180px; }
    .value { color: #222; }
    .section { background-color: #f9f9f9; padding: 15px; margin: 15px 0; border-left: 4px solid #d32f2f; }
    .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; }
    .important { background-color: #fff3e0; padding: 15px; border-left: 4px solid #ff9800; margin: 15px 0; }
  </style>
</head>
<body>
  <div class="header">
    <h2>🚒 SPRINKLER MODIFICATION WORK PERMIT</h2>
    <p>File Safety Compliance Form</p>
  </div>
  
  <div class="content">
    <div class="section">
      <h3>📋 Form Information</h3>
      <div class="field"><span class="label">Form ID:</span> <span class="value">${formId}</span></div>
      <div class="field"><span class="label">Submission Time:</span> <span class="value">${timestamp.toLocaleString()}</span></div>
      <div class="field"><span class="label">Scroll No:</span> <span class="value">${data.scroll_no || ''}</span></div>
    </div>
    
    <div class="section">
      <h3>👤 Owner Details</h3>
      <div class="field"><span class="label">Owner Name:</span> <span class="value">${data.owner_name || ''}</span></div>
      <div class="field"><span class="label">Flat No:</span> <span class="value">${data.tower || ''} - ${data.flat_no || ''}</span></div>
      <div class="field"><span class="label">Contact:</span> <span class="value">${data.contact_number || ''}</span></div>
      <div class="field"><span class="label">Email:</span> <span class="value">${data.email || ''}</span></div>
    </div>
    
    <div class="section">
      <h3>📍 Extension Locations</h3>
      <div class="value">${data.extension_locations || 'No locations selected'}</div>
    </div>
    
    <div class="important">
      <h4>⚠️ Important Notes:</h4>
      <p>• Advisory Acknowledged: <strong>${data.advisory_acknowledged ? 'YES' : 'NO'}</strong></p>
      <p>• Signature Provided: <strong>${data.signature ? 'YES' : 'NO'}</strong></p>
      <p>• PDF Generated: <strong>${data.pdf_base64 ? 'YES - See attachment' : 'NO - Data only'}</strong></p>
    </div>
  </div>
  
  <div class="footer">
    <p>This is an automated notification from Sprinkler Permit System.</p>
    <p>Form data has been logged to Google Sheets: ${CONFIG.SHEET_NAME}</p>
    <p>© ${new Date().getFullYear()} Team4LifeSpaces - Facilities Management</p>
  </div>
</body>
</html>`;
    
    // Try to attach PDF if available
    let attachments = [];
    if (data.pdf_base64 && data.pdf_base64.length > 100) {
      try {
        const pdfBlob = Utilities.newBlob(
          Utilities.base64Decode(data.pdf_base64), 
          'application/pdf', 
          `Sprinkler_Permit_${data.flat_no}_${formId}.pdf`
        );
        attachments = [pdfBlob];
      } catch (pdfError) {
        console.warn('Could not create PDF attachment:', pdfError);
        // Continue without PDF attachment
      }
    }
    
    // Send the email
    MailApp.sendEmail({
      to: CONFIG.RECIPIENT_EMAIL,
      subject: subject,
      htmlBody: htmlBody,
      name: 'Sprinkler Permit System',
      replyTo: data.email || CONFIG.SENDER_EMAIL,
      attachments: attachments
    });
    
    // Update sheet with email status
    updateEmailStatusInSheet(formId, 'SENT');
    
    return {
      success: true,
      emailId: formId,
      hasAttachment: attachments.length > 0
    };
    
  } catch (error) {
    console.error('Error sending email:', error);
    updateEmailStatusInSheet('UNKNOWN', 'FAILED: ' + error.toString());
    
    // Try to send a simple text email as fallback
    try {
      MailApp.sendEmail({
        to: CONFIG.RECIPIENT_EMAIL,
        subject: 'URGENT: Sprinkler Form Error',
        body: 'Form submission failed. Check Apps Script logs. Data: ' + JSON.stringify(data, null, 2)
      });
    } catch (fallbackError) {
      console.error('Fallback email also failed:', fallbackError);
    }
    
    throw error;
  }
}

// Helper function to get or create spreadsheet
function getOrCreateSpreadsheet() {
  try {
    // Try to find existing spreadsheet
    const files = DriveApp.getFilesByName(CONFIG.SHEET_NAME);
    if (files.hasNext()) {
      return SpreadsheetApp.open(files.next());
    }
    
    // Create new spreadsheet
    const newSpreadsheet = SpreadsheetApp.create(CONFIG.SHEET_NAME);
    
    // Create sheets
    const submissionsSheet = newSpreadsheet.getSheets()[0];
    submissionsSheet.setName('Submissions');
    
    // Create summary sheet
    const summarySheet = newSpreadsheet.insertSheet('Summary');
    setupSummarySheet(summarySheet);
    
    // Create settings sheet
    const settingsSheet = newSpreadsheet.insertSheet('Settings');
    setupSettingsSheet(settingsSheet);
    
    // Move sheets to correct order
    newSpreadsheet.setActiveSheet(submissionsSheet);
    newSpreadsheet.moveActiveSheet(1);
    newSpreadsheet.setActiveSheet(summarySheet);
    newSpreadsheet.moveActiveSheet(2);
    
    return newSpreadsheet;
    
  } catch (error) {
    console.error('Error with spreadsheet:', error);
    throw error;
  }
}

// Setup summary sheet
function setupSummarySheet(sheet) {
  const headers = ['Metric', 'Value', 'Last Updated'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  
  const metrics = [
    ['Total Submissions', '=COUNTA(Submissions!A:A)-1', new Date()],
    ['Today\'s Submissions', '=COUNTIF(Submissions!A:A,">="&TODAY())', new Date()],
    ['PDFs Generated', '=COUNTIF(Submissions!M:M,"YES")', new Date()],
    ['Pending Review', '=COUNTIF(Submissions!N:N,"PENDING")', new Date()]
  ];
  
  sheet.getRange(2, 1, metrics.length, 3).setValues(metrics);
  sheet.autoResizeColumns(1, 3);
}

// Setup settings sheet
function setupSettingsSheet(sheet) {
  const settings = [
    ['CONFIGURATION', '', ''],
    ['System Email', CONFIG.SENDER_EMAIL, ''],
    ['Recipient Email', CONFIG.RECIPIENT_EMAIL, ''],
    ['', '', ''],
    ['LAST UPDATED', new Date().toLocaleString(), ''],
    ['SCRIPT VERSION', '1.0', ''],
    ['', '', ''],
    ['INSTRUCTIONS', '', ''],
    ['1. Do not delete this sheet', '', ''],
    ['2. All form data goes to Submissions sheet', '', ''],
    ['3. Summary updates automatically', '', '']
  ];
  
  sheet.getRange(1, 1, settings.length, 3).setValues(settings);
  sheet.getRange(1, 1, 1, 3).merge().setBackground('#d32f2f').setFontColor('white').setFontWeight('bold');
  sheet.autoResizeColumns(1, 3);
}

// Update email status in sheet
function updateEmailStatusInSheet(formId, status) {
  try {
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = spreadsheet.getSheetByName('Submissions');
    
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === formId) { // Form ID is in column B
        sheet.getRange(i + 1, 14).setValue(status); // Column N
        break;
      }
    }
  } catch (error) {
    console.error('Error updating email status:', error);
  }
}

// Function to test the setup
function testSetup() {
  console.log('Testing Sprinkler Permit System...');
  
  // Test spreadsheet
  const ss = getOrCreateSpreadsheet();
  console.log('Spreadsheet URL:', ss.getUrl());
  
  // Test email (without actually sending)
  console.log('Sender Email:', CONFIG.SENDER_EMAIL);
  console.log('Recipient Email:', CONFIG.RECIPIENT_EMAIL);
  
  // Create a test log entry
  const testData = {
    scroll_no: 'TEST001',
    date: '2026-01-12',
    tower: 'Tower A',
    flat_no: '10F',
    owner_name: 'Test Owner',
    contact_number: '1234567890',
    email: 'test@example.com',
    extension_locations: 'Hall, Kitchen',
    advisory_acknowledged: true,
    signature: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    pdf_base64: 'JVBERi0xLjQKJcOkw7zDtsOfCjIgMCBvYmoKPDwvTGVuZ3RoIDMgMCBSL0ZpbHRlci9GbGF0ZURlY29kZT4+CnN0cmVhbQp4nF2UwY6kIBCG732KXKaPoRvQk5NJMsnshcxhH3'
  };
  
  const result = logToGoogleSheet(testData, new Date());
  console.log('Test log result:', result);
  
  return {
    spreadsheetUrl: ss.getUrl(),
    testPassed: true,
    message: 'Setup test completed successfully'
  };
}

// ============================================
// END OF MAIN SCRIPT
// ============================================