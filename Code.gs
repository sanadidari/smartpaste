/**
 * SmartPaste AI - Intelligent Copy-Paste for Google Sheets
 * GitHub: Sanadidari Ecosystem
 */

function onOpen() {
  createMenu();
}

/**
 * Creates the menu in the Sheets interface
 */
function createMenu() {
  SpreadsheetApp.getUi()
    .createMenu('✨ SmartPaste AI')
    .addItem('Open Assistant', 'showSidebar')
    .addItem('Force Refresh Menu', 'createMenu')
    .addToUi();
}

/**
 * Configuration function (placeholder)
 */
function showConfig() {
  SpreadsheetApp.getUi().alert("🔧 Configuration coming soon.");
}

/**
 * Displays the SmartPaste Sidebar
 */
function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('SmartPaste AI')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Utility function to include HTML files (CSS/JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * specialized Gemini API call for data transmutation (with Fallback 10 models)
 */
function callGeminiForTransmutation(rawData) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY not configured. Run setupApiKey() first.');
  const models = [
    'gemini-2.0-flash', 
    'gemini-pro-latest',
    'gemini-flash-latest',
    'gemini-2.0-flash-lite',
    'gemini-2.0-flash-exp-image-generation',
    'gemini-2.5-flash',
    'gemini-2.5-pro',
    'gemini-1.5-flash',
    'gemini-1.5-pro'
  ];

  const systemPrompt = "You are an expert Data Scientist. Your mission is to take raw unstructured data (text, web snippets, PDF extracts) " +
    "and transform them into a 2D JSON array (Array of Arrays). " +
    "INSTRUCTIONS: " +
    "1. Detect logical column headers. " +
    "2. Clean formats (dates as YYYY-MM-DD, numbers without currency symbols). " +
    "3. Respond ONLY with the 2D array JSON. " +
    "Example: [[\"Name\", \"Price\"], [\"Product A\", 10], [\"Product B\", 20]]. " +
    "Data to process: " + rawData;

  const payload = {
    contents: [{ parts: [{ text: systemPrompt }] }]
  };

  let lastError = "No model could respond.";

  for (let model of models) {
    try {
      const apiUrl = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(apiUrl, options);
      const resText = response.getContentText();
      const content = JSON.parse(resText);
      
      if (response.getResponseCode() === 200 && content.candidates && content.candidates[0].content) {
        let aiText = content.candidates[0].content.parts[0].text;
        const jsonMatch = aiText.match(/\[[\s\S]*\]/); 
        const cleanJson = jsonMatch ? jsonMatch[0] : aiText.trim();
        return JSON.parse(cleanJson);
      } else {
        lastError = "Model " + model + " - " + (content.error ? content.error.message : resText);
        console.warn(lastError);
        continue;
      }
    } catch (e) {
      lastError = "System error with " + model + ": " + e.toString();
      console.error(lastError);
    }
  }
  
  throw new Error("No AI model could provide a valid response.");
}

/**
 * AI Transmutation Engine
 * Transforms raw text into a structured table
 */
function processSmartPaste(rawData) {
  try {
    const email = Session.getActiveUser().getEmail();
    const credits = getUserCredits(email);
    
    if (credits <= 0) {
      return { status: 'error', message: '⚠️ Your credits are exhausted. Please recharge to continue.' };
    }

    // Call Gemini engine
    const cleanedData = callGeminiForTransmutation(rawData);
    
    // Deduct one credit
    updateUserCredits(credits - 1);
    
    return { 
      status: 'success', 
      data: cleanedData,
      credits: credits - 1,
      message: '✨ Data transmuted successfully!' 
    };
  } catch (e) {
    return { status: 'error', message: '❌ Error: ' + e.toString() };
  }
}

/**
 * Inserts cleaned data into the active sheet
 */
function insertDataIntoSheet(dataMatrix) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeRange = sheet.getActiveRange();
    const startRow = activeRange.getRow();
    const startCol = activeRange.getColumn();
    
    sheet.getRange(startRow, startCol, dataMatrix.length, dataMatrix[0].length)
      .setValues(dataMatrix);
      
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// --- CREDIT MANAGEMENT (Shared with Sanadidari Router) ---

function getUserCredits(email) {
  const scriptProps = PropertiesService.getScriptProperties();
  const creditsDB = JSON.parse(scriptProps.getProperty('CREDITS_DB') || "{}");
  return creditsDB[email] === undefined ? 10 : creditsDB[email];
}

function updateCreditsByEmail(email, newCount) {
  const scriptProps = PropertiesService.getScriptProperties();
  const creditsDB = JSON.parse(scriptProps.getProperty('CREDITS_DB') || "{}");
  creditsDB[email] = newCount;
  scriptProps.setProperty('CREDITS_DB', JSON.stringify(creditsDB));
}

function updateUserCredits(newCount) {
  updateCreditsByEmail(Session.getActiveUser().getEmail(), newCount);
}

/**
 * Handles PayPal payment notifications (Via Sanadidari Hub Router)
 */
function doPost(e) {
  try {
    const params = e.parameter;
    const paymentStatus = params.payment_status;
    const payerEmail = params.payer_email; // Sent by Hub Router
    const itemNumber = params.item_number || ""; // Ex: SP_PACK_50

    if (paymentStatus === "Completed" && itemNumber.indexOf("SP_") === 0) {
      let creditsToAdd = 0;
      if (itemNumber === "SP_PACK_10" || itemNumber.includes("10")) creditsToAdd = 10;
      if (itemNumber === "SP_PACK_50" || itemNumber.includes("50")) creditsToAdd = 50;
      if (itemNumber === "SP_PACK_100" || itemNumber.includes("100")) creditsToAdd = 100;
      if (itemNumber === "SP_PACK_500" || itemNumber.includes("500")) creditsToAdd = 500;

      if (creditsToAdd > 0 && payerEmail) {
        const currentCredits = getCreditsByEmail(payerEmail);
        updateCreditsByEmail(payerEmail, currentCredits + creditsToAdd);
        console.log(`Payment validated: +${creditsToAdd} credits for ${payerEmail}`);
      }
      return ContentService.createTextOutput("SUCCESS");
    }
  } catch (err) {
    console.error("SmartPaste IPN Error: " + err.toString());
    return ContentService.createTextOutput("ERROR");
  }
  return ContentService.createTextOutput("IGNORED");
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Retrieves credits for a specific email
 */
function getCreditsByEmail(email) {
  const scriptProps = PropertiesService.getScriptProperties();
  const creditsDB = JSON.parse(scriptProps.getProperty('CREDITS_DB') || "{}");
  return creditsDB[email] === undefined ? 10 : creditsDB[email];
}

/**
 * Retrieves credits for the current user
 */
function getCurrentUserCredits() {
  return getCreditsByEmail(getUserEmail());
}

/**
 * Run once from the SmartPaste AI menu or Apps Script editor to store your Gemini API key.
 */
function setupApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Setup', 'Enter your Gemini API Key:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() === ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', result.getResponseText().trim());
    ui.alert('✅ API Key saved securely.');
  }
}
