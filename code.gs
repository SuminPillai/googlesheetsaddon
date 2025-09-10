/**
 * @OnlyCurrentDoc
 *
 * The above tag is required to give this script access to the current spreadsheet only.
 * This is a best practice for security and privacy.
 */

/**
 * Runs when the spreadsheet is opened.
 * This function creates a custom menu in the Google Sheets UI.
 * The menu will contain an item to show the add-on's sidebar.
 */
function onOpen(e) {
  // Get the UI instance for the active spreadsheet.
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Show Add-on', 'showSidebar')
    .addToUi();
}

/**
 * Runs when the add-on's icon is clicked in the side panel.
 * This function is defined in the "sheets" section of the manifest.
 */
function onSheetsHomepage() {
  showSidebar();
}

/**
 * Displays the add-on's sidebar in the Google Sheets UI.
 * This function will load an HTML file that contains the add-on's UI.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('20 Percent Price Data');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Fetches data from the backend and writes it to the active sheet.
 * This function will be called from the `Sidebar.html` file.
 * @param {object} formData - An object containing the user's input.
 */
function fetchDataFromBackend(formData) {
  const { tickers, fromDate, toDate, columns, startCell } = formData;

  if (!tickers || tickers.length === 0 || !fromDate || !toDate || !columns || columns.length === 0) {
    return { success: false, message: "Please provide all required inputs." };
  }

  const backendUrl = "https://excel-addin-backend-o5molvd7pa-el.a.run.app";
  const tickerList = Array.isArray(tickers) ? tickers : tickers.split(',').map(t => t.trim()).filter(t => t.length > 0);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Determine starting row and column from startCell parameter
    let startRow = 1;
    let startCol = 1;
    if (startCell) {
      try {
        const range = sheet.getRange(startCell);
        startRow = range.getRow();
        startCol = range.getColumn();
      } catch (e) {
        return { success: false, message: `Invalid start cell: ${startCell}. Please provide a valid cell reference (e.g., A2).` };
      }
    }
    let currentRow = startRow;

    for (const ticker of tickerList) {
      const encodedColumns = encodeURIComponent(columns.join(','));
      const queryParams = `?symbol=${encodeURIComponent(ticker)}&from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}&columns=${encodedColumns}`;
      const fullUrl = `${backendUrl}/stocks${queryParams}`;

      Logger.log(`Sending GET request for ${ticker}: ${fullUrl}`);

      try {
        const response = UrlFetchApp.fetch(fullUrl);
        const data = JSON.parse(response.getContentText());

        if (data.length === 0) {
          sheet.getRange(currentRow, startCol).setValue(`No data found for ${ticker}.`);
          currentRow += 2; // Leave a blank row
        } else {
          // Add "Symbol" to headers and data
          const originalHeaders = Object.keys(data[0]);
          const headersWithSymbol = ["Symbol", ...originalHeaders];
          const dataRowsWithSymbol = data.map(item => {
            const row = originalHeaders.map(header => item[header]);
            return [ticker, ...row];
          });

          // Write header row
          sheet.getRange(currentRow, startCol, 1, headersWithSymbol.length).setValues([headersWithSymbol]);
          sheet.getRange(currentRow, startCol, 1, headersWithSymbol.length).setFontWeight("bold");
          currentRow++;

          // Write data rows
          sheet.getRange(currentRow, startCol, dataRowsWithSymbol.length, headersWithSymbol.length).setValues(dataRowsWithSymbol);
          currentRow += dataRowsWithSymbol.length + 2;
        }
      } catch (innerError) {
        // Log and write error for a single ticker, then continue
        Logger.log(`Error fetching data for ${ticker}: ${innerError.message}`);
        sheet.getRange(currentRow, startCol).setValue(`Error fetching data for ${ticker}: ${innerError.message}`);
        currentRow += 2;
      }
    }

    return { success: true, message: `Successfully fetched and displayed data for all requested tickers.` };
  } catch (e) {
    // This will catch errors in the initial setup (e.g., getting the sheet)
    return { success: false, message: `An unexpected error occurred: ${e.message}.` };
  }
}

/**
 * Analyzes stock data using the Gemini AI API based on user-selected analysis type.
 * @param {object} requestData The form data plus the analysis type and custom question.
 * @return {string} A text summary of the stock's performance.
 */
function analyzeStockPerformance(requestData) {
  const { tickers, fromDate, toDate, columns, analysisType, customQuestion, startCell } = requestData;

  // 1. Re-fetch the data for the first ticker to be analyzed
  if (!tickers || tickers.length === 0) {
    return "Error: No ticker provided for analysis.";
  }
  const ticker = tickers[0];

  const backendUrl = "https://excel-addin-backend-o5molvd7pa-el.a.run.app";
  const encodedColumns = encodeURIComponent(columns.join(','));
  const queryParams = `?symbol=${encodeURIComponent(ticker)}&from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}&columns=${encodedColumns}`;
  const fullUrl = `${backendUrl}/stocks${queryParams}`;

  let stockData;
  try {
    const response = UrlFetchApp.fetch(fullUrl);
    stockData = JSON.parse(response.getContentText());
    if (stockData.length === 0) {
      return `No data found for ${ticker} to analyze.`;
    }
  } catch (e) {
    Logger.log(`Error re-fetching data for analysis: ${e.message}`);
    return `Could not fetch data for ${ticker} to analyze.`;
  }

  // 2. Call the Gemini API with a dynamic prompt
  try {
    const apiKey = _getApiKey(); // Securely get the API key
    const dataString = JSON.stringify(stockData);
    let prompt;

    if (customQuestion && customQuestion.trim() !== '') {
      prompt = `You are a helpful financial assistant. Given the following daily stock data for ${ticker} from ${fromDate} to ${toDate}: ${dataString}. Please provide a clear and concise answer to the following question: "${customQuestion}"`;
    } else {
      switch (analysisType) {
        case 'swot':
          prompt = `You are a financial analyst. Based on the following daily stock data for ${ticker} from ${fromDate} to ${toDate}, generate a brief SWOT analysis (Strengths, Weaknesses, Opportunities, Threats). Strengths and weaknesses should be based on the provided data (e.g., price trends, volume). Opportunities and threats can be more general market considerations. Data: ${dataString}`;
          break;
        case 'outlook':
          prompt = `You are a financial analyst. Based on the trends in the following daily stock data for ${ticker} from ${fromDate} to ${toDate}, provide a brief, speculative future outlook. Mention key support or resistance levels if identifiable from the data. Data: ${dataString}`;
          break;
        default: // 'summary'
          prompt = `You are a financial analyst. Analyze the following daily stock data for ${ticker} from ${fromDate} to ${toDate} and provide a concise, one-paragraph summary of its performance, highlighting key trends in price and volume. Do not start with "Here is an analysis". Just provide the analysis. Data: ${dataString}`;
          break;
      }
    }

    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    const payload = { "contents": [{"parts": [{"text": prompt}]}] };
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = JSON.parse(response.getContentText());

    if (responseData.error) {
      Logger.log(`Gemini API Error: ${JSON.stringify(responseData.error)}`);
      return `AI Error: ${responseData.error.message}`;
    }

    const analysis = responseData.candidates[0].content.parts[0].text.trim();

    // 3. Write the analysis to the sheet
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const range = sheet.getRange(startCell);
      // Place the analysis 2 rows below the header of the data block it analyzed
      const outputRow = range.getRow() + stockData.length + 1;
      sheet.getRange(outputRow, range.getColumn()).setValue(analysis).setFontStyle('italic');
    } catch (e) {
      Logger.log(`Error writing AI analysis to sheet: ${e.message}`);
      // Don't fail the whole function if writing to the sheet fails, just log it.
    }

    return analysis;
  } catch (e) {
    Logger.log(`Error calling Gemini API: ${e.toString()}`);
    return `Error: Could not connect to the AI service. ${e.message}`;
  }
}

/**
 * Retrieves the value of the currently selected cell.
 * @return {string} The value of the cell.
 */
function getCellValue() {
  const cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
  return cell.getValue();
}

/**
 * Retrieves the A1 notation of the currently selected cell.
 * @return {string} The A1 notation of the cell (e.g., 'A1').
 */
function getActiveCellA1Notation() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getA1Notation();
}

/**
 * Retrieves the values of a given range.
 * @param {string} rangeA1Notation The A1 notation of the range (e.g., 'A1:A10').
 * @return {Array<Array<any>>} The values of the cells in the range.
 */
function getCellRangeValues(rangeA1Notation) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  try {
    const range = sheet.getRange(rangeA1Notation);
    return range.getValues();
  } catch (e) {
    throw new Error(`Invalid range: ${rangeA1Notation}. Please provide a valid range.`);
  }
}

/**
 * This function is an entry point for the web app deployment.
 * It's not used by the add-on, but it's essential for the deployment process.
 */
function doGet() {
  return HtmlService.createHtmlOutput('Web App is running successfully!');
}

/**
 * This function is an entry point for the web app deployment.
 * It's not used by the add-on, but it's essential for the deployment process.
 */
function doPost(e) {
  const requestBody = JSON.parse(e.postData.contents);
  const { symbols, from, to, columns } = requestBody;

  // Simulate calling the fetchDataFromBackend function
  const formData = {
    tickers: symbols,
    fromDate: from,
    toDate: to,
    columns: columns
  };
  
  const result = fetchDataFromBackend(formData);
  
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Tests the connection to the backend service.
 */
function testBackendConnection() {
  const backendUrl = "https://excel-addin-backend-o5molvd7pa-el.a.run.app";
  const testPayload = {
    symbol: "RCDL", // singular
    from: "2025-01-01",
    to: "2025-01-05",
    columns: ["Open", "Close"]
  };

  const queryParams = `?symbol=${encodeURIComponent(testPayload.symbol)}&from=${encodeURIComponent(testPayload.from)}&to=${encodeURIComponent(testPayload.to)}&columns=${encodeURIComponent(testPayload.columns.join(','))}`;
  const fullUrl = `${backendUrl}/stocks${queryParams}`;

  try {
    Logger.log("Testing backend connection with a sample GET request...");
    const response = UrlFetchApp.fetch(fullUrl);
    const responseText = response.getContentText();
    Logger.log(`Backend responded with: ${responseText}`);
    return { success: true, message: `Successfully connected to backend. Response: ${responseText}` };
  } catch (e) {
    Logger.log(`Failed to connect to backend: ${e.message}`);
    return { success: false, message: `Failed to connect to backend: ${e.message}` };
  }
}

/**
 * A helper function to be run ONCE from the script editor to set the API key.
 * @param {string} key The Gemini API key.
 */
function _setApiKey(key) {
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
}

/**
 * A helper function to retrieve the API key from script properties.
 * @returns {string} The API key.
 */
function _getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('API key not found. Please set the GEMINI_API_KEY in script properties.');
  }
  return apiKey;
}
