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
function fetchDataAndAnalyze(requestData) {
  const { tickers, fromDate, toDate, columns, startCell, includeAiAnalysis, analysisType, customQuestion } = requestData;

  if (!tickers || tickers.length === 0 || !fromDate || !toDate || !columns || columns.length === 0) {
    return { success: false, message: "Please provide all required inputs." };
  }

  const backendUrl = "https://excel-addin-backend-o5molvd7pa-el.a.run.app";
  const tickerList = Array.isArray(tickers) ? tickers : tickers.split(',').map(t => t.trim()).filter(t => t.length > 0);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    let startRow = 1;
    let startCol = 1;
    if (startCell) {
      try {
        const range = sheet.getRange(startCell);
        startRow = range.getRow();
        startCol = range.getColumn();
      } catch (e) {
        return { success: false, message: `Invalid start cell: ${startCell}.` };
      }
    }
    let currentRow = startRow;
    let singleAnalysisResult = null;

    for (const ticker of tickerList) {
      const encodedColumns = encodeURIComponent(columns.join(','));
      const queryParams = `?symbol=${encodeURIComponent(ticker)}&from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}&columns=${encodedColumns}`;
      const fullUrl = `${backendUrl}/stocks${queryParams}`;

      try {
        const response = UrlFetchApp.fetch(fullUrl);
        const data = JSON.parse(response.getContentText());

        if (data.length === 0) {
          sheet.getRange(currentRow, startCol).setValue(`No data found for ${ticker}.`);
          currentRow += 2;
        } else {
          const originalHeaders = Object.keys(data[0]);
          const headersWithSymbol = ["Symbol", ...originalHeaders];
          const dataRowsWithSymbol = data.map(item => {
            const row = originalHeaders.map(header => item[header]);
            return [ticker, ...row];
          });

          sheet.getRange(currentRow, startCol, 1, headersWithSymbol.length).setValues([headersWithSymbol]).setFontWeight("bold");
          currentRow++;
          sheet.getRange(currentRow, startCol, dataRowsWithSymbol.length, headersWithSymbol.length).setValues(dataRowsWithSymbol);
          currentRow += dataRowsWithSymbol.length;

          // Conditional AI Analysis
          if (includeAiAnalysis) {
            const analysis = _getAiAnalysis(ticker, data, fromDate, toDate, analysisType, customQuestion);
            sheet.getRange(currentRow, startCol).setValue(analysis).setFontStyle('italic');

            if (tickerList.length === 1) {
              singleAnalysisResult = analysis;
            }
            currentRow++;
          }

          currentRow++; // Add a blank row between tickers
        }
      } catch (innerError) {
        Logger.log(`Error fetching data for ${ticker}: ${innerError.message}`);
        sheet.getRange(currentRow, startCol).setValue(`Error fetching data for ${ticker}: ${innerError.message}`);
        currentRow += 2;
      }
    }

    return { success: true, message: `Successfully fetched and displayed data.`, aiAnalysis: singleAnalysisResult };
  } catch (e) {
    return { success: false, message: `An unexpected error occurred: ${e.message}.` };
  }
}

/**
 * Calls the Gemini API to get a financial analysis for a given set of stock data.
 * @param {string} ticker The stock ticker symbol.
 * @param {Array<Object>} stockData The array of stock data for the ticker.
 * @param {string} fromDate The start date for the analysis period.
 * @param {string} toDate The end date for the analysis period.
 * @param {string} analysisType The type of analysis to perform (e.g., 'summary', 'swot').
 * @param {string} customQuestion A custom question for the AI.
 * @return {string} The AI-generated analysis text, or an error message.
 */
function _getAiAnalysis(ticker, stockData, fromDate, toDate, analysisType, customQuestion) {
  try {
    const apiKey = _getApiKey();
    const dataString = JSON.stringify(stockData);
    let prompt;

    if (customQuestion && customQuestion.trim() !== '') {
      prompt = `You are a helpful financial assistant. All monetary values are in Indian Rupees (INR). Given the following daily stock data for ${ticker} from ${fromDate} to ${toDate}: ${dataString}. Please provide a clear and concise answer to the following question: "${customQuestion}"`;
    } else {
      switch (analysisType) {
        case 'swot':
          prompt = `You are a financial analyst. All monetary values are in Indian Rupees (INR). Based on the following daily stock data for ${ticker} from ${fromDate} to ${toDate}, generate a brief SWOT analysis (Strengths, Weaknesses, Opportunities, Threats). Strengths and weaknesses should be based on the provided data (e.g., price trends, volume). Opportunities and threats can be more general market considerations. Data: ${dataString}`;
          break;
        case 'outlook':
          prompt = `You are a financial analyst. All monetary values are in Indian Rupees (INR). Based on the trends in the following daily stock data for ${ticker} from ${fromDate} to ${toDate}, provide a brief, speculative future outlook. Mention key support or resistance levels if identifiable from the data. Data: ${dataString}`;
          break;
        default: // 'summary'
          prompt = `You are a financial analyst. All monetary values are in Indian Rupees (INR). Analyze the following daily stock data for ${ticker} from ${fromDate} to ${toDate} and provide a concise, one-paragraph summary of its performance, highlighting key trends in price and volume. Do not start with "Here is an analysis". Just provide the analysis. Data: ${dataString}`;
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
 * Gets the value from the active cell and formats it as YYYY-MM-DD.
 * Returns an empty string if the cell value is not a valid date.
 * @returns {string} The formatted date string.
 */
function _getCellValueAsFormattedDate() {
  const cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
  const value = cell.getValue();
  if (value instanceof Date) {
    return Utilities.formatDate(value, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd");
  }
  // Attempt to parse if it's a string or number
  try {
    const date = new Date(value);
    // Check if it's a valid date
    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd");
    }
  } catch (e) {
    // Ignore parsing errors
  }
  return ""; // Return empty if not a valid date
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
