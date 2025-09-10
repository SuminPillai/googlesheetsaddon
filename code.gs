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
          currentRow += dataRowsWithSymbol.length + 2; // Move to the next position, leaving a blank row
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
