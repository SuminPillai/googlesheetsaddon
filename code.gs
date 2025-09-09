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
  const { tickers, fromDate, toDate, columns } = formData;
  
  if (!tickers || tickers.length === 0 || !fromDate || !toDate || !columns || columns.length === 0) {
    return { success: false, message: "Please provide all required inputs." };
  }

  // Use the correct, publicly accessible Cloud Run URL.
  const backendUrl = "https://excel-addin-backend-o5molvd7pa-el.a.run.app";

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let row = 1;

    // Split tickers into an array if they are passed as a single string.
    const tickerList = Array.isArray(tickers) ? tickers : tickers.split(',').map(t => t.trim()).filter(t => t.length > 0);
    
    // Construct the URL with query parameters for a GET request
    const queryParams = `?symbols=${tickerList.join(',')}&from=${fromDate}&to=${toDate}&columns=${columns.join(',')}`;
    const fullUrl = `${backendUrl}/stocks${queryParams}`;


    // Make the GET request to the backend
    Logger.log(`Sending GET request to backend: ${fullUrl}`);
    const response = UrlFetchApp.fetch(fullUrl); // Method is GET by default
    const data = JSON.parse(response.getContentText());

    if (data.length === 0) {
      sheet.getRange(row, 1).setValue(`No data found for the selected tickers.`);
      row += 2;
    } else {
      // Write header row
      const headers = Object.keys(data[0]);
      sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(row, 1, 1, headers.length).setFontWeight("bold");
      row++;

      // Write data rows
      const dataRows = data.map(item => headers.map(header => item[header]));
      sheet.getRange(row, 1, dataRows.length, headers.length).setValues(dataRows);
      row += dataRows.length + 2;
    }
    
    return { success: true, message: `Successfully fetched and displayed data.` };
  } catch (e) {
    return { success: false, message: `Error: ${e.message}. Please make sure your backend server is running and the URL is correct.` };
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
    symbols: ["RCDL"],
    from: "2025-01-01",
    to: "2025-01-05",
    columns: ["Open", "Close"]
  };

  // Construct the URL with query parameters for a GET request
  const queryParams = `?symbols=${testPayload.symbols.join(',')}&from=${testPayload.from}&to=${testPayload.to}&columns=${testPayload.columns.join(',')}`;
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
