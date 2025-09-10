20 Percent Price Data: Excel & Google Sheets Add-in
This project contains a financial data add-in for both Microsoft Excel and Google Sheets. The add-in is designed to fetch real-time and historical stock data, including financial metrics and technical indicators, directly into your spreadsheets.

The solution is powered by a backend service hosted on Google Cloud Run and uses a Google Cloud SQL instance for data storage.

Features
Multi-Platform Support: Works as a task pane add-in in Microsoft Excel and as a sidebar add-on in Google Sheets.

Comprehensive Financial Data: Fetches price data and key financial metrics from various data providers and the yfinance library.

Technical Indicators: Supports a wide range of technical indicators for advanced analysis.

Cloud-Based Backend: The data fetching and processing logic are handled by a scalable, serverless backend on Google Cloud Run.

Database Integration: Connects to a Google Cloud SQL instance for persistent data storage.

Technology Stack
Frontend: HTML, CSS, and JavaScript for the user interface.

Add-in Frameworks: Microsoft Office Add-in Platform and Google Apps Script.

Backend: Python application running on Google Cloud Run.

Database: Google Cloud SQL (SQL Server instance)

Data Sources: yfinance library and proprietary data providers.

APIs: Excel JavaScript API, Google Sheets API, and UrlFetchApp.

Setup and Deployment
This guide covers the steps required to get the add-in running for both platforms.

1. Google Cloud Configuration
A user-managed Google Cloud Platform (GCP) project is required to host the backend and manage API access.

Project ID: plus-percent

Project Number: 1088354707719

Cloud Run Service: excel-addin-backend-o5molvd7pa-el.a.run.app

Cloud SQL Instance: plus-percent:us-central1:stock-data-server-2

OAuth Consent Screen:
The OAuth consent screen must be configured as External.

APIs to Enable:
The following APIs must be enabled in your GCP project:

Google Sheets API

Google Workspace Marketplace SDK

2. Google Sheets Add-on (Code.gs and Sidebar.html)
The Google Sheets add-on is built using Google Apps Script.

appsscript.json
The manifest file must explicitly whitelist the backend URL for security.

{
  "timeZone": "Asia/Kolkata",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "[https://www.googleapis.com/auth/script.container.ui](https://www.googleapis.com/auth/script.container.ui)",
    "[https://www.googleapis.com/auth/spreadsheets.currentonly](https://www.googleapis.com/auth/spreadsheets.currentonly)",
    "[https://www.googleapis.com/auth/script.external_request](https://www.googleapis.com/auth/script.external_request)"
  ],
  "urlFetchWhitelist": [
    "[https://excel-addin-backend-o5molvd7pa-el.a.run.app](https://excel-addin-backend-o5molvd7pa-el.a.run.app)"
  ],
  "addOns": {
    "common": {
      "name": "20 Percent Price Data",
      "logoUrl": "[https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/20percent_logo.png](https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/20percent_logo.png)",
      "homepageTrigger": {
        "runFunction": "onOpen"
      }
    },
    "sheets": {
      "homepageTrigger": {
        "runFunction": "onSheetsHomepage"
      }
    }
  }
}

Code.gs
This file contains the server-side logic for the add-on, including UI creation, data retrieval, and the backend API call.

#### Core Logic: Fetching Data

The primary function for interacting with the backend is `fetchDataFromBackend`. After a series of debugging steps, this function was refactored to correctly match the backend API's contract. Here is how it works:

1.  **Iterates Through Tickers**: The backend is designed to handle one stock symbol per request. The function loops through the list of tickers provided by the user.
2.  **Constructs a GET Request**: For each ticker, it constructs a unique URL. The data (symbol, from date, to date, columns) is passed as URL query parameters.
3.  **Encodes Parameters**: All query parameter values are wrapped in `encodeURIComponent` to handle special characters (like `%`) and prevent errors.
4.  **Fetches Data**: It uses `UrlFetchApp.fetch()` to make a `GET` request to the backend.
5.  **Processes and Writes Data**: The JSON response is parsed, and the data is written to the active spreadsheet, with headers for each ticker's data block.
6.  **Robust Error Handling**: If a request for a single ticker fails, the error is logged and written to the sheet, and the loop continues with the next ticker.

Here is the final, corrected code for the function:

```javascript
function fetchDataFromBackend(formData) {
  const { tickers, fromDate, toDate, columns } = formData;

  if (!tickers || tickers.length === 0 || !fromDate || !toDate || !columns || columns.length === 0) {
    return { success: false, message: "Please provide all required inputs." };
  }

  const backendUrl = "https://excel-addin-backend-o5molvd7pa-el.a.run.app";
  const tickerList = Array.isArray(tickers) ? tickers : tickers.split(',').map(t => t.trim()).filter(t => t.length > 0);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let currentRow = 1; // Start writing from the first row

    for (const ticker of tickerList) {
      const encodedColumns = encodeURIComponent(columns.join(','));
      const queryParams = `?symbol=${encodeURIComponent(ticker)}&from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}&columns=${encodedColumns}`;
      const fullUrl = `${backendUrl}/stocks${queryParams}`;

      Logger.log(`Sending GET request for ${ticker}: ${fullUrl}`);

      try {
        const response = UrlFetchApp.fetch(fullUrl);
        const data = JSON.parse(response.getContentText());

        if (data.length === 0) {
          sheet.getRange(currentRow, 1).setValue(`No data found for ${ticker}.`);
          currentRow += 2; // Leave a blank row
        } else {
          // Write header row
          const headers = Object.keys(data[0]);
          sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
          sheet.getRange(currentRow, 1, 1, headers.length).setFontWeight("bold");
          currentRow++;

          // Write data rows
          const dataRows = data.map(item => headers.map(header => item[header]));
          sheet.getRange(currentRow, 1, dataRows.length, headers.length).setValues(dataRows);
          currentRow += dataRows.length + 2; // Move to the next position, leaving a blank row
        }
      } catch (innerError) {
        // Log and write error for a single ticker, then continue
        Logger.log(`Error fetching data for ${ticker}: ${innerError.message}`);
        sheet.getRange(currentRow, 1).setValue(`Error fetching data for ${ticker}: ${innerError.message}`);
        currentRow += 2;
      }
    }

    return { success: true, message: `Successfully fetched and displayed data for all requested tickers.` };
  } catch (e) {
    // This will catch errors in the initial setup (e.g., getting the sheet)
    return { success: false, message: `An unexpected error occurred: ${e.message}.` };
  }
}
```

Sidebar.html
This file contains the HTML, CSS, and client-side JavaScript for the add-on's user interface.

[Immersive content redacted for brevity.]

3. Excel Add-in (manifest.xml)
The Excel add-in is defined by a manifest file that points to the web-based backend and UI assets.

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="[http://schemas.microsoft.com/office/appforoffice/1.1](http://schemas.microsoft.com/office/appforoffice/1.1)" xmlns:xsi="[http://www.w3.org/2001/XMLSchema-instance](http://www.w3.org/2001/XMLSchema-instance)" xmlns:bt="[http://schemas.microsoft.com/office/officeappbasictypes/1.0](http://schemas.microsoft.com/office/officeappbasictypes/1.0)" xmlns:ov="[http://schemas.microsoft.com/office/taskpaneappversionoverrides](http://schemas.microsoft.com/office/taskpaneappversionoverrides)" xsi:type="TaskPaneApp">
  <Id>c6f9e7a9-6b7c-4a3d-8fba-2d7a6d3b1c9b</Id>
  <Version>1.0.1.0</Version>
  <ProviderName>Plus Percent</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="20 Percent Price Data"/>
  <Description DefaultValue="Advanced Indian stock market analytics with 2,500+ tickers, 100+ technical indicators, and 10+ years of historical data. Superior to Yahoo Finance with institutional-grade analytics."/>
  <IconUrl DefaultValue="[https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/icon-32.png](https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/icon-32.png)"/>
  <HighResolutionIconUrl DefaultValue="[https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/icon-64.png](https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/icon-64.png)"/>
  <SupportUrl DefaultValue="[https://excel-addin-backend-o5molvd7pa-el.a.run.app/help](https://excel-addin-backend-o5molvd7pa-el.a.run.app/help)"/>
  <AppDomains>
    <AppDomain>[https://excel-addin-backend-o5molvd7pa-el.a.run.app](https://excel-addin-backend-o5molvd7pa-el.a.run.app)</AppDomain>
    <AppDomain>[https://storage.googleapis.com](https://storage.googleapis.com)</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="ExcelApi" MinVersion="1.7"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="[https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/taskpane.html](https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/taskpane.html)"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="[http://schemas.microsoft.com/office/taskpaneappversionoverrides](http://schemas.microsoft.com/office/taskpaneappversionoverrides)" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.1">
        <bt:Set Name="ExcelApi" MinVersion="1.7"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>IndianStockDataTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="[https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/icon-16.png](https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/icon-16.png)"/>
        <bt:Image id="Icon.32x32" DefaultValue="[https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/icon-32.png](https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/icon-32.png)"/>
        <bt:Image id="Icon.80x80" DefaultValue="[https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/icon-80.png](https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/assets/icon-80.png)"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="[https://go.microsoft.com/fwlink/?LinkId=276812](https://go.microsoft.com/fwlink/?LinkId=276812)"/>
        <bt:Url id="Commands.Url" DefaultValue="[https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/commands.html](https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/commands.html)"/>
        <bt:Url id="Taskpane.Url" DefaultValue="[https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/taskpane.html](https://storage.googleapis.com/20pluspercentpricedata/pricedataexceladdin/taskpane.html)"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Welcome to Indian Stock Market Analytics!"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="Indian Stock Data"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Get Stock Data"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Access 2,500+ Indian stock tickers with 100+ technical indicators and 10+ years of historical data. Click 'Get Stock Data' to begin advanced market analysis."/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open Indian stock market analytics with NSE/BSE data, technical indicators, and backtesting capabilities"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
