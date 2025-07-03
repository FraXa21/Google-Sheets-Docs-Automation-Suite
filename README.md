# Google Sheets & Docs Automation Suite
This repository contains Google Apps Script functions designed to automate common tasks involving Google Sheets and Google Docs, including data transfer, mail merge for document generation, and PDF combination.

## Features
- **Data Transfer by Header**: Transfers data from a source Google Sheet to a target Google Sheet, matching columns based on their header names (case-insensitive). It also handles date formatting.

- **Mail Merge to Docs & PDFs**: Generates individual Google Docs files and then converts them to PDFs based on a Google Docs template and data from a Google Sheet. It dynamically replaces placeholders in the template with data from each row.

- **Combine PDFs**: Merges all generated PDF files from a specified source folder into a single PDF file, naming it based on data from the spreadsheet.

- **Automated Cleanup**: Automatically deletes the individual generated Google Docs and PDF files after the combination process.

## Setup
To use these scripts, you'll need a Google Account and access to Google Sheets, Google Docs, and Google Drive.

1. **Open Google Apps Script:**
   - Open your Google Sheet (the "source" sheet).
   - Go to Extensions > Apps Script. This will open the Google Apps Script editor.
2. **Create Script Files:**
   - In the Apps Script editor, you'll see a Code.gs file by default. You can paste all the provided code into this file, or create separate .gs files for better organization (e.g., DataTransfer.gs, MailMerge.gs, PDFCombine.gs).
3. **Enable Advanced Google Services:**
   - In the Apps Script editor, on the left sidebar, click on "Services" (the + icon next to "Services").
   - Search for and add the following services:
     - Google Drive API
     - Google Docs API
   - Click "Add".

4. **Prepare Your Google Sheet (Source Data):**
   - Ensure your source Google Sheet (named "Sheet1" by default in the script) contains the data you want to process. The first row should contain your headers.
5. **Prepare Your Google Sheet (Target Data - for transferDataByHeader):**
   - Create or identify the target Google Sheet where data will be transferred.
   - Note its Spreadsheet ID (from the URL: https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit).
   - Ensure the target sheet (named "Sheet1" by default in the script) has headers that match the source sheet's headers for the columns you wish to transfer.
6. **Prepare Your Google Docs Template (for mailMergeToDocsAndPDFs):**
   - Create a Google Docs template file.
   - Use placeholders in the format {{HeaderName}} for each piece of data you want to insert. For example, if you have a column named "Product" in your sheet, use {{Product}} in your template.
   - Note its Document ID (from the URL: https://docs.google.com/document/d/YOUR_DOCUMENT_ID/edit).
7. **Create Google Drive Folders:**
   - Create a Google Drive folder where the individual generated Docs and PDFs will be saved temporarily (e.g., "Mail Merge Output").
   - Create another Google Drive folder where the combined PDF will be saved (e.g., "Combined PDFs").
   - Note the Folder IDs for both (from the URL: https://drive.google.com/drive/folders/YOUR_FOLDER_ID).

**Configuration**
Before running the scripts, you need to update the IDs and names within the code:
- transferDataByHeader function:
  - targetSpreadsheetId: Replace "CHANGE_YOUR_ID" with your target Google Sheet ID.
  - targetSheetName: If your target sheet has a different name, update "Sheet1".
  - dateKeywords: Customize the array ["dated", "tanggal pembelian"] if your date columns have different header names.

- mailMergeToDocsAndPDFs function:
  - templateId: Replace 'CHANGE_YOUR_ID' with your Google Docs template ID.
  - folderId: Replace 'CHANGE_YOUR_ID' with the ID of the folder where individual Docs/PDFs will be temporarily stored.

- combinePDFs function:
  - sourceFolderId: Replace 'CHANGE_YOUR_ID' with the ID of the folder containing the individual PDFs to be combined (this should be the same as folderId in mailMergeToDocsAndPDFs).
  - targetFolderId: Replace 'CHANGE_YOUR_ID' with the ID of the folder where the final combined PDF will be saved.

- getLastDataRow function:
  - keyCol: By default, it checks column 3 (index 2) for data to determine the last row. Adjust 3 if your key column is different.

## Usage
**1. Transfer Data**
To transfer data from your source sheet to the target sheet:
- In the Apps Script editor, select the transferDataByHeader function from the dropdown menu at the top.
- Click the "Run" button (play icon).
- The first time you run it, you will be prompted to authorize the script. Follow the on-screen instructions to grant the necessary permissions.

**2. Perform Mail Merge**
To generate individual documents and PDFs:
- In the Apps Script editor, select the mailMergeToDocsAndPDFs function from the dropdown menu.
- Click the "Run" button.
- This function will automatically call combinePDFs after generating all individual files.

**Helper Function:** getLastDataRow
This function is used internally by transferDataByHeader and mailMergeToDocsAndPDFs to accurately determine the last row containing data in your spreadsheet, preventing the processing of empty rows. It checks a specified keyCol (defaulting to column 3) to find the last non-empty cell.

**Important Notes**
- **Permissions:** When you first run these scripts, Google will ask for permissions to access your Google Sheets, Google Docs, and Google Drive. Grant these permissions for the scripts to function correctly.
- **PDF-lib:** The combinePDFs function uses the pdf-lib library, which is loaded directly from a CDN using UrlFetchApp.fetch().getContentText(). This is a common practice in Google Apps Script for external libraries but be aware of potential security implications if the CDN source is compromised (though cdnjs.cloudflare.com is generally reliable).
- **Rate Limits:** Be mindful of Google Apps Script execution limits and Google API rate limits, especially if you are processing a very large number of rows or documents. For extremely large datasets, you might need to implement batch processing or introduce delays.
- **Cleanup:** The combinePDFs function includes a cleanup step that trashes the individual Docs and PDFs created during the mail merge process from the sourceFolderId. This helps keep your Drive organized.
- **Date Formats:** The script handles date formatting for specific headers (dated, tanggal pembelian). Ensure your date columns in the source sheet are actual Date objects for proper formatting.

