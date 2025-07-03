function transferDataByHeader() {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const targetSpreadsheetId = "CHANGE_YOUR_ID";
  const targetSheetName = "Sheet1";
  const targetSheet = SpreadsheetApp.openById(targetSpreadsheetId).getSheetByName(targetSheetName);

  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];

  const headerMap = {};
  sourceHeaders.forEach((header, idx) => {
    const lowerHeader = header.toLowerCase();
    let targetIdx = targetHeaders.map(h => h.toLowerCase()).indexOf(lowerHeader);
    if (targetIdx !== -1) {
      headerMap[idx] = targetIdx;
    }
  });

  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
  }

  // Deteksi kolom tanggal di target
  const dateFormat = "d-MMM-yy";
  const dateKeywords = ["dated", "tanggal pembelian"];
  const dateTargetIndices = targetHeaders
    .map((h, i) => ({ h: h.toLowerCase(), i }))
    .filter(obj => dateKeywords.includes(obj.h))
    .map(obj => obj.i);

  const lastDataRow = getLastDataRow(sourceData);

  for (let i = 1; i <= lastDataRow; i++) {
    const sourceRow = sourceData[i];
    if (sourceRow.join("").trim() === "" || !sourceRow[0]) {
      continue;
    }

    const targetRow = new Array(targetHeaders.length).fill("");

    for (let sourceIdx in headerMap) {
      const idx = parseInt(sourceIdx);
      const targetIdx = headerMap[idx];
      let value = sourceRow[idx];

      if (dateTargetIndices.includes(targetIdx) && value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), dateFormat);
      }

      targetRow[targetIdx] = value;
    }

    targetSheet.appendRow(targetRow);
  }
}

function mailMergeToDocsAndPDFs() {
  const templateId = 'CHANGE_YOUR_ID'; // ID Google Docs template
  const folderId = 'CHANGE_YOUR_ID'; // Folder untuk output individual docs & pdf
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const outputFolder = DriveApp.getFolderById(folderId);
  const lastDataRow = getLastDataRow(data)
  for (let i = 1; i <= lastDataRow; i++) {
    const row = data[i];
    if (row.join("").trim()===""||!row[0]){
      continue;
    }
    const newFile = DriveApp.getFileById(templateId).makeCopy(`Ketentuan Smartphone - ${row[0]}`, outputFolder);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();
    const header = doc.getHeader();
    const footer = doc.getFooter();

    headers.forEach((headerText, j) => {
      const placeholder = `{{${headerText}}}`;
      let value = row[j];

      if (headerText.toLowerCase() === "dated" && value instanceof Date) {
        const day = value.getDate();
        const month = value.getMonth() + 1;
        const year = value.getFullYear();
        value = `${month}/${day}/${year}`;
      }

      body.replaceText(placeholder, value);
      if (header) header.replaceText(placeholder, value);
      if (footer) footer.replaceText(placeholder, value);
    });

    doc.saveAndClose();

    
    const pdf = DriveApp.getFileById(newFile.getId()).getAs(MimeType.PDF);
    outputFolder.createFile(pdf).setName(`Ketentuan Smartphone - ${row[0]}.pdf`);
  }

  combinePDFs();
}

function getLastDataRow (data,keyCol = 3){
  for(let i= data.length-1;i>0;i--){
    if(data[i][keyCol]&&data[i]
    [keyCol].toString().trim()!==""){
      return i;
    }
  }
  return 1;
}

async function combinePDFs() {
  const sourceFolderId = 'CHANGE_YOUR_ID'; // Folder berisi PDF hasil mail merge
  const targetFolderId = 'CHANGE_YOUR_ID'; // Folder untuk menyimpan hasil gabungan PDF

  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  const files = sourceFolder.getFiles();

  // Ambil nomor urut awal dan akhir dari spreadsheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();
  const firstNumber = data[1][20]; // baris kedua (index 1), kolom pertama (index 0)
  const lastNumber = data[data.length - 2][20]; // baris terakhir

  // Gunakan nama sesuai format diminta
  const mergedFileName = `Form Serah Terima HP Samsung ${firstNumber} - ${lastNumber}.pdf`;

  // Load PDF-lib dari CDN
  const cdnjs = "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js";
  eval(UrlFetchApp.fetch(cdnjs).getContentText());
  const setTimeout = function(f, t) {
    Utilities.sleep(t);
    return f();
  }

  const pdfDoc = await PDFLib.PDFDocument.create();

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === 'application/pdf') {
      const pdfData = await PDFLib.PDFDocument.load(new Uint8Array(file.getBlob().getBytes()));
      const pages = await pdfDoc.copyPages(pdfData, [...Array(pdfData.getPageCount())].map((_, i) => i));
      pages.forEach(page => pdfDoc.addPage(page));
    }
  }

  const bytes = await pdfDoc.save();
  const mergedFile = targetFolder.createFile(Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, mergedFileName));

  // Hapus semua file PDF dan DOC dari sourceFolder
const cleanupFiles = sourceFolder.getFiles();
while (cleanupFiles.hasNext()) {
  const file = cleanupFiles.next();
  const mimeType = file.getMimeType();
  if (
    mimeType === 'application/pdf' || 
    mimeType === 'application/vnd.google-apps.document'
  ) {
    file.setTrashed(true);
  }
}

  // Buka otomatis file hasil merge
  const mergedUrl = mergedFile.getUrl();
  const html = HtmlService.createHtmlOutput(`<script>window.open('${mergedUrl}', '_blank');google.script.host.close();</script>`)
    .setWidth(100).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, "Membuka File Hasil Merge...");
}
