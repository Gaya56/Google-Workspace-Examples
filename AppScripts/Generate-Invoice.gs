/**
 * Copyright 2025 Peter Horner
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * This script was created by Peter Horner.
 * Check out my YouTube channel for more Apps Script tutorials:
 * https://www.youtube.com/@PeterHornerGoogleTech 
 */

function onOpen(){
  const ui = SpreadsheetApp.getUi()
  ui.createMenu('Mail Merge')
    .addItem('Run Mail Merge', 'mailMerge')
    .addToUi()
}

function copyDocument(sourceDocId, copyName, folderId) {
  const sourceDoc = DriveApp.getFileById(sourceDocId)
  const folder = DriveApp.getFolderById(folderId)

  // Make a copy of the document
  const copiedDoc = sourceDoc.makeCopy(copyName, folder)

  return copiedDoc.getId()
}

function replacePlaceholder(docId, replacements) {
  const doc = DocumentApp.openById(docId)
  const body = doc.getBody()

  for (var key in replacements) {
    body.replaceText("{" + key + "}", replacements[key])
  }
  doc.saveAndClose(); 

}

function mailMerge() {
  // Get the active spreadsheet and the data sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const sheet = ss.getSheetByName('Sheet1'); 

  // Get all data from the sheet, separating headers from the rest
  const [head, ...data] = sheet.getDataRange().getValues(); 

  // Define the source document ID and the destination folder ID
  const sourceDocId = '<Your Template Doc ID>'; // Update with your source doc ID
  const folderId = '<Your Destination Folder>'; // Update with your folder ID
  const destinationFolder = DriveApp.getFolderById(folderId);

  const statusColIndex = head.indexOf('Status') + 1
  const linkColIndex = head.indexOf('Invoice Link') + 1


  // Log the data to the console (for debugging)
  console.log(data); 

  // Iterate over each row of data
  data.forEach((row, i) => { 
    const currentRow = i + 2; 

    const currentStatus = sheet.getRange(currentRow, statusColIndex).getValue();
    if (currentStatus) {
      return
    }
    // Create a unique name for each copied document
    const copyName = 'invoice' + (i+1); 
    const invoiceNumber = row[head.indexOf('InvoiceNumber')]
    const clientEmail = row[head.indexOf('ClientEmail')]
    const clientName = row[head.indexOf('ClientName')]


    // Copy the source document to the destination folder
    const newDoc = copyDocument(sourceDocId, copyName, folderId); 
    console.log(i,newDoc)

    // Create an object with placeholders and their corresponding values from the current row
    const replacements = {
      "ClientName": row[head.indexOf('ClientName')],
      "ClientEmail": row[head.indexOf('ClientEmail')],
      "InvoiceNumber": row[head.indexOf('InvoiceNumber')],
      "InvoiceDate": row[head.indexOf('InvoiceDate')].toLocaleDateString(), // Formats the date
      "DueDate": row[head.indexOf('DueDate')].toLocaleDateString(), // Formats the date
      "ItemDescription": row[head.indexOf('ItemDescription')],
      "Amount": row[head.indexOf('Amount')].toFixed(2) // Formats the amount to 2 decimal places
    };

    // Replace the placeholders in the new document with the values from the replacements object
    replacePlaceholder(newDoc, replacements); 
    
    // Create a PDF blob from the Google Doc
    const newDocFile = DriveApp.getFileById(newDoc);
    
    // Create the final PDF file in the destination folder
    const pdfFile = destinationFolder.createFile(newDocFile.getAs('application/pdf')).setName(copyName + ".pdf");

    // Define the email subject and body
    const subject = `Invoice from Your Company: #${invoiceNumber}`;
    const emailBody = `Hi ${clientName},\n\nPlease find your invoice #${invoiceNumber} attached.\n\nThank you!\n\nBest regards,\nYour Name`;

    // Send the email with the PDF as an attachment
    GmailApp.sendEmail(clientEmail, subject, emailBody, {
      attachments: [pdfFile],
      name: 'Your Company Name' // Optional: Sets the sender's name
    });

    sheet.getRange(currentRow, statusColIndex).setValue(`Emailed on ${new Date().toLocaleDateString()}`);

    sheet.getRange(currentRow, linkColIndex).setValue(pdfFile.getUrl());
  });

  SpreadsheetApp.getUi().alert('Success!', 'All invoices have been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}




