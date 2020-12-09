/*
  # Merge Document and Spreadsheet

  This code will create a copy of a document and copy the template of the original document for each row found in the sheet.
  For each row the {placeholders} will be replaced with the values from the sheet.
  In the document there will be a page break between each template.
  
  Some prerequisites:
  - Placeholder keys should correspond to the column key/values in the first row of the sheet
  - To make sure everything goes fine, provide a clean spreadsheet: one sheet, delete obsolete rows and columns, first row = header
  - This script can probably only handle paragraphs (no images, tables, lists, etc...)
  
  #1 | name   | address   |
  #2 | Jeroen | Terlinden |
  
  Placeholders you can use in the document: {name} and {address}
  
*/

var sheetId = '1wc3B1CInln1H5kwoBRwXWk3CvLwn3gzEPNtxJv7FXFg';

function mergeDocumentAndSheet() {
  var doc = DocumentApp.getActiveDocument();
  var originalBody = doc.getBody();
  var originalBodyParagraphs = originalBody.getParagraphs();
  
  var newDocument = createIdenticalCopy(doc, getDataSetName());
  newDocumentBody = newDocument.getBody();
  
  var data = getDataSet();

  data.forEach(function createThemAll(row, index) {
    var isFirst = index === 0;
    var newCopy = originalBodyParagraphs.map(function createNewParagraph(originalParagraph) {
      return fillTemplate(originalParagraph.copy(), row);
    });
    appendToDocument(newDocumentBody, newCopy, isFirst);
  });
  
}

function fillTemplate(newParagraph, row) {
  Object.keys(row).forEach(function replacePlaceholder(placeholderKey) {
    newParagraph.replaceText(toPlaceholder(placeholderKey), row[placeholderKey]);
  });

  return newParagraph;
}

function appendToDocument(body, newCopy, isFirst) {
  if (!isFirst) body.appendPageBreak();
  newCopy.forEach(function addParagraphs(paragraph) {
    body.appendParagraph(paragraph);
  });
}

/*
  Create and return a new document with the same layout as the original
*/
function createIdenticalCopy(originalDoc, mergedDataSetName) {
  var newDocument = DocumentApp.create(originalDoc.getName() + ' - ' + mergedDataSetName);
  
  var height = originalDoc.getBody().getPageHeight();
  var width = originalDoc.getBody().getPageWidth();
  
  var marginTop = originalDoc.getBody().getMarginTop();
  var marginRight = originalDoc.getBody().getMarginRight();
  var marginBottom = originalDoc.getBody().getMarginBottom();
  var marginLeft = originalDoc.getBody().getMarginLeft();
  
  newDocument.getBody().setPageHeight(height);
  newDocument.getBody().setPageWidth(width);
  
  newDocument.getBody().setMarginTop(marginTop)
  newDocument.getBody().setMarginRight(marginRight)
  newDocument.getBody().setMarginBottom(marginBottom)
  newDocument.getBody().setMarginLeft(marginLeft)
  
  return newDocument;
}

function getDataSetName() {
  return SpreadsheetApp.openById(sheetId).getName();
}

/*
  Return Array<Object> of all rows in the sheet where the keys of the object are the keys used for the placeholders

  return [{
    "name": "Jeroen",
    "address": "Terlinden"
  }]
*/
function getDataSet() {
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadsheet.getActiveSheet();
  var rangeData = sheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow(); 
  var dataRange = sheet.getRange(2, 1, lastRow-1, lastColumn);
  var rangeValues = dataRange.getValues();
  
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  
  var dataSet = rangeValues.map(function createDataSet(row) {
    var obj = {};
    
    headers.forEach(function addValuePerHeader(header, index) {
      obj[toKey(header)] = row[index];
    });
    
    return obj;
  });
  
  return dataSet;
}

function toKey(string) {
  return string.toLowerCase();
}

function toPlaceholder(placeholderKey) {
  return '{{' + placeholderKey + '}}';
}
