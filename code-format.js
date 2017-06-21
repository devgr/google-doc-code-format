// In Google Docs, go to the Tools menu > Script Editor, and paste in this code.
// Author: George Darling - Decision Source Inc.
// License: MIT

function onOpen(){
  var ui = DocumentApp.getUi();
  ui.createMenu('Format Code')
  .addItem('Format selection as code', 'setStyleCode')
  .addItem('Format `...` as code', 'replaceMarkdown')
  .addItem('Format ``` code blocks', 'replaceBlockCode')
  .addToUi();
}

function setStyleCode(){
  var selection = DocumentApp.getActiveDocument().getSelection();
  if(selection) {
    var elements = selection.getRangeElements();
    for(var i = 0; i < elements.length; i++) {
      var rangeElem = elements[i];
      
      // Only modify elements that can be edited as text; skip images and other non-text elements.
      var elem = rangeElem.getElement();
      if(elem.editAsText) {
        if(rangeElem.isPartial()) {
          formatElement(elem, rangeElem.getStartOffset(), rangeElem.getEndOffsetInclusive());
        } else {
          formatElement(elem);
        }
      }
    }
  }
}

function replaceMarkdown() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  for(var i = 0, len = paragraphs.length; i < len; i++) {
    var paragraph = paragraphs[i];
    var match = paragraph.findText('`[^`].*?`'); // match `anything`
    while(match !== null) {
      var elem = match.getElement();
      if (elem.editAsText) { 
        // There are lots of types of elements, only some implement editAsText
        // https://developers.google.com/apps-script/reference/document/element
        var text = elem.editAsText();
        var start = match.getStartOffset();
        text.deleteText(start, start);
        var end = match.getEndOffsetInclusive();
        text.deleteText(end-1, end-1);
        formatElement(elem, start, end-2);
      }
      match = paragraph.findText('`[^`].*?`', match);
    }
  }
}

function formatElement(textElement, optStart, optEnd) {
  var text = textElement.editAsText();
  
  // Edit the selected part of the element, or the full element if it's completely selected.
  if(optStart !== undefined && optEnd !== undefined) {
    text.setFontFamily(optStart, optEnd, 'Consolas');
    text.setFontSize(optStart, optEnd, 10);
    text.setBackgroundColor(optStart, optEnd, '#efefef');
  } else {
    text.setFontFamily('Consolas');
    text.setFontSize(10);
    text.setBackgroundColor('#efefef');
  }
}

function replaceBlockCode() {
  var body = DocumentApp.getActiveDocument().getBody();
  var allText = body.editAsText();
  var giantStr = allText.getText();
  var findExp = /```\n([\s\S]*?)\n```/gm; // match everything between ``` and ```
  var result = null;
  var matches = 0;
  while(result = findExp.exec(giantStr)) {
    // grab the block of code as a normal string
    var codeStr = result[1]; // just the code, not the ```
    codeStr = codeStr.replace(/\n/g, '\r'); // in Google Docs, \n is a new paragraph, \r is a new line
    // remove that block from the document
    var indexCorrection = matches * 8; // for the ```\n\n``` that was removed
    allText.deleteText(result.index - indexCorrection, findExp.lastIndex - 1 - indexCorrection);
    // add it in as a single paragraph
    allText.insertText(result.index - indexCorrection, codeStr);
    
    // find the paragraph that was just added. kind of messy, though I can't find a better way.
    var rangeElem = null;
    while (rangeElem = body.findElement(DocumentApp.ElementType.PARAGRAPH, rangeElem)) {
      if(rangeElem.getElement().getText() === codeStr) {
        rangeElem.getElement().setLineSpacing(1); // change to single spacing rather than 1.15
        break;
      }
    }
    // then format it
    formatElement(allText, result.index - indexCorrection, result.index + codeStr.length - 1 - indexCorrection);
    matches++;
  }
}

function clearLogs() {
  Logger.clear();
}