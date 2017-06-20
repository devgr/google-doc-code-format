// In Google Docs, go to the Tools menu > Script Editor, and paste in this code.
// Author: George Darling
// License: MIT

function onOpen(){
  var ui = DocumentApp.getUi();
  ui.createMenu('Format Code')
  .addItem('Format selection as code', 'setStyleCode')
  .addItem('Format `...` as code', 'replaceMarkdown')
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