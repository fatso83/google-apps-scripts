// Ripped from some stackoverflow answer and modified to add tables around the code
// Usage: start the script editor in Google Docs and copy-paste this script
//        On subsequent loads of the document you will have an extra menu called Extras
//        with a menu entry for formatting code

// is called by google docs when a document is open
// adds a menu with a menu item that applies a style to the currently selected text
function onOpen() {
  DocumentApp.getUi()
  .createMenu('Extras')
  .addItem('Apply code style', 'applyCodeStyle')
  .addToUi();
}

var backgroundColor =  "#DDDDDD";

// definition of a style to be applied
var style = {
  bold: false,
  backgroundColor: backgroundColor,
  fontFamily: DocumentApp.FontFamily.CONSOLAS,
  fontSize: 9
};

// helper function that strips the selected element and passes it to a handler
function withElement(processPartial, processFull) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      if (element.getElement().editAsText) {
        var text = element.getElement().editAsText();
        var trueElement =  element.getElement();
        var parent = trueElement.getParent();
        
        Logger.log("parent: " + parent.getType());
        Logger.log("Type of element:" + trueElement.getType());

        if (element.isPartial()) {
          var from = element.getStartOffset();
          var to = element.getEndOffsetInclusive();
          processPartial(element, text, from, to);
        } else {          
          processFull(element, text);
          
          if(trueElement.getType() === DocumentApp.ElementType.PARAGRAPH && parent.getType() === DocumentApp.ElementType.BODY_SECTION){
            // insert table
            var index = parent.getChildIndex(trueElement);
            Logger.log("Index of element: " + index);
            var removed = parent.removeChild(trueElement);

            var newTable = parent.insertTable(index, [[trueElement.getText()]]);
            var tableStyle = {};
            tableStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
            tableStyle[DocumentApp.Attribute.FONT_FAMILY] = style.fontFamily;
            tableStyle[DocumentApp.Attribute.FONT_SIZE] = style.fontSize;

            newTable.setAttributes(tableStyle);
            
            var cell0 = newTable.getChild(0).getChild(0);
            Logger.log(cell0 + ": " + cell0.getBackgroundColor());
            cell0.setBackgroundColor(style.backgroundColor);
          }

        }
      }
    }
  }
}

// called in response to the click on a menu item
function applyCodeStyle() {
  return withElement(
    applyPartialStyle.bind(this, style),
    applyFullStyle.bind(this, style)
  );
}

// applies the style to a selected text range
function applyPartialStyle(style, element, text, from, to) {
  text.setFontFamily(from, to, style.fontFamily);
  text.setBackgroundColor(from, to, style.backgroundColor);
  text.setBold(from, to, style.bold);
  text.setFontSize(from, to, style.fontSize);
}

// applies the style if the entire element is selected
function applyFullStyle(style, element, text) {
  text.setFontFamily(style.fontFamily);
  text.setBackgroundColor(style.backgroundColor);
  text.setBold(style.bold); 
  text.setFontSize(style.fontSize);
}

