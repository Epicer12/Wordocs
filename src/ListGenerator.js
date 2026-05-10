/**
 * List Generator
 * Generates lists of tables and figures
 */

var ListGenerator = (function() {
  
  /**
   * Generates a list of all tables in the document
   * @return {Object} Result object with success status and message
   */
  function generateListOfTables() {
    try {
      var doc = DocumentApp.getActiveDocument();
      var body = doc.getBody();
      var cursor = doc.getCursor();
      
      if (!cursor) {
        return {
          success: false,
          message: 'Please place your cursor where you want to insert the list.'
        };
      }
      
      // Find all table captions
      var tableCaptions = findCaptions('Table');
      
      if (tableCaptions.length === 0) {
        return {
          success: false,
          message: 'No table captions found in the document.'
        };
      }
      
      // Insert heading
      var element = cursor.getElement();
      var insertIndex = element.getParent().getChildIndex(element);
      var parent = element.getParent();
      
      var heading = parent.insertParagraph(insertIndex, 'List of Tables');
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      
      // Insert each table caption
      for (var i = 0; i < tableCaptions.length; i++) {
        var captionPara = parent.insertParagraph(insertIndex + i + 1, tableCaptions[i]);
        captionPara.setIndentFirstLine(36); // Indent for better formatting
      }
      
      // Add blank line after the list
      parent.insertParagraph(insertIndex + tableCaptions.length + 1, '');
      
      return {
        success: true,
        message: 'List of Tables inserted successfully with ' + tableCaptions.length + ' item(s).'
      };
      
    } catch (error) {
      Logger.log('Error in generateListOfTables: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }
  
  /**
   * Generates a list of all figures in the document
   * @return {Object} Result object with success status and message
   */
  function generateListOfFigures() {
    try {
      var doc = DocumentApp.getActiveDocument();
      var body = doc.getBody();
      var cursor = doc.getCursor();
      
      if (!cursor) {
        return {
          success: false,
          message: 'Please place your cursor where you want to insert the list.'
        };
      }
      
      // Find all figure captions
      var figureCaptions = findCaptions('Figure');
      
      if (figureCaptions.length === 0) {
        return {
          success: false,
          message: 'No figure captions found in the document.'
        };
      }
      
      // Insert heading
      var element = cursor.getElement();
      var insertIndex = element.getParent().getChildIndex(element);
      var parent = element.getParent();
      
      var heading = parent.insertParagraph(insertIndex, 'List of Figures');
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      
      // Insert each figure caption
      for (var i = 0; i < figureCaptions.length; i++) {
        var captionPara = parent.insertParagraph(insertIndex + i + 1, figureCaptions[i]);
        captionPara.setIndentFirstLine(36); // Indent for better formatting
      }
      
      // Add blank line after the list
      parent.insertParagraph(insertIndex + figureCaptions.length + 1, '');
      
      return {
        success: true,
        message: 'List of Figures inserted successfully with ' + figureCaptions.length + ' item(s).'
      };
      
    } catch (error) {
      Logger.log('Error in generateListOfFigures: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }
  
  /**
   * Finds all captions of a specific type in the document
   * @param {string} prefix - The caption prefix (e.g., "Table" or "Figure")
   * @return {Array<string>} Array of caption texts
   */
  function findCaptions(prefix) {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    var captions = [];
    
    var captionRegex = new RegExp('^' + prefix + '\\s+\\d+:', 'i');
    
    for (var i = 0; i < paragraphs.length; i++) {
      var text = paragraphs[i].getText();
      if (captionRegex.test(text)) {
        captions.push(text);
      }
    }
    
    return captions;
  }
  
  /**
   * Searches and replaces text in the document
   * Used for updating cross-references
   * @param {string} searchPattern - Text to search for
   * @param {string} replacement - Text to replace with
   * @return {number} Number of replacements made
   */
  function searchAndReplace(searchPattern, replacement) {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    
    var searchResult = body.findText(searchPattern);
    var count = 0;
    
    while (searchResult !== null) {
      var foundElement = searchResult.getElement();
      var start = searchResult.getStartOffset();
      var end = searchResult.getEndOffsetInclusive();
      
      foundElement.asText().deleteText(start, end);
      foundElement.asText().insertText(start, replacement);
      
      count++;
      searchResult = body.findText(searchPattern, searchResult);
    }
    
    return count;
  }
  
  // Public API
  return {
    generateListOfTables: generateListOfTables,
    generateListOfFigures: generateListOfFigures,
    findCaptions: findCaptions,
    searchAndReplace: searchAndReplace
  };
  
})();
