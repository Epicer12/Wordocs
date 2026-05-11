/**
 * List Generator
 * Generates lists of tables and figures with bookmark links
 */

var ListGenerator = (function () {

  /**
   * Generates a list of all tables in the document
   * Clears existing list if found, then creates new one with links
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

      // Find all table captions with their bookmarks
      var tableCaptions = findCaptionsWithBookmarks('Table');

      if (tableCaptions.length === 0) {
        return {
          success: false,
          message: 'No table captions found in the document.'
        };
      }

      // Check if "List of Tables" already exists and remove it
      clearExistingList('List of Tables');

      // Insert at cursor position
      var element = cursor.getElement();
      var parent = element.getParent();
      var insertIndex = parent.getChildIndex(element);

      // Insert heading
      var heading = parent.insertParagraph(insertIndex, 'List of Tables');
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      insertIndex++;

      // Insert each table caption as a linked entry
      for (var i = 0; i < tableCaptions.length; i++) {
        var caption = tableCaptions[i];
        var listItem = parent.insertParagraph(insertIndex, '');

        // Add the caption text
        var text = listItem.editAsText();
        text.appendText(caption.text);

        // If there's a bookmark, make it a link
        if (caption.bookmarkId) {
          var linkUrl = '#bookmark=' + caption.bookmarkId;
          text.setLinkUrl(0, caption.text.length - 1, linkUrl);
        }

        // Format the list item
        listItem.setIndentFirstLine(36);

        insertIndex++;
      }

      // Add blank line after the list
      parent.insertParagraph(insertIndex, '');

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
   * Clears existing list if found, then creates new one with links
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

      // Find all figure captions with their bookmarks
      var figureCaptions = findCaptionsWithBookmarks('Figure');

      if (figureCaptions.length === 0) {
        return {
          success: false,
          message: 'No figure captions found in the document.'
        };
      }

      // Check if "List of Figures" already exists and remove it
      clearExistingList('List of Figures');

      // Insert at cursor position
      var element = cursor.getElement();
      var parent = element.getParent();
      var insertIndex = parent.getChildIndex(element);

      // Insert heading
      var heading = parent.insertParagraph(insertIndex, 'List of Figures');
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      insertIndex++;

      // Insert each figure caption as a linked entry
      for (var i = 0; i < figureCaptions.length; i++) {
        var caption = figureCaptions[i];
        var listItem = parent.insertParagraph(insertIndex, '');

        // Add the caption text
        var text = listItem.editAsText();
        text.appendText(caption.text);

        // If there's a bookmark, make it a link
        if (caption.bookmarkId) {
          var linkUrl = '#bookmark=' + caption.bookmarkId;
          text.setLinkUrl(0, caption.text.length - 1, linkUrl);
        }

        // Format the list item
        listItem.setIndentFirstLine(36);

        insertIndex++;
      }

      // Add blank line after the list
      parent.insertParagraph(insertIndex, '');

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
   * Clears an existing list by heading name
   * @param {string} headingText - The heading to search for (e.g., "List of Tables")
   */
  function clearExistingList(headingText) {
    try {
      var doc = DocumentApp.getActiveDocument();
      var body = doc.getBody();
      var paragraphs = body.getParagraphs();

      var foundHeading = false;
      var itemsToRemove = [];

      for (var i = 0; i < paragraphs.length; i++) {
        var para = paragraphs[i];
        var text = para.getText();

        // Check if this is the heading we're looking for
        if (text === headingText && para.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
          foundHeading = true;
          itemsToRemove.push(para);
          continue;
        }

        // If we found the heading, remove subsequent list items
        if (foundHeading) {
          // Stop at the next heading or empty line followed by non-indented text
          if (para.getHeading() !== DocumentApp.ParagraphHeading.NORMAL ||
            (text.trim() === '' && itemsToRemove.length > 1)) {
            // Found the end of the list
            break;
          }

          // Remove list items (indented paragraphs)
          if (para.getIndentFirstLine() > 0 || text.trim() === '') {
            itemsToRemove.push(para);
          } else if (text.trim() !== '') {
            // Found non-indented text, stop here
            break;
          }
        }
      }

      // Remove all collected items
      for (var j = 0; j < itemsToRemove.length; j++) {
        itemsToRemove[j].removeFromParent();
      }

    } catch (error) {
      Logger.log('Error in clearExistingList: ' + error);
    }
  }

  /**
   * Finds all captions of a specific type with their bookmark IDs
   * @param {string} prefix - The caption prefix (e.g., "Table" or "Figure")
   * @return {Array<Object>} Array of caption objects with text and bookmarkId
   */
  function findCaptionsWithBookmarks(prefix) {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    var captions = [];
    var bookmarks = doc.getBookmarks();

    // Create a map of position to bookmark ID
    var bookmarkMap = {};
    for (var i = 0; i < bookmarks.length; i++) {
      var bookmark = bookmarks[i];
      var position = bookmark.getPosition();
      var element = position.getElement();

      // Get the paragraph containing this bookmark
      var para = element.getType() === DocumentApp.ElementType.PARAGRAPH ?
        element.asParagraph() :
        element.getParent().asParagraph();

      if (para) {
        bookmarkMap[para.getText()] = bookmark.getId();
      }
    }

    var captionRegex = new RegExp('^' + prefix + '\\s+\\d+:', 'i');

    for (var j = 0; j < paragraphs.length; j++) {
      var text = paragraphs[j].getText();

      if (captionRegex.test(text)) {
        captions.push({
          text: text,
          bookmarkId: bookmarkMap[text] || null
        });
      }
    }

    return captions;
  }

  /**
   * Finds all captions of a specific type (legacy - for backward compatibility)
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
    findCaptionsWithBookmarks: findCaptionsWithBookmarks,
    searchAndReplace: searchAndReplace
  };

})();