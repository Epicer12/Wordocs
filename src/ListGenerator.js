/**
 * List Generator
 * Generates lists of tables and figures with bookmark links
 */

var ListGenerator = (function () {

  var LIST_HEADINGS = ['List of Tables', 'List of Figures'];

  /**
   * Escapes special regex characters in a string
   * @param {string} str
   * @return {string}
   */
  function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  /**
   * Builds a caption regex for a given prefix
   * @param {string} prefix
   * @return {RegExp}
   */
  function buildCaptionRegex(prefix) {
    return new RegExp('^' + escapeRegex(prefix) + '\\s+\\d+:', 'i');
  }

  /**
   * Determines if a paragraph marks the end of a list block
   * @param {Paragraph} para
   * @param {string} text
   * @return {boolean}
   */
  function isEndOfListBlock(para, text) {
    var heading = para.getHeading();
    if (heading === DocumentApp.ParagraphHeading.HEADING1 ||
      heading === DocumentApp.ParagraphHeading.HEADING2) {
      return true;
    }
    if (text.trim() !== '' && para.getIndentFirstLine() === 0) {
      return true;
    }
    return false;
  }

  /**
   * Returns a map of paragraph indices that belong inside list-of blocks
   * @param {Array<Paragraph>} paragraphs
   * @return {Object<number, boolean>}
   */
  function getListBlockIndices(paragraphs) {
    var indices = {};

    for (var h = 0; h < LIST_HEADINGS.length; h++) {
      var headingText = LIST_HEADINGS[h];
      var inBlock = false;

      for (var i = 0; i < paragraphs.length; i++) {
        var para = paragraphs[i];
        var text = para.getText();

        if (text === headingText &&
          para.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
          inBlock = true;
          indices[i] = true;
          continue;
        }

        if (inBlock) {
          if (isEndOfListBlock(para, text)) {
            break;
          }
          indices[i] = true;
        }
      }
    }

    return indices;
  }

  /**
   * Checks if a paragraph is a real caption (not a list row or list heading)
   * @param {Paragraph} paragraph
   * @param {string} prefix
   * @param {Object<number, boolean>} listBlockIndices
   * @param {number} paragraphIndex
   * @return {boolean}
   */
  /**
   * Resolves a cursor to a container that supports insertParagraph (Body, TableCell, etc.)
   * @param {Position} cursor
   * @return {Object|null} { container, index } or null
   */
  function resolveCursorInsertLocation(cursor) {
    var current = cursor.getElement();
    var paragraph = null;

    while (current) {
      if (current.getType() === DocumentApp.ElementType.PARAGRAPH) {
        paragraph = current.asParagraph();
        break;
      }
      current = current.getParent();
    }

    if (!paragraph) {
      return null;
    }

    var container = paragraph.getParent();
    while (container && typeof container.insertParagraph !== 'function') {
      paragraph = container;
      container = container.getParent();
    }

    if (!container) {
      return null;
    }

    return {
      container: container,
      index: container.getChildIndex(paragraph)
    };
  }

  function isRealCaptionParagraph(paragraph, prefix, listBlockIndices, paragraphIndex) {
    var text = paragraph.getText();
    var captionRegex = buildCaptionRegex(prefix);

    if (!captionRegex.test(text)) {
      return false;
    }
    if (listBlockIndices[paragraphIndex]) {
      return false;
    }
    if (paragraph.getIndentFirstLine() > 0) {
      return false;
    }
    return true;
  }

  /**
   * Generates a list of all tables in the document
   * Clears existing list if found, then creates new one with links
   * @return {Object} Result object with success status and message
   */
  function generateListOfTables() {
    try {
      var doc = DocumentApp.getActiveDocument();
      var cursor = doc.getCursor();
      var prefix = CaptionManager.getTablePrefix();

      if (!cursor) {
        return {
          success: false,
          message: 'Please place your cursor where you want to insert the list.'
        };
      }

      clearExistingList('List of Tables');

      var tableCaptions = findCaptionsWithBookmarks(prefix);

      if (tableCaptions.length === 0) {
        return {
          success: false,
          message: 'No table captions found in the document.'
        };
      }

      var insertLocation = resolveCursorInsertLocation(cursor);
      if (!insertLocation) {
        return {
          success: false,
          message: 'Could not determine where to insert the list. Place your cursor in the document body.'
        };
      }

      var parent = insertLocation.container;
      var insertIndex = insertLocation.index;

      var heading = parent.insertParagraph(insertIndex, 'List of Tables');
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      insertIndex++;

      for (var i = 0; i < tableCaptions.length; i++) {
        var caption = tableCaptions[i];
        var listItem = parent.insertParagraph(insertIndex, '');

        var text = listItem.editAsText();
        text.appendText(caption.text);

        if (caption.bookmarkId) {
          var linkUrl = '#bookmark=' + caption.bookmarkId;
          text.setLinkUrl(0, caption.text.length - 1, linkUrl);
        }

        listItem.setIndentFirstLine(36);
        insertIndex++;
      }

      parent.insertParagraph(insertIndex, '');

      return {
        success: true,
        count: tableCaptions.length,
        message: 'List of Tables updated with ' + tableCaptions.length + ' item(s).'
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
      var cursor = doc.getCursor();
      var prefix = CaptionManager.getFigurePrefix();

      if (!cursor) {
        return {
          success: false,
          message: 'Please place your cursor where you want to insert the list.'
        };
      }

      clearExistingList('List of Figures');

      var figureCaptions = findCaptionsWithBookmarks(prefix);

      if (figureCaptions.length === 0) {
        return {
          success: false,
          message: 'No figure captions found in the document.'
        };
      }

      var insertLocation = resolveCursorInsertLocation(cursor);
      if (!insertLocation) {
        return {
          success: false,
          message: 'Could not determine where to insert the list. Place your cursor in the document body.'
        };
      }

      var parent = insertLocation.container;
      var insertIndex = insertLocation.index;

      var heading = parent.insertParagraph(insertIndex, 'List of Figures');
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      insertIndex++;

      for (var i = 0; i < figureCaptions.length; i++) {
        var caption = figureCaptions[i];
        var listItem = parent.insertParagraph(insertIndex, '');

        var text = listItem.editAsText();
        text.appendText(caption.text);

        if (caption.bookmarkId) {
          var linkUrl = '#bookmark=' + caption.bookmarkId;
          text.setLinkUrl(0, caption.text.length - 1, linkUrl);
        }

        listItem.setIndentFirstLine(36);
        insertIndex++;
      }

      parent.insertParagraph(insertIndex, '');

      return {
        success: true,
        count: figureCaptions.length,
        message: 'List of Figures updated with ' + figureCaptions.length + ' item(s).'
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

        if (text === headingText &&
          para.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
          foundHeading = true;
          itemsToRemove.push(para);
          continue;
        }

        if (foundHeading) {
          if (isEndOfListBlock(para, text)) {
            break;
          }
          itemsToRemove.push(para);
        }
      }

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
    var listBlockIndices = getListBlockIndices(paragraphs);

    var bookmarkMap = {};
    for (var i = 0; i < bookmarks.length; i++) {
      var bookmark = bookmarks[i];
      var position = bookmark.getPosition();
      var element = position.getElement();

      var para = element.getType() === DocumentApp.ElementType.PARAGRAPH ?
        element.asParagraph() :
        element.getParent().asParagraph();

      if (para) {
        bookmarkMap[para.getText()] = bookmark.getId();
      }
    }

    for (var j = 0; j < paragraphs.length; j++) {
      var captionPara = paragraphs[j];
      var text = captionPara.getText();

      if (isRealCaptionParagraph(captionPara, prefix, listBlockIndices, j)) {
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
    var listBlockIndices = getListBlockIndices(paragraphs);

    for (var i = 0; i < paragraphs.length; i++) {
      var para = paragraphs[i];
      if (isRealCaptionParagraph(para, prefix, listBlockIndices, i)) {
        captions.push(para.getText());
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

  /**
   * Returns paragraph elements that are real captions for a prefix
   * @param {string} prefix
   * @return {Array<Paragraph>}
   */
  function getCaptionParagraphs(prefix) {
    var doc = DocumentApp.getActiveDocument();
    var paragraphs = doc.getBody().getParagraphs();
    var listBlockIndices = getListBlockIndices(paragraphs);
    var result = [];

    for (var i = 0; i < paragraphs.length; i++) {
      if (isRealCaptionParagraph(paragraphs[i], prefix, listBlockIndices, i)) {
        result.push(paragraphs[i]);
      }
    }

    return result;
  }

  return {
    generateListOfTables: generateListOfTables,
    generateListOfFigures: generateListOfFigures,
    findCaptions: findCaptions,
    findCaptionsWithBookmarks: findCaptionsWithBookmarks,
    getCaptionParagraphs: getCaptionParagraphs,
    searchAndReplace: searchAndReplace
  };

})();
