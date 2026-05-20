/**
 * List Generator
 * Generates lists of tables and figures with bookmark links
 */

var ListGenerator = (function () {

  var LIST_HEADINGS = ['List of Tables', 'List of Figures'];

  function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function buildCaptionRegex(prefix) {
    return new RegExp('^' + escapeRegex(prefix) + '\\s+\\d+:', 'i');
  }

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

  function getParagraphLocator(para) {
    var parent = para.getParent();
    return {
      parent: parent,
      childIndex: parent.getChildIndex(para)
    };
  }

  function findParagraphIndexByLocator(paragraphs, locator) {
    for (var i = 0; i < paragraphs.length; i++) {
      var p = paragraphs[i];
      var parent = p.getParent();
      if (parent === locator.parent && parent.getChildIndex(p) === locator.childIndex) {
        return i;
      }
    }
    return -1;
  }

  function getParagraphFromBookmark(bookmark) {
    try {
      var position = bookmark.getPosition();
      var element = position.getElement();
      var node = element;

      while (node) {
        if (node.getType() === DocumentApp.ElementType.PARAGRAPH) {
          return node.asParagraph();
        }
        node = node.getParent();
      }
    } catch (error) {
      Logger.log('getParagraphFromBookmark: ' + error);
    }

    return null;
  }

  function findParagraphIndexByStructuralKey(paragraphs, targetKey) {
    for (var i = 0; i < paragraphs.length; i++) {
      if (CaptionManager.paragraphStructuralKey(paragraphs[i]) === targetKey) {
        return i;
      }
    }
    return -1;
  }

  /**
   * Maps paragraph indices to bookmark IDs using stable paragraph locators
   * @param {Document} doc
   * @param {Array<Paragraph>} paragraphs
   * @return {Object<number, string>}
   */
  function buildBookmarkIdByIndex(doc, paragraphs) {
    var map = {};
    var bookmarks = doc.getBookmarks();

    for (var i = 0; i < bookmarks.length; i++) {
      var bookmarkPara = getParagraphFromBookmark(bookmarks[i]);
      if (!bookmarkPara) {
        continue;
      }
      var idx = findParagraphIndexByStructuralKey(
        paragraphs,
        CaptionManager.paragraphStructuralKey(bookmarkPara)
      );
      if (idx >= 0) {
        map[idx] = bookmarks[i].getId();
      }
    }

    return map;
  }

  /**
   * Finds paragraph indices of all H1 list blocks (sorted descending for safe updates)
   * @param {string} headingText
   * @return {Array<number>}
   */
  function findAllListBlockIndices(headingText) {
    var paragraphs = DocumentApp.getActiveDocument().getBody().getParagraphs();
    var indices = [];

    for (var i = 0; i < paragraphs.length; i++) {
      var para = paragraphs[i];
      if (para.getText() === headingText &&
        para.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
        indices.push(i);
      }
    }

    indices.sort(function (a, b) {
      return b - a;
    });

    return indices;
  }

  /**
   * Finds all H1 list blocks for a given heading title
   * @param {string} headingText
   * @return {Array<number>} Paragraph indices of headings
   */
  function findAllListBlocks(headingText) {
    return findAllListBlockIndices(headingText);
  }

  /**
   * Returns a map of paragraph indices inside any list-of block
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
   * Inserts linked list item paragraphs after a given index in a container
   * @param {ContainerElement} parent
   * @param {number} insertIndex
   * @param {Array<Object>} captions
   * @return {number} Next insert index after items
   */
  function insertListItems(parent, insertIndex, captions) {
    for (var i = 0; i < captions.length; i++) {
      var caption = captions[i];
      var listItem = parent.insertParagraph(insertIndex, '');

      var text = listItem.editAsText();
      text.appendText(caption.text);

      if (caption.bookmarkId && caption.text.length > 0) {
        var linkUrl = '#bookmark=' + caption.bookmarkId;
        text.setLinkUrl(0, caption.text.length - 1, linkUrl);
      }

      listItem.setIndentFirstLine(36);
      insertIndex++;
    }

    return insertIndex;
  }

  /**
   * Replaces list rows under a heading with fresh caption entries
   * @param {number} headingIndex - Index in body.getParagraphs()
   * @param {Array<Object>} captions
   */
  function refreshListBlockAtIndex(headingIndex, captions) {
    var body = DocumentApp.getActiveDocument().getBody();
    var paragraphs = body.getParagraphs();

    if (headingIndex < 0 || headingIndex >= paragraphs.length) {
      return;
    }

    var headingPara = paragraphs[headingIndex];
    var itemsToRemove = [];

    for (var j = headingIndex + 1; j < paragraphs.length; j++) {
      var para = paragraphs[j];
      var text = para.getText();
      if (isEndOfListBlock(para, text)) {
        break;
      }
      itemsToRemove.push(para);
    }

    for (var k = itemsToRemove.length - 1; k >= 0; k--) {
      itemsToRemove[k].removeFromParent();
    }

    var parent = headingPara.getParent();
    var insertIndex = parent.getChildIndex(headingPara) + 1;
    insertIndex = insertListItems(parent, insertIndex, captions);
    parent.insertParagraph(insertIndex, '');
  }

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

      CaptionManager.autoRenumberTablesIfNeeded();

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

      insertIndex = insertListItems(parent, insertIndex, tableCaptions);
      parent.insertParagraph(insertIndex, '');

      return {
        success: true,
        count: tableCaptions.length,
        message: 'Inserted new List of Tables with ' + tableCaptions.length + ' item(s).'
      };

    } catch (error) {
      Logger.log('Error in generateListOfTables: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }

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

      CaptionManager.autoRenumberFiguresIfNeeded();

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

      insertIndex = insertListItems(parent, insertIndex, figureCaptions);
      parent.insertParagraph(insertIndex, '');

      return {
        success: true,
        count: figureCaptions.length,
        message: 'Inserted new List of Figures with ' + figureCaptions.length + ' item(s).'
      };

    } catch (error) {
      Logger.log('Error in generateListOfFigures: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }

  function updateAllListsOfTables() {
    try {
      var prefix = CaptionManager.getTablePrefix();
      CaptionManager.autoRenumberTablesIfNeeded();

      var captions = findCaptionsWithBookmarks(prefix);
      if (captions.length === 0) {
        return {
          success: false,
          message: 'No table captions found in the document.'
        };
      }

      var blockIndices = findAllListBlockIndices('List of Tables');
      if (blockIndices.length === 0) {
        return {
          success: false,
          message: 'No List of Tables found — use Insert List of Tables first.'
        };
      }

      for (var i = 0; i < blockIndices.length; i++) {
        refreshListBlockAtIndex(blockIndices[i], captions);
      }

      return {
        success: true,
        count: captions.length,
        blocksUpdated: blockIndices.length,
        message: 'Updated ' + blockIndices.length + ' List of Tables block(s) with ' +
          captions.length + ' item(s) each.'
      };

    } catch (error) {
      Logger.log('Error in updateAllListsOfTables: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }

  function updateAllListsOfFigures() {
    try {
      var prefix = CaptionManager.getFigurePrefix();
      CaptionManager.autoRenumberFiguresIfNeeded();

      var captions = findCaptionsWithBookmarks(prefix);
      if (captions.length === 0) {
        return {
          success: false,
          message: 'No figure captions found in the document.'
        };
      }

      var blockIndices = findAllListBlockIndices('List of Figures');
      if (blockIndices.length === 0) {
        return {
          success: false,
          message: 'No List of Figures found — use Insert List of Figures first.'
        };
      }

      for (var j = 0; j < blockIndices.length; j++) {
        refreshListBlockAtIndex(blockIndices[j], captions);
      }

      return {
        success: true,
        count: captions.length,
        blocksUpdated: blockIndices.length,
        message: 'Updated ' + blockIndices.length + ' List of Figures block(s) with ' +
          captions.length + ' item(s) each.'
      };

    } catch (error) {
      Logger.log('Error in updateAllListsOfFigures: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }

  function findCaptionsWithBookmarks(prefix) {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    var captions = [];
    var listBlockIndices = getListBlockIndices(paragraphs);
    var bookmarkIdsByIndex = buildBookmarkIdByIndex(doc, paragraphs);

    for (var j = 0; j < paragraphs.length; j++) {
      var captionPara = paragraphs[j];
      var text = captionPara.getText();

      if (isRealCaptionParagraph(captionPara, prefix, listBlockIndices, j)) {
        var bookmarkId = bookmarkIdsByIndex[j];
        if (!bookmarkId) {
          bookmarkId = CaptionManager.ensureBookmarkOnParagraph(doc, captionPara);
        }

        captions.push({
          text: captionPara.getText(),
          bookmarkId: bookmarkId
        });
      }
    }

    return captions;
  }

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
    updateAllListsOfTables: updateAllListsOfTables,
    updateAllListsOfFigures: updateAllListsOfFigures,
    findCaptions: findCaptions,
    findCaptionsWithBookmarks: findCaptionsWithBookmarks,
    getCaptionParagraphs: getCaptionParagraphs,
    findAllListBlocks: findAllListBlocks,
    searchAndReplace: searchAndReplace
  };

})();
