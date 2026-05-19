/**
 * Caption Manager
 * Handles caption insertion, numbering, and management
 */

var CaptionManager = (function () {

  var PROPERTY_KEY_TABLE_PREFIX = 'CAPTION_TABLE_PREFIX';
  var PROPERTY_KEY_FIGURE_PREFIX = 'CAPTION_FIGURE_PREFIX';

  function getDocProperties() {
    return PropertiesService.getDocumentProperties();
  }

  function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function buildCaptionNumberRegex(prefix) {
    return new RegExp('^' + escapeRegex(prefix) + '\\s+(\\d+):', 'i');
  }

  function buildCaptionTextExtractRegex(prefix) {
    return new RegExp('^' + escapeRegex(prefix) + '\\s+\\d+:\\s*(.*)$', 'i');
  }

  function getTableCount() {
    return ListGenerator.findCaptions(getTablePrefix()).length;
  }

  function getFigureCount() {
    return ListGenerator.findCaptions(getFigurePrefix()).length;
  }

  function getTablePrefix() {
    var props = getDocProperties();
    return props.getProperty(PROPERTY_KEY_TABLE_PREFIX) || 'Table';
  }

  function getFigurePrefix() {
    var props = getDocProperties();
    return props.getProperty(PROPERTY_KEY_FIGURE_PREFIX) || 'Figure';
  }

  /**
   * Returns true if caption numbers are not sequential 1..n in document order
   * @param {string} prefix
   * @return {boolean}
   */
  function needsRenumbering(prefix) {
    var paragraphs = ListGenerator.getCaptionParagraphs(prefix);
    var numRegex = buildCaptionNumberRegex(prefix);

    for (var i = 0; i < paragraphs.length; i++) {
      var match = paragraphs[i].getText().match(numRegex);
      var num = match ? parseInt(match[1], 10) : 0;
      if (num !== i + 1) {
        return true;
      }
    }
    return false;
  }

  /**
   * Stable identity for a paragraph (object references from getParagraphs() are not reliable)
   * @param {Paragraph} para
   * @return {Object}
   */
  function getParagraphLocator(para) {
    var parent = para.getParent();
    return {
      parent: parent,
      childIndex: parent.getChildIndex(para)
    };
  }

  /**
   * @param {Paragraph} para
   * @param {Object} locator
   * @return {boolean}
   */
  function paragraphMatchesLocator(para, locator) {
    var parent = para.getParent();
    return parent === locator.parent && parent.getChildIndex(para) === locator.childIndex;
  }

  /**
   * Ensures a bookmark exists on a caption paragraph; reuses existing if present
   * @param {Document} doc
   * @param {Paragraph} para
   * @return {string} Bookmark ID
   */
  function ensureBookmarkOnParagraph(doc, para) {
    var locator = getParagraphLocator(para);
    var bookmarks = doc.getBookmarks();

    for (var i = 0; i < bookmarks.length; i++) {
      var bookmark = bookmarks[i];
      var position = bookmark.getPosition();
      var element = position.getElement();
      var bookmarkPara = element.getType() === DocumentApp.ElementType.PARAGRAPH ?
        element.asParagraph() :
        element.getParent().asParagraph();

      if (bookmarkPara && paragraphMatchesLocator(bookmarkPara, locator)) {
        return bookmark.getId();
      }
    }

    return doc.addBookmark(doc.newPosition(para, 0)).getId();
  }

  /**
   * Renumbers all captions for a prefix in document order
   * @param {string} prefix
   * @return {number} Number of captions renumbered
   */
  function renumberCaptions(prefix) {
    var doc = DocumentApp.getActiveDocument();
    var extractRegex = buildCaptionTextExtractRegex(prefix);
    var paragraphs = ListGenerator.getCaptionParagraphs(prefix);
    var count = 0;

    for (var i = 0; i < paragraphs.length; i++) {
      var para = paragraphs[i];
      var text = para.getText();
      count++;

      var match = text.match(extractRegex);
      var captionText = match ? match[1] : '';
      var newCaption = prefix + ' ' + count + ': ' + captionText;

      para.clear();
      para.setText(newCaption);
      para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      var boldEnd = prefix.length + count.toString().length + 1;
      para.editAsText().setBold(0, boldEnd, true);

      ensureBookmarkOnParagraph(doc, para);
    }

    return count;
  }

  /**
   * Renumbers table captions only if needed
   * @return {Object} { changed: boolean, count: number }
   */
  function autoRenumberTablesIfNeeded() {
    var prefix = getTablePrefix();
    if (!needsRenumbering(prefix)) {
      return { changed: false, count: getTableCount() };
    }
    return { changed: true, count: renumberCaptions(prefix) };
  }

  /**
   * Renumbers figure captions only if needed
   * @return {Object} { changed: boolean, count: number }
   */
  function autoRenumberFiguresIfNeeded() {
    var prefix = getFigurePrefix();
    if (!needsRenumbering(prefix)) {
      return { changed: false, count: getFigureCount() };
    }
    return { changed: true, count: renumberCaptions(prefix) };
  }

  /**
   * Renumbers both table and figure captions when needed
   * @return {Object} { changed, tableCount, figureCount }
   */
  function autoRenumberIfNeeded() {
    var tableResult = autoRenumberTablesIfNeeded();
    var figureResult = autoRenumberFiguresIfNeeded();
    return {
      changed: tableResult.changed || figureResult.changed,
      tableCount: tableResult.count,
      figureCount: figureResult.count
    };
  }

  function addTableCaption(captionText) {
    try {
      var doc = DocumentApp.getActiveDocument();
      var cursor = doc.getCursor();

      if (!cursor) {
        return {
          success: false,
          message: 'Please place your cursor in the document.'
        };
      }

      var element = cursor.getElement();
      var table = findParentTable(element);

      if (!table) {
        return {
          success: false,
          message: 'Please place your cursor inside a table.'
        };
      }

      autoRenumberTablesIfNeeded();

      var tableNum = getTableCount() + 1;
      var prefix = getTablePrefix();
      var fullCaption = prefix + ' ' + tableNum + ': ' + captionText;

      var parent = table.getParent();
      var tableIndex = parent.getChildIndex(table);
      var captionParagraph = parent.insertParagraph(tableIndex + 1, fullCaption);

      captionParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      captionParagraph.editAsText().setBold(0, prefix.length + tableNum.toString().length + 1, true);

      ensureBookmarkOnParagraph(doc, captionParagraph);

      return {
        success: true,
        message: 'Table caption added successfully!'
      };

    } catch (error) {
      Logger.log('Error in addTableCaption: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }

  function addFigureCaption(captionText) {
    try {
      var doc = DocumentApp.getActiveDocument();
      var cursor = doc.getCursor();

      if (!cursor) {
        return {
          success: false,
          message: 'Please place your cursor in the document.'
        };
      }

      var element = cursor.getElement();
      var image = findParentImage(element);

      if (!image) {
        return {
          success: false,
          message: 'Please place your cursor near an image or select an image.'
        };
      }

      autoRenumberFiguresIfNeeded();

      var figureNum = getFigureCount() + 1;
      var prefix = getFigurePrefix();
      var fullCaption = prefix + ' ' + figureNum + ': ' + captionText;

      var imageParent = image.getParent();
      var captionParagraph;

      if (imageParent.getType() === DocumentApp.ElementType.PARAGRAPH) {
        var imageParagraph = imageParent.asParagraph();
        var body = imageParagraph.getParent();
        var paraIndex = body.getChildIndex(imageParagraph);
        captionParagraph = body.insertParagraph(paraIndex + 1, fullCaption);
      } else {
        var imageIndex = imageParent.getChildIndex(image);
        captionParagraph = imageParent.insertParagraph(imageIndex + 1, fullCaption);
      }

      captionParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      captionParagraph.editAsText().setBold(0, prefix.length + figureNum.toString().length + 1, true);

      ensureBookmarkOnParagraph(doc, captionParagraph);

      return {
        success: true,
        message: 'Figure caption added successfully!'
      };

    } catch (error) {
      Logger.log('Error in addFigureCaption: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }

  function findParentTable(element) {
    while (element) {
      if (element.getType() === DocumentApp.ElementType.TABLE) {
        return element.asTable();
      }
      if (element.getParent) {
        element = element.getParent();
      } else {
        break;
      }
    }
    return null;
  }

  function findParentImage(element) {
    if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      return element.asInlineImage();
    }

    var parent = element.getParent();
    if (parent && parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var numChildren = parent.getNumChildren();
      for (var i = 0; i < numChildren; i++) {
        var child = parent.getChild(i);
        if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
          return child.asInlineImage();
        }
      }
    }

    return null;
  }

  function updateAllCaptions() {
    try {
      var tableCount = renumberCaptions(getTablePrefix());
      var figureCount = renumberCaptions(getFigurePrefix());

      return {
        success: true,
        message: 'Updated ' + tableCount + ' table caption(s) and ' + figureCount + ' figure caption(s).'
      };

    } catch (error) {
      Logger.log('Error in updateAllCaptions: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }

  return {
    addTableCaption: addTableCaption,
    addFigureCaption: addFigureCaption,
    updateAllCaptions: updateAllCaptions,
    autoRenumberIfNeeded: autoRenumberIfNeeded,
    autoRenumberTablesIfNeeded: autoRenumberTablesIfNeeded,
    autoRenumberFiguresIfNeeded: autoRenumberFiguresIfNeeded,
    needsRenumbering: needsRenumbering,
    getTableCount: getTableCount,
    getFigureCount: getFigureCount,
    getTablePrefix: getTablePrefix,
    getFigurePrefix: getFigurePrefix,
    ensureBookmarkOnParagraph: ensureBookmarkOnParagraph,
    getParagraphLocator: getParagraphLocator
  };

})();
