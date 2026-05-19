/**
 * Caption Manager
 * Handles caption insertion, numbering, and management
 */

var CaptionManager = (function () {

  var PROPERTY_KEY_TABLE_PREFIX = 'CAPTION_TABLE_PREFIX';
  var PROPERTY_KEY_FIGURE_PREFIX = 'CAPTION_FIGURE_PREFIX';
  var PROPERTY_KEY_STYLE_TABLE = 'CAPTION_STYLE_TABLE';
  var PROPERTY_KEY_STYLE_FIGURE = 'CAPTION_STYLE_FIGURE';

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

  function getBuiltInDefaultCaptionStyle() {
    return {
      alignment: 'CENTER',
      labelBold: false,
      descriptionItalic: false,
      fontSize: 10,
      fontFamily: 'Arial',
      spacingBefore: 0,
      spacingAfter: 6
    };
  }

  function normalizeCaptionStyle(style) {
    var defaults = getBuiltInDefaultCaptionStyle();
    var normalized = {};
    var key;

    for (key in defaults) {
      if (defaults.hasOwnProperty(key)) {
        normalized[key] = defaults[key];
      }
    }

    if (style) {
      if (style.alignment === 'LEFT' || style.alignment === 'CENTER' || style.alignment === 'RIGHT') {
        normalized.alignment = style.alignment;
      }
      if (typeof style.labelBold === 'boolean') {
        normalized.labelBold = style.labelBold;
      }
      if (typeof style.descriptionItalic === 'boolean') {
        normalized.descriptionItalic = style.descriptionItalic;
      }
      if (style.fontSize) {
        normalized.fontSize = Math.max(8, Math.min(18, parseInt(style.fontSize, 10) || defaults.fontSize));
      }
      if (style.fontFamily) {
        normalized.fontFamily = String(style.fontFamily);
      }
      if (style.spacingBefore !== undefined && style.spacingBefore !== null) {
        normalized.spacingBefore = Math.max(0, Math.min(36, parseInt(style.spacingBefore, 10) || 0));
      }
      if (style.spacingAfter !== undefined && style.spacingAfter !== null) {
        normalized.spacingAfter = Math.max(0, Math.min(36, parseInt(style.spacingAfter, 10) || 0));
      }
    }

    return normalized;
  }

  function getStylePropertyKey(type) {
    return type === 'figure' ? PROPERTY_KEY_STYLE_FIGURE : PROPERTY_KEY_STYLE_TABLE;
  }

  function getDefaultCaptionStyle(type) {
    var props = getDocProperties();
    var raw = props.getProperty(getStylePropertyKey(type));

    if (!raw) {
      return getBuiltInDefaultCaptionStyle();
    }

    try {
      return normalizeCaptionStyle(JSON.parse(raw));
    } catch (e) {
      Logger.log('Invalid caption style JSON: ' + e);
      return getBuiltInDefaultCaptionStyle();
    }
  }

  function saveDefaultCaptionStyle(type, style) {
    var props = getDocProperties();
    var normalized = normalizeCaptionStyle(style);
    props.setProperty(getStylePropertyKey(type), JSON.stringify(normalized));
    return normalized;
  }

  function alignmentToHorizontalAlignment(alignment) {
    if (alignment === 'LEFT') {
      return DocumentApp.HorizontalAlignment.LEFT;
    }
    if (alignment === 'RIGHT') {
      return DocumentApp.HorizontalAlignment.RIGHT;
    }
    return DocumentApp.HorizontalAlignment.CENTER;
  }

  function getLabelEndIndex(prefix, number) {
    return prefix.length + 1 + number.toString().length;
  }

  function getDescriptionStartIndex(prefix, number) {
    return getLabelEndIndex(prefix, number) + 2;
  }

  /**
   * Applies caption formatting to a paragraph
   * @param {Paragraph} para
   * @param {string} prefix
   * @param {number} number
   * @param {string} descriptionText
   * @param {Object} style
   */
  function applyCaptionStyle(para, prefix, number, descriptionText, style) {
    var normalized = normalizeCaptionStyle(style);
    var fullCaption = prefix + ' ' + number + ': ' + descriptionText;

    para.clear();
    para.setText(fullCaption);
    para.setAlignment(alignmentToHorizontalAlignment(normalized.alignment));
    para.setSpacingBefore(normalized.spacingBefore);
    para.setSpacingAfter(normalized.spacingAfter);

    var text = para.editAsText();
    var len = fullCaption.length;

    if (len === 0) {
      return;
    }

    text.setFontFamily(0, len - 1, normalized.fontFamily);
    text.setFontSize(0, len - 1, normalized.fontSize);
    text.setBold(0, len - 1, false);
    text.setItalic(0, len - 1, false);

    var labelEnd = getLabelEndIndex(prefix, number);
    if (normalized.labelBold && labelEnd < len) {
      text.setBold(0, labelEnd, true);
    }

    var descStart = getDescriptionStartIndex(prefix, number);
    if (normalized.descriptionItalic && descriptionText && descStart < len) {
      text.setItalic(descStart, len - 1, true);
    }
  }

  function applyStyleToExistingParagraph(para, prefix, style) {
    var text = para.getText();
    var extractRegex = buildCaptionTextExtractRegex(prefix);
    var numRegex = buildCaptionNumberRegex(prefix);
    var textMatch = text.match(extractRegex);
    var numMatch = text.match(numRegex);

    if (!textMatch || !numMatch) {
      return;
    }

    applyCaptionStyle(para, prefix, parseInt(numMatch[1], 10), textMatch[1], style);
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

  function getStyleForPrefix(prefix) {
    if (prefix === getTablePrefix()) {
      return getDefaultCaptionStyle('table');
    }
    return getDefaultCaptionStyle('figure');
  }

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

  function getParagraphLocator(para) {
    var parent = para.getParent();
    return {
      parent: parent,
      childIndex: parent.getChildIndex(para)
    };
  }

  function paragraphMatchesLocator(para, locator) {
    var parent = para.getParent();
    return parent === locator.parent && parent.getChildIndex(para) === locator.childIndex;
  }

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

  function renumberCaptions(prefix) {
    var doc = DocumentApp.getActiveDocument();
    var extractRegex = buildCaptionTextExtractRegex(prefix);
    var paragraphs = ListGenerator.getCaptionParagraphs(prefix);
    var style = getStyleForPrefix(prefix);
    var count = 0;

    for (var i = 0; i < paragraphs.length; i++) {
      var para = paragraphs[i];
      var text = para.getText();
      count++;

      var match = text.match(extractRegex);
      var captionText = match ? match[1] : '';

      applyCaptionStyle(para, prefix, count, captionText, style);
      ensureBookmarkOnParagraph(doc, para);
    }

    return count;
  }

  function autoRenumberTablesIfNeeded() {
    var prefix = getTablePrefix();
    if (!needsRenumbering(prefix)) {
      return { changed: false, count: getTableCount() };
    }
    return { changed: true, count: renumberCaptions(prefix) };
  }

  function autoRenumberFiguresIfNeeded() {
    var prefix = getFigurePrefix();
    if (!needsRenumbering(prefix)) {
      return { changed: false, count: getFigureCount() };
    }
    return { changed: true, count: renumberCaptions(prefix) };
  }

  function autoRenumberIfNeeded() {
    var tableResult = autoRenumberTablesIfNeeded();
    var figureResult = autoRenumberFiguresIfNeeded();
    return {
      changed: tableResult.changed || figureResult.changed,
      tableCount: tableResult.count,
      figureCount: figureResult.count
    };
  }

  function applyStyleToAllCaptions(type) {
    try {
      var prefix = type === 'figure' ? getFigurePrefix() : getTablePrefix();
      var style = getDefaultCaptionStyle(type);
      var doc = DocumentApp.getActiveDocument();
      var paragraphs = ListGenerator.getCaptionParagraphs(prefix);
      var count = 0;

      if (paragraphs.length === 0) {
        return {
          success: false,
          message: 'No ' + (type === 'figure' ? 'figure' : 'table') + ' captions found in the document.'
        };
      }

      for (var i = 0; i < paragraphs.length; i++) {
        applyStyleToExistingParagraph(paragraphs[i], prefix, style);
        ensureBookmarkOnParagraph(doc, paragraphs[i]);
        count++;
      }

      return {
        success: true,
        count: count,
        message: 'Applied style to ' + count + ' caption(s).'
      };

    } catch (error) {
      Logger.log('Error in applyStyleToAllCaptions: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }

  function addTableCaption(captionText, style) {
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
      var captionStyle = normalizeCaptionStyle(style || getDefaultCaptionStyle('table'));

      var parent = table.getParent();
      var tableIndex = parent.getChildIndex(table);
      var captionParagraph = parent.insertParagraph(tableIndex + 1, '');

      applyCaptionStyle(captionParagraph, prefix, tableNum, captionText, captionStyle);
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

  function addFigureCaption(captionText, style) {
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
      var captionStyle = normalizeCaptionStyle(style || getDefaultCaptionStyle('figure'));

      var imageParent = image.getParent();
      var captionParagraph;

      if (imageParent.getType() === DocumentApp.ElementType.PARAGRAPH) {
        var imageParagraph = imageParent.asParagraph();
        var body = imageParagraph.getParent();
        var paraIndex = body.getChildIndex(imageParagraph);
        captionParagraph = body.insertParagraph(paraIndex + 1, '');
      } else {
        var imageIndex = imageParent.getChildIndex(image);
        captionParagraph = imageParent.insertParagraph(imageIndex + 1, '');
      }

      applyCaptionStyle(captionParagraph, prefix, figureNum, captionText, captionStyle);
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
    getDefaultCaptionStyle: getDefaultCaptionStyle,
    saveDefaultCaptionStyle: saveDefaultCaptionStyle,
    applyStyleToAllCaptions: applyStyleToAllCaptions,
    ensureBookmarkOnParagraph: ensureBookmarkOnParagraph,
    getParagraphLocator: getParagraphLocator
  };

})();
