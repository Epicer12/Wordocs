/**
 * Caption Manager
 * Handles caption insertion, numbering, and management
 */

var CaptionManager = (function () {

  var PROPERTY_KEY_TABLE_PREFIX = 'CAPTION_TABLE_PREFIX';
  var PROPERTY_KEY_FIGURE_PREFIX = 'CAPTION_FIGURE_PREFIX';
  var PROPERTY_KEY_STYLE_TABLE = 'CAPTION_STYLE_TABLE';
  var PROPERTY_KEY_STYLE_FIGURE = 'CAPTION_STYLE_FIGURE';
  var PROPERTY_KEY_TABLE_POSITION = 'CAPTION_TABLE_POSITION';
  var DEFAULT_TABLE_PREFIX = 'Table';
  var DEFAULT_FIGURE_PREFIX = 'Figure';
  var DEFAULT_TABLE_POSITION = 'above';
  var MAX_PREFIX_LENGTH = 20;

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

  function refreshCaptionCharacterStyles(para, prefix, number, descriptionText, style) {
    var normalized = normalizeCaptionStyle(style);
    var text = para.editAsText();
    var len = para.getText().length;

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

  function applyCaptionParagraphStyles(para, style) {
    var normalized = normalizeCaptionStyle(style);
    para.setAlignment(alignmentToHorizontalAlignment(normalized.alignment));
    para.setSpacingBefore(normalized.spacingBefore);
    para.setSpacingAfter(normalized.spacingAfter);
  }

  /**
   * Updates only the caption number digits — preserves bookmarks on the paragraph.
   */
  function renumberCaptionNumberInPlace(para, prefix, newNumber, style) {
    var numRegex = buildCaptionNumberRegex(prefix);
    var extractRegex = buildCaptionTextExtractRegex(prefix);
    var text = para.getText();
    var numMatch = text.match(numRegex);

    if (!numMatch) {
      return;
    }

    var oldNumber = parseInt(numMatch[1], 10);
    if (oldNumber === newNumber) {
      return;
    }

    var numStart = prefix.length + 1;
    var numEnd = numStart + numMatch[1].length - 1;
    var textEl = para.editAsText();

    textEl.deleteText(numStart, numEnd);
    textEl.insertText(numStart, String(newNumber));

    applyCaptionParagraphStyles(para, style);

    var updatedText = para.getText();
    var descMatch = updatedText.match(extractRegex);
    var descriptionText = descMatch ? descMatch[1] : '';
    refreshCaptionCharacterStyles(para, prefix, newNumber, descriptionText, style);
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
    var currentText = para.getText();

    if (currentText !== fullCaption) {
      var textElement = para.editAsText();
      if (currentText.length > 0) {
        textElement.deleteText(0, currentText.length - 1);
      }
      if (fullCaption.length > 0) {
        textElement.insertText(0, fullCaption);
      }
    }

    para.setAlignment(alignmentToHorizontalAlignment(normalized.alignment));
    para.setSpacingBefore(normalized.spacingBefore);
    para.setSpacingAfter(normalized.spacingAfter);

    refreshCaptionCharacterStyles(para, prefix, number, descriptionText, style);
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
    return props.getProperty(PROPERTY_KEY_TABLE_PREFIX) || DEFAULT_TABLE_PREFIX;
  }

  function getFigurePrefix() {
    var props = getDocProperties();
    return props.getProperty(PROPERTY_KEY_FIGURE_PREFIX) || DEFAULT_FIGURE_PREFIX;
  }

  /**
   * Where new table captions are inserted relative to the table
   * @return {string} 'above' or 'below'
   */
  function getTablePosition() {
    var stored = getDocProperties().getProperty(PROPERTY_KEY_TABLE_POSITION);
    return stored === 'below' ? 'below' : DEFAULT_TABLE_POSITION;
  }

  function normalizePosition(value) {
    return value === 'below' ? 'below' : 'above';
  }

  function normalizePrefix(value) {
    var raw = value === undefined || value === null ? '' : String(value);
    return raw.replace(/\s+/g, ' ').trim();
  }

  /**
   * @param {string} prefix - Normalized prefix
   * @param {string} name - 'Table' or 'Figure', used in the message
   * @return {string|null} Error message, or null when valid
   */
  function validatePrefix(prefix, name) {
    if (!prefix) {
      return name + ' label cannot be empty.';
    }
    if (prefix.length > MAX_PREFIX_LENGTH) {
      return name + ' label must be ' + MAX_PREFIX_LENGTH + ' characters or fewer.';
    }
    if (prefix.indexOf(':') !== -1) {
      return name + ' label cannot contain a colon.';
    }
    return null;
  }

  /**
   * Current caption label settings for the Settings sidebar
   * @return {{tablePrefix: string, figurePrefix: string, defaults: Object}}
   */
  function getSettings() {
    return {
      tablePrefix: getTablePrefix(),
      figurePrefix: getFigurePrefix(),
      tablePosition: getTablePosition(),
      defaults: {
        tablePrefix: DEFAULT_TABLE_PREFIX,
        figurePrefix: DEFAULT_FIGURE_PREFIX,
        tablePosition: DEFAULT_TABLE_POSITION
      }
    };
  }

  /**
   * Replaces the label text on existing captions, keeping number and description.
   * Google Docs relocates a bookmark when the label at the start of its paragraph is edited,
   * so each caption is rebuilt instead: a detached copy (which carries no bookmark) gets the
   * new label and replaces the original, then receives a fresh bookmark. The returned map
   * lets cross-reference links be repointed from the old bookmark IDs.
   * @return {Object<string, string>} Old bookmark ID -> new bookmark ID
   */
  function renameCaptionPrefix(oldPrefix, newPrefix, style) {
    var doc = DocumentApp.getActiveDocument();
    var paragraphs = ListGenerator.getCaptionParagraphs(oldPrefix);
    var oldNumRegex = buildCaptionNumberRegex(oldPrefix);
    var newExtractRegex = buildCaptionTextExtractRegex(newPrefix);
    var remap = {};
    var oldIdSets = [];
    var i;

    // Snapshot every caption's bookmarks first: once a paragraph is replaced, Docs moves
    // its leftover bookmarks onto neighbouring paragraphs, which would blur who owned what
    for (i = 0; i < paragraphs.length; i++) {
      oldIdSets.push(collectBookmarkIds(doc, paragraphs[i]));
    }

    var rebuiltParagraphs = [];

    // Phase 1: rebuild every caption, last to first so child indices still to process do
    // not shift. No bookmarks yet: Docs relocates bookmarks created earlier in the same run
    // when the document structure changes afterwards.
    for (i = paragraphs.length - 1; i >= 0; i--) {
      var para = paragraphs[i];
      var numMatch = para.getText().match(oldNumRegex);
      if (!numMatch) {
        continue;
      }

      var parent = para.getParent();
      var index = parent.getChildIndex(para);

      var copy = para.copy();
      var copyText = copy.editAsText();
      copyText.insertText(0, newPrefix);
      copyText.deleteText(newPrefix.length, newPrefix.length + oldPrefix.length - 1);

      var rebuilt = parent.insertParagraph(index + 1, copy);
      var descMatch = rebuilt.getText().match(newExtractRegex);
      refreshCaptionCharacterStyles(
        rebuilt, newPrefix, parseInt(numMatch[1], 10), descMatch ? descMatch[1] : '', style
      );

      para.removeFromParent();
      rebuiltParagraphs.push({ index: i, paragraph: rebuilt });
    }

    // Phase 2: with the structure settled, bookmark each rebuilt caption and map old IDs
    for (var r = 0; r < rebuiltParagraphs.length; r++) {
      var newId = doc.addBookmark(doc.newPosition(rebuiltParagraphs[r].paragraph, 0)).getId();
      var oldIds = oldIdSets[rebuiltParagraphs[r].index];
      for (var k = 0; k < oldIds.length; k++) {
        remap[oldIds[k]] = newId;
      }
    }

    return remap;
  }

  /**
   * Best-effort cleanup of bookmarks left behind by a rebuilt caption. Docs does not always
   * allow removal, and a leftover is harmless once links point at the new bookmark.
   */
  function removeBookmarksQuietly(doc, bookmarkIds) {
    var failed = 0;

    for (var i = 0; i < bookmarkIds.length; i++) {
      try {
        var bookmark = doc.getBookmark(bookmarkIds[i]);
        if (bookmark) {
          bookmark.remove();
        }
      } catch (error) {
        failed++;
      }
    }

    if (failed > 0) {
      Logger.log('Could not remove ' + failed + ' stale bookmark(s); they are unused.');
    }
  }

  /**
   * All bookmark IDs currently on a caption paragraph (creating one if there is none)
   * @return {Array<string>}
   */
  function collectBookmarkIds(doc, para) {
    ensureBookmarkOnParagraph(doc, para);
    var bookmarks = findBookmarkIdsOnParagraph(doc, paragraphStructuralKey(para));
    var ids = [];

    for (var i = 0; i < bookmarks.length; i++) {
      ids.push(bookmarks[i].getId());
    }

    return ids;
  }

  /**
   * Saves caption labels and renames existing captions, cross-references and lists.
   * @param {string} tablePrefix
   * @param {string} figurePrefix
   * @return {Object}
   */
  function saveSettings(tablePrefix, figurePrefix, tablePosition) {
    try {
      var newPosition = tablePosition === undefined ? getTablePosition() : normalizePosition(tablePosition);
      var positionChanged = newPosition !== getTablePosition();

      var newTable = normalizePrefix(tablePrefix);
      var newFigure = normalizePrefix(figurePrefix);
      var validationError = validatePrefix(newTable, 'Table') || validatePrefix(newFigure, 'Figure');

      if (validationError) {
        return { success: false, message: validationError };
      }
      if (newTable.toLowerCase() === newFigure.toLowerCase()) {
        return { success: false, message: 'Table and figure labels must be different.' };
      }

      var oldTable = getTablePrefix();
      var oldFigure = getFigurePrefix();
      var tableChanged = newTable !== oldTable;
      var figureChanged = newFigure !== oldFigure;

      if (!tableChanged && !figureChanged) {
        if (positionChanged) {
          getDocProperties().setProperty(PROPERTY_KEY_TABLE_POSITION, newPosition);
          return {
            success: true,
            changed: false,
            message: 'Saved. New table captions will go ' + newPosition + ' the table.',
            settings: getSettings()
          };
        }
        return { success: true, changed: false, message: 'Settings are already up to date.', settings: getSettings() };
      }

      if ((tableChanged && newTable.toLowerCase() === oldFigure.toLowerCase()) ||
        (figureChanged && newFigure.toLowerCase() === oldTable.toLowerCase())) {
        return {
          success: false,
          message: 'A label cannot take the other type\'s current label. Change one label at a time.'
        };
      }

      var bookmarkRemap = {};
      var oldId;
      var partial;
      if (tableChanged) {
        partial = renameCaptionPrefix(oldTable, newTable, getDefaultCaptionStyle('table'));
        for (oldId in partial) {
          bookmarkRemap[oldId] = partial[oldId];
        }
      }
      if (figureChanged) {
        partial = renameCaptionPrefix(oldFigure, newFigure, getDefaultCaptionStyle('figure'));
        for (oldId in partial) {
          bookmarkRemap[oldId] = partial[oldId];
        }
      }

      var props = getDocProperties();
      props.setProperty(PROPERTY_KEY_TABLE_PREFIX, newTable);
      props.setProperty(PROPERTY_KEY_FIGURE_PREFIX, newFigure);
      props.setProperty(PROPERTY_KEY_TABLE_POSITION, newPosition);

      try {
        CrossRef.updateAllReferences({ table: oldTable, figure: oldFigure }, bookmarkRemap);
      } catch (refError) {
        Logger.log('Cross-ref update after label change: ' + refError);
      }

      var staleIds = [];
      for (oldId in bookmarkRemap) {
        staleIds.push(oldId);
      }
      removeBookmarksQuietly(DocumentApp.getActiveDocument(), staleIds);

      // Lists are refreshed by a separate call (see refreshListsAfterLabelChange): list
      // rebuilding changes the document structure, which would relocate the bookmarks
      // created above if it ran in this same execution.
      return {
        success: true,
        changed: true,
        tableChanged: tableChanged,
        figureChanged: figureChanged,
        message: 'Labels saved. Existing captions and references were updated.',
        settings: getSettings()
      };

    } catch (error) {
      Logger.log('Error in saveSettings: ' + error);
      return { success: false, message: 'Error: ' + error.toString() };
    }
  }

  /**
   * @return {Object}
   */
  function resetSettings() {
    return saveSettings(DEFAULT_TABLE_PREFIX, DEFAULT_FIGURE_PREFIX, DEFAULT_TABLE_POSITION);
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

  function paragraphContainingBookmarkElement(element) {
    var node = element;
    while (node) {
      if (node.getType() === DocumentApp.ElementType.PARAGRAPH) {
        return node.asParagraph();
      }
      node = node.getParent();
    }
    return null;
  }

  function paragraphStructuralKey(paragraph) {
    var segments = [];
    var node = paragraph;

    while (node) {
      var parentNode = node.getParent();
      if (!parentNode) {
        break;
      }
      segments.unshift(parentNode.getChildIndex(node));
      node = parentNode;

      if (segments.length >= 64) {
        break;
      }
    }

    return segments.join(':');
  }

  function findBookmarkIdsOnParagraph(doc, targetKey) {
    var matching = [];
    var bookmarks = doc.getBookmarks();

    for (var i = 0; i < bookmarks.length; i++) {
      var bookmark = bookmarks[i];

      try {
        var bookmarkPara = paragraphContainingBookmarkElement(
          bookmark.getPosition().getElement()
        );

        if (!bookmarkPara || paragraphStructuralKey(bookmarkPara) !== targetKey) {
          continue;
        }

        matching.push(bookmark);
      } catch (error) {
        Logger.log('findBookmarkIdsOnParagraph: ' + error);
      }
    }

    return matching;
  }

  function ensureBookmarkOnParagraph(doc, para) {
    var targetKey = paragraphStructuralKey(para);

    if (!targetKey) {
      return addCaptionBookmark(doc, para);
    }

    var matching = findBookmarkIdsOnParagraph(doc, targetKey);

    if (matching.length > 0) {
      var keepId = matching[0].getId();
      for (var i = matching.length - 1; i > 0; i--) {
        try {
          matching[i].remove();
        } catch (error) {
          Logger.log('ensureBookmarkOnParagraph remove duplicate: ' + error);
        }
      }
      return keepId;
    }

    return addCaptionBookmark(doc, para);
  }

  /**
   * Adds a bookmark at the start of a caption paragraph. Edits after the first character
   * (renumbering) leave it in place; replacing the label itself does not, which is why
   * label changes rebuild the paragraph instead (see renameCaptionPrefix).
   * @return {string} New bookmark ID
   */
  function addCaptionBookmark(doc, para) {
    return doc.addBookmark(doc.newPosition(para, 0)).getId();
  }

  function getParagraphLocator(para) {
    var parent = para.getParent();
    return {
      parent: parent,
      childIndex: parent.getChildIndex(para)
    };
  }

  function renumberCaptions(prefix) {
    var doc = DocumentApp.getActiveDocument();
    var style = getStyleForPrefix(prefix);
    var paragraphs = ListGenerator.getCaptionParagraphs(prefix);

    for (var i = 0; i < paragraphs.length; i++) {
      renumberCaptionNumberInPlace(paragraphs[i], prefix, i + 1, style);
    }

    for (var j = 0; j < paragraphs.length; j++) {
      ensureBookmarkOnParagraph(doc, paragraphs[j]);
    }

    try {
      CrossRef.updateAllReferences();
    } catch (error) {
      Logger.log('Cross-ref update after renumber: ' + error);
    }

    return paragraphs.length;
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
      var initialParagraphs = ListGenerator.getCaptionParagraphs(prefix);

      if (initialParagraphs.length === 0) {
        return {
          success: false,
          message: 'No ' + (type === 'figure' ? 'figure' : 'table') + ' captions found in the document.'
        };
      }

      var processed = 0;
      while (true) {
        var freshParagraphs = ListGenerator.getCaptionParagraphs(prefix);
        if (processed >= freshParagraphs.length) {
          break;
        }
        applyStyleToExistingParagraph(freshParagraphs[processed], prefix, style);
        ensureBookmarkOnParagraph(doc, freshParagraphs[processed]);
        processed++;
      }

      return {
        success: true,
        count: processed,
        message: 'Applied style to ' + processed + ' caption(s).'
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
      var insertIndex = getTablePosition() === 'above' ? tableIndex : tableIndex + 1;
      var captionParagraph = parent.insertParagraph(insertIndex, '');

      applyCaptionStyle(captionParagraph, prefix, tableNum, captionText, captionStyle);
      ensureBookmarkOnParagraph(doc, captionParagraph);

      return {
        success: true,
        message: prefix + ' ' + tableNum + ' caption added successfully!',
        captionCount: getTableCount(),
        insertedNumber: tableNum
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
      var image = findTargetImage(doc);

      if (!image) {
        return {
          success: false,
          message: 'Please select the image, or place your cursor right next to it.'
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
        message: prefix + ' ' + figureNum + ' caption added successfully!',
        captionCount: getFigureCount(),
        insertedNumber: figureNum
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

  /**
   * Finds the inline image in a paragraph, preferring the one just before the cursor offset
   * @param {Paragraph} paragraph
   * @param {number=} offset - Child index the cursor sits at
   * @return {InlineImage|null}
   */
  function findImageInParagraph(paragraph, offset) {
    var numChildren = paragraph.getNumChildren();

    if (offset !== undefined && offset > 0 && offset <= numChildren) {
      var before = paragraph.getChild(offset - 1);
      if (before.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
        return before.asInlineImage();
      }
    }

    for (var i = 0; i < numChildren; i++) {
      var child = paragraph.getChild(i);
      if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
        return child.asInlineImage();
      }
    }

    return null;
  }

  function findParentImage(element) {
    if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      return element.asInlineImage();
    }

    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      return findImageInParagraph(element.asParagraph());
    }

    var parent = element.getParent();
    if (parent && parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      return findImageInParagraph(parent.asParagraph());
    }

    return null;
  }

  /**
   * The image a figure caption should go under: a selected image, or the image next to
   * the cursor. Docs reports no cursor while an image is selected, and reports the cursor
   * as a paragraph position when it sits directly beside an image.
   * @param {Document} doc
   * @return {InlineImage|null}
   */
  function findTargetImage(doc) {
    var selection = doc.getSelection();
    if (selection) {
      var rangeElements = selection.getRangeElements();
      for (var i = 0; i < rangeElements.length; i++) {
        var selected = findParentImage(rangeElements[i].getElement());
        if (selected) {
          return selected;
        }
      }
    }

    var cursor = doc.getCursor();
    if (!cursor) {
      return null;
    }

    var element = cursor.getElement();
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      return findImageInParagraph(element.asParagraph(), cursor.getOffset());
    }
    return findParentImage(element);
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
    getTablePosition: getTablePosition,
    getSettings: getSettings,
    saveSettings: saveSettings,
    resetSettings: resetSettings,
    getDefaultCaptionStyle: getDefaultCaptionStyle,
    saveDefaultCaptionStyle: saveDefaultCaptionStyle,
    applyStyleToAllCaptions: applyStyleToAllCaptions,
    ensureBookmarkOnParagraph: ensureBookmarkOnParagraph,
    getParagraphLocator: getParagraphLocator,
    paragraphStructuralKey: paragraphStructuralKey
  };

})();
