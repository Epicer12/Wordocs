/**
 * Cross Reference Manager
 * In-text bookmark links to table/figure captions
 */

var CrossRef = (function() {

  var BOOKMARK_LINK_PREFIX = '#bookmark=';
  var SNIPPET_MAX_LEN = 40;

  function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function getPrefixForType(type) {
    return type === 'figure' ? CaptionManager.getFigurePrefix() : CaptionManager.getTablePrefix();
  }

  function buildLabel(prefix, number) {
    return prefix + ' ' + number;
  }

  function buildCaptionNumberRegex(prefix) {
    return new RegExp('^' + escapeRegex(prefix) + '\\s+(\\d+):', 'i');
  }

  function buildLabelOnlyRegex(prefix) {
    return new RegExp('^' + escapeRegex(prefix) + '\\s+\\d+$', 'i');
  }

  function parseNumberFromCaptionText(prefix, text) {
    var match = text.match(buildCaptionNumberRegex(prefix));
    return match ? parseInt(match[1], 10) : null;
  }

  function extractSnippet(prefix, fullText) {
    var extractRegex = new RegExp('^' + escapeRegex(prefix) + '\\s+\\d+:\\s*(.*)$', 'i');
    var match = fullText.match(extractRegex);
    var desc = match && match[1] ? match[1].trim() : '';
    if (!desc) {
      return '';
    }
    if (desc.length <= SNIPPET_MAX_LEN) {
      return desc;
    }
    return desc.substring(0, SNIPPET_MAX_LEN) + '…';
  }

  function applyBookmarkLink(text, start, end, bookmarkId) {
    text.setLinkUrl(start, end, BOOKMARK_LINK_PREFIX + bookmarkId);
  }

  function findParentParagraph(element) {
    var node = element;
    while (node) {
      if (node.getType() === DocumentApp.ElementType.PARAGRAPH) {
        return node.asParagraph();
      }
      node = node.getParent();
    }
    return null;
  }

  function getChildTextLength(child) {
    if (child.getType() === DocumentApp.ElementType.TEXT) {
      return child.asText().getText().length;
    }
    return 1;
  }

  function getParagraphChildTextOffset(paragraph, childIndex) {
    var textOffset = 0;
    var limit = Math.min(childIndex, paragraph.getNumChildren());
    for (var i = 0; i < limit; i++) {
      textOffset += getChildTextLength(paragraph.getChild(i));
    }
    return textOffset;
  }

  function getTextOffsetBeforeElement(paragraph, target) {
    var textOffset = 0;
    var numChildren = paragraph.getNumChildren();
    for (var i = 0; i < numChildren; i++) {
      var child = paragraph.getChild(i);
      if (child === target) {
        return textOffset;
      }
      textOffset += getChildTextLength(child);
    }
    var paraText = paragraph.editAsText().getText();
    return paraText.length > 0 ? paraText.length - 1 : 0;
  }

  /**
   * Resolves cursor to a Text insert point (TEXT node or containing paragraph).
   * @param {Position} cursor
   * @return {{text: Text, offset: number}|null}
   */
  function resolveCursorTextInsert(cursor) {
    var element = cursor.getElement();
    var offset = cursor.getOffset();

    if (element.getType() === DocumentApp.ElementType.TEXT) {
      return { text: element.asText(), offset: offset };
    }

    var paragraph = findParentParagraph(element);
    if (!paragraph) {
      return null;
    }

    var text = paragraph.editAsText();
    var insertOffset;

    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      insertOffset = getParagraphChildTextOffset(paragraph, offset);
    } else {
      insertOffset = getTextOffsetBeforeElement(paragraph, element);
    }

    return { text: text, offset: insertOffset };
  }

  function scanParagraphLinks(textElement, callback) {
    var len = textElement.getText().length;
    var i = 0;

    while (i < len) {
      var url = textElement.getLinkUrl(i);
      if (!url || url.indexOf(BOOKMARK_LINK_PREFIX) !== 0) {
        i++;
        continue;
      }

      var bookmarkId = url.substring(BOOKMARK_LINK_PREFIX.length);
      var j = i + 1;
      while (j < len && textElement.getLinkUrl(j) === url) {
        j++;
      }

      callback({
        start: i,
        end: j - 1,
        bookmarkId: bookmarkId,
        linkText: textElement.getText().substring(i, j)
      });

      i = j;
    }
  }

  function buildBookmarkLabelMap() {
    var map = {};
    var types = [
      { type: 'table', prefix: CaptionManager.getTablePrefix() },
      { type: 'figure', prefix: CaptionManager.getFigurePrefix() }
    ];

    for (var t = 0; t < types.length; t++) {
      var prefix = types[t].prefix;
      var captions = ListGenerator.findCaptionsWithBookmarks(prefix);

      for (var i = 0; i < captions.length; i++) {
        var caption = captions[i];
        var number = parseNumberFromCaptionText(prefix, caption.text);
        if (number === null || !caption.bookmarkId) {
          continue;
        }
        map[caption.bookmarkId] = {
          label: buildLabel(prefix, number),
          prefix: prefix
        };
      }
    }

    return map;
  }

  /**
   * @param {string} type - 'table' or 'figure'
   * @return {Array<Object>}
   */
  function getTargets(type) {
    var prefix = getPrefixForType(type);
    var captions = ListGenerator.findCaptionsWithBookmarks(prefix);
    var targets = [];

    for (var i = 0; i < captions.length; i++) {
      var caption = captions[i];
      var number = parseNumberFromCaptionText(prefix, caption.text);
      if (number === null || !caption.bookmarkId) {
        continue;
      }

      targets.push({
        number: number,
        label: buildLabel(prefix, number),
        snippet: extractSnippet(prefix, caption.text),
        bookmarkId: caption.bookmarkId
      });
    }

    return targets;
  }

  /**
   * @param {string} type
   * @param {number} number
   * @return {Object}
   */
  function insertReference(type, number) {
    try {
      var doc = DocumentApp.getActiveDocument();
      var cursor = doc.getCursor();

      if (!cursor) {
        return {
          success: false,
          message: 'Click in the document where you want the reference, then click Insert reference.'
        };
      }

      var insertPoint = resolveCursorTextInsert(cursor);
      if (!insertPoint) {
        return {
          success: false,
          message: 'Could not insert here. Place the cursor in document body text (not a header or footer).'
        };
      }

      var targets = getTargets(type);
      var target = null;

      for (var i = 0; i < targets.length; i++) {
        if (targets[i].number === number) {
          target = targets[i];
          break;
        }
      }

      if (!target) {
        var kind = type === 'figure' ? 'figure' : 'table';
        return {
          success: false,
          message: 'No ' + kind + ' caption found for that number.'
        };
      }

      var text = insertPoint.text;
      var offset = insertPoint.offset;
      var refText = target.label;

      text.insertText(offset, refText);
      applyBookmarkLink(text, offset, offset + refText.length - 1, target.bookmarkId);

      return {
        success: true,
        message: 'Inserted link to ' + refText + '.',
        label: refText
      };
    } catch (error) {
      Logger.log('Error in insertReference: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }

  function updateLinkIfCrossRef(textElement, link, labelMap) {
    var entry = labelMap[link.bookmarkId];
    if (!entry) {
      return false;
    }

    if (!buildLabelOnlyRegex(entry.prefix).test(link.linkText)) {
      return false;
    }

    if (link.linkText === entry.label) {
      return false;
    }

    textElement.deleteText(link.start, link.end);
    textElement.insertText(link.start, entry.label);
    applyBookmarkLink(
      textElement,
      link.start,
      link.start + entry.label.length - 1,
      link.bookmarkId
    );
    return true;
  }

  function updateCrossRefsInParagraph(textElement, labelMap) {
    var updated = 0;

    scanParagraphLinks(textElement, function(link) {
      if (updateLinkIfCrossRef(textElement, link, labelMap)) {
        updated++;
      }
    });

    return updated;
  }

  /**
   * Refreshes in-text cross-reference labels after caption renumbering
   * @return {Object}
   */
  function updateAllReferences() {
    try {
      var labelMap = buildBookmarkLabelMap();
      var paragraphs = DocumentApp.getActiveDocument().getBody().getParagraphs();
      var updated = 0;

      for (var p = 0; p < paragraphs.length; p++) {
        updated += updateCrossRefsInParagraph(paragraphs[p].editAsText(), labelMap);
      }

      return {
        success: true,
        count: updated,
        message: updated > 0 ?
          'Updated ' + updated + ' cross-reference(s).' :
          'All cross-references are up to date.'
      };
    } catch (error) {
      Logger.log('Error in updateAllReferences: ' + error);
      return {
        success: false,
        message: 'Error: ' + error.toString()
      };
    }
  }

  return {
    getTargets: getTargets,
    insertReference: insertReference,
    updateAllReferences: updateAllReferences
  };

})();
