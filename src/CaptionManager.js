/**
 * Caption Manager
 * Handles caption insertion, numbering, and management
 */

var CaptionManager = (function () {

  // Property keys for caption prefixes (kept for customization)
  var PROPERTY_KEY_TABLE_PREFIX = 'CAPTION_TABLE_PREFIX';
  var PROPERTY_KEY_FIGURE_PREFIX = 'CAPTION_FIGURE_PREFIX';

  /**
   * Gets document properties
   * @return {PropertiesService.Properties}
   */
  function getDocProperties() {
    return PropertiesService.getDocumentProperties();
  }

  /**
   * Gets the current table count by scanning the document
   * @return {number}
   */
  function getTableCount() {
    return ListGenerator.findCaptions(getTablePrefix()).length;
  }

  /**
   * Gets the current figure count by scanning the document
   * @return {number}
   */
  function getFigureCount() {
    return ListGenerator.findCaptions(getFigurePrefix()).length;
  }

  /**
   * Gets the table caption prefix (default: "Table")
   * @return {string}
   */
  function getTablePrefix() {
    var props = getDocProperties();
    return props.getProperty(PROPERTY_KEY_TABLE_PREFIX) || 'Table';
  }

  /**
   * Gets the figure caption prefix (default: "Figure")
   * @return {string}
   */
  function getFigurePrefix() {
    var props = getDocProperties();
    return props.getProperty(PROPERTY_KEY_FIGURE_PREFIX) || 'Figure';
  }

  /**
   * Adds a caption to a table
   * @param {string} captionText - The caption text
   * @return {Object} Result object with success status and message
   */
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

      // Get next table number by scanning document
      var tableNum = getTableCount() + 1;

      // Create caption text
      var prefix = getTablePrefix();
      var fullCaption = prefix + ' ' + tableNum + ': ' + captionText;

      // Insert caption after the table
      var parent = table.getParent();
      var tableIndex = parent.getChildIndex(table);
      var captionParagraph = parent.insertParagraph(tableIndex + 1, fullCaption);

      // Format the caption
      captionParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      captionParagraph.editAsText().setBold(0, prefix.length + tableNum.toString().length + 1, true);

      // Add bookmark for cross-referencing
      var position = doc.newPosition(captionParagraph, 0);
      doc.addBookmark(position);

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

  /**
   * Adds a caption to a figure/image
   * @param {string} captionText - The caption text
   * @return {Object} Result object with success status and message
   */
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

      // Get next figure number by scanning document
      var figureNum = getFigureCount() + 1;

      // Create caption text
      var prefix = getFigurePrefix();
      var fullCaption = prefix + ' ' + figureNum + ': ' + captionText;

      // Insert caption after the image (inline images live inside a Paragraph)
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

      // Format the caption
      captionParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      captionParagraph.editAsText().setBold(0, prefix.length + figureNum.toString().length + 1, true);

      // Add bookmark for cross-referencing
      var position = doc.newPosition(captionParagraph, 0);
      doc.addBookmark(position);

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

  /**
   * Finds the parent table of an element
   * @param {Element} element
   * @return {Table|null}
   */
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
   * Finds the parent image of an element
   * @param {Element} element
   * @return {InlineImage|null}
   */
  function findParentImage(element) {
    // Check if element itself is an image
    if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      return element.asInlineImage();
    }

    // Check parent paragraph for images
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

  /**
   * Updates all caption numbering in the document
   * Scans the document and renumbers all captions sequentially
   * Preserves bookmarks and formatting
   * @return {Object} Result object with success status and message
   */
  function updateAllCaptions() {
    try {
      var doc = DocumentApp.getActiveDocument();

      var tableCount = 0;
      var figureCount = 0;

      var tablePrefix = getTablePrefix();
      var figurePrefix = getFigurePrefix();

      var tableMatchRegex = new RegExp('^' + tablePrefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s+\\d+:\\s*(.*)$', 'i');
      var figureMatchRegex = new RegExp('^' + figurePrefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s+\\d+:\\s*(.*)$', 'i');

      var tableParagraphs = ListGenerator.getCaptionParagraphs(tablePrefix);
      for (var i = 0; i < tableParagraphs.length; i++) {
        var para = tableParagraphs[i];
        var text = para.getText();
        tableCount++;

        var match = text.match(tableMatchRegex);
        var captionText = match ? match[1] : '';
        var newCaption = tablePrefix + ' ' + tableCount + ': ' + captionText;

        para.clear();
        para.setText(newCaption);
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        var boldEnd = tablePrefix.length + tableCount.toString().length + 1;
        para.editAsText().setBold(0, boldEnd, true);

        var position = doc.newPosition(para, 0);
        doc.addBookmark(position);
      }

      var figureParagraphs = ListGenerator.getCaptionParagraphs(figurePrefix);
      for (var j = 0; j < figureParagraphs.length; j++) {
        var figPara = figureParagraphs[j];
        var figText = figPara.getText();
        figureCount++;

        var figMatch = figText.match(figureMatchRegex);
        var figCaptionText = figMatch ? figMatch[1] : '';
        var newFigCaption = figurePrefix + ' ' + figureCount + ': ' + figCaptionText;

        figPara.clear();
        figPara.setText(newFigCaption);
        figPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        var figBoldEnd = figurePrefix.length + figureCount.toString().length + 1;
        figPara.editAsText().setBold(0, figBoldEnd, true);

        var figPosition = doc.newPosition(figPara, 0);
        doc.addBookmark(figPosition);
      }

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

  // Public API
  return {
    addTableCaption: addTableCaption,
    addFigureCaption: addFigureCaption,
    updateAllCaptions: updateAllCaptions,
    getTableCount: getTableCount,
    getFigureCount: getFigureCount,
    getTablePrefix: getTablePrefix,
    getFigurePrefix: getFigurePrefix
  };

})();