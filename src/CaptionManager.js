/**
 * Caption Manager
 * Handles caption insertion, numbering, and management
 */

var CaptionManager = (function() {
  
  // Property keys for storing caption metadata
  var PROPERTY_KEY_TABLE_COUNT = 'CAPTION_TABLE_COUNT';
  var PROPERTY_KEY_FIGURE_COUNT = 'CAPTION_FIGURE_COUNT';
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
   * Gets the current table count
   * @return {number}
   */
  function getTableCount() {
    var props = getDocProperties();
    return parseInt(props.getProperty(PROPERTY_KEY_TABLE_COUNT) || '0');
  }
  
  /**
   * Gets the current figure count
   * @return {number}
   */
  function getFigureCount() {
    var props = getDocProperties();
    return parseInt(props.getProperty(PROPERTY_KEY_FIGURE_COUNT) || '0');
  }
  
  /**
   * Sets the table count
   * @param {number} count
   */
  function setTableCount(count) {
    getDocProperties().setProperty(PROPERTY_KEY_TABLE_COUNT, count.toString());
  }
  
  /**
   * Sets the figure count
   * @param {number} count
   */
  function setFigureCount(count) {
    getDocProperties().setProperty(PROPERTY_KEY_FIGURE_COUNT, count.toString());
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
      
      // Increment table count
      var tableNum = getTableCount() + 1;
      setTableCount(tableNum);
      
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
      
      // Increment figure count
      var figureNum = getFigureCount() + 1;
      setFigureCount(figureNum);
      
      // Create caption text
      var prefix = getFigurePrefix();
      var fullCaption = prefix + ' ' + figureNum + ': ' + captionText;
      
      // Insert caption after the image
      var parent = image.getParent();
      var imageIndex = parent.getChildIndex(image);
      var captionParagraph = parent.insertParagraph(imageIndex + 1, fullCaption);
      
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
   * @return {Object} Result object with success status and message
   */
  function updateAllCaptions() {
    try {
      var doc = DocumentApp.getActiveDocument();
      var body = doc.getBody();
      
      var tableCount = 0;
      var figureCount = 0;
      
      var tablePrefix = getTablePrefix();
      var figurePrefix = getFigurePrefix();
      
      // Regex patterns to match captions
      var tableCaptionRegex = new RegExp('^' + tablePrefix + '\\s+\\d+:', 'i');
      var figureCaptionRegex = new RegExp('^' + figurePrefix + '\\s+\\d+:', 'i');
      
      // Search for all paragraphs
      var paragraphs = body.getParagraphs();
      
      for (var i = 0; i < paragraphs.length; i++) {
        var para = paragraphs[i];
        var text = para.getText();
        
        // Check if it's a table caption
        if (tableCaptionRegex.test(text)) {
          tableCount++;
          // Update the number
          var newText = text.replace(/^(Table\s+)\d+:/, '$1' + tableCount + ':');
          para.setText(newText);
          // Re-apply formatting
          para.editAsText().setBold(0, tablePrefix.length + tableCount.toString().length + 1, true);
        }
        // Check if it's a figure caption
        else if (figureCaptionRegex.test(text)) {
          figureCount++;
          // Update the number
          var newText = text.replace(/^(Figure\s+)\d+:/, '$1' + figureCount + ':');
          para.setText(newText);
          // Re-apply formatting
          para.editAsText().setBold(0, figurePrefix.length + figureCount.toString().length + 1, true);
        }
      }
      
      // Update stored counts
      setTableCount(tableCount);
      setFigureCount(figureCount);
      
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
