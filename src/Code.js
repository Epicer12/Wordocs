/**
 * Google Docs Word Features Extension
 * Main entry point and menu setup
 * 
 * @author Hasun Tisera
 * @license MIT
 */

/**
 * RUN THIS ONCE MANUALLY
 * This forces Google authorization
 */
function authorizeApp() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  Logger.log(body.getText());
}

function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when the document is opened
 * Creates the custom menu in Google Docs
 */
function onOpen(e) {
  Logger.log("onOpen triggered");

  DocumentApp.getUi()
    .createMenu('Wordocs')
    .addItem('Add Table Caption', 'showTableCaptionDialog')
    .addItem('Add Figure Caption', 'showFigureCaptionDialog')
    .addSeparator()
    .addItem('Insert List of Tables', 'insertListOfTables')
    .addItem('Insert List of Figures', 'insertListOfFigures')
    .addSeparator()
    .addItem('Update All Numbering', 'updateAllNumbering')
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addItem('Help', 'showHelp')
    .addToUi();
}

/**
 * Shows the table caption dialog sidebar
 */
function showTableCaptionDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Add Table Caption')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Shows the figure caption dialog sidebar
 */
function showFigureCaptionDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Add Figure Caption')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Shows the settings dialog
 */
function showSettings() {
  var ui = DocumentApp.getUi();
  ui.alert('Settings', 'Settings feature coming soon!', ui.ButtonSet.OK);
}

/**
 * Shows the help dialog
 */
function showHelp() {
  var ui = DocumentApp.getUi();
  var helpText = 'Google Docs Word Features Extension\n\n' +
    'Features:\n' +
    '• Add captions to tables and figures\n' +
    '• Generate list of tables and figures\n' +
    '• Automatic numbering\n\n' +
    'Visit our GitHub for more information:\n' +
    'github.com/Epicer12/Wordocs';
  ui.alert('Help', helpText, ui.ButtonSet.OK);
}

/**
 * Inserts a list of all tables in the document
 */
function insertListOfTables() {
  try {
    var result = ListGenerator.generateListOfTables();
    if (result.success) {
      DocumentApp.getUi().alert('Success', result.message, DocumentApp.getUi().ButtonSet.OK);
    } else {
      DocumentApp.getUi().alert('Error', result.message, DocumentApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    Logger.log('Error in insertListOfTables: ' + error);
    DocumentApp.getUi().alert('Error', 'Failed to insert list of tables: ' + error, DocumentApp.getUi().ButtonSet.OK);
  }
}

/**
 * Inserts a list of all figures in the document
 */
function insertListOfFigures() {
  try {
    var result = ListGenerator.generateListOfFigures();
    if (result.success) {
      DocumentApp.getUi().alert('Success', result.message, DocumentApp.getUi().ButtonSet.OK);
    } else {
      DocumentApp.getUi().alert('Error', result.message, DocumentApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    Logger.log('Error in insertListOfFigures: ' + error);
    DocumentApp.getUi().alert('Error', 'Failed to insert list of figures: ' + error, DocumentApp.getUi().ButtonSet.OK);
  }
}

/**
 * Updates all caption numbering in the document
 */
function updateAllNumbering() {
  try {
    var result = CaptionManager.updateAllCaptions();
    if (result.success) {
      DocumentApp.getUi().alert('Success', result.message, DocumentApp.getUi().ButtonSet.OK);
    } else {
      DocumentApp.getUi().alert('Error', result.message, DocumentApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    Logger.log('Error in updateAllNumbering: ' + error);
    DocumentApp.getUi().alert('Error', 'Failed to update numbering: ' + error, DocumentApp.getUi().ButtonSet.OK);
  }
}

/**
 * Server-side functions for sidebar integration
 */

/**
 * Gets current caption counts for the sidebar
 * @return {Object} Object with table and figure counts
 */
function getCaptionCounts() {
  return {
    tables: CaptionManager.getTableCount(),
    figures: CaptionManager.getFigureCount()
  };
}

/**
 * Adds a table caption from the sidebar
 * @param {string} captionText - The caption text
 * @return {Object} Result object
 */
function addTableCaptionFromSidebar(captionText) {
  return CaptionManager.addTableCaption(captionText);
}

/**
 * Adds a figure caption from the sidebar
 * @param {string} captionText - The caption text
 * @return {Object} Result object
 */
function addFigureCaptionFromSidebar(captionText) {
  return CaptionManager.addFigureCaption(captionText);
}