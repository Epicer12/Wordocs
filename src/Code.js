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

  var ui = DocumentApp.getUi();
  var captionMenu = ui.createMenu('Caption')
    .addItem('Add Table Caption', 'showTableCaptionDialog')
    .addItem('Add Figure Caption', 'showFigureCaptionDialog');

  ui.createMenu('Wordocs')
    .addSubMenu(captionMenu)
    .addSeparator()
    .addItem('Update All Numbering', 'updateAllNumbering')
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addItem('Help', 'showHelp')
    .addToUi();
}

/**
 * Shows the table caption sidebar
 */
function showTableCaptionDialog() {
  var tpl = HtmlService.createTemplateFromFile('Sidebar');
  tpl.mode = 'table';
  DocumentApp.getUi().showSidebar(
    tpl.evaluate().setTitle('Table Caption').setWidth(300)
  );
}

/**
 * Shows the figure caption sidebar
 */
function showFigureCaptionDialog() {
  var tpl = HtmlService.createTemplateFromFile('Sidebar');
  tpl.mode = 'figure';
  DocumentApp.getUi().showSidebar(
    tpl.evaluate().setTitle('Figure Caption').setWidth(300)
  );
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
    '• Add captions to tables and figures (Wordocs > Caption)\n' +
    '• Generate list of tables and figures from the sidebar\n' +
    '• Automatic numbering\n\n' +
    'Visit our GitHub for more information:\n' +
    'github.com/Epicer12/Wordocs';
  ui.alert('Help', helpText, ui.ButtonSet.OK);
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
    tables: ListGenerator.findCaptions(CaptionManager.getTablePrefix()).length,
    figures: ListGenerator.findCaptions(CaptionManager.getFigurePrefix()).length
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

/**
 * Inserts or refreshes the list of tables from the sidebar
 * @return {Object} Result object
 */
function insertListOfTablesFromSidebar() {
  return ListGenerator.generateListOfTables();
}

/**
 * Inserts or refreshes the list of figures from the sidebar
 * @return {Object} Result object
 */
function insertListOfFiguresFromSidebar() {
  return ListGenerator.generateListOfFigures();
}
