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

  try {
    CaptionManager.autoRenumberIfNeeded();
  } catch (error) {
    Logger.log('Auto-renumber on open: ' + error);
  }

  var ui = DocumentApp.getUi();
  var captionMenu = ui.createMenu('Caption')
    .addItem('Add Table Caption', 'showTableCaptionDialog')
    .addItem('Add Figure Caption', 'showFigureCaptionDialog')
    .addSeparator()
    .addItem('Insert Cross-Reference', 'showCrossRefSidebar');

  ui.createMenu('Wordocs')
    .addSubMenu(captionMenu)
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addItem('Help', 'showHelp')
    .addToUi();
}

/**
 * Shows the table caption sidebar
 */
function showTableCaptionDialog() {
  try {
    CaptionManager.autoRenumberTablesIfNeeded();
  } catch (error) {
    Logger.log('Auto-renumber tables on sidebar open: ' + error);
  }

  var tpl = HtmlService.createTemplateFromFile('Sidebar');
  tpl.mode = 'table';
  DocumentApp.getUi().showSidebar(
    tpl.evaluate().setTitle('Table Caption').setWidth(320)
  );
}

/**
 * Shows the figure caption sidebar
 */
function showFigureCaptionDialog() {
  try {
    CaptionManager.autoRenumberFiguresIfNeeded();
  } catch (error) {
    Logger.log('Auto-renumber figures on sidebar open: ' + error);
  }

  var tpl = HtmlService.createTemplateFromFile('Sidebar');
  tpl.mode = 'figure';
  DocumentApp.getUi().showSidebar(
    tpl.evaluate().setTitle('Figure Caption').setWidth(320)
  );
}

/**
 * Shows the cross-reference sidebar
 */
function showCrossRefSidebar() {
  try {
    CaptionManager.autoRenumberIfNeeded();
  } catch (error) {
    Logger.log('Auto-renumber on cross-ref sidebar open: ' + error);
  }

  DocumentApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('CrossRefSidebar')
      .evaluate()
      .setTitle('Cross Reference')
      .setWidth(280)
  );
}

/**
 * Returns caption targets for the cross-reference sidebar dropdown
 * @param {string} type - 'table' or 'figure'
 * @return {Object}
 */
function getCrossRefTargetsForSidebar(type) {
  var refType = type === 'figure' ? 'figure' : 'table';

  if (refType === 'table') {
    CaptionManager.autoRenumberTablesIfNeeded();
  } else {
    CaptionManager.autoRenumberFiguresIfNeeded();
  }

  return {
    targets: CrossRef.getTargets(refType),
    tablePrefix: CaptionManager.getTablePrefix(),
    figurePrefix: CaptionManager.getFigurePrefix()
  };
}

/**
 * Inserts a cross-reference at the cursor from the sidebar
 * @param {string} type
 * @param {number} number
 * @return {Object}
 */
function insertCrossRefFromSidebar(type, number) {
  var refType = type === 'figure' ? 'figure' : 'table';
  return CrossRef.insertReference(refType, number);
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
    '• Automatic caption numbering when you open the document\n' +
    '• Customizable caption formatting from the sidebar\n' +
    '• Insert and update lists of tables and figures from the sidebar\n' +
    '• Insert cross-references to captions (Wordocs > Caption > Insert Cross-Reference)\n\n' +
    'Visit our GitHub for more information:\n' +
    'github.com/Epicer12/Wordocs';
  ui.alert('Help', helpText, ui.ButtonSet.OK);
}

/**
 * Gets caption style defaults and count for the sidebar
 * @param {string} mode - 'table' or 'figure'
 * @return {Object}
 */
function getCaptionStyleForSidebar(mode) {
  var type = mode === 'figure' ? 'figure' : 'table';

  if (type === 'table') {
    CaptionManager.autoRenumberTablesIfNeeded();
  } else {
    CaptionManager.autoRenumberFiguresIfNeeded();
  }

  return {
    style: CaptionManager.getDefaultCaptionStyle(type),
    captionCount: type === 'table' ? CaptionManager.getTableCount() : CaptionManager.getFigureCount()
  };
}

/**
 * Saves caption style defaults from the sidebar
 * @param {string} mode
 * @param {Object} style
 * @return {Object}
 */
function saveCaptionStyleFromSidebar(mode, style) {
  try {
    var type = mode === 'figure' ? 'figure' : 'table';
    var saved = CaptionManager.saveDefaultCaptionStyle(type, style);
    return { success: true, style: saved };
  } catch (error) {
    Logger.log('Error in saveCaptionStyleFromSidebar: ' + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Applies saved style to all captions of the current type
 * @param {string} mode
 * @return {Object}
 */
function applyCaptionStyleToAllFromSidebar(mode) {
  var type = mode === 'figure' ? 'figure' : 'table';
  return CaptionManager.applyStyleToAllCaptions(type);
}

/**
 * Adds a table caption from the sidebar
 * @param {string} captionText
 * @param {Object} style
 * @return {Object}
 */
function addTableCaptionFromSidebar(captionText, style) {
  return CaptionManager.addTableCaption(captionText, style);
}

/**
 * Adds a figure caption from the sidebar
 * @param {string} captionText
 * @param {Object} style
 * @return {Object}
 */
function addFigureCaptionFromSidebar(captionText, style) {
  return CaptionManager.addFigureCaption(captionText, style);
}

/**
 * Inserts a new list of tables at the cursor from the sidebar
 * @return {Object}
 */
function insertListOfTablesFromSidebar() {
  CaptionManager.autoRenumberTablesIfNeeded();
  return ListGenerator.generateListOfTables();
}

/**
 * Inserts a new list of figures at the cursor from the sidebar
 * @return {Object}
 */
function insertListOfFiguresFromSidebar() {
  CaptionManager.autoRenumberFiguresIfNeeded();
  return ListGenerator.generateListOfFigures();
}

/**
 * Updates every existing List of Tables block in the document
 * @return {Object}
 */
function updateListOfTablesFromSidebar() {
  return ListGenerator.updateAllListsOfTables();
}

/**
 * Updates every existing List of Figures block in the document
 * @return {Object}
 */
function updateListOfFiguresFromSidebar() {
  return ListGenerator.updateAllListsOfFigures();
}
