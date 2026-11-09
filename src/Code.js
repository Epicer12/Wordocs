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
    .addItem('Refresh All', 'showRefreshAll')
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
 * Shows the Refresh All dialog. The dialog runs the two refresh steps as separate server
 * calls, and being modal it keeps the document from being edited while they run.
 */
function showRefreshAll() {
  DocumentApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('RefreshAllDialog')
      .evaluate()
      .setWidth(340)
      .setHeight(190),
    'Wordocs'
  );
}

/**
 * Refresh All, step 1: renumbers every caption and updates the cross-reference labels
 * @return {Object}
 */
function refreshAllCaptionsAndReferences() {
  var captions = CaptionManager.updateAllCaptions();
  if (!captions.success) {
    return captions;
  }

  var refs = CrossRef.updateAllReferences();
  if (!refs.success) {
    return refs;
  }

  return {
    success: true,
    message: captions.message + ' ' + refs.message
  };
}

/**
 * Refresh All, step 2: rebuilds every existing list of tables and figures.
 * Missing lists are not an error: a document may have no lists at all.
 * @return {Object}
 */
function refreshAllLists() {
  var refreshed = 0;

  try {
    if (ListGenerator.updateAllListsOfTables().success) {
      refreshed++;
    }
    if (ListGenerator.updateAllListsOfFigures().success) {
      refreshed++;
    }
  } catch (error) {
    Logger.log('Refresh all lists: ' + error);
    return { success: false, message: 'Captions were updated, but lists could not be refreshed: ' + error };
  }

  return {
    success: true,
    refreshed: refreshed,
    message: refreshed > 0 ? 'Lists updated.' : 'No lists to update.'
  };
}

/**
 * Shows the settings sidebar
 */
function showSettings() {
  DocumentApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('SettingsSidebar')
      .evaluate()
      .setTitle('Wordocs Settings')
      .setWidth(280)
  );
}

/**
 * Returns the caption label settings for the settings sidebar
 * @return {Object}
 */
function getSettingsForSidebar() {
  return CaptionManager.getSettings();
}

/**
 * Saves caption labels from the settings sidebar
 * @param {string} tablePrefix
 * @param {string} figurePrefix
 * @param {string} tablePosition - 'above' or 'below'
 * @return {Object}
 */
function saveSettingsFromSidebar(tablePrefix, figurePrefix, tablePosition) {
  return CaptionManager.saveSettings(tablePrefix, figurePrefix, tablePosition);
}

/**
 * Restores the default Table / Figure labels from the settings sidebar
 * @return {Object}
 */
function resetSettingsFromSidebar() {
  return CaptionManager.resetSettings();
}

/**
 * Refreshes existing lists of tables/figures after a label change. Runs as its own call,
 * after the labels are saved, because rebuilding lists changes the document structure.
 * @param {boolean} tableChanged
 * @param {boolean} figureChanged
 * @return {Object}
 */
function refreshListsAfterLabelChange(tableChanged, figureChanged) {
  var refreshed = 0;

  try {
    if (tableChanged && ListGenerator.updateAllListsOfTables().success) {
      refreshed++;
    }
    if (figureChanged && ListGenerator.updateAllListsOfFigures().success) {
      refreshed++;
    }
  } catch (error) {
    Logger.log('List refresh after label change: ' + error);
    return { success: false, message: 'Labels saved, but lists could not be refreshed: ' + error };
  }

  return { success: true, refreshed: refreshed };
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
    '• Insert cross-references to captions (Wordocs > Caption > Insert Cross-Reference)\n' +
    '• Refresh everything at once (Wordocs > Refresh All)\n' +
    '• Change the Table / Figure labels and table caption position (Wordocs > Settings)\n\n' +
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
