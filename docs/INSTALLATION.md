# Installation Guide

## Quick Start (For Users)

### Option 1: Google Workspace Marketplace (Coming Soon)
The easiest way to install this extension will be through the Google Workspace Marketplace once published.

### Option 2: Manual Installation (Current Method)

Follow these steps to manually install the extension in your Google Docs:

#### Step 1: Open Apps Script Editor

1. Open a new or existing Google Docs document
2. Click **Extensions** → **Apps Script**
3. You'll see the Apps Script editor with a default `Code.gs` file

#### Step 2: Add the Script Files

1. **Delete** the default code in `Code.gs`
2. **Copy and paste** the contents of `src/Code.js` from this repository
3. Click the **+** button next to Files → **Script** to create new script files
4. Create the following files and paste their contents:
   - `CaptionManager.gs` (copy from `src/CaptionManager.js`)
   - `ListGenerator.gs` (copy from `src/ListGenerator.js`)
   - `CrossRef.gs` (copy from `src/CrossRef.js`)

#### Step 3: Add the HTML Files

1. Click the **+** button next to Files → **HTML**
2. Name it `Sidebar` (without .html extension)
3. Paste the contents of `src/Sidebar.html`
4. Repeat for `Styles` HTML file:
   - Click **+** → **HTML**
   - Name it `Styles`
   - Paste the contents of `src/Styles.html`

#### Step 4: Save and Authorize

1. Click the **Save** icon (💾) or press `Ctrl+S` / `Cmd+S`
2. Give your project a name (e.g., "Word Features Extension")
3. Close the Apps Script editor
4. **Reload** your Google Docs document (refresh the page)
5. You should now see a **"Word Features"** menu in the menu bar
6. Click on any menu item - you'll be prompted to authorize the script
7. Click **Continue** → **Select your Google account**
8. Click **Advanced** → **Go to [Project Name] (unsafe)** (this is safe - it's your own script)
9. Click **Allow**

#### Step 5: Start Using!

Once authorized, you can use all the features from the **Word Features** menu!

## Developer Installation (Using clasp)

If you're a developer and want to contribute or modify the code:

### Prerequisites

- [Node.js](https://nodejs.org/) installed
- Google account
- Basic knowledge of Google Apps Script

### Setup Steps

1. **Clone the repository**
   ```bash
   git clone https://github.com/Epicer12/Wordocs.git
   cd Wordocs
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Login to your Google account**
   ```bash
   npx clasp login
   ```

4. **Push the code to Apps Script**
   ```bash
   npm run push
   ```

6. **Open in the Apps Script editor**
   ```bash
   clasp open
   ```

7. **Test in Google Docs**
   - Open a Google Doc
   - Go to Extensions → Apps Script
   - Select your project
   - Run and test

### Development Workflow

1. Make changes to the code locally
2. Push changes: `clasp push`
3. Test in Google Docs
4. Pull any changes made in the web editor: `clasp pull`

## Troubleshooting

### "Word Features" menu doesn't appear
- Make sure you refreshed the Google Docs page after saving
- Check if the script is properly saved in Apps Script editor
- Try closing and reopening the document

### Authorization errors
- Make sure you clicked "Allow" in all authorization prompts
- Try removing the script authorization and re-authorizing:
  - Go to [myaccount.google.com/permissions](https://myaccount.google.com/permissions)
  - Remove the script and re-authorize

### Captions not inserting
- Make sure your cursor is in the right place:
  - For tables: cursor must be inside the table
  - For figures: cursor must be near or on the image
- Check the Apps Script logs for errors (View → Logs in Apps Script editor)

### Script timeout errors
- This can happen with very large documents
- Try breaking your document into smaller sections
- Contact us if this persists

## Updating the Extension

### Manual Installation
Simply replace the code in each file with the new version from the repository.

### developer Installation
```bash
git pull origin main
npm install
npm run push
```

## Need Help?

- 📖 Check the [README](../README.md) for usage instructions
- 🐛 Report bugs in [GitHub Issues](https://github.com/Epicer12/Wordocs/issues)
- 💬 Ask questions in [GitHub Discussions](https://github.com/Epicer12/Wordocs/discussions)