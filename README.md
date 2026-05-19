# Wordocs Extension

> Bringing Microsoft Word's professional document features to Google Docs

## 🎯 Overview

This Google Docs extension adds essential professional document features that are standard in Microsoft Word but missing in Google Docs, including:

- **Table Captions** - Add numbered captions to tables
- **Figure Captions** - Add numbered captions to images
- **List of Tables** - Auto-generated table of tables with references
- **List of Figures** - Auto-generated table of figures with references
- **Cross-references** - Reference tables and figures throughout your document
- **Auto-numbering** - Automatic sequential numbering with updates

Perfect for academic papers, technical documentation, reports, and any professional document.

## ✨ Features

### Current Features (v1.0)
- ✅ Add captions to tables and images
- ✅ Automatic sequential numbering (Table 1, Table 2, Figure 1, etc.)
- ✅ Generate List of Tables
- ✅ Generate List of Figures
- ✅ Automatic caption numbering on document open
- ✅ Update all lists of tables/figures from the sidebar
- ✅ Customizable caption styles
- ✅ Basic Cross-references

### Roadmap
- 🔄 Enhanced Cross-reference management (links, auto-updates)
- 🔄 Custom numbering formats (1.1, A, I, Roman numerals)
- 🔄 Multi-level numbering (Chapter-based)
- 🔄 Caption templates and presets

## 🚀 Installation

### From Google Workspace Marketplace (Coming Soon)
1. Visit the Google Workspace Marketplace
2. Click "Install"
3. Grant necessary permissions

## 📖 Usage

### Adding a Caption to a Table

1. Click on the table you want to caption
2. Go to **Wordocs > Caption > Add Table Caption**
3. Enter your caption text in the sidebar
4. Click **Insert Caption**
5. The caption will appear below your table as "**Table 1:** Your caption text"

### Adding a Caption to an Image

1. Click on the image you want to caption
2. Go to **Wordocs > Caption > Add Figure Caption**
3. Enter your caption text
4. Click **Insert Caption**
5. The caption will appear below your image as "**Figure 1:** Your caption text"

### Generating List of Tables

1. Place your cursor where you want a new list
2. Open **Wordocs > Caption > Add Table Caption** (sidebar)
3. Click **Insert List of Tables** — adds a new list at the cursor (you can have multiple lists)
4. Click **Update List of Tables** — refreshes every existing List of Tables in the document

### Generating List of Figures

1. Place your cursor where you want a new list
2. Open **Wordocs > Caption > Add Figure Caption** (sidebar)
3. Click **Insert List of Figures** — adds a new list at the cursor
4. Click **Update List of Figures** — refreshes every existing List of Figures in the document

### Automatic Caption Numbering

Caption numbers (Table 1, Table 2, Figure 1, etc.) update automatically when you:
- Open the document
- Open a caption sidebar
- Insert a new caption

If you delete or reorder captions, reopen the document or open the sidebar to renumber. Then use **Update List of Tables/Figures** to refresh list links.

### Caption Formatting

In the caption sidebar, use the formatting toolbar to set document-wide defaults:

- **Alignment** — left, center, or right
- **Bold label** — bold the `Table N:` / `Figure N:` prefix (default: off)
- **Italic description** — italic text after the colon
- **Font size** — default 10 pt
- **Font family** — Arial, Times New Roman, and more
- **Spacing** — space before and after the caption paragraph

Changes save automatically for the document. Use **Apply style to all captions** to update existing captions of that type. Table and figure sidebars keep separate style defaults.

## 🛠️ Development

### Prerequisites
- Google Account
- Basic knowledge of Google Apps Script (JavaScript)
- Familiarity with Google Docs API

### Setup for Development

1. **Clone the repository**
   ```bash
   git clone https://github.com/Epicer12/Wordocs.git
   cd Wordocs
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Configure clasp**
   ```bash
   cp .clasp.json.example .clasp.json
   ```
   Edit `.clasp.json` and set your Apps Script project ID.

4. **Login to Google**
   ```bash
   npx clasp login
   ```

5. **Push code to Apps Script**
   ```bash
   npm run push
   ```

6. **Open in Apps Script editor**
   ```bash
   npm run open
   ```

### Project Structure

```
Wordocs/
├── src/
│   ├── Code.js              # Main entry point, menu setup
│   ├── CaptionManager.js    # Caption insertion & numbering logic
│   ├── ListGenerator.js     # Generate lists of tables/figures
│   ├── CrossRef.js          # Cross-reference handling
│   ├── Sidebar.html         # Main UI sidebar
│   ├── Styles.html          # CSS styling
│   └── appsscript.json      # Apps Script manifest
├── docs/
│   ├── INSTALLATION.md      # Guide for the Development/Testing Setup
│   ├── DEVELOPMENT.md       # Guide for Contributing
│   └── CONTRIBUTING.md      # Contribution guidelines
├── .github/
│   └── workflows/           # GitHub Actions
├── .clasp.json              # Clasp configuration
├── package.json             # NPM dependencies and scripts
├── .gitignore               # Git ignore file
├── LICENSE
└── README.md
```

## 🤝 Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](docs/CONTRIBUTING.md) for details on our code of conduct and the process for submitting pull requests.

### How to Contribute

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Inspired by Microsoft Word's caption and reference features
- Built with Google Apps Script
- Community feedback and contributions

## 📧 Contact

- **Issues**: [GitHub Issues](https://github.com/Epicer12/Wordocs/issues)
- **Discussions**: [GitHub Discussions](https://github.com/Epicer12/Wordocs/discussions)

## ⭐ Show Your Support

If this project helped you, please give it a ⭐ on GitHub!

---

**Made with ❤️ for better document creation**