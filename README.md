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
- ✅ Update all numbering with one click
- ✅ Customizable caption styles

### Roadmap
- 🔄 Cross-reference insertion (see Table 1, Figure 2, etc.)
- 🔄 Custom numbering formats (1.1, A, I, Roman numerals)
- 🔄 Multi-level numbering (Chapter-based)
- 🔄 Caption templates and presets

## 🚀 Installation

### From Google Workspace Marketplace (Coming Soon)
1. Visit the [Google Workspace Marketplace](#)
2. Click "Install"
3. Grant necessary permissions

## 📖 Usage

### Adding a Caption to a Table

1. Click on the table you want to caption
2. Go to **Word Features > Add Table Caption**
3. Enter your caption text in the sidebar
4. Click **Insert Caption**
5. The caption will appear below your table as "**Table 1:** Your caption text"

### Adding a Caption to an Image

1. Click on the image you want to caption
2. Go to **Word Features > Add Figure Caption**
3. Enter your caption text
4. Click **Insert Caption**
5. The caption will appear below your image as "**Figure 1:** Your caption text"

### Generating List of Tables

1. Place your cursor where you want the list
2. Go to **Word Features > Insert List of Tables**
3. A formatted list of all tables will be inserted

### Generating List of Figures

1. Place your cursor where you want the list
2. Go to **Word Features > Insert List of Figures**
3. A formatted list of all figures will be inserted

### Updating Numbering

If you add, delete, or rearrange captioned items:
1. Go to **Word Features > Update All Numbering**
2. All captions and lists will be renumbered automatically

## 🛠️ Development

### Prerequisites
- Google Account
- Basic knowledge of Google Apps Script (JavaScript)
- Familiarity with Google Docs API

### Setup for Development

1. **Clone the repository**
   ```bash
   git clone https://github.com/Epicer12/Wordocs.git
   ```

2. **Install clasp (Google Apps Script CLI)**
   ```bash
   npm install -g @google/clasp
   clasp login
   ```

3. **Create a new Apps Script project**
   ```bash
   clasp create --type docs --title "Wordocs"
   ```

4. **Push code to Apps Script**
   ```bash
   clasp push
   ```

5. **Open in Apps Script editor**
   ```bash
   clasp open
   ```

### Project Structure

```
Wordocs/
├── src/
│   ├── Code.gs              # Main entry point, menu setup
│   ├── CaptionManager.gs    # Caption insertion & numbering logic
│   ├── ListGenerator.gs     # Generate lists of tables/figures
│   ├── CrossRef.gs          # Cross-reference handling
│   ├── Sidebar.html         # Main UI sidebar
│   ├── Styles.html          # CSS styling
|   └── appsscript.json      # Apps Script manifest
├── docs/
|   ├── INSTALLATION.md      # Guide for the Development/Testing Setup
│   ├── DEVELOPMENT.md       # Guide for Contributing
│   └── CONTRIBUTING.md      # Contribution guidelines
├── .github/
│   └── workflows/           # GitHub Actions
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