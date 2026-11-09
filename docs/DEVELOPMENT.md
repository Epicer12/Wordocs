# Development Guide

Quick reference for team members working on this project.

## First Time Setup

```bash
git clone https://github.com/Epicer12/Wordocs.git
cd Wordocs
npm install
cp .clasp.json.example .clasp.json
# Edit .clasp.json with your Apps Script project ID
npx clasp login
npm run push
npm run open
```

## Daily Workflow

```bash
git checkout dev && git pull origin dev # Get latest dev
npm install                             # Ensure dependencies are up to date
git checkout -b feature/your-feature    # New branch off dev
# ... make changes ...
npm run validate                        # Lint JS and HTML (same as CI)
npm run push && npm run open            # Test
git add . && git commit -m "message"    # Commit
git push origin feature/your-feature    # Push
# Open a PR into dev on GitHub (main only receives release merges from dev)
```

## Quick Commands

```bash
npm run push        # Upload code to Apps Script
npm run open        # Open in browser
npm run pull        # Download from Apps Script (rarely needed)
npx clasp logs      # View execution logs
```

## Project Structure
src/
├── Code.js              # Menu & entry point
├── CaptionManager.js    # Caption logic
├── ListGenerator.js     # List generation
├── CrossRef.js          # Cross-references
├── Sidebar.html         # Caption sidebar (table/figure) with its own styles
├── CrossRefSidebar.html # Cross-reference sidebar
└── appsscript.json      # Apps Script manifest

## Testing Checklist

- [ ] Add table caption (cursor in table)
- [ ] Add figure caption (cursor near image)
- [ ] Insert and update List of Tables
- [ ] Insert and update List of Figures
- [ ] Delete or reorder a caption, reopen the doc, and confirm renumbering
- [ ] Insert a cross-reference, renumber, and confirm the label updates
- [ ] Caption formatting toolbar and "Apply style to all captions"
- [ ] Multiple tables/figures and multiple lists
- [ ] Edge cases (empty doc, cursor outside table/image)

## Common Issues

**"clasp: command not found"**
```bash
npm install -g @google/clasp
```

**"Apps Script API not enabled"**
Go to https://script.google.com/home/usersettings

**Git conflicts**
```bash
git pull origin dev
# Fix conflicts in editor
git add . && git commit
```

## Code Style

- camelCase for functions/variables
- PascalCase for classes
- JSDoc comments on all functions
- Test before committing

See CONTRIBUTING.md for detailed guidelines.