# Development Guide

Quick reference for team members working on this project.

## First Time Setup

```bash
git clone https://github.com/Epicer12/Wordocs.git
npm install -g @google/clasp
clasp login
clasp create --type standalone --title "Wordocs"
clasp push
clasp open
```

## Daily Workflow

```bash
git pull origin main                    # Get latest
git checkout -b feature/your-feature    # New branch
# ... make changes ...
clasp push && clasp open                # Test
git add . && git commit -m "message"    # Commit
git push origin feature/your-feature    # Push
# Create PR on GitHub
```

## Quick Commands

```bash
clasp push          # Upload code to Apps Script
clasp open          # Open in browser
clasp pull          # Download from Apps Script (rarely needed)
clasp logs          # View execution logs
```

## Project Structure
src/
├── Code.gs              # Menu & entry point
├── CaptionManager.gs    # Caption logic
├── ListGenerator.gs     # List generation
├── CrossRef.gs          # Cross-references
├── Sidebar.html         # UI
└── Styles.html          # CSS

## Testing Checklist

- [ ] Add table caption (cursor in table)
- [ ] Add figure caption (cursor near image)
- [ ] Generate List of Tables
- [ ] Generate List of Figures
- [ ] Update All Numbering
- [ ] Multiple tables/figures
- [ ] Edge cases (empty doc, etc.)

## Common Issues

**"clasp: command not found"**
```bash
npm install -g @google/clasp
```

**"Apps Script API not enabled"**
Go to https://script.google.com/home/usersettings

**Git conflicts**
```bash
git pull origin main
# Fix conflicts in editor
git add . && git commit
```

## Code Style

- camelCase for functions/variables
- PascalCase for classes
- JSDoc comments on all functions
- Test before committing

See CONTRIBUTING.md for detailed guidelines.