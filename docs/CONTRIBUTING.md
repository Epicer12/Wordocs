# Contributing to Wordocs

Thanks for your interest in contributing! 🎉

## Getting Started

1. Fork the repository
2. Follow setup in [DEVELOPMENT.md](DEVELOPMENT.md)
3. Create a feature branch off `dev`
4. Make your changes and run `npm run validate`
5. Open a Pull Request into `dev` (not `main`)

## Code Standards

### JavaScript
```javascript
/**
 * Clear JSDoc comment describing what the function does
 * @param {string} paramName - What this parameter is
 * @return {Object} What gets returned
 */
function myFunction(paramName) {
  // Implementation
}
```

- Use camelCase for variables/functions
- Keep functions focused and small
- Add comments for complex logic

### HTML/CSS
- Semantic HTML
- Styles live in a `<style>` block inside each sidebar HTML file; follow the Google Docs sidebar look
- Accessibility-friendly

## Commit Messages

Use conventional commit format:
feat: Add cross-reference feature
fix: Caption numbering bug with deleted tables
docs: Update installation guide

## Pull Request Process

1. **Test thoroughly** in Google Docs
2. **Update docs** if adding features
3. **Fill out PR template** completely
4. **Respond to review feedback**

## Reporting Bugs

Use the bug report template and include:
- Steps to reproduce
- Expected vs actual behavior
- Screenshots if helpful
- Browser/environment info

## Suggesting Features

Use the feature request template and explain:
- The problem it solves
- Your proposed solution
- Any alternatives considered

## Questions?

Open an issue with the `question` label.

---

**Read [DEVELOPMENT.md](DEVELOPMENT.md) for setup and daily workflow details.**