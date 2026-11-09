# Google Workspace Marketplace Publishing Checklist

Wordocs is an **editor add-on** (custom menu and sidebars), so its Marketplace listing is configured in the Google Cloud project and Marketplace SDK, not in an `addOns` block of `appsscript.json` (that block is only for card-based Workspace add-ons).

## In the repository (done)

- `src/appsscript.json` requests the narrowest scopes:
  - `https://www.googleapis.com/auth/documents.currentonly` — only the open document
  - `https://www.googleapis.com/auth/script.container.ui` — menus and sidebars
- `onOpen` only auto-renumbers when the add-on has full authorization (`e.authMode === FULL`), and always builds the menu.
- Privacy policy: [PRIVACY.md](PRIVACY.md)

## Still to do (needs the Google account owner)

1. **Standard Google Cloud project.** Apps Script editor → Project Settings → change the project to a standard Cloud project, and note its number.
2. **OAuth consent screen.** Set app name, support email and logo. Confirm in the scope list that `documents.currentonly` and `script.container.ui` are non-sensitive; if so, no OAuth verification review is needed (brand verification may still apply).
3. **Enable APIs:** Google Workspace Marketplace SDK.
4. **Marketplace SDK → App Configuration:** visibility Public, installation settings, Docs extension, and the Apps Script **deployment ID** (Deploy → New deployment → Add-on; use a versioned deployment).
5. **Store listing:**
   - Name, short description, detailed description
   - Icon 128×128 and banner (see Google's current size requirements)
   - Screenshots (at least one, 1280×800) of: caption sidebar, list of tables, cross-reference sidebar, Settings
   - Category: Productivity
   - Support URL: the GitHub issues page
   - Privacy policy URL: a public link to `docs/PRIVACY.md` (GitHub link works)
   - Terms of service URL (can point to the MIT `LICENSE` or a short terms page)
6. **Test as an installed add-on** before publishing: use Test deployments → install in a test doc; check menu appears, authorization prompt, auto-renumber on reopen, all sidebars.
7. **Submit for review**, then publish.

## Release flow

When `dev` is ready: merge `dev` into `main`, tag `v1.0.0` — `release.yml` creates the GitHub release.
