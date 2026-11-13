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
2. **OAuth consent screen.** Set app name, support email, logo, and the home page, privacy and terms links on the Vercel site: https://wordocs.vercel.app, https://wordocs.vercel.app/privacy and https://wordocs.vercel.app/terms. Add `wordocs.vercel.app` as an authorised domain. In the console `documents.currentonly` is **non-sensitive** but `script.container.ui` is **sensitive**, so expect **OAuth app verification** (see below).
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

## OAuth app verification (because `script.container.ui` is sensitive)

Before submitting, verify `wordocs.vercel.app` in Google Search Console (URL-prefix property `https://wordocs.vercel.app/`, HTML-file method: add the file Google gives you to `docs/`, merge it so Vercel redeploys, then click Verify). The site is the `docs/` folder deployed on Vercel (Root Directory `docs`, Framework Preset `Other`). The home page, privacy policy and terms links must be on that site.

**Scope justifications** (paste into the verification form):

- `https://www.googleapis.com/auth/documents.currentonly` — "Wordocs adds numbered table and figure captions, lists of tables and figures, and cross-references to the Google Doc the user has open. It reads caption paragraphs, bookmarks and links in that document and inserts or updates them. This narrow scope is used instead of the full Documents scope so the add-on can only touch the current document."
- `https://www.googleapis.com/auth/script.container.ui` — "Wordocs shows its custom menu, sidebars and dialogs (caption, cross-reference, Settings and Refresh All) inside Google Docs. This scope is required to display that interface."

**Demo video** (unlisted YouTube, about 2 minutes, English): show (1) the Google consent screen with the app name and requested permissions, (2) the Wordocs menu, (3) adding a table caption from the sidebar, (4) inserting a cross-reference and a list of tables, (5) Settings and Refresh All. Narrate which scope each step uses.

## Release flow

When `dev` is ready: merge `dev` into `main`, tag `v1.0.0` — `release.yml` creates the GitHub release.
