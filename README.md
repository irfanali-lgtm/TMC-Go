# TMC Go — all files

Your Apps Script link is already built into every file. Nothing to edit.

| File | Use |
|---|---|
| `index.html` | The app |
| `404.html` | Copy of the app (GitHub Pages: refresh never shows an error) |
| `sw.js` | Lets the app open instantly and be installed |
| `manifest.webmanifest` | App name + icons for "Install app" |
| `tmc-go.json` | The backend link |
| `icon-192.png`, `icon-512.png`, `icon-maskable-512.png` | App icons |
| `.nojekyll` | Empty file for GitHub Pages (hidden on most computers) |
| `tmc-go.html` | The app as a **HubSpot page template** |
| `Code.gs` | The backend — paste into Google Apps Script |

**GitHub / Netlify:** upload the first nine files together (same folder, exact names).
**HubSpot page:** Design Manager → New file → HTML & HubL → Page template → paste `tmc-go.html` → Publish changes → create a Website page with the TMC Go template.
**HubSpot Files:** upload `index.html` (and the PNGs if you want them hosted too).
**Apps Script:** paste `Code.gs` → Deploy → Manage deployments → ✎ → New version → Deploy.
