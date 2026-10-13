# TMC Go — handover

Field-visit logging, kilometres and coverage for TMC account executives, with **Go**, a Gemini assistant.
The page runs on GitHub Pages (or Netlify / HubSpot); the data lives in the Google Sheet behind an Apps Script web app.
Your Apps Script `/exec` link is already built into every page file.

## Files
| File | Where it goes |
|---|---|
| `index.html`, `404.html`, `sw.js`, `manifest.webmanifest`, `tmc-go.json`, `icon-192.png`, `icon-512.png`, `icon-maskable-512.png`, `.nojekyll` | GitHub repo root (or Netlify), exact names |
| `tmc-go.html` | HubSpot Design Manager → page template (alternative host) |
| `Code.gs` | Google Sheet → Extensions → Apps Script |
| `Visit_Intelligence_Sheet_Seed.xlsx` | Only for a brand-new sheet (File → Import → Replace) |

## One-time setup (in this order)
1. **Sheet → Settings tab → `adminPin`:** set at least 6 digits. `0000` and simple PINs are refused for sign-in.
2. **Apps Script:** paste `Code.gs` → pick **checkSetup** → **Run** → **Allow** (Gemini, Maps, sending sign-in emails). Then **Deploy → Manage deployments → ✎ → New version → Deploy** (Execute as **Me**, access **Anyone**).
3. **Upload the page files** to GitHub (Add file → Upload files → Commit). Pages: Settings → Pages → main / root.
4. **You first:** open the app → your name → *I'm a manager* → admin PIN. You're now a manager.
5. **Give each AE a way in:** Settings → Team → ✉️ (their email — codes arrive automatically) or 🔑 Code (one-time code, 24 h, paste into WhatsApp).
6. **Go:** Settings tab `aiKey` = a Gemini key (aistudio.google.com/app/apikey). Settings → Go → 🩺 Test Go & Maps should be green.

## Who can do what
- **AE:** logs, edits and deletes **their own** visits; edits their own photo/title/phone. Sign-in email is set by managers.
- **Manager:** everything — anyone's visits, notes, team, roles, sign-in codes, signing devices out, settings. No PIN once signed in.
- Wrong codes / PINs: 8 tries, then 15-minute lockout. Sessions last 60 days per device.

## Updating later
- New page files → upload with the same names → Commit (installed phones update by themselves).
- New `Code.gs` → paste → Deploy → Manage deployments → ✎ → **New version**.

## If something's off
The app shows a plain message with *Details for your manager*. Settings → Go → 🩺 Test Go & Maps checks Gemini, Maps, sign-in email and the admin PIN. For install problems, Settings → Install shows exactly which file is missing.
