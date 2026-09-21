# BBSS Google Sheets + Apps Script Setup

## Files
- `bbss_savings_pro.html` — main BBSS website (filename stays unchanged)
- `bbss_config.js` — contains only the Apps Script Web App URL
- `Code.gs` — Google Apps Script backend

## 1. Create the Google Sheet
Create a new Google Sheet, for example **BBSS Online Database**.

Open **Extensions → Apps Script**, delete the default code, paste the full content of `Code.gs`, and save.

Run the function **`setupBBSS`** once from the Apps Script editor. Google will ask for permission because the script needs to read/write the Sheet. After it finishes, open **Execution log** and copy the line:

`BBSS WRITE KEY: ...`

Keep this key private. Do not put it in GitHub or `bbss_config.js`.

The setup function automatically creates these tabs:

- README
- Meta
- Members
- Deposits
- Funds
- AnnualReview
- Investments
- PersonalInfo
- Transactions
- ProfileRequests
- PasswordResetRequests
- Notes
- Audit
- State
- Sessions

`State` is the canonical app data. Large JSON is automatically split across safe-size rows, so it also works with profile photos and a growing ledger. The other tabs are readable mirrors for checking the data. Do not manually edit the `State` rows.

## 2. Deploy Apps Script as a Web App
In Apps Script:

1. **Deploy → New deployment**
2. Type: **Web app**
3. Execute as: **Me**
4. Who has access: **Anyone**
5. Deploy
6. Copy the URL ending in `/exec`

The web app uses `doGet`/`doPost`; Google Apps Script supports deploying scripts this way and returning JSON through ContentService.

## 3. Connect the website
Open `bbss_config.js` and replace:

`PASTE_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE`

with your `/exec` URL.

Upload both `bbss_savings_pro.html` and `bbss_config.js` to the same GitHub Pages folder.

## 4. First database upload (bootstrap)
Before the online database has any BBSS data:

1. Open the website.
2. Login as Admin using the current local Admin PIN.
3. Go to **Admin → Backup / Sync**.
4. Paste the private **Write Key** from `setupBBSS`.
5. Click **Initial/Recovery Save** (or Server Save when no remote session exists).

This uploads the existing BBSS state to Google Sheets and creates the readable mirror tabs.

## 5. Normal use after bootstrap
After bootstrap, reload the site.

- Admin login is checked by the Apps Script backend.
- Member login is checked by the Apps Script backend.
- Admin edits are saved to Google Sheets automatically when a remote Admin session is active.
- Members get the latest central data when they login/reload, and the page refreshes remote data periodically.
- Member profile update requests and member password changes are sent to the backend.

## Security notes
- Keep the Apps Script **Write Key** private. It is only for initial setup/recovery and is saved only in the Admin browser's local storage.
- Do not put the Write Key in GitHub.
- `bbss_config.js` contains only the public Web App URL; that is expected to be visible.
- The Sessions tab stores hashes of session tokens rather than the raw tokens.
- Use HTTPS GitHub Pages.

## Updating Code.gs later
After changing `Code.gs`, create a **new version/deployment** or edit the existing deployment so the `/exec` endpoint uses the latest code.


## Configured Web App URL

The current `bbss_config.js` is already configured with:

`https://script.google.com/macros/s/AKfycbxFz9snA2W3y7DVE6ET4VPFz2G8OPa-QTRlTp-98y0DXe5iR58YImfPBu_uk18JfGmT/exec`

Do not put the private Write Key in `bbss_config.js` or GitHub. Enter it only in the Admin → Backup / Sync screen for the initial/recovery save.
