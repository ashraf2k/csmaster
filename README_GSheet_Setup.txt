# EduPlan — Google Sheets Setup

This package switches storage from browser localStorage to Google Sheets.

## What you'll set up
- **A Google Sheet** with a tab named `Plans`.
- **An Apps Script Web App** bound to that sheet to handle read/write requests.
- **Your site files** configured with your `SHEET_ID` and `WEB_APP_URL`.

---

## 1) Create the Google Sheet
1. Create a new Google Sheet (any name). Copy its **SHEET ID** from the URL:
   - The sheet URL looks like: `https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_SHEET_ID/edit#gid=0`
2. Rename (or add) a tab named **Plans**.
3. In row 1, create these headers exactly:
   ```
   Week | Day | Lesson | Subject | Class | Materials | Textbook Pages | Quiz | ID | Saved At
   ```

## 2) Add the Apps Script
1. In the Sheet, go to **Extensions → Apps Script**.
2. Paste the contents of `google_apps_script/Code.gs` into the editor (replace any starter code).
3. **Save**.

## 3) Deploy as Web App
1. Click **Deploy → Manage deployments → New deployment**.
2. Choose **Web app**.
3. **Execute as**: Me (your account).
4. **Who has access**: Anyone.
5. Deploy, then **copy the Web app URL**.

## 4) Configure the site
1. Open `config.js` and set:
   ```js
   window.EduPlanConfig = {
     SHEET_ID: "PUT_YOUR_GOOGLE_SHEET_ID_HERE",
     WEB_APP_URL: "PASTE_WEB_APP_URL_HERE",
     SHEET_TAB_NAME: "Plans"
   };
   ```
2. Upload the site files to your host (GitHub Pages, Netlify, etc.).

## 5) Test
- Open `teacher.html` → add a sample plan → Save.
- See it appear in **Saved Lesson Plans**.
- Open `parent.html` → choose Week + Class → **Display Weekly Grid**.

### Notes
- If your browser blocks cross‑origin requests, ensure the Web App is deployed with **Who has access: Anyone** (or try re‑deploying a new version).
- **Editing/Deleting** uses the unique `ID` per row. When you click **Edit** and re‑save, we call an **update** on the row with the same `ID`.
- You can still **Export Excel** in Teacher Portal — the page builds an `.xlsx` from the live data.
