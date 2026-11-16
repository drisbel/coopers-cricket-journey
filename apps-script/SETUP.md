# Google Sheets backend setup

The backend code is prepared but **not deployed**. The public GitHub website continues to use browser storage until configured.

## 1. Create the spreadsheet
1. Create a blank Google Sheet named **Cooper's Cricket Journey — Database**.
2. Extensions → Apps Script.
3. Replace Code.gs with [Code.gs](Code.gs), and add an HTML file named **Editor** containing [Editor.html](Editor.html).
4. Run **setupDatabase** once and authorise the script. This creates Players, Seasons, Matches, Deliveries, ScorecardSources and AuditLog tabs.
5. In Project Settings → Script Properties set `EDITOR_EMAIL` to your own Google account email. Never place credentials in GitHub.

## 2. Deploy
Use **Deploy → New deployment → Web app**.

**Public read endpoint**:
- Execute as **Me**.
- Who has access: **Anyone**.
- Copy its `/exec` URL. Public JSONP: `?action=publicMatches&callback=callbackName`.

**Private editor**:
- Create a **second deployment** of the same script.
- Execute as **User accessing the web app**.
- Who has access: **Anyone with Google account**.
- Editor URL: `/exec?action=editor`.
- Only `EDITOR_EMAIL` is authorised by `assertEditor_`. Google identity may not be available for every account configuration; test the editor with your account and an unauthorised account before use. The editing Google account must have edit access to the Sheet.

**Important**: The public deployment runs as the owner, so the editor endpoint there will reject requests (active-user identity is unavailable). Do not weaken `assertEditor_` to get around this. <https://developers.google.com/apps-script/guides/web>

## 3. Next integration step
The GitHub frontend must be configured with the public endpoint URL and private editor URL before cloud data will appear. The public API is read-only. The private editor uses `google.script.run` inside the Google-authenticated Apps Script page; the GitHub site never receives a write token.

## 4. Existing browser records
Export a JSON backup from the existing site **before** enabling cloud sync. No automatic migration has occurred. Existing browser records must be imported separately to avoid duplicates.

## 5. Privacy and reliability
Public viewing means opponent names, match dates and scores are visible to anyone. Keep private information (child contact details, full names, personal notes) out of public fields. Google Apps Script quotas apply. Keep regular copies of the spreadsheet.

## Current limitations
- No live automatic PlayCricket fetching.
- No cross-device cloud sync until URLs are configured and tested.
- No automatic import of existing browser records yet.
- Private editor is basic manual match entry; scorecard comparison currently remains in the browser prototype.
