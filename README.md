# Coffee Intelligence Research System

Private web system for collaborative academic research management and competitive intelligence collection in the global coffee market.

## Architecture

- Frontend: GitHub Pages static app using `index.html`, `style.css`, and `script.js`.
- Backend: Google Apps Script Web App API in `GoogleAppsScript_Backend.gs`.
- Database: Google Sheets.
- API transport: hidden iframe POST bridge with `postMessage` from GitHub Pages to Apps Script.

Important: `google.script.run` only works when the page is served by Google Apps Script HTML Service. Because this project is deployed on GitHub Pages, the frontend calls the Apps Script Web App URL through a POST bridge instead.

## Official Links

- Google Sheets Database: https://docs.google.com/spreadsheets/d/1cZ7iit2zpPsE_gDcJyi2h2UBvVK64TOlMuuBR8xFflg/edit?usp=sharing
- Google Apps Script Web App: https://script.google.com/macros/s/AKfycbwwoLfaAVBdrH9l7myWTZ3rvWlvO0NZRi1cwXISK4_2RO1DV5CxpjfBlo2qRF8kMsz_/exec

## Files

```text
IC/
|-- index.html
|-- style.css
|-- script.js
|-- GoogleAppsScript_Backend.gs
|-- style.html
|-- script.html
`-- README.md
```

`style.html` and `script.html` are synchronized helper partials for Apps Script compatibility, but GitHub Pages uses `style.css` and `script.js`.

## Deploy Backend

1. Open Google Apps Script.
2. Create or open the Apps Script project connected to the Web App URL above.
3. Replace `Code.gs` with the contents of `GoogleAppsScript_Backend.gs`.
4. Run `initializeSpreadsheet()` once from the Apps Script editor.
5. Deploy as Web App:
   - Execute as: Me
   - Who has access: Anyone
6. Keep the Web App deployment URL equal to the URL configured in `script.js`.

## Deploy Frontend To GitHub Pages

1. Create a GitHub repository.
2. Upload at least these files:
   - `index.html`
   - `style.css`
   - `script.js`
3. In the repository, open Settings > Pages.
4. Select deploy from branch.
5. Choose the branch and root folder.
6. Open the generated `https://username.github.io/repository/` URL.

## Default Access

If the Users sheet is empty, the backend creates:

```text
username: admin
password: admin123
role: admin
```

Change this password immediately after first login.

## Main Features

- Required login with username and password.
- Session token stored in browser until logout or expiry.
- Admin and researcher permissions.
- Dashboard with uploads, insights, countries, users, recent activity, and simple charts.
- Uploads registry for news, reports, PDFs, spreadsheets, scientific articles, references, observations, Google Drive links, and external links.
- Modern filters by country, category, type, author, tag, and period.
- Collaborative insights with comments.
- Country monitoring with related upload and insight counts.
- Activity timeline for logins, uploads, insights, comments, and user actions.
- Admin-only user management.
- Responsive layout and dark mode.

## Google Sheets Tabs

The backend creates and normalizes these tabs automatically:

- `Users`
- `Sessions`
- `Uploads`
- `Insights`
- `Comments`
- `Activity`
- `Countries`

Existing older tabs are preserved. Missing columns are appended without deleting current data.

## Notes For GitHub Pages

- The frontend cannot use `google.script.run` on GitHub Pages.
- A hidden iframe + form POST bridge is used so the browser can call Apps Script without CORS errors.
- The Apps Script response uses `postMessage` to return data to the static frontend.
- The backend still supports JSON/JSONP from `doGet` for lightweight status checks, but the app itself uses POST for login and mutations.

## Security Notes

- Passwords created or changed by the new backend are stored as SHA-256 hashes.
- Existing plaintext passwords are accepted once and migrated to a hash after successful login.
- Only admins can create users, change roles, deactivate users, and reset passwords.
- Protect direct access to the Google Sheet and share it only with authorized collaborators.

## Local Preview

You can open `index.html` directly in a browser for a quick visual check, but API calls require the deployed Apps Script Web App to be reachable.
