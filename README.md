# Student Management System

A Google Sheets-backed student records and dynamic public-form manager.

## Stack
- GitHub Pages frontend (index.html)
- Google Apps Script backend (Code.gs)
- Google Sheets database

## Important setup

### 1. Apps Script
1. Open the target spreadsheet: https://docs.google.com/spreadsheets/d/1ZnzqvJqoWTOzJ6mD25yl_YO_UzUcmP9mU6xUOa5oDlM/edit
2. Open Extensions → Apps Script.
3. Replace the editor contents with the repository's Code.gs.
4. In Project Settings → Script Properties, add ADMIN_PASSWORD and set a new, unique password. Do not reuse a password posted in chat.
5. Run setup() and authorize the script.
6. Deploy → New deployment → Web app; Execute as Me; Who has access Anyone.
7. Copy the deployed Web App URL ending in /exec.

### 2. Configure frontend
In index.html, replace PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE with the Web App URL. Commit the change.

### 3. GitHub Pages
1. Open repository Settings → Pages.
2. Choose Deploy from a branch.
3. Select main and /(root), then Save.
4. Wait for Pages deployment.

Site URL: https://smsvedang.github.io/student-management-system/

## Spreadsheet behavior
The script selects the first non-system tab as the main student sheet and detects its header row from the first 10 rows. It preserves existing columns and blank spacer columns. When a form introduces a new field, it appends a data column with one blank spacer column before it.

System tabs created by setup(): Form_Config, Form_Responses, Admin_Logs.

## Notes
- Never put the admin password in frontend code or repository files.
- Public form links work only for forms marked Published.
- Current features: dashboard, student search/edit/delete, CSV export, dynamic form creation/publication, public submissions, admin activity logging, and missing data audit with 1-click gap form generation.
- Test with a copy of the spreadsheet before bulk edits or deletions.
