# INNOV v1.0 — Starter Web App

INNOV is a Google Apps Script web app backed by your existing Google Sheets spreadsheet. This is a beginner-friendly starter that includes an Editorial Board portal, a Members portal, staffing tracking, and manual acknowledgement receipt sending.

## Included in this starter

- Editorial Board: dashboard, review requests, approve/reject flow.
- Members: view approved events and volunteer for one of four coverage roles.
- Staffing: view lineups and vacancies.
- Receipts: manual send action only; the board confirms before sending when roles are vacant.
- Google Sheets: uses the existing spreadsheet and creates `INNOV_Events`, `INNOV_Members`, `INNOV_Assignments`, and `INNOV_EmailLog` tabs.
- Outlook: Microsoft Graph integration, disabled until configured.

## Important limitations of v1

1. **Authentication and authorization are not yet production-ready.** The navigation separates the two portals for the starter demo, but it is not secure role-based access. Do not deploy publicly or put sensitive data in it until you add and test access controls. Restrict the Apps Script deployment to your organization's Google accounts where possible.
2. The member portal currently asks a volunteer to type their name and email. A production version should verify identity and role against a managed Members list.
3. Form header mapping is heuristic. After first setup, verify the field mapping in `syncFormResponses()` against the actual headers in your Google Form response sheet.
4. A role is considered filled if at least one member volunteers for it. If your workflow needs multiple photographers or writers, add a required headcount per role.
5. Outlook sending requires a Microsoft Entra app registration and admin-approved Microsoft Graph `Mail.Send` application permission. If your school IT team cannot grant this, use an approved organizational email integration instead.

## Part A — Connect the project to Google Sheets

1. Open your existing spreadsheet:
   `https://docs.google.com/spreadsheets/d/1spanhY8DgDMqQb9-NOvf7-mKeU-BTF2rwkk8T7rSH5I/edit`
2. Choose **Extensions → Apps Script**. This creates a script project attached to your spreadsheet.
3. In the Apps Script editor, create these files:
   - `Code.gs` (replace its contents with the `Code.gs` in this folder)
   - `Index.html`
   - `Styles.html`
   - `Client.html`
4. Copy each matching file's entire contents into the Apps Script editor. The file names must match exactly (HTML files should be created as HTML files).
5. Save the project.
6. Run `getAppData` from the Apps Script editor once. Review the authorization prompt and approve the spreadsheet access. If Google warns that the app is unverified, only continue if this is your own script and you understand the permissions requested.
7. Run `syncFormResponses` once from the editor. Check the execution log for any errors. The first run creates the four `INNOV_...` tabs and imports source rows it can map.
8. Open the existing response sheet and verify that the event title, organization, contact email, date, and venue imported correctly. If the headers differ, edit the candidate header lists in `syncFormResponses()`.

### Find the correct source tab

The provided spreadsheet URL contains `gid=71866861`. The script searches for that tab ID. If the form responses are on another tab, update `RESPONSE_SHEET_GID` near the top of `Code.gs`.

### Test the workflow before using real requests

- Import a test request from the form.
- Approve it; verify it appears in Members portal.
- Volunteer for a role; verify the assignment is written to `INNOV_Assignments`.
- In Staffing & receipts, try Send receipt with vacancies. Confirm the warning appears.
- Do not configure Outlook until you're ready to test email sending with a test recipient.

## Part B — Deploy as a web app

1. In Apps Script, click **Deploy → New deployment**.
2. Select the gear icon and choose **Web app**.
3. Description: `INNOV v1.0`.
4. **Execute as:** `Me` (the script owner).
5. **Who has access:** choose the most restrictive option available for your school/team, ideally your organization's Google Workspace domain. Avoid “Anyone” for a real deployment.
6. Click **Deploy** and authorize if prompted.
7. Copy the Web app URL and open it in a browser. That is your INNOV web app.
8. When code changes later, choose **Deploy → Manage deployments → Edit** and deploy a new version.

## Part C — Configure Outlook (Microsoft Graph)

Email sending is intentionally unavailable until all properties below are configured.

Ask your Microsoft 365 administrator for help if you do not manage your organization's Entra tenant.

1. Open Microsoft Entra admin center and create an app registration for INNOV.
2. Record the **Directory (tenant) ID** and **Application (client) ID**.
3. Create a client secret and store it securely. Do not put it in HTML, frontend JavaScript, GitHub, or a shared document.
4. Under API permissions, add Microsoft Graph **Application permission** `Mail.Send` and have an administrator grant consent. Application permission can be powerful; follow your organization's policy and restrict the app to an approved sender mailbox where possible.
5. In Apps Script, open **Project Settings → Script properties → Add script property** and add:

   | Property | Value |
   |---|---|
   | `MS_TENANT_ID` | Entra tenant ID |
   | `MS_CLIENT_ID` | Application client ID |
   | `MS_CLIENT_SECRET` | Client secret |
   | `MS_SENDER_EMAIL` | Approved Outlook mailbox address |

6. Save properties and redeploy the web app.
7. Test using a controlled test request and a mailbox you own. Check the Outlook Sent Items folder and `INNOV_EmailLog`.

The script requests an app-only Microsoft Graph token each time it sends a message. If your organization prohibits app-only mail sending, do not bypass that policy; ask IT for an approved alternative.

## Data tabs

- `INNOV_Events`: imported requests plus editorial decision and receipt state.
- `INNOV_Members`: reserved for a verified member directory in a future iteration.
- `INNOV_Assignments`: event-to-member role signups.
- `INNOV_EmailLog`: record of successfully sent messages.

The source Google Form response tab is not edited by INNOV.

## Workflow rules

- New requests begin at `Pending Review`.
- Approving changes status to `Approved for Staffing`, which makes the event visible in the Members portal.
- Rejecting requires a reason and sends a rejection email through Outlook (when configured).
- Acknowledgement receipts are never automatic.
- The Editorial Board manually starts receipt sending.
- If roles remain vacant, the board must explicitly confirm sending anyway.
- Successful receipt sending updates the event status and logs the email.
- Duplicate receipts are blocked once `receiptSent` is `Yes`.

## Troubleshooting

- **Could not find the form-response tab:** confirm `RESPONSE_SHEET_GID`.
- **Could not identify the event title column:** update the header candidates in `syncFormResponses()`.
- **Outlook is not configured:** add all four Script Properties and check Graph permissions.
- **An event did not import:** verify it has a title value and was not already imported. The script uses the source row number to avoid importing the same row twice.
- **Page fails to load:** confirm the file names `Index`, `Styles`, and `Client` in Apps Script (without `.html` in the file list) and save all files.
