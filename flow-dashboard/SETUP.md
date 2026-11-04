# FLOW REALTY x Mask360 | Live Campaign Dashboard

A client-facing dashboard that reads the Campaign Master Sheet live.

- FLOW sees status, roadmap, deliverables, inputs and updates.
- FLOW can leave feedback on any deliverable, and approve anything marked "With FLOW for review".
- Feedback and approvals write back into the sheet (Deliverables column I, plus the Updates tab).
- The Updates tab fills itself: any stage, status, link or milestone change in the sheet adds a timed entry, and the Overview "Last updated" stamp refreshes.

Master sheet: https://docs.google.com/spreadsheets/d/1TUeUwXPxSGD3Gh9ZoHyC08Pn_jIspX3HojbvvzOj4lI/edit

## Switch it on (about 5 minutes, one time)

1. Open the master sheet. Click **Extensions > Apps Script**.
2. In the editor, delete everything in `Code.gs`. Paste the full contents of `Code.gs` from this folder. Save.
3. Click **+** next to Files, choose **HTML**, name it `Index` (exactly). Paste the full contents of `Index.html`. Save.
4. Click **Deploy > New deployment**. Click the gear icon, choose **Web app**.
   - Description: `FLOW dashboard`
   - Execute as: **Me**
   - Who has access: **Anyone with the link** (so Nikita can open it without a Google login)
5. Click **Deploy**. Approve the permissions prompt (Advanced > Go to project > Allow).
6. Copy the **Web app URL**. That is the link to send FLOW.

The self-updating log works as soon as the code is saved (step 2). No extra setup.

## Adding the FLOW REALTY logo

In `Code.gs`, paste a public image link into `CLIENT_LOGO: ''`. Save, then **Deploy > Manage deployments > Edit > Version: New version > Deploy**. Until then the dashboard shows "FLOW REALTY" as text.

## Daily use

- Change stages in **3. Deliverables** and **2. Roadmap**, statuses in **4. Client Inputs**. Each change logs itself in **5. Updates** and shows on the dashboard.
- For a daily note, type it in column B of the next empty row in **5. Updates**. Date and owner fill in by themselves.
- The dashboard refreshes every 60 seconds, and on the Refresh button.

## After editing the code

Any change to `Code.gs` or `Index.html` needs **Deploy > Manage deployments > Edit > Version: New version > Deploy**. The link stays the same.
