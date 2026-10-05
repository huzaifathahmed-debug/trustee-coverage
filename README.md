# Trustee Coverage

A personal insurance-sales CRM for work at **Amãna Takaful (Maldives) PLC**, built with **Google Sheets + Google Apps Script**.

This repository is based on the earlier Google Apps Script CRM and extends it with a faster dashboard, follow-up and renewal queues, a dedicated motor-policy register, and a separate underwriting workspace that can be shared with underwriting staff without exposing the full personal sales CRM.

## What it includes

- Sales dashboard with follow-ups, renewals, underwriting and motor renewal counts
- Customer / prospect search and editing
- Lead, quotation, document and policy statuses
- Overdue / today / upcoming follow-up queues
- 60-day policy renewal queue
- Multiple motor policies per customer
- Separate underwriting workbook with sync back to the CRM
- Stable customer IDs so motor and underwriting records remain linked
- Batched sheet writes and short-lived caching for a smoother web interface
- Mobile-friendly Apps Script web UI

## Google Sheets created / used

### `CRM`

The original CRM columns **A:X are preserved**:

1. Company Name
2. Business Name
3. Registration No.
4. Contact Person
5. Contact Number
6. Current Insurer
7. Lead Status
8. Products
9. Last Contact
10. Next Follow-up
11. Action Required
12. Quotation Status
13. Quotation Date
14. Policy Status
15. Policy Start Date
16. Policy Expiry Date
17. UW Status
18. Documents Status
19. Priority
20. UW Remarks
21. Remarks / Notes
22. Days Since Contact
23. Follow-up Status
24. Renewal Status

The CRM appends four system fields in **Y:AB**:

- Customer ID
- Updated At
- Amana Policy Check
- UW Last Updated

### `MOTOR POLICIES`

Stores multiple vehicles / policies against one CRM customer. It tracks registration, vehicle type, cover, insurer, dates, premium, quote / policy / UW / document statuses, next action and remarks.

### Separate underwriting workbook

Run `setupUnderwritingWorkspace()` or create it from the web app. The script creates **Trustee Coverage - Underwriting Workbench** as a separate spreadsheet.

This is intentionally a different Google Sheet file. Google Sheets permissions apply to the entire spreadsheet, so a single tab inside the personal CRM cannot be privately shared with underwriting while the other tabs remain hidden.

The underwriting workbook includes:

- Customer and product
- Sales request
- UW status
- Whether the customer already has a policy with Amãna
- UW response
- Reviewer
- Next action
- Due date
- Sales notes

Use **Sync UW → CRM** in the interface to pull the latest underwriting response into the matching CRM customer.

## Installation

1. Open the Google Sheet that contains your existing CRM data.
2. Open **Extensions → Apps Script**.
3. Add these files to the Apps Script project:
   - `Code.gs`
   - `MotorPolicies.gs`
   - `Underwriting.gs`
   - `Index.html`
4. Copy the contents from this repository into those files.
5. Set the Apps Script project timezone to **Indian/Maldives** if it is not already set.
6. Run `setupTrusteeCoverage()` once and approve the requested Google permissions.
7. If you still have the old six-column `Sheet1` and the new CRM is empty, optionally run `importExistingData()` once.
8. Deploy the app with **Deploy → New deployment → Web app**.
9. Execute as yourself. Choose access appropriate for your personal/work environment.

## Updating an existing CRM

`setupTrusteeCoverage()` is designed to preserve the existing A:X CRM layout and append the new system fields instead of replacing customer rows. It also creates the motor sheet and repairs the formulas / dropdowns used by the current version.

Before running a major update against production customer data, make a copy of the Google Sheet as a backup.

## Underwriting workflow

1. Open a customer in Trustee Coverage.
2. Click **Send to Underwriting**.
3. If the UW workbook does not exist yet, create it from the **Underwriting** page first.
4. Share the UW workbook with the underwriting colleague(s) who need to update it.
5. They update UW Status, Amana Policy Check, UW Response, UW Reviewer and Next Action.
6. In Trustee Coverage, open **Underwriting → Sync UW → CRM**.
7. The latest UW entry for each Customer ID updates the main CRM.

## Motor workflow

1. Save the customer in the CRM first.
2. Open the customer and click **Add Motor Policy**, or use the Motor Policies page.
3. Add each motorcycle, car, pickup or other vehicle separately.
4. Track policy expiry, quotation, documents, underwriting and next actions per vehicle.

## Important

This is a personal sales-management tool. Do not put unnecessary sensitive medical or identity data into the CRM. Use your organization’s approved systems and handling rules for official policy documents and confidential customer information.
