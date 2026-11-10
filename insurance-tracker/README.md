# Insurance & Credentialing Tracker

Blank template for tracking payer enrollment, credentialing documents, claims, authorizations, superbills, and payer calls.

**Never commit real data here.** Patient and business data lives only in the Google Sheet in the BAA-covered Google Workspace account.

## Files
- `Bella256_Insurance_Credentialing_Tracker.xlsx`: the blank template
- `build_tracker.py`: generates the template. Edit this file to change tabs, columns, dropdowns, or formulas.

## Use it
1. Sign in to the BAA-covered Google Workspace account, then open Drive and upload the `.xlsx`.
2. Right-click the file → Open with → Google Sheets → File → Save as Google Sheets.
3. Delete the uploaded `.xlsx` so only one copy exists.

## Change the template
```bash
cd insurance-tracker
python3 build_tracker.py Bella256_Insurance_Credentialing_Tracker.xlsx
```
Template changes do **not** reach the live Google Sheet. Either make the same change directly in the Sheet, or re-import the new template and copy your data over.
