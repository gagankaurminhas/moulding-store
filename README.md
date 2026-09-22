# Moulding Store Online App

This is the first live-connected version of the Moulding Store delivery-management frontend.

Architecture:
- GitHub Pages: hosts the custom HTML/CSS/JavaScript.
- Microsoft Entra ID: signs employees in.
- Microsoft Graph: reads/writes the Excel workbook.
- OneDrive: stores `MouldingStoreDatabase.xlsx`.

## One-time workbook setup

1. Make sure `MouldingStoreDatabase.xlsx` is in OneDrive -> My files (root), or change `workbookFileName` in `config.js`.
2. Open the app while signed in as the workbook owner.
3. Go to Settings.
4. Click **Find My Workbook**.
5. Copy the Drive ID and Item ID into `config.js`.
6. Commit the updated `config.js` to GitHub.
7. Reload the GitHub Pages site.
8. Click **Test Excel Connection**.

## Employee access

Share the Excel workbook with the Microsoft 365 accounts that should use the system, with edit permission as appropriate.
Do not use an anonymous public "Anyone with the link can edit" share link.

The app uses the configured Drive ID + Item ID and the signed-in employee's delegated Microsoft Graph permission to access the same workbook.

## Important

Do not add client secrets, passwords, or API secrets to this repository.
The Client ID and Tenant ID are public application configuration identifiers.
