# Mail Merge for Batch Email

A Google Apps Script to send batch emails with rich-text, inline images, QR codes, attachments, and dynamic previews, integrated with Google Sheets and Drive.  
**(Click [this direct link](https://docs.google.com/spreadsheets/u/0/d/1u-99RviC_9bjV_RnYVvloFawD2VCyq1AR9X-Z8meRzc/copy) to make a copy of the pre-configured Google Sheets file which scripts)**

[![Version](https://img.shields.io/badge/version-9.0.2-F1C40F)](https://docs.google.com/spreadsheets/u/0/d/1u-99RviC_9bjV_RnYVvloFawD2VCyq1AR9X-Z8meRzc/copy)  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Charlotte%20Lau-0288D1?logo=linkedin)](https://www.linkedin.com/in/charlotte-lau-hk/)  
[![Website](https://img.shields.io/badge/Website-syclau.hk-D81B60)](https://www.syclau.hk)  
*Last Updated: 2025-04-15 (bug fix)*

## Overview

`MailMergeForBatchEmail` is a powerful tool for sending personalized batch emails using Google Sheets as a data source and Google Drive for attachments. It supports rich-text email bodies (via Markdown), inline images, QR codes, and attachments from a selected folder. The script includes a dynamic folder picker, email preview with navigation, and automatic rerun for large batches, making it ideal for educators, administrators, or anyone needing to send customized emails in bulk.

## What’s New in v9.0

Version 9.0 introduces several enhancements to improve functionality and user experience:

- **Added "Preview Merged Emails" Feature**: Users can now preview emails before sending, with a user-friendly interface that includes navigation to review each recipient’s email, ensuring accuracy and confidence in the batch email process.
- **Enhanced UI for Menu and Folder Picker**: The custom menu now includes emoji icons for better visual clarity, and the Folder Picker has been redesigned with a tree-like hierarchy using ideographic spaces, improving readability and navigation of folder structures.
- **Improved Auto-Rerun Mechanism**: The auto-rerun feature has been upgraded to set a trigger that resumes execution after a 6-minute timeout (applicable to all users), ensuring reliable completion of large email batches.
- **Introduced "Initialize Sheets" Feature**: A new menu option allows users to set up all required worksheets (`Dashboard`, `Datasheet`, `Settings`, `Template1`, `Usage`) with default content and sample data, simplifying the initial setup process.

## Setup

1. **Make a Copy of the Google Sheet**: Click [this link](https://docs.google.com/spreadsheets/u/0/d/1u-99RviC_9bjV_RnYVvloFawD2VCyq1AR9X-Z8meRzc/copy) to make a copy of the pre-configured Google Sheet, which includes all necessary scripts and templates.
2. **Run "Check Mail Quota" to Authorize**:
Open the copied Google Sheet, then go to the menu "Mail Merge for Batch Email" > "📬 Check Mail Quota" to activate the authorization process and grant the necessary permissions (e.g., access to Google Drive, Gmail, and Sheets).
3. **Initialize Sheets**:
From the Google Sheet, select "Mail Merge for Batch Email" > "📑 Initialize Sheets" to set up the necessary worksheets (`Dashboard`, `Datasheet`, `Settings`, `Template1`, `Usage`). The `Datasheet` will contain three lines of sample data, which you can remove, but keep the first 6 columns (A to F) for proper functionality.
4. **Configure Settings**:
Edit the `Settings` sheet (see [Settings](#settings) below).
5. **Send Emails**:
Use the menu options to pick a folder, preview emails, and send them.

## Menu Options

The script adds a custom menu to your Google Sheet:

- **📑 Initialize Sheets**: Sets up the required worksheets with default content.
- **📁 Pick Folder for Attachment**: Opens a folder picker to select a Google Drive folder for attachments, and the folder ID is automatically filled on the `Settings` sheet.
- **🔍 Preview Merged Emails**: Displays a preview of emails to be sent, with navigation to review each email.
- **📧 Send Emails Now!**: Sends emails to recipients marked in the `Datasheet`, with confirmation prompts and quota checks.
- **📬 Check Mail Quota**: Shows the remaining daily email quota.

## Settings

Edit these in the `Settings` sheet:
- **Cc**: Comma-separated list of email addresses to Cc (e.g., `alice@example.com,bob@example.com`); leave blank to disable.
- **Bcc**: Comma-separated list of email addresses to Bcc; leave blank to disable.
- **Folder ID**: ID of the Google Drive folder containing attachments; set via the folder picker or manually.
- **Send as**: The name to display as the sender (e.g., `School Admin`); leave blank to use your default name.
- **Reply to**: Email address for replies (e.g., `admin@example.com`); leave blank to use your email.
- **No reply**: Set to `TRUE` to prevent replies; `FALSE` to allow replies.
- **QR API URL**: URL for generating QR codes (default: `https://qrcode.tec-it.com/API/QRCode?quietzone=2&dpi=150&&data=`); can be customized.

**Example `Settings` Sheet:**
| Settings   | Values                          | Remarks                              |
|------------|---------------------------------|--------------------------------------|
| Cc         |                                 | comma-separated list of email addresses |
| Bcc        |                                 | comma-separated list of email addresses |
| Folder ID  | 1TLhS8KsSIBbmR9rTYQ-dCeB-hFTQm9da | ID of template folder, required for attachment |
| Send as    | School Admin                    | The name to show as sender           |
| Reply to   | admin@example.com               | Email address to receive user reply  |
| No reply   | FALSE                           | No reply: TRUE or FALSE              |
| QR API URL | https://qrcode.tec-it.com/API/QRCode?quietzone=2&dpi=150&&data= | Default QR API URL                   |

### Special Note for Datasheet

When editing the `Datasheet`:
- **Keep Required Columns**: The first 6 columns (From A to F) are required and should not be removed. Add data columns after them.
- **Field Names Must Be Unique**: Ensure all column headers (field names) are unique to avoid data conflicts.
- **Inline Images and QR Codes**: To activate inline images or QR code generation, field names must start with:
  - `imglink` for images from a URL (e.g., `imglink1`, `imglink2`).
  - `imgfile` for images from the attachment folder (e.g., `imgfile1`, `imgfile2`).
  - `qrdata` for QR code generation (e.g., `qrdata1`, `qrdata2`).

## Features
- **Rich-Text Emails**: Write email bodies in Markdown (e.g., bold, tables) for rich formatting.
- **Inline Images and QR Codes**: Embed images from URLs (`imglink`), Drive files (`imgfile`), or generate QR codes (`qrdata`).
- **Attachments**: Attach files from a selected Google Drive folder, with support for subfolders.
- **Dynamic Folder Picker**: Select attachment folders with a tree-like hierarchy display using ideographic spaces and symbols.
- **Email Preview**: Preview emails before sending, with navigation to review each recipient’s email.
- **Automatic Rerun**: Handles large batches by setting triggers for rerun after timeouts.
- **Quota Management**: Checks email quotas before sending and updates the `Dashboard` with remaining quota.
- **Detailed Logging**: Sends a completion email with logs in HTML format.

## Limitations
- **Email Quotas**: Limited to 100 emails/day for consumer accounts, 1500/day for Google Workspace Education accounts.
- **Timeout**: Limited to 6 minutes for all users as of today, but the auto-rerun feature will complete the job automatically.
- **Quotas and Runtime Reference**: See [Google’s Quotas Page](https://developers.google.com/apps-script/guides/services/quotas) for details on mail quotas and script runtime limits.
- **Attachments**: Files must exist in the selected folder; missing files are skipped with an error logged.
- **QR Codes**: Requires an internet connection to fetch QR codes via the API URL.
- **Markdown Support**: Limited to features supported by Showdown v2.1.0 (e.g., tables, underline).

## Usage Scenarios
- Send personalized exam results to students with attached PDFs and QR codes for verification.
- Distribute newsletters with inline images and attachments to a mailing list.
- Notify staff of events with rich-text emails and links to shared Drive folders.

## Credits
- Uses [Showdown v2.1.0](https://github.com/showdownjs/showdown) for Markdown-to-HTML conversion, licensed under [MIT](https://opensource.org/licenses/MIT).
- Folder picker inspired by Google Apps Script tutorials on Drive integration.
- *Note*: The email merge logic was originally adapted from Google's "Mail Merge Tutorial" in 2016, but the reference (`https://developers.google.com/apps-script/articles/mail_merge`) is no longer available (404 error as of 2025).

## License
This project is licensed under [MIT](LICENSE).

## Contributing
Suggestions and bug reports are welcome! Please open an [issue](https://github.com/charlotte-lau-hk/MailMergeForBatchEmail/issues) or submit a pull request.

