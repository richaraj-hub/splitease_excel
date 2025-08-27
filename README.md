# Splitease Excel

SplitEase is a Google Sheets-based tool designed to manage group expenses easily and transparently. Whether you're planning a group vacation, trip, or any event where expenses are shared among members, this template helps everyone keep track and settle up.

## Features

- **Simple Expense Entry Form:** Members can quickly enter their expenses using a mobile-friendly form.
- **Automated Settlement Summary:** Instantly see who owes whom and how much.
- **Detailed Expense Breakdown:** Review expense details with clear split information.
- **Mobile-Ready:** Once initialized, all routine operations can be performed from mobile devices.
- **Easy Setup:** One-time setup and sheet initialization by the sheet owner from a desktop; afterward, everything runs smoothly Apon mobile.
- **Google Apps Script Integration:** Custom menu and automation through Google Apps Script.

## How to Use

### 1. Download the Template

- Copy the Google Sheet template - Splitease_template.xlsx provided in this repository.
- Save it in your Google Drive.
- Open it using Google Sheet
- Select File --> Save as Google Sheet.
- Rename the sheet

### 2. Set Up the Google Apps Script

1. Open the sheet.
2. Go to `Extensions > Apps Script`.
3. REPLACE anything there with the contents of `splitease_template_code.gs` from this repository into the Apps Script editor.
4. Save the script.
5. And close the App Script editor.

### 3. Initialize the Sheet

1. Close and reopen the Google Sheet.
2. You’ll now see a new menu item called **Expense Splitter**.
3. Click **Expense Splitter > Initialize Sheets**. You will be asked to provide Authorisation required (A script attached to this document needs your permission to run.) - This google ask your permission to the script code. Please click "OK". YOu will be navigated to multiple Autorisation windows.
4. Click Advanced at "Google hasn’t verified this app" clid "Go to Untitled project (unsafe)"
5. Click Continue to "Google will allow Untitled project to access this info about you"
6. Select All and then Continue at screen "This app hasn’t been verified by Google. Because this app is requesting ....."
7. Multiple sheets will be created for managing expenses and members.
8. A pop-up will confirm that Initialization is successful. Click Ok

### 4. Add Members

1. Go to the `Members` sheet and update the list of members.
2. Click **Expense Splitter > Refresh Forms**.
3. The list of members will be updated in both the entry form and the expense sheet.

### 5. Start Tracking Expenses

- Members can now use the form (even from mobile devices) to record expenses.
- The sheet will automatically calculate who owes whom and provide a summary for settlements.

## Prerequisites

- A Google account.
- Access to Google Sheets (Desktop required for initial setup; mobile supported thereafter).

## File Structure

- `splitwise_excel.gs`: The Google Apps Script file for menu and automation.
- Google Sheet Template: Provided in the repository or as a downloadable link.

## Contributing

Feel free to suggest improvements or report issues by opening an issue or pull request!

## License


---

**Created by richaraj-hub**
