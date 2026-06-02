# 🚀 HR Payroll & Attendance Automation System  
### Zoho People + Google Sheets + Google Apps Script

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-Automation-blue)
![Zoho People](https://img.shields.io/badge/Zoho%20People-API%20Integration-red)
![Google Sheets](https://img.shields.io/badge/Google%20Sheets-HR%20Dashboard-green)
![Status](https://img.shields.io/badge/Status-Active-success)

---

## 📌 Project Overview

This project is a complete **HR Attendance, Payroll, Reimbursement, Compensation, and Intern Salary Automation System** built using **Google Apps Script**, **Google Sheets**, and **Zoho People APIs**.

The main purpose of this project is to reduce manual HR work by automatically fetching employee attendance, leave, on-duty, reimbursement, compensation, and intern data, then converting everything into structured salary sheets.

It helps HR teams manage monthly payroll calculations in a faster, cleaner, and more reliable way.

---

## ✨ Key Features

### 👨‍💼 Employee Attendance Automation

- Fetches monthly employee attendance from Zoho People.
- Supports selected month and year.
- Handles active employees.
- Handles resigned employees using Date of Exit.
- Creates a clean attendance sheet in Google Sheets.
- Automatically formats attendance statuses with colors.

### 🗓 Attendance Status Handling

The system supports multiple attendance cases, such as:

- Present
- Absent
- Weekend
- Holiday
- Casual Leave
- Sick Leave
- Earned Leave
- Leave Without Pay
- On Duty
- Optional WFH
- Compensatory Off
- Half-day combinations
- Weekend/Present combinations
- Holiday/Present combinations
- OD/Leave combinations

Example statuses:

```text
P
A
W
H
CL
SL
EL
LWP
OD
W/P
H/P
OD/CL
0.5CL/0.5P
0.5OD/0.5CL
```

---

## 💰 Draft Calculation Sheet Automation

The project automatically builds a monthly **Draft Calculation** using attendance and salary master data.

It calculates:

- Expected payable days
- Weekends
- Mandatory offs
- Expected worked days
- Total present days
- On Duty days
- Paid leaves
- Unpaid leaves
- Total payable days
- Monthly fixed salary
- Compensation amount
- Reimbursement amount
- ESI/PF deduction
- Remarks

---

## 🧾 Salary Sheet

The system also creates a final payroll sheet named:

```text
Salary Sheet
```

This sheet includes:

- Employee entity
- Employee ID
- Employee name
- Total days of month
- Total payable days
- Monthly fixed salary
- Basic Pay
- HRA
- Transport Allowance
- Special Allowance
- Month payable amount
- Compensation request
- Reimbursement
- One-time amount
- Arrears
- Referral
- Gross payable amount
- ESI/PF
- TDS
- Advance salary / other deductions
- Net payable amount
- Status
- Remarks

Manual fields are backed up month-wise, so important edited data is not lost after rebuilding the sheet.

---

## 🧑‍🎓 Intern Salary Automation

This project also supports intern payroll processing.

It can:

- Copy intern attendance from another Google Sheet.
- Copy selected intern master data.
- Add DOJ automatically.
- Sort interns by joining date.
- Build a final intern salary sheet.
- Calculate payable days.
- Handle paid leave and unpaid leave.
- Save manual intern salary fields.
- Highlight intern code mismatches.

Final intern sheet name:

```text
Stipend Sheet
```

Intern-related sheets:

```text
Intern Attendance Record-FloTrack
Intern Stipend Details
Stipend Sheet
```

---

## 🔁 Auto Backup System

The project saves important manual data into Google Drive JSON files.

This prevents data loss when sheets are refreshed.

Backup is used for:

- Draft Calculation sheet remarks
- Salary Sheet manual columns
- Stipend Sheet manual columns

Example backup folders:

```text
Draft_Calculation_remarks
Final_Employee_SS_Data
Final_Interns_data
```

---

## 🧩 Main Sheets Used

| Sheet Name | Purpose |
|---|---|
| `Attendance Record-Zoho` | Employee attendance fetched from Zoho |
| `Salary and PF Details` | Employee salary master data |
| `Draft Calculation` | Salary calculation sheet |
| `Reimbursement-Zoho` | Reimbursement records from Zoho |
| `Overtime Compensation-Zoho` | Compensation / overtime records from Zoho |
| `Salary Sheet` | Final employee salary sheet |
| `Intern Attendance Record-FloTrack` | Intern attendance copied from Google Sheet |
| `Intern Stipend Details` | Intern master data |
| `Stipend Sheet` | Final intern salary sheet |

---

## 🛠 Technologies Used

- **Google Apps Script**
- **Google Sheets**
- **Google Drive**
- **Zoho People API**
- **OAuth 2.0**
- **JavaScript**
- **JSON backup files**

---

# 📖 User Guide: How to Set Up This Project

This is the most important section if you want to use this automation in your own Google Sheet.

---

## ✅ Step 1: Create a Google Sheet

Create a new Google Sheet where this automation will run.

Recommended sheet names:

```text
Attendance Record-Zoho
Salary and PF Details
Salary sheet
Reimbursement-Zoho
Overtime Compensation-Zoho
Draft Calculation
Intern Attendance Record-FloTrack
Intern Stipend Details
Stipend Sheet
```

Some sheets can be created automatically by the script, but it is better to keep the names consistent.

---

## ✅ Step 2: Open Apps Script

In your Google Sheet:

1. Click **Extensions**
2. Click **Apps Script**
3. Delete the default code
4. Paste the project code
5. Save the project
---
## ✅ Step 3:🔐 OAuth 2.0 Scopes Required

The following Zoho People OAuth scopes are required for this automation project to work properly (Zoho API Console):

| Scope | Purpose |
|---|---|
| `ZOHOPEOPLE.attendance.all` | Access attendance records |
| `ZOHOPEOPLE.forms.READ` | Read employee form data such as Reimbursement, Compensation, etc|
| `ZOHOPEOPLE.employee.READ` | Read employee records |
| `ZOHOPEOPLE.leave.all` | Read leave balances, booked leaves, and leave type definitions i.e. CL, EL, etc |

> **Note:** After adding or changing OAuth scopes, generate a new authorization code and refresh token. Old refresh tokens will not automatically include newly added scopes.
---

## ✅ Step 4: Add Zoho API Credentials

Go to Apps Script:

1. Click **Project Settings**
2. Scroll to **Script Properties**
3. Add the following properties:

```text
CLIENT_ID
CLIENT_SECRET
REFRESH_TOKEN
```

Example:

```text
CLIENT_ID = your_zoho_client_id
CLIENT_SECRET = your_zoho_client_secret
REFRESH_TOKEN = your_zoho_refresh_token
```

⚠️ Do not hardcode these values directly in the code. Keep them in Script Properties for safety.

---

## ✅ Step 5: Add Backup Folder ID

Create a folder in Google Drive where backup JSON files will be stored.

Copy the folder ID from the folder URL.

Example Drive URL:

```text
https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQr
```

Folder ID:

```text
1AbCdEfGhIjKlMnOpQr
```

Add this in Script Properties:

```text
Hr_backup_files = your_google_drive_folder_id
```

This folder will store month-wise backups for remarks, final salary data, and intern salary data.

---

## ✅ Step 6: Add Intern Sheet URL

If you want to use intern automation, add your intern attendance source sheet URL in Script Properties:

```text
INTERN_SHEET_URL = your_intern_attendance_google_sheet_url
```

The script uses this URL to copy intern attendance into the main payroll sheet.

---

## ✅ Step 7: Install Triggers

After pasting the code, run this function once from Apps Script:

```javascript
installTriggers()
```

This creates automatic triggers for:

- Sheet edit events
- Sheet change events

These triggers help auto-save manual changes and sync employee rows.

---

## ✅ Step 8: Reload the Google Sheet

After setup, reload your Google Sheet.

You should see a custom menu:

```text
HR All Automation
```

This menu includes:

```text
Set Month & Year
Intern Sheet Selection
Run Employee Master
Run Intern Master
```

---

## ✅ Step 9: Select Month and Year

Use:

```text
HR All Automation → Set Month & Year
```

Choose the month and year for which you want to generate attendance and salary data.

---

## ✅ Step 10: Run the Master Automation

Use:

```text
HR All Automation → Run Employee Master
```

This runs the complete process:

```text
Attendance Record-Zoho
↓
Reimbursement-Zoho
↓
Overtime Compensation-Zoho
↓
Draft Calculation
↓
Salary Sheet

Run Intern Master
↓
Intern Attendance Record-FloTrack
↓
Intern Stipend Details
↓
Stipend Sheet
↓
Mismatch Highlighting
```

---

# 🔄 Automation Flow

```text
Start
  ↓
Select Month & Year
  ↓
Fetch Zoho Attendance
  ↓
Fetch Leave and OD Data
  ↓
Process Attendance Status
  ↓
Build Attendance Sheet
  ↓
Fetch Reimbursement Records
  ↓
Fetch Compensation Records
  ↓
Build Draft Calculation Sheet
  ↓
Build Salary Sheet
  ↓
Copy Intern Attendance
  ↓
Copy Intern Master Data
  ↓
Build Final Intern Salary Sheet
  ↓
Highlight Intern Code Mismatches
  ↓
Save Manual Data to Drive JSON
  ↓
End
```

---

## 🧠 Important Functions

| Function Name | Purpose |
|---|---|
| `master()` | Runs the full automation |
| `getAccessToken()` | Generates Zoho access token |
| `attendance()` | Fetches attendance from Zoho |
| `reimbursement()` | Fetches reimbursement data |
| `compensation()` | Fetches compensation data |
| `buildSalarySheet()` | Builds salary calculation sheet |
| `buildFinalEmployeeSS()` | Builds final employee salary sheet |
| `copyInternAttendanceGoogleSheet()` | Copies intern attendance |
| `copySelectedColumnsFromInternSheet()` | Copies intern master data |
| `buildFinalInterns()` | Builds final intern salary sheet |
| `installTriggers()` | Installs edit/change triggers |
| `saveFinalSSData()` | Saves final employee manual data |
| `saveFinalInternsData()` | Saves final intern manual data |

---

## 🔐 Required Zoho API Access

You need Zoho People API access for:

- Attendance data
- Leave records
- On Duty requests
- Employee records
- Reimbursement records
- Compensation records
- Resigned employee 
Make sure your Zoho OAuth app has the required permissions/scopes for these modules.

---

## ⚠️ Important Notes

- Do not change sheet names unless you also update the `CONFIG` object in the script.
- Do not delete the backup folder from Google Drive.
- Do not manually rename important headers.
- Always select the correct month and year before running automation.
- Run `installTriggers()` only when setting up or resetting triggers.
- Run `initSnapshot()` once after first setup.
- Keep Zoho credentials private.

---

## 🧪 Troubleshooting

### ❌ Invalid OAuth Token

Check:

```text
CLIENT_ID
CLIENT_SECRET
REFRESH_TOKEN
```

Also make sure your Zoho API scopes are correct.

---

### ❌ Sheet Not Found

Check whether the sheet name exactly matches the name used in the script.

Example:

```text
Salary and PF Details
```

is different from:

```text
Salary PF Details
```

---

### ❌ Intern Data Not Coming

Check:

```text
INTERN_SHEET_URL
```

Also make sure the source sheet is accessible by the Google account running the script.

---

### ❌ Manual Data Lost After Refresh

Check whether your Drive backup folder property is correct:

```text
Hr_backup_files
```

The project saves manual data into JSON files inside this folder.

---

## 📁 Suggested Project Structure

```text
HR_Automation_Report/HR_Salary_Processing/    (1st Section)
├── Config_Triggers.gs
├── Sync_Registry.gs
├── Zoho_Auth_Fetch.gs
├── Attendance.gs
├── Salary.gs
├── Backup_Drive.gs
├── Interns.gs
└── Master_UI_Helpers.gs

HR_Automation_Report/FNF_Automation/          (2nd Section)
├── Code.gs

HR_Automation_Code_Report.docx
HR_Column_Reference.docx
README.md
```

You can keep everything in one Apps Script file, but separating code by feature makes the project easier to maintain.

---

## 👤 Who Can Use This Project?

This project is useful for:

- HR teams
- Payroll teams
- Startups
- Small and medium businesses
- Companies using Zoho People
- Teams using Google Sheets for payroll processing

---

## 🎯 Benefits

- Saves manual HR time
- Reduces salary calculation mistakes
- Keeps monthly payroll records organized
- Automatically handles employee and intern data
- Stores manual changes safely
- Makes salary processing faster and cleaner

---

## 📌 Project Status

This project is actively developed and customized for HR payroll automation using Zoho People and Google Sheets.

---

## ⭐ Support

If this project helps you, you can give the repository a star on GitHub.
