# Bulk Email Sender (Local, Outlook Desktop)

A simple local Streamlit tool for sending personalized bulk emails through the
Outlook desktop app already logged in on your computer. No AI, no external
email API, no cloud service — everything stays on this machine.

## Requirements

- **Windows** with **Microsoft Outlook desktop** installed and an account
  already signed in (this app automates Outlook the same way a macro would).
- Python 3.9+

## Setup

1. Open a terminal in this folder.
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Make sure Outlook desktop is open and logged into the account you want to
   send from.
4. Run the app:
   ```
   streamlit run app.py
   ```
5. Your browser will open automatically at `http://localhost:8501`.

## How to use it

1. **Upload Excel** — upload the `.xlsx` file with your recipient list.
2. **Pick the email column** — tell the app which column holds recipient
   addresses.
3. **CC (optional)** — one fixed CC address applied to every email.
4. **Write the subject and body** — use `<<ColumnName>>` anywhere to insert
   that row's value from the matching Excel column, e.g.:
   ```
   Subject: Meeting with <<Company>>
   Body: Dear <<Name>>, ...
   ```
5. **Attachments (optional)** — upload one or more files; they're attached to
   every email sent.
6. **Validate** — the app checks:
   - the email column exists
   - every `<<Placeholder>>` matches a real column
   - each row has a non-empty, correctly formatted email address
   - each row has values for every placeholder used
   Rows with problems are listed with the reason and automatically skipped.
7. **Preview** — pick any valid row to see exactly what that recipient's
   email will look like before sending anything.
8. **Choose scope** — send to all valid rows, the first N rows, or a specific
   row range (e.g. rows 50–100).
9. **Confirm & Send** — a two-step confirmation (Prepare → Confirm & Send)
   prevents accidental bulk sends. Progress, Sent/Failed/Remaining counts,
   and failure reasons are shown live as it sends.
10. **Report** — after sending, review the full log in-app and download it as
    a CSV for your records.

## Notes

- This app only works on Windows with Outlook desktop, because it uses
  Outlook's COM automation interface (`pywin32`) to drive the already
  logged-in account — there is no server-side email sending and no API keys
  involved.
- Emails are sent one at a time using `Outlook.Application`, exactly as if
  you clicked "Send" yourself, so Outlook's own send/receive rules, signature
  settings, and account apply normally.
- If Outlook shows a security prompt when the app tries to send (some
  antivirus/Outlook security add-ins do this), click "Allow" — this is
  standard behavior for any program (including VBA macros) that sends mail
  through Outlook.
