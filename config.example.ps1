# ─────────────────────────────────────────────────────────────────
#  config.example.ps1 — SMTP Credentials Template
#
#  HOW TO USE:
#    1. Copy this file and rename it to: config.ps1
#    2. Fill in your real credentials below
#    3. config.ps1 is in .gitignore — it will NOT be uploaded to GitHub
#
#  Gmail App Password:
#    Go to https://myaccount.google.com/apppasswords
#    Create a password for "Mail" and paste the 16-char result below
# ─────────────────────────────────────────────────────────────────

$SmtpUser = "sender@gmail.com"           # Gmail account used to SEND the report
$SmtpPass = "xxxx xxxx xxxx xxxx"        # Gmail App Password (not your main password)
$SmtpTo   = "recipient@yourdomain.com"   # Where the hardware report is SENT to