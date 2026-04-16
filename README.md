# Microsoft 365 Silent Installer & System Provisioner

A fully automated PowerShell setup suite. Run one command on any machine — no USB, no manual steps.

---

## 🚀 One-Command Install

Open **PowerShell as Administrator**, then run:

```powershell
irm https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/Run-Setup.ps1 | iex
```

> Replace `YOUR_USERNAME/YOUR_REPO` with your actual GitHub repo path.

---

## 📋 What It Does

| Phase | Description |
|-------|-------------|
| **1. Credentials** | Prompts for M365 email & password (once, at the start) |
| **2. Debloat** | Removes bloatware, disables telemetry, Cortana, ad tracking |
| **3. Restore Point** | Creates a system restore point before any changes |
| **4. Hardware Report** | Collects RAM, CPU, disk, battery, GPU info → emails HTML report |
| **5. Install M365** | Silent install via Office Deployment Tool (ODT) |
| **6. Configure Account** | Pre-fills credentials, stores in Credential Manager, sets up SSO |

---

## ⚙️ Configuration (Required Before Use)

Edit the `$CONFIG` block at the top of **`Run-Setup.ps1`**:

```powershell
$CONFIG = @{
    RepoBase     = "https://raw.githubusercontent.com/dievs-vana/auto-installer/main"
    ReportEmail  = "you@yourdomain.com"       # Where to receive hardware reports
    SMTPServer   = "smtp.gmail.com"
    SMTPPort     = 587
    SMTPUser     = "sender@gmail.com"
    SMTPPass     = "xxxx xxxx xxxx xxxx"      # Gmail App Password
    ...
}
```

### Gmail App Password Setup
1. Go to [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
2. Create an App Password for "Mail"
3. Paste the 16-character password into `SMTPPass`

---

## 🗂️ Repository Structure

```
├── Run-Setup.ps1              ← Entry point (the one command you run)
├── Config/
│   └── ODT-Config.xml         ← Office Deployment Tool configuration
└── Modules/
    ├── 1-Debloat.ps1          ← Windows debloat
    ├── 2-RestorePoint.ps1     ← System restore point
    ├── 3-HardwareReport.ps1   ← Hardware diagnostics + email report
    ├── 4-InstallM365.ps1      ← Silent M365 installation via ODT
    └── 5-ConfigureM365.ps1    ← Account pre-configuration & SSO setup
```

---

## 🔐 Auto-Login Notes

Microsoft 365 uses **OAuth 2.0 / modern authentication**, which means fully silent SSO depends on the device type:

| Device Type | Auto-Login Behavior |
|-------------|-------------------|
| **Azure AD / Entra ID Joined** | Full SSO — completely silent after setup |
| **Workplace Registered** | Semi-silent — token cached after first login |
| **Workgroup / Local** | Email pre-filled — password required **once** on first Office launch |

For most standalone machines, the user will see a Microsoft sign-in screen with their email **already filled in**. After entering the password once, Office stays signed in automatically.

---

## 🛠️ Customizing the Office Install

Edit `Config/ODT-Config.xml` to change what gets installed:

- **Change channel**: `Current` (monthly) → `MonthlyEnterprise` (stable monthly) → `SemiAnnual` (6-month)
- **Add/remove apps**: Comment/uncomment `<ExcludeApp>` lines
- **Change language**: Replace `en-us` with your locale (e.g., `fil-PH`, `en-GB`)
- **Product ID options**: `O365BusinessRetail`, `O365ProPlusRetail`, `O365HomePremRetail`

---

## 📦 GitHub Setup Instructions

1. Create a new **public** GitHub repository
2. Upload all files maintaining the folder structure above
3. Update `RepoBase` in `Run-Setup.ps1` with your repo URL
4. Test by running the one-liner in a PowerShell Admin window

---

## ⚠️ Requirements

- Windows 10 / 11
- PowerShell 5.1+ (pre-installed on all modern Windows)
- Run as Administrator
- Internet connection
- Valid Microsoft 365 license/subscription

---

## 🔒 Security Notes

- Credentials entered at the prompt are **never written to disk** — held in memory only
- SMTP password in `Run-Setup.ps1` is stored in the script — consider using a **dedicated sender email** with limited permissions
- The repo should be **private** if it contains your SMTP credentials — or better yet, prompt for SMTP credentials at runtime too
