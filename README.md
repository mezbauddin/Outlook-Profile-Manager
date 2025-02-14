# 📩 Outlook Profile & Settings Manager

## 📌 Overview

This PowerShell script **automates the re-creation of Outlook profiles** while recovering essential configurations from the old profile. It simplifies the process for IT admins by:

- **Automatically re-creating or updating an Outlook profile**
- **Recovering user settings and key configuration files from the old profile**
- **Ensuring necessary Outlook data is available for a smooth experience**
- **Assigning default configurations on a per-account basis**
- **Setting the new profile as default and bypassing the profile selection prompt**

## 🔧 Features

✅ Automates Outlook profile re-creation for rapid troubleshooting  
✅ Recovers crucial user settings (old configurations, preferences, etc.)  
✅ Supports importing a `.prf` file for corporate deployments  
✅ Backs up and restores core Outlook configurations to maintain consistency  
✅ Applies per-account **new-message and reply/forward** defaults  
✅ Prevents profile selection prompts for a smooth user experience

## 📥 Installation & Usage

### **1️⃣ Run the Script**

Save the script and execute it in **PowerShell (Admin Mode)**:

```powershell
.\Outlook-Profile-Settings-Manager.ps1
The script will:

Close Outlook (if running).
Create or import a new Outlook profile.
Recover critical settings from the old profile.
Ensure the required configuration files are in place.
Set per-account default options.
Launch Outlook with the updated profile.
2️⃣ Profile File (.PRF) Support
If you have a .prf file (ideal for corporate deployment scenarios), place it at:

plaintext
Copy
Edit
C:\Deployment\CorporateExchangeProfile.prf
Otherwise, the script will create a basic profile using registry settings.

📌 Requirements
Administrator privileges
Windows with Outlook 2016/2019/365 installed
PowerShell 5.1 or later
🔄 Example Output
plaintext
Copy
Edit
✅ Closing Outlook if running...
✅ Importing Outlook profile...
✅ Recovering settings from C:\Temp\OldProfileBackup\user
✅ Ensuring required Outlook data is available
✅ Setting default configurations for user@example.com
✅ Launching Outlook...
🎉 Done! Old profile settings recovered and applied successfully.
🛠 Troubleshooting
Issue: Default configurations not applying?
🔹 Verify via File → Options → Mail → Settings in Outlook

Issue: No .prf file found?
🔹 The script falls back to registry-based profile creation

📜 License
This project is licensed under the MIT License.
