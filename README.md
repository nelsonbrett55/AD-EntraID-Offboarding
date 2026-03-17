# AD-EntraID-Offboarding
**Automated offboarding for Active Directory and Microsoft Entra ID (Azure AD).**
*This PowerShell script provides a GUI-based workflow to help system administrators seamlessly offboard employees from both on-prem and cloud environments.*

# ⚙️ Prerequisites
Before running the script, ensure the following PowerShell modules and assemblies are available:

**Required PowerShell Modules**
```
Install-Module ExchangeOnlineManagement
Install-Module Microsoft.Graph
Install-Module ActiveDirectory
```
**Required .NET Assemblies**
```
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
Add-Type -AssemblyName PresentationFramework
```
These are needed to support the Windows Forms-based user interface.

# 🔧 Required Configuration
Update the following variable to match your environment:
```
$FormerEmployeeOU = "OU=Former Employees,DC=CONTOSO,DC=LOCAL"
```
This specifies the Organizational Unit (OU) where offboarded employees should be moved.

# 🧩 Function Overview
| **Function**                   | **Description**                                                                 |
|--------------------------------|-------------------------------------------------------------------------------|
| `Load-User`                     | Loads an AD user and retrieves cloud data if connected.                        |
| `Disable-User`                  | Disables the user account in Active Directory.                                  |
| `Rename-User`                   | Appends `(Former Employee)` to the user's display name.                         |
| `Remove-Groups`                 | Removes user from AD groups and cloud groups (Microsoft 365 / Teams / DG).     |
| `Move-OU`                       | Moves the user to the "Former Employees" OU.                                    |
| `Revoke-Sessions`               | Revokes Microsoft 365 active sessions.                                         |
| `Remove-CloudGroups`            | Removes user from M365 cloud groups; logs groups requiring manual cleanup.     |
| `Convert-Mailbox`               | Converts the user’s mailbox to a shared mailbox.                               |
| `Remove-Licenses`               | Removes all directly assigned Microsoft 365 licenses.                           |
| `Delegate-Mailbox`              | Grants the former manager full access to the mailbox.                           |
| `Setup-AutoReply`               | Sets up an Exchange auto-reply for incoming emails with a customizable template.|
| `Verify-FormerOU`               | Audits the Former Employees OU for remaining groups, licenses, and display name.|
| `Export-Report`                 | Exports the offboarding log to a CSV file on the Desktop.                       |
| `Run-All`                       | Executes the full offboarding workflow for a selected user.                     |

# 🖥 Workflow
1. **Search & Load User** – Type a name in the search box to find and select a user.
2. **Connect to Microsoft 365** – Establish Graph and Exchange Online session.
3. **Check M365 Groups** – Scan for former employees still present in mail-enabled groups.
4. **Run Offboarding** – Perform AD and cloud cleanup:
   - Disable AD account
   - Rename account to mark former employee
   - Remove from groups (AD & cloud)
   - Move to Former Employees OU
   - Revoke M365 sessions
   - Convert mailbox to shared
   - Remove licenses
   - Delegate mailbox to manager
5. **Setup Auto Reply** – Optional: configure a custom Exchange auto-reply.
6. **Verify Former OU** – Audit the OU to ensure no residual permissions or licenses.
7. **Export Report** – Save a CSV log to Desktop for auditing.

# 💡 Notes
- All **cloud operations** require an active Microsoft 365 connection.
- Designed for **hybrid environments** using on-prem Active Directory and Microsoft 365.
- Uses **Windows Forms GUI** for intuitive button-based interactions.
- Logs actions with **colored text** for immediate feedback in the GUI.
- Includes **retry mechanisms** for transient errors during cloud operations.
- AutoReply template dynamically fills **user first name, location phone, and manager email**.

# 🛡 Disclaimer
Use at your own risk. Always test in a **development environment** before running in production. Make sure all scripts are approved by your IT security policies.
