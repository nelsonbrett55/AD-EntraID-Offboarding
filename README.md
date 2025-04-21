# AD-EntraID-Offboarding
**Automated offboarding for Active Directory and Microsoft Entra ID (Azure AD).**
*This PowerShell script provides a GUI-based workflow to help system administrators seamlessly offboard employees from both on-prem and cloud environments.*

# ⚙️ Prerequisites
Before running the script, ensure the following PowerShell modules and assemblies are available:

**Required PowerShell Modules**
```
Install-Module ExchangeOnlineManagement
Install-Module Microsoft.Graph
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
$FormerEmployeeOU = "OU=Former Employees,OU=UserAccounts,DC=FABRIKAM,DC=COM"
```
This specifies the Organizational Unit (OU) where offboarded employees should be moved.

# 🧩 Function Overview
| **Function**                | **Description**                                                               |
|-----------------------------|-------------------------------------------------------------------------------|
| `Disable-AdAccount`         | Disables the user account in Active Directory.                                |
| `Update-DisplayName`        | Appends `"(Former Employee)"` to the user's display name.                     |
| `Remove-FromSecurityGroups` | Removes the user from all Active Directory security groups.                   |
| `Clear-Manager`             | Captures and clears the user’s manager attribute.                             |
| `Move-AdUserToOU`           | Moves the user to the "Former Employees" OU.                                  |
| `Revoke-Office365Sessions`  | Revokes all active Microsoft 365 sessions.                                    |
| `Remove-FromExchangeGroups` | Removes the user from Exchange distribution and mail-enabled security groups. |
| `Convert-MailboxToShared`   | Converts the user’s mailbox to a shared mailbox.                              |
| `Remove-MicrosoftLicenses`  | Removes all assigned Microsoft 365 licenses.                                  |
| `Delegate-MailboxAccess`    | Grants the former manager access to the user's mailbox.                       |
| `Grant-OneDriveAccess`      | Grants the former manager access to the user’s OneDrive.                      |
| `Show-HideButtons`          | Shows or hides UI buttons in the PowerShell form.                             |
| `Enable-DisableButtons`     | Enables or disables UI buttons based on user actions or conditions.           |



# 📌 Notes
The script uses a combination of Microsoft Graph and Exchange Online PowerShell for cloud operations.

All actions are initiated via a GUI with button-based interactions.

Designed for hybrid environments using both on-prem Active Directory and Microsoft 365 (Entra ID).

# 🛡 Disclaimer
Use at your own risk. Always test in a development environment before deploying in production.
