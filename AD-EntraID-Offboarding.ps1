# ==========================================
# Enterprise Employee Offboarding Console
# ==========================================

Import-Module ActiveDirectory -ErrorAction Stop
#Import-Module Microsoft.Graph -ErrorAction Stop
Import-Module ExchangeOnlineManagement -ErrorAction Stop

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$FormerEmployeeOU = "OU=Former Employees,DC=CONTOSO,DC=LOCAL"

function Setup-AutoReply {

    if (-not $script:Context.Connected) {
        Add-Log "Auto Reply" "ERROR" "Not connected to Microsoft 365"
        return
    }

    if (-not $script:Context.ADUser) {
        Add-Log "Auto Reply" "ERROR" "No user loaded"
        return
    }

    $user = $script:Context.ADUser
    $manager = $script:Context.ADManager

    $dn = $user.DistinguishedName

    # -----------------------------------
    # Determine Location and Phone by OU
    # -----------------------------------

    $LocationPhone = "123-456-0001" #default
    $LocationName  = "Location1"

    if ($dn -like "*OU=Location2*") {
        $LocationPhone = "123-456-0002"
        $LocationName  = "Location2"
    }
    elseif ($dn -like "*OU=Location3*") {
        $LocationPhone = "123-456-0003"
        $LocationName  = "Location3"
    }

    $firstName = $user.GivenName
    $managerEmail = if ($manager) { $manager.UserPrincipalName }

    # -----------------------------------
    # Default Message Template
    # -----------------------------------

    $defaultMessage = @"
Thank you for your email.

$firstName is no longer with the company. Could you please call our office at $LocationPhone or email their manager at $managerEmail?

Thanks and have a great day!
"@

    # -----------------------------------
    # Popup Form
    # -----------------------------------

    $formAR = New-Object System.Windows.Forms.Form
    $formAR.Text = "Set Auto Reply - $($user.DisplayName)"
    $formAR.Size = "600,400"
    $formAR.StartPosition = "CenterParent"

    $txtBox = New-Object System.Windows.Forms.RichTextBox
    $txtBox.Dock = "Fill"
    $txtBox.Text = $defaultMessage
    $formAR.Controls.Add($txtBox)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Bottom"
    $panel.Height = 50
    $formAR.Controls.Add($panel)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Apply"
    $btnOK.Width = 100
    $btnOK.Height = 30
    $btnOK.Left = 350
    $btnOK.Top = 10
    $panel.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Width = 100
    $btnCancel.Height = 30
    $btnCancel.Left = 460
    $btnCancel.Top = 10
    $panel.Controls.Add($btnCancel)

    $btnCancel.Add_Click({ $formAR.Close() })

    $btnOK.Add_Click({

        try {
            Set-MailboxAutoReplyConfiguration `
                -Identity $user.UserPrincipalName `
                -AutoReplyState Enabled `
                -InternalMessage $txtBox.Text `
                -ExternalMessage $txtBox.Text `
                -ExternalAudience All

            Add-Log "Auto Reply" "SUCCESS" "Enabled for $($user.DisplayName) $LocationName"
        }
        catch {
            Add-Log "Auto Reply" "ERROR" $_
        }

        $formAR.Close()
    })

    $formAR.ShowDialog()
}
# ==========================================
# Context Engine
# ==========================================

$script:Context = [PSCustomObject]@{
    ADUser    = $null
    ADManager = $null
    MGUser    = $null
    MGGroups  = @()
    Licenses  = @()
    EXMailbox = $null
    Connected = $false
    Log       = @()
}

# ==========================================
# Logging System
# ==========================================

function Add-Log {
    param($Action,$Status,$Details)

    $entry = [PSCustomObject]@{
        Time = Get-Date
        Action = $Action
        Status = $Status
        Details = $Details
    }

    $script:Context.Log += $entry

    $timestamp = $entry.Time.ToString("HH:mm:ss")
    $line = "[$timestamp] [$Status] $Action — $Details`r`n"

    # Choose color based on status
    switch ($Status.ToUpper()) {
        "SUCCESS" { $color = [System.Drawing.Color]::LimeGreen }
        "WARNING" { $color = [System.Drawing.Color]::Orange }
        "ERROR"   { $color = [System.Drawing.Color]::Red }
        "INFO"    { $color = [System.Drawing.Color]::LightGray }
        default   { $color = [System.Drawing.Color]::White }
    }

    # Append colored text
    $LogBox.SelectionStart = $LogBox.TextLength
    $LogBox.SelectionLength = 0
    $LogBox.SelectionColor = $color
    $LogBox.AppendText($line)
    $LogBox.SelectionColor = $LogBox.ForeColor
    $LogBox.ScrollToCaret()
}

function Invoke-WithRetry {
    param([scriptblock]$Script,[int]$Retries=3)

    for ($i=1;$i -le $Retries;$i++) {
        try { return & $Script }
        catch {
            if ($i -eq $Retries) { throw }
            Start-Sleep 2
        }
    }
}

# ==========================================
# Cloud Connection
# ==========================================

function Connect-Cloud {
    try {
        Add-Log "Cloud Connect" "INFO" "Connecting to Microsoft 365..."
        Connect-MgGraph
        Connect-ExchangeOnline -DisableWAM
        $script:Context.Connected = $true
        Add-Log "Cloud Connect" "SUCCESS" "Connected to Microsoft 365"
        $CheckMailGroupsBtn.Enabled = $true
        $ConnectBtn.Enabled = $false
        $ExportBtn.Enabled = $true
        $btnVerify.Enabled = $true
    }
    catch {
        Add-Log "Cloud Connect" "FAILED" $_
        $CheckMailGroupsBtn.Enabled = $false
        $ConnectBtn.Enabled = $true
        $RunAllBtn.Enabled = $false
        $ExportBtn.Enabled = $false
        $btnVerify.Enabled = $false
    }
    Refresh-UI
}

# ==========================================
# Checks for former employees in mail groups
# ==========================================
function Check-MailGroups {
    $CheckMailGroupsBtn.Enabled = $false
    
    if (-not $script:Context.Connected) {
        Add-Log "Check Mail Groups" "ERROR" "Not connected to Microsoft 365"
        return
    }

    Add-Log "Check Mail Groups" "INFO" "Scanning all mail-enabled groups (Distribution + M365 + Teams)..."

    # Get former employees from OU
    $formerUsers = Get-ADUser -Filter * `
        -SearchBase $FormerEmployeeOU `
        -Properties DisplayName,UserPrincipalName

    if ($formerUsers.Count -eq 0) {
        Add-Log "Check Mail Groups" "INFO" "No users found in Former Employees OU"
        return
    }

    # Build lookup table
    $formerLookup = @{}
    foreach ($u in $formerUsers) {
        if ($u.UserPrincipalName) {
            $formerLookup[$u.UserPrincipalName.ToLower()] = $u.DisplayName
        }
    }

    $groupHits = @{}

    # ======================================================
    # 1 Distribution + Mail-Enabled Security Groups
    # ======================================================

    $distGroups = Get-DistributionGroup -ResultSize Unlimited

    foreach ($group in $distGroups) {

        try {
            $members = Get-DistributionGroupMember `
                -Identity $group.Identity `
                -ResultSize Unlimited | Select PrimarySmtpAddress

            foreach ($member in $members) {

                $smtp = $member.PrimarySmtpAddress.ToString().ToLower()

                if ($formerLookup.ContainsKey($smtp)) {

                    if (-not $groupHits.ContainsKey($group.DisplayName)) {
                        $groupHits[$group.DisplayName] = 0
                    }

                    $groupHits[$group.DisplayName]++
                }
            }
        }
        catch {
            Add-Log "Check Mail Groups" "WARNING" "Failed checking DG $($group.DisplayName)"
        }
    }

    # ======================================================
    # 2 Microsoft 365 / Public / Teams Groups
    # ======================================================

    $unifiedGroups = Get-UnifiedGroup -ResultSize Unlimited

    foreach ($group in $unifiedGroups) {

        try {
            $members = Get-UnifiedGroupLinks `
                -Identity $group.Identity `
                -LinkType Members `
                -ResultSize Unlimited

            foreach ($member in $members) {

                $smtp = $member.PrimarySmtpAddress.ToString().ToLower()

                if ($formerLookup.ContainsKey($smtp)) {

                    if (-not $groupHits.ContainsKey($group.DisplayName)) {
                        $groupHits[$group.DisplayName] = 0
                    }

                    $groupHits[$group.DisplayName]++
                }
            }
        }
        catch {
            Add-Log "Check Mail Groups" "WARNING" "Failed checking M365 group $($group.DisplayName)"
        }
    }

    # ======================================================
    # Final Output
    # ======================================================

    if ($groupHits.Count -eq 0) {
        Add-Log "Check Mail Groups" "SUCCESS" "No mail-enabled groups contain former employees"
        return
    }

    Add-Log "Check Mail Groups" "WARNING" "Mail-enabled groups containing former employees:"

    foreach ($g in $groupHits.GetEnumerator() | Sort-Object Name) {

        $count = $g.Value
        $label = if ($count -eq 1) { "former employee" } else { "former employees" }

        Add-Log "Mail Group" "WARNING" "$($g.Name) ($count $label)"
    }

    $CheckMailGroupsBtn.Enabled = $true
}

# ==========================================
# Load User
# ==========================================

function Load-User {
    param($Sam)

    try {
        $ctx = $script:Context

        $ctx.ADUser = Get-ADUser $Sam -Properties *
        $ctx.ADManager = if ($ctx.ADUser.Manager) {
            Get-ADUser $ctx.ADUser.Manager -Properties *
        }

        if ($ctx.Connected) {
            $upn = $ctx.ADUser.UserPrincipalName
            $ctx.MGUser = Invoke-WithRetry { Get-MgUser -UserId $upn }
            $ctx.MGGroups = Invoke-WithRetry { Get-MgUserMemberOf -UserId $upn }
            $ctx.Licenses = Invoke-WithRetry { Get-MgUserLicenseDetail -UserId $upn }
            $ctx.EXMailbox = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
        }

        Add-Log "Load User" "SUCCESS" $ctx.ADUser.DisplayName
        $RunAllBtn.Enabled = $true
        $AutoReplyBtn.Enabled = $true
    }
    catch {
        Add-Log "Load User" "FAILED" $_
        $RunAllBtn.Enabled = $false
        $AutoReplyBtn.Enabled = $false
    }

    Refresh-UI
}

# ==========================================
# Offboarding Actions
# ==========================================
# === Verify OU function ===
function Verify-FormerOU {

    $RunAllBtn.Enabled = $false

    $users = Get-ADUser -Filter * -SearchBase $FormerEmployeeOU -Properties *

    if ($users.Count -eq 0) {
        Add-Log "Verify Former OU" "INFO" "No users found in Former Employees OU"
        return
    }

    $total = $users.Count
    $current = 0

    Add-Log "Verify Former OU" "INFO" "Checking $total users in Former Employees OU"

    foreach ($u in $users) {

        $current++
        $upn = $u.UserPrincipalName
        $issues = @()

        # -----------------------------
        # AD Groups
        # -----------------------------
        $adGroups = Get-ADPrincipalGroupMembership $u |
            Where-Object { $_.Name -ne "Domain Users" }

        if ($adGroups.Count -gt 0) {
            $issues += "AD Groups: $($adGroups.Name -join ', ')"
        }

        # -----------------------------
        # Licenses (Graph)
        # -----------------------------
        try {
            if ($script:Context.Connected -and $upn) {
                $mgUser = Get-MgUser -UserId $upn -Property AssignedLicenses
                if ($mgUser.AssignedLicenses.Count -gt 0) {
                    $issues += "Licenses still assigned"
                }
            }
        } catch {}

        # -----------------------------
        # Cloud Groups
        # -----------------------------
        try {
            if ($script:Context.Connected -and $upn) {
                $mgGroups = Get-MgUserMemberOf -UserId $upn -All |
                    Where-Object { $_.'@odata.type' -eq "#microsoft.graph.group" }

                if ($mgGroups.Count -gt 0) {
                    $issues += "Cloud Groups: $($mgGroups.AdditionalProperties.displayName -join ', ')"
                }
            }
        } catch {}

        # -----------------------------
        # Display Name Check
        # -----------------------------
        if ($u.DisplayName -notlike "*Former Employee*") {
            $issues += "Display name not marked Former Employee"
        }

        # -----------------------------
        # Output
        # -----------------------------
        $displayName = "$($u.GivenName) $($u.Surname)"

        if ($issues.Count -eq 0) {

            Add-Log "$current/$total" "SUCCESS" `
                "$displayName has no remaining AD groups, licenses, or M365 groups."

        } else {

            Add-Log "$current/$total" "WARNING" `
                "$displayName still has: $($issues -join '; ')"
        }
    }

    Add-Log "Verify Former OU" "INFO" "Audit complete."

    $RunAllBtn.Enabled = $true
}
function Disable-User {
    $u = $script:Context.ADUser
    if ($u.Enabled) {
        Disable-ADAccount $u
        Add-Log "Disable AD" "SUCCESS" $u.SamAccountName
    }
}

function Rename-User {
    $u = $script:Context.ADUser
    if ($u.DisplayName -notlike "*Former Employee*") {
        $new = "$($u.DisplayName) (Former Employee)"
        Set-ADUser $u -DisplayName $new
        Add-Log "Rename User" "SUCCESS" $new
    }
}

function Remove-Groups {
    $u = $script:Context.ADUser
    $upn = $u.UserPrincipalName

    # 1 On-prem AD groups (not synced, not Domain Users)
    $adGroups = Get-ADPrincipalGroupMembership $u |
        Where-Object { $_.Name -ne "Domain Users" -and -not $_.IsCriticalSystemObject -and -not $_.DistinguishedName -like "*CN=Sync*" }

    foreach ($g in $adGroups) {
        try {
            Remove-ADGroupMember -Identity $g -Members $u -Confirm:$false -ErrorAction Stop
            Add-Log "Remove AD Group" "SUCCESS" "Removed $($u.SamAccountName) from $($g.Name)"
        }
        catch {
            Add-Log "Remove AD Group" "ERROR" "Failed to remove $($u.SamAccountName) from $($g.Name): $_"
        }
    }

    # 2 Azure / mail-enabled / synced groups
    if ($script:Context.Connected) {
        $mgGroups = Get-MgUserMemberOf -UserId $upn | Where-Object { $_.'@odata.type' -eq "#microsoft.graph.group" }

        foreach ($g in $mgGroups) {
            try {
                # Skip read-only synced groups
                if ($g.SecurityEnabled -and -not $g.GroupTypes.Contains("Unified")) {
                    Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($g.Id)/members/$($u.Id)/\$ref"
                    Add-Log "Remove Cloud Group" "SUCCESS" "Removed $($u.DisplayName) from $($g.DisplayName)"
                }
                else {
                    Add-Log "Remove Cloud Group" "INFO" "Skipped synced/mail group $($g.DisplayName)"
                }
            }
            catch {
                Add-Log "Remove Cloud Group" "ERROR" "Failed to remove $($u.DisplayName) from $($g.DisplayName): $_"
            }
        }
    }

    Add-Log "Remove Groups" "INFO" "Finished processing"
}

function Move-OU {
    $u = $script:Context.ADUser
    if ($u.DistinguishedName -notlike "*Former Employees*") {
        Move-ADObject $u.DistinguishedName -TargetPath $FormerEmployeeOU
        Add-Log "Move OU" "SUCCESS" ""
    }
}

function Revoke-Sessions {
    $upn = $script:Context.ADUser.UserPrincipalName
    Invoke-WithRetry { Revoke-MgUserSignInSession -UserId $upn }
    Add-Log "Revoke Sessions" "SUCCESS" ""
}

function Remove-CloudGroups {
    $ctx = $script:Context
    $manualRemoval = @()  # groups that could not be removed automatically

    foreach ($g in $ctx.MGGroups) {

        # Skip on-prem synced groups
        if ($g.AdditionalProperties.onPremisesSyncEnabled -eq $true) {
            Add-Log "Skip Synced Group" "INFO" $g.AdditionalProperties.displayName
            continue
        }

        $groupName = $g.AdditionalProperties.displayName
        $upn = $ctx.MGUser.UserPrincipalName
        $removed = $false

        # ---------------------------
        # 1. Try Microsoft Graph removal (Unified Groups / Cloud Security Groups)
        # ---------------------------
        if (-not $g.AdditionalProperties.mailEnabled -or $g.GroupTypes -contains "Unified") {
            try {
                Invoke-WithRetry {
                    Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $ctx.MGUser.Id
                }
                Add-Log "Remove Cloud Group" "SUCCESS" $groupName
                $removed = $true
            }
            catch {
                Add-Log "Remove Cloud Group" "WARNING" "Graph removal failed for $($ctx.MGUser.DisplayName) in $groupName : $_"
            }
        }

        # ---------------------------
        # 2. Try Exchange Online removal (Distribution / Mail-Enabled Security Groups)
        # ---------------------------
        if (-not $removed -and $g.AdditionalProperties.mailEnabled -eq $true) {
            try {
                # Use Remove-DistributionGroupMember for mail-enabled groups
                Remove-DistributionGroupMember -Identity $groupName -Member $upn -Confirm:$false -ErrorAction Stop
                Add-Log "Remove DG Member" "SUCCESS" $groupName
                $removed = $true
            }
            catch {
                Add-Log "Remove DG Member" "WARNING" "Exchange removal failed for $($ctx.MGUser.DisplayName) in $groupName : $_"
            }
        }

        # ---------------------------
        # 3. If still not removed, add to manual cleanup
        # ---------------------------
        if (-not $removed) {
            $manualRemoval += $groupName
        }
    }

    # ---------------------------
    # Report manual cleanup required
    # ---------------------------
    if ($manualRemoval.Count -gt 0) {
        Add-Log "Manual Cleanup Required" "WARNING" "User still in groups: $($manualRemoval -join ', ')"
    } else {
        Add-Log "Manual Cleanup Required" "SUCCESS" "User removed from all cloud groups successfully"
    }
}

function Convert-Mailbox {

    $upn = $script:Context.ADUser.UserPrincipalName

    if (-not $script:Context.EXMailbox) {
        Add-Log "Convert Mailbox" "WARNING" "No mailbox found"
        return $false
    }

    try {

        $mb = Get-Mailbox -Identity $upn -ErrorAction Stop

        if ($mb.RecipientTypeDetails -eq "SharedMailbox") {
            Add-Log "Convert Mailbox" "INFO" "Already a Shared Mailbox"
            return $true
        }

        Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop

        Start-Sleep 3

        $verify = Get-Mailbox -Identity $upn

        if ($verify.RecipientTypeDetails -eq "SharedMailbox") {
            Add-Log "Convert Mailbox" "SUCCESS" "Converted to Shared Mailbox"
            return $true
        }
        else {
            Add-Log "Convert Mailbox" "ERROR" "Mailbox conversion did not complete"
            return $false
        }
    }
    catch {
        Add-Log "Convert Mailbox" "ERROR" $_
        return $false
    }
}

function Remove-Licenses {

    $ctx = $script:Context
    $upn = $ctx.ADUser.UserPrincipalName

    # Get current licenses
    $user = Get-MgUser `
        -UserId $upn `
        -Property AssignedLicenses,LicenseAssignmentStates

    if (-not $user.AssignedLicenses -or $user.AssignedLicenses.Count -eq 0) {
        Add-Log "Remove Licenses" "INFO" "No licenses assigned"
        return
    }

    # Only remove directly assigned licenses
    $direct = @()

    foreach ($state in $user.LicenseAssignmentStates) {
        if (-not $state.AssignedByGroup) {
            $direct += $state.SkuId
        }
    }

    if ($direct.Count -eq 0) {
        Add-Log "Remove Licenses" "INFO" "Only group-based licenses detected"
        return
    }

    Add-Log "Remove Licenses" "INFO" "Removing $($direct.Count) license(s)"

    # Build raw JSON payload
    $payload = @{
        addLicenses    = @()
        removeLicenses = $direct
    } | ConvertTo-Json -Depth 5

    try {

        Invoke-WithRetry {
            Invoke-MgGraphRequest `
                -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/users/$upn/assignLicense" `
                -Body $payload `
                -ContentType "application/json"
        }

        Start-Sleep 2

        $verify = (Get-MgUser `
            -UserId $upn `
            -Property AssignedLicenses).AssignedLicenses

        if ($verify.Count -lt $user.AssignedLicenses.Count) {
            Add-Log "Remove Licenses" "SUCCESS" "Direct licenses removed"
        }
        else {
            Add-Log "Remove Licenses" "WARNING" "License removal did not change state"
        }

    }
    catch {
        Add-Log "Remove Licenses" "ERROR" $_
    }
}

function Delegate-Mailbox {
    $ctx = $script:Context
    if ($ctx.ADManager) {
        Add-MailboxPermission `
            -Identity $ctx.ADUser.UserPrincipalName `
            -User $ctx.ADManager.UserPrincipalName `
            -AccessRights FullAccess `
            -AutoMapping $true
        Add-Log "Delegate Mailbox" "SUCCESS" ""
    }
}

function Run-All {
    if (-not $script:Context.ADUser) { return }

    $confirm = [Windows.Forms.MessageBox]::Show(
        "Offboard $($script:Context.ADUser.DisplayName)?",
        "Confirm",
        "YesNo"
    )

    if ($confirm -ne "Yes") { return }

    Disable-User
    Rename-User
    Remove-Groups
    Move-OU

    if ($script:Context.Connected) {
        Revoke-Sessions
        Remove-CloudGroups
        Convert-Mailbox
        Remove-Licenses
        Delegate-Mailbox
    }

    Add-Log "OFFBOARDING COMPLETE" "SUCCESS" ""
}

# ==========================================
# Export Report
# ==========================================

function Export-Report {
    $path = "$env:USERPROFILE\Desktop\OffboardingReport.csv"
    $script:Context.Log | Export-Csv $path -NoTypeInformation
    Add-Log "Export Report" "SUCCESS" $path
}

# ==========================================
# UI State Control
# ==========================================

function Refresh-UI {
    $hasUser = $script:Context.ADUser -ne $null
    $hasCloud = $script:Context.Connected

    $ActionButtons | ForEach { $_.Enabled = $hasUser }
    $CloudButtons  | ForEach { $_.Enabled = ($hasUser -and $hasCloud) }

    $StatusLabel.Text =
        if (-not $hasUser) {"No user loaded"}
        elseif (-not $hasCloud) {"User loaded — Cloud not connected"}
        else {"Ready"}
}

# ==========================================
# AD Autocomplete Search
# ==========================================

function Search-AD {
    param($text)

    $UserList.Items.Clear()
    if ([string]::IsNullOrWhiteSpace($text)) { return }

    Get-ADUser -Filter "DisplayName -like '*$text*'" |
        Select Name,SamAccountName |
        ForEach {
            $item = New-Object Windows.Forms.ListViewItem($_.Name)
            $item.SubItems.Add($_.SamAccountName) | Out-Null
            $UserList.Items.Add($item)
        }

    $UserList.Visible = $UserList.Items.Count -gt 0
}

# ==========================================
# Setup AutoReply
# ==========================================


# ==========================================
# UI Layout
# ==========================================

$form = New-Object Windows.Forms.Form
$form.Text = "Employee Offboarding Console"
$form.Size = "900,650"
$form.Font = [System.Drawing.Font]::new("Segoe UI", 12)
$form.StartPosition = "CenterScreen"

$SearchBox = New-Object Windows.Forms.TextBox
$SearchBox.Location = "20,20"
$SearchBox.Width = 400
$form.Controls.Add($SearchBox)

$UserList = New-Object Windows.Forms.ListView
$UserList.Location = "20,50"
$UserList.Size = "300,150"
$UserList.View = "Details"
$UserList.Columns.Add("Name",200) | Out-Null
$UserList.Columns.Add("Sam",0) | Out-Null
$UserList.Visible = $false
$form.Controls.Add($UserList)

$SearchBox.Add_TextChanged({ Search-AD $SearchBox.Text })

$UserList.Add_DoubleClick({
    $sam = $UserList.SelectedItems[0].SubItems[1].Text
    Load-User $sam
    $UserList.Visible = $false
})

$ConnectBtn = New-Object Windows.Forms.Button
$ConnectBtn.Text = "Connect to M365"
$ConnectBtn.Location = "440,20"
$ConnectBtn.Size = New-Object System.Drawing.Size(200,30)
$ConnectBtn.Add_Click({ Connect-Cloud })
$form.Controls.Add($ConnectBtn)

$CheckMailGroupsBtn = New-Object Windows.Forms.Button
$CheckMailGroupsBtn.Text = "Check M365 Groups"
$CheckMailGroupsBtn.Location = "660,20"
$CheckMailGroupsBtn.Size = New-Object System.Drawing.Size(200,30)
$CheckMailGroupsBtn.Add_Click({ check-mailgroups })
$CheckMailGroupsBtn.Enabled = $false
$form.Controls.Add($CheckMailGroupsBtn)

$RunAllBtn = New-Object Windows.Forms.Button
$RunAllBtn.Text = "Run Offboarding"
$RunAllBtn.Location = "440,60"
$RunAllBtn.Size = New-Object System.Drawing.Size(200,30)
$RunAllBtn.Add_Click({ Run-All })
$RunAllBtn.Enabled = $false
$form.Controls.Add($RunAllBtn)

$AutoReplyBtn = New-Object System.Windows.Forms.Button
$AutoReplyBtn.Text = "Setup Auto Reply"
$AutoReplyBtn.Location = "660,60"
$AutoReplyBtn.Size = New-Object System.Drawing.Size(200,30)
$AutoReplyBtn.Add_Click({ Setup-AutoReply })
$AutoReplyBtn.Enabled = $false
$form.Controls.Add($AutoReplyBtn)

$ExportBtn = New-Object Windows.Forms.Button
$ExportBtn.Text = "Export Report"
$ExportBtn.Location = "440,100"
$ExportBtn.Size = New-Object System.Drawing.Size(200,30)
$ExportBtn.Add_Click({ Export-Report })
$ExportBtn.Enabled = $false
$form.Controls.Add($ExportBtn)

$btnVerify = New-Object System.Windows.Forms.Button
$btnVerify.Text = "Verify Former OU"
$btnVerify.Location = "440,140"
$btnVerify.Size = New-Object System.Drawing.Size(200,30)
$btnVerify.Add_Click({Verify-FormerOU})
$btnVerify.Enabled = $false
$form.Controls.Add($btnVerify)

$StatusLabel = New-Object Windows.Forms.Label
$StatusLabel.Location = "440,180"
$StatusLabel.Width = 400
$form.Controls.Add($StatusLabel)

$LogBox = New-Object System.Windows.Forms.RichTextBox
$LogBox.Multiline = $true
$LogBox.ReadOnly = $true
$LogBox.Location = "20,220"
$LogBox.Size = "840,360"
$LogBox.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
$LogBox.ForeColor = "White"
$LogBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($LogBox)

$ActionButtons = @($RunAllBtn)
$CloudButtons = @()

Refresh-UI

$form.ShowDialog()
