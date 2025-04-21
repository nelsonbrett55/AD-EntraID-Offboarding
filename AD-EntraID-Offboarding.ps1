#Import-Module ExchangeOnlineManagement
#Import-Module Microsoft.Graph

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
Add-Type -AssemblyName PresentationFramework

$FormerEmployeeOU = "OU=Former Employees,OU=UserAccounts,DC=FABRIKAM,DC=COM"

function Disable-AdAccount {
    if ($global:AdUser) {
        $global:AdUser | Disable-ADAccount
        Write-Host -F Cyan "$($global:ADName)" -NoNewline
        Write-Host -F White "'s account is now " -NoNewline
        Write-Host -F Red "Disabled"
    } else {
        write-Host -F Red "Disable-AdAccount: There's no AD User"
    }
}
function Update-DisplayName {
    if ($global:AdUser) {
        $newDisplayName = "$($global:AdUser.DisplayName) (Former Employee)"
        Set-ADUser -Identity $global:ADSam -DisplayName $newDisplayName
        Write-Host -F Cyan "$($global:ADName)" -NoNewLine
        Write-Host "'s DisplayName is now " -NoNewline
        Write-Host -F Red $newDisplayName
    } else {
        write-Host -F Red "Update-DisplayName: There's no AD User"
    }
}
function Remove-FromSecurityGroups {
    if ($global:AdUser) {
        $AllGroups = Get-ADGroup -Filter *
        foreach ($group in $AllGroups) {
            $members = Get-ADGroupMember -Identity $group
            if ($members | Where-Object { $_.SamAccountName -eq $global:ADSam }) {
                if ($group.Name -ne "Domain Users"){
                    Remove-ADGroupMember -Identity $group -Members $global:ADSam -Confirm:$false
                    Write-Host -F Cyan "$($global:ADName) was " -NoNewline
                    Write-Host -F Red "Red" -noNewLine
                    Write-Host -F White " from " -NoNewline
                    Write-Host -F Cyan $group.Name
                }
            }
        }
    } else {
        write-Host -F Red "Remove-FromSecurityGroups: There's no AD User"
    }
}
function Clear-Manager {
    if ($global:AdUser) {
        Set-ADUser -Identity $global:ADSam -Manager $null
        Write-Host -F Cyan "$($global:ADName)" -NoNewline
        Write-Host -F White "'s previous manager was " -NoNewline
        Write-Host -F Green "$($global:ADManager.DisplayName)" -NoNewline
        Write-Host -F White " but is now " -nonewLine
        write-Host -F Red "NULL" -NoNewline
    } else {
        write-Host -F Red "Clear-Manager: There's no AD User"
    }
}
function Move-AdUserToOU {
    $SourceOU = ($global:AdUser.DistinguishedName -replace "CN=.*?,", "" -replace "OU=", "\" -replace ",DC=", "\" -replace "DKB\\LOCAL", "DKB.LOCAL" -replace ",\\", "\")
    $components = $SourceOU.TrimStart('\').Split('\') | ForEach-Object { $_.Trim() }
    [array]::Reverse($components)
    $SourceOU = $components -join "\"

    $targetOU = $global:FormerEmployeeOU
    Move-ADObject -Identity $global:AdUser.DistinguishedName -TargetPath $targetOU
    
    $targetOU = ($targetOU -replace "CN=.*?,", "" -replace "OU=", "\" -replace ",DC=", "\" -replace "DKB\\LOCAL", "DKB.LOCAL" -replace ",\\", "\")
    $components = $targetOU.TrimStart('\').Split('\') | ForEach-Object { $_.Trim() }
    [array]::Reverse($components)
    $targetOU = $components -join "\"
    Write-Host -F Cyan "$($global:ADName)" -NoNewline
    Write-Host -F White " was " -NoNewline
    Write-Host -F Red "moved" -NoNewline
    Write-Host -F White " from " -NoNewline
    Write-Host -F Cyan $SourceOU -NoNewline
    Write-Host " to " -NoNewline
    Write-Host -F Red $targetOU
}
function Revoke-Office365Sessions {
    $null = Revoke-MgUserSignInSession -UserId $global:ADUser.UserPrincipalName
    Write-Host -F Cyan "$($global:ADName)" -NoNewline
    Write-Host -F White "'s Office 365 sessions were " -NoNewline
    Write-Host -F Red "Revoked"
}
function Remove-FromExchangeGroups {
    $groupMemberships = $MGGroups

    foreach ($groupMembership in $groupMemberships) {
        if ($groupMembership.AdditionalProperties.mailEnabled) {
            # ===================================================================
            # You'll have to log into Microsoft Exchange Manually and Remove them
            # ===================================================================

            #Remove-MgGroupM -GroupId $groupMembership.Id -UserId $global:MGUser.Id
            Write-Host -F Cyan "$($global:ADName)" -NoNewline
            Write-Host -F Red " could NOT be removed" -NoNewline
            Write-Host -F White " from " -NoNewline
            Write-Host -F Cyan $groupMembership.AdditionalProperties.displayName -NoNewline
            Write-Host -F White " because Microsoft doesn't allow changes to MailEnabled Groups through Graph"
        } else {
            Remove-MgGroupMemberByRef -GroupId $groupMembership.Id -DirectoryObjectId $global:MGUser.Id
            Write-Host -F Cyan "$($global:ADName) was " -NoNewline
            Write-Host -F Red "Removed" -noNewLine
            Write-Host -F White " from " -NoNewline
            Write-Host -F Cyan $groupMembership.AdditionalProperties.displayName
        }
    }
}
function Convert-MailboxToShared {
    Set-Mailbox -Identity $global:ADUser.UserPrincipalName -Type Shared
    Write-Host -F Cyan "$($global:ADName)" -NoNewline
    Write-Host -F White "'s Mailbox was changed to a " -NoNewline
    Write-Host -F Red "Shared Mailbox"
}
function Remove-MicrosoftLicenses {
    $licenses = $MGlicenses
    foreach ($license in $licenses) {
        $NULL = Set-MgUserLicense -UserId $global:ADUser.UserPrincipalName -RemoveLicenses @($license.SkuId) -AddLicenses @{}
        Write-Host -F Cyan "$($global:ADName)'s " -NoNewline
        Write-Host -F Cyan $license.SkuPartNumber -NoNewline
        Write-Host -F White " was " -NoNewline
        Write-Host -F Red "removed"
    }
}
function Delegate-MailboxAccess {
    $NULL = Add-MailboxPermission -Identity $global:ADUser.UserPrincipalName -User $global:ADManager.UserPrincipalName -AccessRights FullAccess -AutoMapping $true
    Write-Host -F Green "$($global:ADManager.DisplayName)" -NoNewline
    Write-Host -F White " now has Delegate rights on " -NoNewline
    Write-Host -F Cyan "$($global:ADName)" -NoNewline
    Write-Host -F White "'s Mailbox"
}
function Grant-OneDriveAccess {
    $Body = @{
        roles = @("owner")
        grantedTo = @{
            user = @{
                email = $global:ADUser.UserP
            }
        }
    }
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/drives/$($MgUserId)/permissions" -Body $Body
    Write-Host -F Green "$($global:ADManager.DisplayName)" -NoNewline
    Write-Host -F White " now has access to " -NoNewline
    Write-Host -F Cyan "$($global:ADName)" -NoNewline
    Write-Host -F White "'s OneDrive"
}
function Show-HideButtons {
    param (
        [bool]$show
    )
    $Button1.Visible = $show
    $Button2.Visible = $show
    $Button3.Visible = $show
    $Button4.Visible = $show
    $Button5.Visible = $show
    $Button6.Visible = $show
    $Button7.Visible = $show
    $Button8.Visible = $show
    $Button9.Visible = $show
    $Button10.Visible = $show
    $Button11.Visible = $show
    $label1.Visible = $show
    $label2.Visible = $show
    $label3.Visible = $show
    $label4.Visible = $show
    $label5.Visible = $show
    $label6.Visible = $show
    $label7.Visible = $show
    $label8.Visible = $show
    $label9.Visible = $show
    $label10.Visible = $show
    $label11.Visible = $show
}
function enable-disableButtons {
    param (
        [bool]$enable
    )
    if ($global:AdUser) {
        write-host $global:AdUser.Name
        write-host $global:AdUser.DisplayName
        write-host $global:AdUser.DistinguishedName
        write-host $global:AdUser.Manager
        if($global:AdUser.Enabled -ne $false){
            write-host "Button 1 - $enable"
            $Button1.Enabled = $enable
            $label1.Enabled = $enable
        } else {
            write-host "Button 1 - Disabled"
            $Button1.Enabled = $false
            $label1.Enabled = $false
        }
        
        if($global:AdUser.DisplayName.IndexOf("(Former Employee)") -eq -1){
            write-host "Button 2 - $enable"
            $Button2.Enabled = $enable
            $label2.Enabled = $enable
        } else {
            write-host "Button 2 - Disabled"
            $Button2.Enabled = $false
            $label2.Enabled = $false
        }
        if($global:AdUser.MemberOf.Count -gt 0){
            write-host "Button 3 - $enable"
            $Button3.Enabled = $enable
            $label3.Enabled = $enable
        } else {
            write-host "Button 3 - Disabled"
            $Button3.Enabled = $false
            $label3.Enabled = $false
        }
        if($global:ADManager -ne $null){
            write-host "Button 4 - $enable"
            $Button4.Enabled = $enable
            $label4.Enabled = $enable
        } else {
            write-host "Button 4 - Disabled"
            $Button4.Enabled = $false
            $label4.Enabled = $false
        }
        if($global:AdUser.DistinguishedName.IndexOf("OU=Former Employees") -eq -1){
            write-host "Button 5 - $enable"
            $Button5.Enabled = $enable
            $label5.Enabled = $enable
        } else {
            write-host "Button 5 - Disabled"
            $Button5.Enabled = $false
            $label5.Enabled = $false
        }

        $Button6.Enabled = $enable
        $label6.Enabled = $enable
        

        if($global:MGGroups.Count -gt 0){
            $Button7.Enabled = $enable
            $label7.Enabled = $enable
        } else {
            write-host "Button 7 - Disabled"
            $Button7.Enabled = $false
            $label7.Enabled = $false
        }

        if($global:EXUser.RecipientTypeDetails -ne "SharedMailbox"){
            $Button8.Enabled = $enable
            $label8.Enabled = $enable
        } else {
            write-host "Button 8 - Disabled"
            $Button8.Enabled = $false
            $label8.Enabled = $false
        }
        

        if($global:MGlicenses.Count -gt 0){
            $Button9.Enabled = $enable
            $label9.Enabled = $enable
        } else {
            write-host "Button 9 - Disabled"
            $Button9.Enabled = $false
            $label9.Enabled = $false
        }
        if($global:ADManager){
            if($EXAlreadyDelegated -eq $false){
                $Button10.Enabled = $enable
                $label10.Enabled = $enable
            } else {
                write-host "Button 9 - Disabled"
                $Button10.Enabled = $false
                $label10.Enabled = $false
            }
        } else {
            write-host "Button 9 - Disabled"
            $Button10.Enabled = $false
            $label10.Enabled = $false
        }

        if(1 -ne 1){
            $Button11.Enabled = $enable
            $label11.Enabled = $enable
        } else {
            $Button11.Enabled = $false
            $label11.Enabled = $false
        }
    } else {
        $Button1.Enabled = $enable
        $Button2.Enabled = $enable
        $Button3.Enabled = $enable
        $Button4.Enabled = $enable
        $Button5.Enabled = $enable
        $Button6.Enabled = $enable
        $Button7.Enabled = $enable
        $Button8.Enabled = $enable
        $Button9.Enabled = $enable
        $Button10.Enabled = $enable
        $Button11.Enabled = $enable
        $label1.Enabled = $enable
        $label2.Enabled = $enable
        $label3.Enabled = $enable
        $label4.Enabled = $enable
        $label5.Enabled = $enable
        $label6.Enabled = $enable
        $label7.Enabled = $enable
        $label8.Enabled = $enable
        $label9.Enabled = $enable
        $label10.Enabled = $enable
        $label11.Enabled = $enable
    }
}

# Create GUI Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Employee Off-boarding"
$form.Size = New-Object System.Drawing.Size(500,600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.Font = [System.Drawing.Font]::new("Segoe UI", 12)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# Add TextBox for employee username
$textbox = New-Object System.Windows.Forms.TextBox
$textbox.Location = New-Object System.Drawing.Point(10,40)
$textbox.Size = New-Object System.Drawing.Size(200,24)
$form.Controls.Add($textbox)

# Add Off Board Button to trigger off-boarding process
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Point(10,80)
$Button.Size = New-Object System.Drawing.Size(110,31)
$Button.Text = "365 Login"
$Button.Add_Click({
    $Button.Enabled = $false
    enable-disableButtons -enable $false

    # Start the background job to run the connection commands
    $job = Start-Job -ScriptBlock {
        write-host "Connecting MgGraph..." -nonewline
        $MgGraph = Connect-MgGraph -NoWelcome
        write-host -f green " Done"

        write-host "Connecting Exchange$Online..." -nonewline
        $ExchangeOnline = Connect-ExchangeOnline
        write-host -f green " Done"
    }

    write-host "Job started, entering wait loop..."

    # Wait loop for checking job completion
    while ($true) {
        $jobState = (Get-Job -Id $job.Id).State
        [System.Windows.Forms.Application]::DoEvents()

        if ($jobState -eq 'Completed') {
            $Button.Text = "Connected"
            $Button.Enabled = $false
            write-host "Job completed!"
            if ($global:ADUser) {
                $global:MGUser = Get-MgUser -UserId $global:ADUser.UserPrincipalName -Property *
                $global:MGGroups = Get-MgUserMemberOf -UserId $global:ADUser.UserPrincipalName | Select *
                $global:MGlicenses = Get-MgUserLicenseDetail -UserId $global:ADUser.UserPrincipalName | Select *
                $global:EXUser = Get-Mailbox -Filter "UserPrincipalName -eq '$($global:ADUser.UserPrincipalName)'" | Select *
                $global:EXPrem = Get-MailboxPermission -Identity $global:ADUser.UserPrincipalName

                enable-disableButtons -enable $true
                Show-HideButtons -show $true
            }
            Receive-Job -Job $job
            Remove-Job -Job $job
            break  # Exit the loop
        } elseif ($jobState -eq 'Failed') {
            write-host "Job failed."
            $Button.Text = "Login Failed"
            $Button.Enabled = $true
            enable-disableButtons -enable $false
            Remove-Job -Job $job
            break  # Exit the loop
        } else {
            $mod2 = [Math]::Round((Get-Date).ToFileTime() / 10000000) % 2
            if($mod2 -eq 0) {
                if ($Button.Text -eq "Please Wait"){
                    $Button.Text = "Please Wait."
                } elseif ($Button.Text -eq "Please Wait."){
                    $Button.Text = "Please Wait.."
                } elseif ($Button.Text -eq "Please Wait.."){
                    $Button.Text = "Please Wait..."
                } else{
                    $Button.Text = "Please Wait"
                }
            }
        }

        Start-Sleep -Milliseconds  100  # Sleep for 1 second before checking again
    }
})
$form.Controls.Add($Button)

# 0. Show-Logs
$Button0 = New-Object System.Windows.Forms.Button
$Button0.Location = New-Object System.Drawing.Point(130,80)
$Button0.Size = New-Object System.Drawing.Size(80, 31)
$Button0.Text = "Logs"
$Button0.Enabled = $false
$Button0.Add_Click({
})
$form.Controls.Add($Button0)

$tooltip = New-Object System.Windows.Forms.ToolTip
# 1. Disable-AdAccount

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(50, 120)
$label1.Size = New-Object System.Drawing.Size(325, 24)
$label1.Text = "Disable AD Account"
# Add a tooltip to the button
$tooltip.SetToolTip($label1, "Click for more info")
$form.Controls.Add($label1)

$Button1 = New-Object System.Windows.Forms.Button
$Button1.Location = New-Object System.Drawing.Point(375, 120)
$Button1.Size = New-Object System.Drawing.Size(95, 31)
$Button1.Text = "Disable"
$Button1.Add_Click({
    Disable-AdAccount
    $Button1.Enabled = $false
})
$form.Controls.Add($Button1)

# 2. Update-DisplayName
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(50, 160)
$label2.Size = New-Object System.Drawing.Size(325, 24)
$label2.Text = "Update Display Name with (Former Employee)"
$form.Controls.Add($label2)

$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Point(375, 160)
$Button2.Size = New-Object System.Drawing.Size(95, 31)
$Button2.Text = "Update"
$Button2.Add_Click({ 
    Update-DisplayName
    $Button2.Enabled = $false
})
$form.Controls.Add($Button2)

# 3. Remove From Groups
$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(50, 200)
$label3.Size = New-Object System.Drawing.Size(325, 24)
$label3.Text = "Remove From Groups"
$form.Controls.Add($label3)

$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Point(375, 200)
$Button3.Size = New-Object System.Drawing.Size(95, 31)
$Button3.Text = "Remove"
$Button3.Add_Click({ 
    Remove-FromSecurityGroups
    $Button3.Enabled = $false
})
$form.Controls.Add($Button3)

# 4. Clear Manager
$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(50, 240)
$label4.Size = New-Object System.Drawing.Size(325, 24)
$label4.Text = "Clear Manager"
$form.Controls.Add($label4)

$Button4 = New-Object System.Windows.Forms.Button
$Button4.Location = New-Object System.Drawing.Point(375, 240)
$Button4.Size = New-Object System.Drawing.Size(95, 31)
$Button4.Text = "Clear"
$Button4.Add_Click({ 
    Clear-Manager
    $Button4.Enabled = $false
})
$form.Controls.Add($Button4)

# 5. Move To Former Employees OU
$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(50, 280)
$label5.Size = New-Object System.Drawing.Size(325, 24)
$label5.Text = "Move to Former Employees OU"
$form.Controls.Add($label5)

$Button5 = New-Object System.Windows.Forms.Button
$Button5.Location = New-Object System.Drawing.Point(375, 280)
$Button5.Size = New-Object System.Drawing.Size(95, 31)
$Button5.Text = "Move"
$Button5.Add_Click({ 
    Move-AdUserToOU
    $Button5.Enabled = $false
})
$form.Controls.Add($Button5)

# 6. Revoke Office 365 Sessions
$label6 = New-Object System.Windows.Forms.Label
$label6.Location = New-Object System.Drawing.Point(50, 320)
$label6.Size = New-Object System.Drawing.Size(325, 24)
$label6.Text = "Revoke Office 365 Sessions"
$form.Controls.Add($label6)

$Button6 = New-Object System.Windows.Forms.Button
$Button6.Location = New-Object System.Drawing.Point(375, 320)
$Button6.Size = New-Object System.Drawing.Size(95, 31)
$Button6.Text = "Revoke"
$Button6.Add_Click({ 
    Revoke-Office365Sessions
    $Button6.Enabled = $false
})
$form.Controls.Add($Button6)

# 7. Remove from Exchange Groups
$label7 = New-Object System.Windows.Forms.Label
$label7.Location = New-Object System.Drawing.Point(50, 360)
$label7.Size = New-Object System.Drawing.Size(325, 24)
$label7.Text = "Remove from Exchange Groups"
$form.Controls.Add($label7)

$Button7 = New-Object System.Windows.Forms.Button
$Button7.Location = New-Object System.Drawing.Point(375, 360)
$Button7.Size = New-Object System.Drawing.Size(95, 31)
$Button7.Text = "Remove"
$Button7.Add_Click({ 
    Remove-FromExchangeGroups
    $Button7.Enabled = $false
})
$form.Controls.Add($Button7)

# 8. Convert to Shared Mailbox
$label8 = New-Object System.Windows.Forms.Label
$label8.Location = New-Object System.Drawing.Point(50, 400)
$label8.Size = New-Object System.Drawing.Size(325, 24)
$label8.Text = "Convert to Shared Mailbox"
$form.Controls.Add($label8)

$Button8 = New-Object System.Windows.Forms.Button
$Button8.Location = New-Object System.Drawing.Point(375, 400)
$Button8.Size = New-Object System.Drawing.Size(95, 31)
$Button8.Text = "Convert"
$Button8.Add_Click({ 
    Convert-MailboxToShared
    $Button8.Enabled = $false
})
$form.Controls.Add($Button8)

# 9. Remove Microsoft Licenses
$label9 = New-Object System.Windows.Forms.Label
$label9.Location = New-Object System.Drawing.Point(50, 440)
$label9.Size = New-Object System.Drawing.Size(325, 24)
$label9.Text = "Remove Microsoft Licenses"
$form.Controls.Add($label9)

$Button9 = New-Object System.Windows.Forms.Button
$Button9.Location = New-Object System.Drawing.Point(375, 440)
$Button9.Size = New-Object System.Drawing.Size(95, 31)
$Button9.Text = "Remove"
$Button9.Add_Click({ 
    Remove-MicrosoftLicenses
    $Button9.Enabled = $false
})
$form.Controls.Add($Button9)

# 10. Delegate Mailbox Access
$label10 = New-Object System.Windows.Forms.Label
$label10.Location = New-Object System.Drawing.Point(50, 480)
$label10.Size = New-Object System.Drawing.Size(325, 24)
$label10.Text = "Delegate Mailbox Access"
$form.Controls.Add($label10)

$Button10 = New-Object System.Windows.Forms.Button
$Button10.Location = New-Object System.Drawing.Point(375, 480)
$Button10.Size = New-Object System.Drawing.Size(95, 31)
$Button10.Text = "Delegate"
$Button10.Add_Click({ 
    Delegate-MailboxAccess 
    $Button10.Enabled = $false
})
$form.Controls.Add($Button10)

# 11. Grant OneDrive Access
$label11 = New-Object System.Windows.Forms.Label
$label11.Location = New-Object System.Drawing.Point(50, 520)
$label11.Size = New-Object System.Drawing.Size(325, 24)
$label11.Text = "Grant OneDrive Access"
$form.Controls.Add($label11)

$Button11 = New-Object System.Windows.Forms.Button
$Button11.Location = New-Object System.Drawing.Point(375, 520)
$Button11.Size = New-Object System.Drawing.Size(95, 31)
$Button11.Text = "Soon..."
$Button11.Add_Click({ 
    Grant-OneDriveAccess
    $Button11.Enabled = $false
})
$form.Controls.Add($Button11)

# Add ListView to display AD users
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(220, 40)
$listView.Size = New-Object System.Drawing.Size(250, 0)
$listView.Visible = $false
$listView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor`
[System.Windows.Forms.AnchorStyles]::Right -bor`
[System.Windows.Forms.AnchorStyles]::Bottom -bor`
[System.Windows.Forms.AnchorStyles]::Left
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$form.Controls.Add($listView)

# Add columns to the ListView
$listView.Columns.Add("Display Name", 170) | Out-Null
$listView.Columns.Add("SamAccount Name", 0) | Out-Null
# $listView.Columns[1].Width = 0  # Set the width of the SamAccount Name column to 0 to hide it

$ADSam =  "cameronm"
$displayName = ""
$ADName = ""
$ADUser = $null
$ADManager = $null
$MGUser = $null
$MGGroups = $null
$MGlicenses = $null
$EXUser = $null
$EXPrem = $null
$EXAlreadyDelegated = $false

# ListView DoubleClick event handler
$listView.add_DoubleClick({
    # Get the selected item from the list
    $selectedItem = $listView.SelectedItems[0]
    $global:displayName = $selectedItem.Text
    $global:ADSam = $selectedItem.SubItems[1].Text
    $textbox.Text = $ADSam
    $listView.Visible = $false
    
    $global:ADUser = Get-ADUser -Filter { SamAccountName -eq $ADSam } -Properties *
    $global:ADName = $global:ADUser.Name
    
    if ($global:ADUser.Manager) {
        $global:ADManager = Get-ADUser -Filter { DistinguishedName -eq $ADUser.Manager } -Properties *
        write-host "Manager's Name: " $global:ADManager.DisplayName
        $EXAlreadyDelegated = ($global:EXPrem | Where-Object { $_.AccessRights -eq "FullAccess" }).User -contains $global:ADManager.UserPrincipalName
    } else {
        $global:ADManager = $null
        write-host "No manager found for the user."
    }
    If($Button.Text -eq "Connected"){
        
        if ($global:ADUser) {
            $global:MGUser = Get-MgUser -UserId $global:ADUser.UserPrincipalName -Property *

            $global:MGGroups = Get-MgUserMemberOf -UserId $global:ADUser.UserPrincipalName | Select *
            $global:MGlicenses = Get-MgUserLicenseDetail -UserId $global:ADUser.UserPrincipalName | Select *
            $global:EXUser = Get-Mailbox -Filter "UserPrincipalName -eq '$($global:ADUser.UserPrincipalName)'" | Select *
            $global:EXPrem = Get-MailboxPermission -Identity $global:ADUser.UserPrincipalName

            # Show buttons after selecting the user
            enable-disableButtons -enable $true
        }
    } else {
        enable-disableButtons -enable $false
    }
    Show-HideButtons -show $true
})

# Function to retrieve AD users based on filter
function Get-ADUsers {
    param (
        [string]$filter
    )
    $searcher = Get-ADUser -Filter "DisplayName -like '*$filter*'"
    $listView.Items.Clear()
    $searcher | ForEach-Object {
        $item = New-Object System.Windows.Forms.ListViewItem($_.Name)
        $item.SubItems.Add($_.SamAccountName) | Out-Null
        $listView.Items.Add($item)
    }
}

# TextBox TextChanged event handler
$textbox.add_TextChanged({
    $filter = $textbox.Text
    if (-not [string]::IsNullOrEmpty($filter)) {
    
        Show-HideButtons -show $false
        $listView.Visible = $true
        $global:AdUsers = Get-ADUsers -filter $filter
        $listView.Height = [math]::Max(70,[math]::Min($form.Height - ($listView.Top * 3),32+($listView.Items.count * 25)))
    } else {
        $listView.Items.Clear()
        $listView.Visible = $false
        Show-HideButtons -show $false
    }
})

$form.Add_Shown({Get-ADUsers -filter "a"})
Show-HideButtons -show $true
enable-disableButtons -enable $true

# Display Form
$form.ShowDialog()
