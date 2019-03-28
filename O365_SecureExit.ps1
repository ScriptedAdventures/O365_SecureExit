<#
.Synopsis
Creates Scheduled Task for running BPUpdater. This runs Daily at 11PM, however this can be changed by changing $ScheduleExecuteTime
#>

<#
.Author

Matthew Russell
https://github.com/ScriptedAdventures
https://www.scriptedadventures.net/
#>

#Environmental
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned 
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

#Form for User Selection
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $UserSelectionForm               = New-Object system.Windows.Forms.Form
    $UserSelectionForm.ClientSize    = '483,633'
    $UserSelectionForm.text          = "Select User to Exit"
    $UserSelectionForm.StartPosition = 'CenterScreen'

    $OKButton                        = New-Object system.Windows.Forms.Button
    $OKButton.text                   = "OK"
    $OKButton.width                  = 76
    $OKButton.height                 = 30
    $OKButton.location               = New-Object System.Drawing.Point(387,54)
    $OKButton.Font                   = 'Consolas,10,style=Bold'
    $OKButton.DialogResult           = [System.Windows.Forms.DialogResult]::OK
    $UserSelectionForm.AcceptButton = $OKButton
    $UserSelectionForm.Controls.Add($OKButton)

    $CancelButton                    = New-Object system.Windows.Forms.Button
    $CancelButton.text               = "Cancel"
    $CancelButton.width              = 78
    $CancelButton.height             = 30
    $CancelButton.location           = New-Object System.Drawing.Point(385,102)
    $CancelButton.Font               = 'Consolas,10,style=Bold'
    $CancelButton.DialogResult       = [System.Windows.Forms.DialogResult]::Cancel
    $UserSelectionForm.CancelButton = $CancelButton
    $UserSelectionForm.Controls.Add($CancelButton)

    $UsersListBox                    = New-Object system.Windows.Forms.ListBox
    $UsersListBox.text               = "listBox"
    $UsersListBox.width              = 342
    $UsersListBox.height             = 555
    $UsersListBox.location           = New-Object System.Drawing.Point(24,54)

    $Label                           = New-Object system.Windows.Forms.Label
    $Label.text                      = "Please Select User"
    $Label.AutoSize                  = $true
    $Label.width                     = 25
    $Label.height                    = 10
    $Label.location                  = New-Object System.Drawing.Point(24,20)
    $Label.Font                      = 'Consolas,13'

    $UserSelectionForm.controls.AddRange(@($OKButton,$CancelButton,$UsersListBox,$Label))

#Credential Request
Write-Host "Please input your Global Adminstrator Credentials"
$script:365GACreds = $host.ui.PromptForCredential("SecureExit Credential Prompt","Please Enter Your Office365 Global Administrator Password","","")

#empty hashtable for storing data later
$Table = @()

$TableParser = [ordered] @{
    "" = ""
}

#ImportRelevant Module
try {
    Import-Module AzureAD
}
catch {
    Write-Host "Module not found, download from https://docs.microsoft.com/en-au/powershell/module/Azuread/?view=azureadps-2.0"
    Pause
    Exit-PSHostProcess
}

function Connect-O365PS {
    $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $script:365GACreds -Authentication Basic -AllowRedirection
    Import-Module (Import-PSSession $O365Session -AllowClobber) -Global
}

#Connect to required Services
try {
    Connect-AzureAD -Credential $script:365GACreds | Out-Null
}
catch {
    Write-Error -Message "Unable to connect to AzureAD Service, please check your credentials are correct"
}
try {
    Connect-O365PS | Out-Null
}
catch {
    Write-Error -Message "Unable to connect to O365 PS Service, please check your credentials are correct"
}


#Gets all AzureADUsers, puts into list for Form to use
$ActiveUsers = Get-AzureADUser -All $true

foreach ($User in $ActiveUsers) {
    [void] $UsersListBox.Items.Add($User.UserPrincipalName)
}

$UserSelectionForm.Controls.Add($UsersListBox)
$UserSelectionForm.TopMost = $true
$result = $UserSelectionForm.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $SelectedUser = $UsersListBox.SelectedItem
}
elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
    Write-Error -Message "Operation Cancelled by User" -Category InvalidArgument
    Pause
    exit
}

#Get Full Account for User - $SeletedUser is only the UPN, as this string gets stripped when run through the form
$SelectedUser = Get-AzureADUser -ObjectID $SelectedUser

#Block Signin for User


#Remove Mobile Devices from Exchange
Remove-MobileDevice -ObjectID $SelectedUser.ObjectID 