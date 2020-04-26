<#
.SYNOPSIS
    Adds a HomeAssistant icon to the system tray that can trigger HomeAssistant automations.

.DESCRIPTION
    This script adds a HomeAssistant icon to the system tray. Clicking on this icon will open
        HomeAssistant in your default browser, while right-clicking on this icon will display
        an alphabetized list of all active Automations in HomeAssistant when the script was run.
        To refresh the list, relaunch the script.

.NOTES
    File Name: HomeAssistantTrayApp.ps1

.NOTES
    Requires CredentialManager module.

    > Install-Module CredentialManager

    > New-StoredCredential -Target HomeAssistant -UserName [server URL] -Password [long-lived API token] -Type Generic -Persist LocalMachine

#>

# Retrieve and parse server and token from Windows Credential Manager
$homeAssistant = Get-StoredCredential -Target "HomeAssistant"
$serverAddress = $homeAssistant.UserName
$headers = @{"Authorization" = "Bearer $($homeAssistant.GetNetworkCredential().Password)"; "Content-Type" = "application/json"}

# Retrieve enabled Automations from HomeAssistant
$states = Invoke-RestMethod -Method GET -Uri "$serverAddress/api/states" -Headers $headers
$automations = $states | Where-Object {$_.entity_id -like "automation.*" -and $_.state -eq "on"} | Sort-Object -Property @{e={$_.attributes.friendly_name}}
 
# Add assemblies
foreach ($assembly in @('System.Windows.Forms','PresentationFramework','System.Drawing','WindowsFormsIntegration')){
    [System.Reflection.Assembly]::LoadWithPartialName($assembly) | Out-Null
}

# Set up icon to display in the systray
$iconBase64 = "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"
$iconBytes  = [Convert]::FromBase64String($iconBase64)
$iconStream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
    $iconStream.Write($iconBytes, 0, $iconBytes.Length)
$iconImage  = [System.Drawing.Image]::FromStream($iconStream, $true)

# Add the systray icon 
$haTool_Icon = New-Object System.Windows.Forms.NotifyIcon
$haTool_Icon.Text = "HomeAssistant Tray Tool"
$haTool_Icon.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $iconStream).GetHIcon())
$haTool_Icon.Visible = $true
 
# Add all menus as context menus
$contextmenu = New-Object System.Windows.Forms.ContextMenuStrip
$haTool_Icon.ContextMenuStrip = $contextmenu

foreach ($automation in $automations) {
    $automationName = $automation.attributes.friendly_name
    $automationID = $($automation.entity_id.split('.')[1])
    $menuItem = $haTool_Icon.ContextMenuStrip.Items.Add($automationName)
    #$menuItem.Text = $automationName
    $menuItem.Tag = $automationID
    $menuItem.add_Click({
        Invoke-RestMethod -Method POST -Body (@{entity_id = "automation.$($this.Tag)"} | ConvertTo-Json) -Uri "$serverAddress/api/services/automation/trigger" -Headers $headers
    })
}

# Add a divider to the menu to separate Automations from app meta functions
$haTool_Icon.ContextMenuStrip.Items.Add("-")

$restart = $haTool_Icon.ContextMenuStrip.Items.Add("Restart")
$restart.Text = "Relaunch Tray App"

$exit = $haTool_Icon.ContextMenuStrip.Items.Add("Exit")
$exit.Text = "Exit"

# Left-click the systray icon to launch HomeAssistant
$haTool_Icon.Add_Click({
 If ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
  [Diagnostics.Process]::Start("$serverAddress")
 }    
})
 
# Action after clicking on the Exit context menu
$exit.add_Click({
    $haTool_Icon.Visible = $false
    $window.Close()
    Stop-Process $pid
})

# Action after clicking on the Restart context menu
$restart.add_Click({
    Start-Process -WindowStyle hidden powershell.exe "$PSCommandPath" 
 
    $haTool_Icon.Visible = $false
    $window.Close()
    Stop-Process $pid
})

# Hide the PowerShell window
$psWindow = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $psWindow -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
 
# Use Garbage Collection to reduce memory usage
[System.GC]::Collect()
 
# Create an application context for it to all run within
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)