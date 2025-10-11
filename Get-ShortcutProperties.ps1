# Define the CSV path and output Excel path
Add-Type -AssemblyName System.Windows.Forms
$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.Title = "Select a Shortcut"
$dlg.Filter = "Shortcut  (*.lnk)|*.lnk"
if ($dlg.ShowDialog() -eq 'OK') {
    $lnkPath = $dlg.FileName
}
else {
    return
}
# Prompt the user for a .lnk file
# $lnkPath = Read-Host "Enter full path to the .lnk file"

if (-not (Test-Path -LiteralPath $lnkPath)) {
    Write-Error "File not found: $lnkPath"
    return
}

# Create the WScript.Shell COM object
$wsh = New-Object -ComObject WScript.Shell

# Open the shortcut
# Note: the CreateShortcut method returns a ShellLink object if the .lnk exists
$shortcut = $wsh.CreateShortcut($lnkPath)

# Display properties
Write-Host "Properties of shortcut file: $lnkPath"
Write-Host "-------------------------------------"
Write-Host "TargetPath        : $($shortcut.TargetPath)"
Write-Host "Arguments         : $($shortcut.Arguments)"
Write-Host "WorkingDirectory  : $($shortcut.WorkingDirectory)"
Write-Host "IconLocation      : $($shortcut.IconLocation)"
Write-Host "WindowStyle       : $($shortcut.WindowStyle)"
Write-Host "Description       : $($shortcut.Description)"
Write-Host "Hotkey            : $($shortcut.Hotkey)"
# There may be other properties (you can list $shortcut | Get-Member)
