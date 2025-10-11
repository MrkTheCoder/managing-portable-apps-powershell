# Prompt for the .lnk file to edit
$lnkPath = Read-Host "Enter full path to the .lnk file to edit"
if (-not (Test-Path $lnkPath)) {
    Write-Error "File not found: $lnkPath"
    return
}

$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($lnkPath)

# Helper function: prompt with default
function Prompt-WithDefault([string] $prompt, [string] $default) {
    if ([string]::IsNullOrEmpty($default)) {
        return Read-Host "$prompt (blank allowed)"
    } else {
        $resp = Read-Host "$prompt [$default]"
        if ([string]::IsNullOrEmpty($resp)) {
            return $default
        } else {
            return $resp
        }
    }
}

Write-Host "Editing shortcut: $lnkPath"
Write-Host "(Leave blank to keep current value)"

# Prompt for each editable property
$newTarget = Prompt-WithDefault "TargetPath" $shortcut.TargetPath
$newArgs = Prompt-WithDefault "Arguments" $shortcut.Arguments
$newWorkDir = Prompt-WithDefault "WorkingDirectory" $shortcut.WorkingDirectory
$newIcon = Prompt-WithDefault "IconLocation" $shortcut.IconLocation
$newDesc = Prompt-WithDefault "Description" $shortcut.Description
$newHotkey = Prompt-WithDefault "Hotkey" $shortcut.Hotkey
$newWindowStyle = Prompt-WithDefault "WindowStyle (numeric)" ($shortcut.WindowStyle)

# Assign new values
$shortcut.TargetPath = $newTarget
$shortcut.Arguments = $newArgs
$shortcut.WorkingDirectory = $newWorkDir
$shortcut.IconLocation = $newIcon
$shortcut.Description = $newDesc
$shortcut.Hotkey = $newHotkey

# WindowStyle is an integer (0 = normal, 1 = minimized, 3 = maximized, etc.)
[int] $wsVal = 0
if ([int]::TryParse($newWindowStyle, [ref] $wsVal)) {
    $shortcut.WindowStyle = $wsVal
}

# Save changes
$shortcut.Save()
Write-Host "Shortcut updated and saved."
