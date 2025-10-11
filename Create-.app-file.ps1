# Create-AppFile.ps1
# Requires: WScript.Shell COM availability
# -------------------------------------
# Helper: Choose folder (FolderDialog)
# -------------------------------------
function Select-FolderDialog {
    param([string]$Title = "Select Folder")

    Add-Type -AssemblyName System.Windows.Forms

    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.FileName = '[Select Folder]'
    $dlg.CheckFileExists = $false
    $dlg.CheckPathExists = $true
    $dlg.DereferenceLinks = $true
    $dlg.ValidateNames = $false
    $dlg.AddExtension = $false
    $dlg.Filter = "Folders|`n"
    $dlg.Title = $Title   # <-- Title applied here

    if ((Show-TopMostDialog $dlg) -eq 'OK') {
        return [System.IO.Path]::GetDirectoryName($dlg.FileName)
    }
    else {
        return $null
    }
}

function Show-TopMostDialog($dialog) {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $form.ShowInTaskbar = $true
    <#
    $form.StartPosition = 'CenterScreen'
    $form.WindowState = 'Minimized'
    #>
    $form.StartPosition = "Manual"
    $form.Location = New-Object System.Drawing.Point(-32000, -32000)  # hidden off-screen
    $form.Size = New-Object System.Drawing.Size(1, 1)

    $form.Show()
    $result = $dialog.ShowDialog($form)
    $form.Dispose()
    return $result
}

# -----------------------------------
# Helper: Prompt for UID (InputBox)
# -----------------------------------
function Prompt-UID {
    param([string]$Title = "Enter Data:", [string]$Message = "A value:")
    Add-Type -AssemblyName Microsoft.VisualBasic
    $uid = [Microsoft.VisualBasic.Interaction]::InputBox($Message, $Title, "")
    return [string]$uid
}


function Clean-Description {
    param (
        [string] $desc
    )

    if ([string]::IsNullOrEmpty($desc)) {
        return ""
    }

    # Trim any surrounding quotes
    $d = $desc.Trim('"')

    # Heuristic: if the string looks like an absolute path (drive letter + colon + backslash or UNC),
    # or contains multiple path separators, or ends with a filename extension, treat it as a path.
    # You can adjust these heuristics as needed.

    # Pattern: starts with letter + ':' + '\' or UNC '\\'
    if ($d -match '^[A-Za-z]:\\') {
        return ""
    }
    if ($d -match '^\\\\') {
        return ""
    }
    # Or if it contains backslash or forward slash more than once, assume a path
    if ($d -match '[\\/].*[\\/]') {
        return ""
    }
    # Or if it ends with a known extension (e.g. .exe, .lnk, .bat, .url)
    if ($d -match '\.(exe|lnk|bat|url|cmd|msi)$') {
        return ""
    }

    # Otherwise keep it
    return $desc
}


function Convert-PathToToken {
    param (
        [string] $fullPath,
        [string] $appPath
    )

    if ([string]::IsNullOrEmpty($fullPath)) {
        return ""
    }

    $fp = $fullPath.Trim('"')
    $ap = $appPath.Trim('"')

    try { $normFull = [System.IO.Path]::GetFullPath($fp) } catch { $normFull = $fp }
    try { $normApp = [System.IO.Path]::GetFullPath($ap) } catch { $normApp = $ap }

    $pos = $normFull.IndexOf($normApp, [System.StringComparison]::InvariantCultureIgnoreCase)
    if ($pos -ge 0) {
        $startOfSuffix = $pos + $normApp.Length
        $suffix = ""
        if ($startOfSuffix -lt $normFull.Length) {
            $suffix = $normFull.Substring($startOfSuffix)
        }
        if ($suffix.StartsWith("\") -or $suffix.StartsWith("/")) {
            $suffix = $suffix.Substring(1)
        }
        return "[.app_path]\" + ($suffix -replace "/", "\")
    }
    else {
        # no matching substring → fallback: attempt replacing system paths or just return normalized full
        # example:
        $sysRoot = $ENV:SystemRoot
        if ($sysRoot) {
            try {
                $fullSys = [System.IO.Path]::GetFullPath($sysRoot)
                if ($normFull.StartsWith($fullSys, [StringComparison]::InvariantCultureIgnoreCase)) {
                    $rest = $normFull.Substring($fullSys.Length)
                    if ($rest.StartsWith("\") -or $rest.StartsWith("/")) {
                        $rest = $rest.Substring(1)
                    }
                    return "%SystemRoot%\" + $rest
                }
            }
            catch { }
        }
        $comSpec = $ENV:ComSpec
        if ($comSpec) {
            try {
                $fullCom = [System.IO.Path]::GetFullPath($comSpec)
                if ($normFull.Equals($fullCom, [StringComparison]::InvariantCultureIgnoreCase)) {
                    return "%ComSpec%"
                }
            }
            catch { }
        }
        # default fallback
        return $normFull
    }
}


function Get-ShortcutProperties {
    param (
        [string]$lnkPath
    )
    $wsh = New-Object -ComObject WScript.Shell
    try {
        $sc = $wsh.CreateShortcut($lnkPath)
    }
    catch {
        return $null
    }
    return @{
        TargetPath       = $sc.TargetPath
        Arguments        = $sc.Arguments
        WorkingDirectory = $sc.WorkingDirectory
        IconLocation     = $sc.IconLocation
        WindowStyle      = $sc.WindowStyle
        Description      = $sc.Description
        Hotkey           = $sc.Hotkey
    }
}

function ConvertTo-JsonSimple {
    param (
        [object] $obj,
        [int] $indent = 0
    )

    if ($null -eq $obj) { return "null" }

    $pad = " " * $indent

    if ($obj -is [hashtable] -or $obj -is [System.Collections.IDictionary]) {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("{")
        $first = $true
        foreach ($k in $obj.Keys) {
            if (-not $first) {
                [void]$sb.AppendLine(",")
            }
            $first = $false
            $escapedKey = ($k -replace '"', '\"')
            [void]$sb.Append($pad + "  `"$escapedKey`": ")
            [void]$sb.Append((ConvertTo-JsonSimple $obj[$k] ($indent + 2)))
        }
        [void]$sb.AppendLine()
        [void]$sb.Append($pad + "}")
        return $sb.ToString()
    }
    elseif ($obj -is [System.Collections.IEnumerable] -and -not ($obj -is [string])) {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("[")
        $first = $true
        foreach ($elem in $obj) {
            if (-not $first) {
                [void]$sb.AppendLine(",")
            }
            $first = $false
            [void]$sb.Append($pad + "  " + (ConvertTo-JsonSimple $elem ($indent + 2)))
        }
        [void]$sb.AppendLine()
        [void]$sb.Append($pad + "]")
        return $sb.ToString()
    }
    else {
        if ($obj -is [string]) {
            $escaped = $obj -replace '\\', '\\' -replace '"', '\"'
            return "`"$escaped`""
        }
        elseif ($obj -is [int] -or $obj -is [double] -or $obj -is [float]) {
            return $obj.ToString()
        }
        elseif ($obj -is [bool]) {
            return $obj.ToString().ToLower()
        }
        else {
            $s = $obj.ToString()
            $escaped = $s -replace '\\', '\\' -replace '"', '\"'
            return "`"$escaped`""
        }
    }
}
# ---- Main ----

$shortcutFolder = Select-FolderDialog "Enter the folder path containing the .lnk shortcuts"
if (-not $shortcutFolder -or (-not (Test-Path -LiteralPath $shortcutFolder))) {
    Write-Error "Shortcut folder not found: $shortcutFolder"
    exit 1
}


$appName = Prompt-UID -Message "App Name:"
$appVersion = Prompt-UID -Message  "App Version:"
$appGroup = Prompt-UID -Message  "App Group:"
$appInstallRegistryData = Prompt-UID -Message  "App Install Registry Key:"
$appStartMenuFolderName = Prompt-UID -Message  "App Start-Menu Folder Name:"
$appDescription = Prompt-UID -Message  "Write a short description (1-2 lines)"

$appPath = Select-FolderDialog "Enter the Portable App path (base folder where .app will be saved)"
if (-not $appPath -or (-not (Test-Path -LiteralPath $appPath))) {
    Write-Error "Portable app path not found: $appPath"
    exit 1
}


$lnkFiles = Get-ChildItem -Path $shortcutFolder -Recurse -Filter "*.lnk" -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer }

$shortcutEntries = @()
foreach ($lnk in $lnkFiles) {
    $props = Get-ShortcutProperties $lnk.FullName
    if ($null -eq $props) {
        Write-Warning "Cannot read shortcut: $($lnk.FullName)"
        continue
    }
    $entry = [ordered]@{}
    $entry["name"] = $lnk.BaseName
    $entry["target"] = Convert-PathToToken $props.TargetPath $appPath
    $entry["arguments"] = Convert-PathToToken $props.Arguments $appPath
    $entry["workingDirectory"] = Convert-PathToToken $props.WorkingDirectory $appPath
    $entry["icon"] = Convert-PathToToken $props.IconLocation $appPath
    $entry["windowStyle"] = $props.WindowStyle
    $entry["description"] = Clean-Description $props.Description
    $shortcutEntries += $entry
}

$appObj = [ordered]@{
    appName                = $appName
    appVersion             = $appVersion
    appGroup               = $appGroup
    appDescription         = $appDescription
    appInstallRegistryData = $appInstallRegistryData
    appStartMenuFolderName = $appStartMenuFolderName
    shortcuts              = $shortcutEntries
}

$jsonText = ConvertTo-JsonSimple $appObj 0

# Save file with ".app" extension
$appFileName = ".app"
$appFilePath = Join-Path -Path $appPath -ChildPath $appFileName

Set-Content -LiteralPath $appFilePath -Value $jsonText -Encoding UTF8

Write-Host "Generated .app file at: $appFilePath" -ForegroundColor Green