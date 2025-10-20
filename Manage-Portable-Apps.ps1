<##
.SYNOPSIS
    Manage Portable Apps - main script that discovers portable apps and shows UI.

.DESCRIPTION
    - Discovers .app metadata files in the configured Portables folder.
    - Builds wrapper objects for each .app and discovers created shortcuts in Start Menu locations.
    - Loads the UI module and calls Show-ManageUI to present results.

.NOTES
    - PowerShell compatibility: aims for PowerShell v2.0+.
    - Requires Manage-Portable-Apps.UI.psm1 to be in the same folder.
    - Put your Portables folder in $scriptDir variable below (or adjust).
    - Uses Read-Host only when explicit input is required (main is GUI-driven).

.SUGGESTED NAME
    Manage-Portable-Apps.ps1
#>

# Script header data
$scriptName = "Manage Portable Apps"
$scriptVersion = "v0.26.202510 Alpha"

Write-Host "Running on PowerShell version: $($PSVersionTable.PSVersion)"
Write-Host "Script: $($scriptName) $($scriptVersion)"


# Function Helper: Reusable Catch helper function
function Show-ErrorReport {
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord] $ErrorRecord
    )
    try {
        $inv = $ErrorRecord.InvocationInfo
        $scriptName = $inv.ScriptName
        $lineNumber = $inv.ScriptLineNumber
        $lineOfCode = $inv.Line
        $commandName = $inv.InvocationName

        Write-Host "`nError in script: '$scriptName'" -ForegroundColor Red
        Write-Host " Command    : $commandName" -ForegroundColor Red
        Write-Host " Line #     : $lineNumber" -ForegroundColor Red
        Write-Host " Line Text  : '$lineOfCode'" -ForegroundColor Red
        Write-Host " Exception  : $($ErrorRecord.Exception.GetType().FullName) - $($ErrorRecord.Exception.Message)" -ForegroundColor Red
        Write-Host " Error ID   : $($ErrorRecord.FullyQualifiedErrorId)" -ForegroundColor Red
        Write-Host " Category   : $($ErrorRecord.CategoryInfo.Category)" -ForegroundColor Red

        if ($ErrorRecord.ScriptStackTrace) {
            Write-Host " StackTrace :" -ForegroundColor Red
            Write-Host $ErrorRecord.ScriptStackTrace -ForegroundColor Red
        }

    }
    catch {
        Write-Host "Error while reporting error: $($_.Exception.Message)" -ForegroundColor Red
    }
}




# --- Robust load of UI module (replace previous Import-Module lines) ---
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$uiModulePath = Join-Path $scriptDir "Manage-Portable-Apps.UI.psm1"

if (-not (Test-Path -LiteralPath $uiModulePath)) {
    Write-Error "UI module not found at '$($uiModulePath)'. Ensure Manage-Portable-Apps.UI.psm1 is in the same folder as this script."
    exit 1
}

# If file is blocked (downloaded), try to unblock it first (harmless if not blocked)
try { Unblock-File -Path $uiModulePath -ErrorAction SilentlyContinue } catch {}

# Try to import the module and fail loudly if it doesn't work
try {
    Import-Module -Name $uiModulePath -Force -ErrorAction Stop
}
catch {
    Write-Warning "Import-Module failed: $($_.Exception.Message)"
    Write-Warning "Attempting to dot-source the module file as a fallback (this will not export module functions)."
    try {
        . $uiModulePath
    }
    catch {
        Write-Error "Failed to load UI module by dot-sourcing: $($_.Exception.Message)"
        Write-Host "Run the following diagnostics in an interactive session to investigate:"
        Write-Host "  1) Test-Path '$uiModulePath'"
        Write-Host "  2) Import-Module -Name '$uiModulePath' -Verbose -ErrorAction Stop"
        Write-Host "  3) Get-Module -ListAvailable | Where-Object { $_.Name -match 'Manage-Portable' }"
        Write-Host "  4) Get-Command -Name Show-ManageUI -All"
        exit 1
    }
}

# Confirm the function is available
if (-not (Get-Command -Name Show-ManageUI -ErrorAction SilentlyContinue)) {
    Write-Warning "Show-ManageUI not found after loading the UI module."
    Write-Host "Available commands from loaded modules (searching for Manage-Portable-Apps.UI):"
    try {
        # list currently loaded module names and commands for debugging
        Get-Module | ForEach-Object {
            Write-Host "Module: $($_.Name) ($($_.Path))"
            Get-Command -Module $_.Name -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name -First 10 | ForEach-Object { Write-Host "  - $_" }
        }
    }
    catch { }
    Write-Host "If the module failed to import, try running:"
    Write-Host "  Import-Module -Name '$uiModulePath' -Verbose -ErrorAction Stop"
    exit 1
}
# --- end import block ---


# ----------- HELPERS (core script) ------------

function Show-ErrorReport {
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord] $ErrorRecord
    )
    try {
        $inv = $ErrorRecord.InvocationInfo
        $scriptNameLocal = $inv.ScriptName
        $lineNumber = $inv.ScriptLineNumber
        $lineOfCode = $inv.Line
        $commandName = $inv.InvocationName

        Write-Host "`nError in script: '$scriptNameLocal'" -ForegroundColor Red
        Write-Host " Command    : $commandName" -ForegroundColor Red
        Write-Host " Line #     : $lineNumber" -ForegroundColor Red
        Write-Host " Line Text  : '$lineOfCode'" -ForegroundColor Red
        Write-Host " Exception  : $($ErrorRecord.Exception.GetType().FullName) - $($ErrorRecord.Exception.Message)" -ForegroundColor Red
        Write-Host " Error ID   : $($ErrorRecord.FullyQualifiedErrorId)" -ForegroundColor Red
        Write-Host " Category   : $($ErrorRecord.CategoryInfo.Category)" -ForegroundColor Red

        if ($ErrorRecord.ScriptStackTrace) {
            Write-Host " StackTrace :" -ForegroundColor Red
            Write-Host $ErrorRecord.ScriptStackTrace -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error while reporting error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-NormalizeRegistryPath {
    param([Parameter(Mandatory)][string]$Path)

    $rootMap = @{
        'HKLM'                = 'HKLM:\'
        'HKEY_LOCAL_MACHINE'  = 'HKLM:\'
        'HKCU'                = 'HKCU:\'
        'HKEY_CURRENT_USER'   = 'HKCU:\'
        'HKCR'                = 'HKCR:\'
        'HKEY_CLASSES_ROOT'   = 'HKCR:\'
        'HKU'                 = 'HKU:\'
        'HKEY_USERS'          = 'HKU:\'
        'HKCC'                = 'HKCC:\'
        'HKEY_CURRENT_CONFIG' = 'HKCC:\'
    }

    foreach ($key in $rootMap.Keys) {
        if ($Path -match "^$key\\") {
            return $Path -replace "^$key\\", $rootMap[$key]
        }
    }
    return $Path
}

function Get-ValidateRegistryExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Name,
        [object]$Value,
        [switch]$UseRegExe
    )

    $Path = Get-NormalizeRegistryPath -Path $Path

    if ($UseRegExe -or $Path -like "HKU:*") {
        $regPath = $Path -replace "^HKU:\\?", "HKU\"
        $regQuery = "reg query `"$regPath`""
        if ($Name) { $regQuery += " /v `"$Name`"" }

        $output = cmd /c $regQuery 2>$null
        if (-not $output) { return $false }

        if ($PSBoundParameters.ContainsKey('Value')) {
            $actual = ($output -split "`r?`n") | Where-Object { $_ -match "^\s*$Name\s+" }
            if (-not $actual) { return $false }
            $parts = $actual -split '\s{2,}'
            $regValue = $parts[-1].Trim()
            return ($regValue -eq "$Value")
        }
        return $true
    }

    try {
        if (-not (Test-Path $Path)) { return $false }
        if ($Name) {
            $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if (-not $item) { return $false }
            if ($PSBoundParameters.ContainsKey('Value')) {
                return ($item.$Name -eq $Value)
            }
            return $true
        }
        return $true
    }
    catch {
        Write-Host "Error while validating registry existence: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Get-RegistryValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $Path = Get-NormalizeRegistryPath -Path $Path
        if (-not (Test-Path $Path)) { return $null }
        $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue
        if ($null -eq $item) { return $null }
        $val = $item.PSObject.Properties[$Name].Value
        if ($null -eq $val) { return $null }
        return "$val"
    }
    catch {
        Write-Host "Error reading registry value from $($Path): $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# ----------------------------------------------------------------------
# Get-CreatedShortcuts: find .app metadata files under Start Menu locations
# ----------------------------------------------------------------------
function Get-CreatedShortcuts {
    [CmdletBinding()]
    param ()

    $list = @()
    $allUsersStart = Join-Path $env:ALLUSERSPROFILE "Microsoft\Windows\Start Menu\Programs"
    $currentUserStart = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"

    $locations = @(
        @{ Root = $allUsersStart; UserType = "AllUsers" },
        @{ Root = $currentUserStart; UserType = "CurrentUser" }
    )

    foreach ($loc in $locations) {
        $rootPath = $loc.Root
        $userType = $loc.UserType
        if (-not (Test-Path $rootPath)) { continue }

        $appFiles = Get-ChildItem -Path $rootPath -Recurse -Filter *.app -ErrorAction SilentlyContinue
        foreach ($f in $appFiles) {
            $shortcutFolder = Split-Path -Parent $f.FullName
            # Load the .app JSON (Use the local Import-AppFile function)
            $jsonObj = Import-AppFile -appFilePath $f.FullName
            if ($null -eq $jsonObj) { continue }
            $obj = [PSCustomObject]@{
                ShortcutPath       = $shortcutFolder
                ShortcutUserType   = $userType
                ShortcutAppName    = $jsonObj.appName
                ShortcutAppVersion = if ($jsonObj.PSObject.Properties.Match("appVersion").Count -gt 0) { $jsonObj.appVersion } else { "" }
            }
            $list += $obj
        }
    }
    return $list
}

# ----------------------------------------------------------------------
# Compatibility: Get-FileHash fallback for PS v2/v3
# - Use this instead of calling Get-FileHash directly in UI code to remain v2 compatible.
# ----------------------------------------------------------------------
function Get-FileHashCompat {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Algorithm = 'MD5'
    )
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        switch ($Algorithm.ToUpper()) {
            'MD5' {
                $md5 = New-Object System.Security.Cryptography.MD5CryptoServiceProvider
                $hash = $md5.ComputeHash($stream)
                $md5.Dispose()
            }
            'SHA1' {
                $sha = New-Object System.Security.Cryptography.SHA1CryptoServiceProvider
                $hash = $sha.ComputeHash($stream)
                $sha.Dispose()
            }
            default {
                # default to MD5
                $md5 = New-Object System.Security.Cryptography.MD5CryptoServiceProvider
                $hash = $md5.ComputeHash($stream)
                $md5.Dispose()
            }
        }
        # convert to hex string
        $sb = New-Object System.Text.StringBuilder
        foreach ($b in $hash) { [void]$sb.Append($b.ToString("x2")) }
        return $sb.ToString()
    }
    finally {
        $stream.Close()
        $stream.Dispose()
    }
}

# ----------------------------------------------------------------------
# JSON loader for .app files (keeps PSv2 compatibility)
# ----------------------------------------------------------------------
function Import-AppFile {
    param([Parameter(Mandatory)][string]$appFilePath)

    $jsonText = $null
    if ($PSVersionTable.PSVersion.Major -ge 3) {
        $jsonText = Get-Content -LiteralPath $appFilePath -Raw -ErrorAction SilentlyContinue
    }
    else {
        $lines = Get-Content -LiteralPath $appFilePath -ErrorAction SilentlyContinue
        if ($lines) { $jsonText = $lines -join "`n" }
    }

    if ([string]::IsNullOrEmpty($jsonText)) { return $null }

    try {
        $obj = ConvertFrom-Json -InputObject $jsonText
    }
    catch {
        Show-ErrorReport $_
        return $null
    }
    return $obj
}

# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------

# Discover created shortcuts
$createdShortcuts = Get-CreatedShortcuts
Write-Host "Found $($createdShortcuts.Count) created shortcuts in Start Menu folders"

# Portables folder: default to scriptDir\Portables (or change)
# For development, you may override below
$portablesRoot = Join-Path $scriptDir "Portables"
# Uncomment for development override:
$portablesRoot = "D:\Portables"

if (-not (Test-Path -LiteralPath $portablesRoot)) {
    Write-Host "Portables directory not found at '$($portablesRoot)'. Update the path and re-run." -ForegroundColor Red
    exit 1
}

$appFiles = Get-ChildItem -Path $portablesRoot -Recurse -Filter *.app -ErrorAction SilentlyContinue
Write-Host "Total Portable Apps: '$($appFiles.Count)'"

$appWrappers = @()
foreach ($f in $appFiles) {
    $json = Import-AppFile -appFilePath $f.FullName
    if ($null -eq $json) { continue }

    $wrapper = New-Object PSObject
    foreach ($prop in $json.PSObject.Properties) {
        $wrapper | Add-Member NoteProperty $prop.Name -Value $prop.Value
    }
    $wrapper | Add-Member NoteProperty "AppFilePath" -Value $f.FullName

    # --- Added logic moved from Populate-Tree ---
    # Shortcut info
    $m = $createdShortcuts | Where-Object { $_.ShortcutAppName -eq $wrapper.appName }
    if ($m) {
        $wrapper | Add-Member -NotePropertyName HasShortcut -NotePropertyValue $true -Force
        $wrapper | Add-Member -NotePropertyName ShortcutUserType -NotePropertyValue $m.ShortcutUserType -Force
        $wrapper | Add-Member -NotePropertyName ShortcutAppVersion -NotePropertyValue $m.ShortcutAppVersion -Force

        try {
            $portableHash = Get-FileHashCompat -Path $wrapper.AppFilePath -Algorithm MD5
            $shortcutAppPath = Join-Path $m.ShortcutPath ".app"
            $shortcutHash = Get-FileHashCompat -Path $shortcutAppPath -Algorithm MD5
            $wrapper | Add-Member -NotePropertyName IsBothSame -NotePropertyValue ($portableHash -eq $shortcutHash) -Force
        }
        catch {
            $wrapper | Add-Member -NotePropertyName IsBothSame -NotePropertyValue $false -Force
        }
    }
    else {
        $wrapper | Add-Member -NotePropertyName HasShortcut -NotePropertyValue $false -Force
        $wrapper | Add-Member -NotePropertyName IsBothSame -NotePropertyValue $false -Force
    }

    # Registry detection for installation
    $wrapper | Add-Member -NotePropertyName IsInstalled -NotePropertyValue $false -Force
    $wrapper | Add-Member -NotePropertyName InstalledVersion -NotePropertyValue "" -Force

    if ($wrapper.appInstallRegistryData) {
        foreach ($root in @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            )) {
            $rp = Join-Path $root $wrapper.appInstallRegistryData
            if (Get-ValidateRegistryExists -Path $rp) {
                $wrapper.IsInstalled = $true
                $ver = Get-RegistryValue -Path $rp -Name "DisplayVersion"
                if ($ver) { $wrapper.InstalledVersion = $ver }
                break
            }
        }
    }
    # --- End moved logic ---

    $appWrappers += $wrapper
}

if ($appWrappers.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No .app files found under $($portablesRoot)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 0
}

# Call the UI (function exported from UI module)
Show-ManageUI -appWrappers $appWrappers -Title "$($scriptName) $($scriptVersion)"