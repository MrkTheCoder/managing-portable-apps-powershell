# Manage-Portable-Apps.ps1

$scriptName = "Manage Portable Apps"
$scriptVersion = "v0.24.202510 Alpha"
Write-Host "Running on PowerShell version: $($PSVersionTable.PSVersion)"
Write-Host "Script: $scriptName $scriptVersion"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------- HELPERS ------------
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
        Show-ErrorReport $_
    }
}

function Get-NormalizeRegistryPath {
    param([Parameter(Mandatory)][string]$Path)

    # Define registry root mappings
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

    # Return unchanged if no match
    return $Path
}

function Get-ValidateRegistryExists {
    <#
        ✅ Usage Examples:
        Get-ValidateRegistryExists -Path "HKCU:\Software\MyApp"
        Get-ValidateRegistryExists -Path "HKCU:\Software\MyApp" -Name "AutoStart"
        Get-ValidateRegistryExists -Path "HKCU:\Software\MyApp" -Name "AutoStart" -Value 1
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Name,
        [object]$Value,
        [switch]$UseRegExe
    )

    $Path = Get-NormalizeRegistryPath -Path $Path

    # Fallback for HKU or forced reg.exe
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

    # Normal PowerShell provider logic
    try {
        if (-not (Test-Path $Path)) { return $false }

        if ($Name) {
            $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if (-not $item) { return $false }

            if ($PSBoundParameters.ContainsKey('Value')) {
                return ($item.$Name -eq $Value)
            }

            return $true  # Name exists
        }

        return $true  # Path exists
    }
    catch {
        Write-Log "Error while validating registry existence: $_" -Color Red
        return $false
    }
}

function Get-RegistryValue {
    <#
        .SYNOPSIS
            Reads a registry value (returns $null if not found).

        .EXAMPLE
            Get-RegistryValue -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Notepad++" -Name "DisplayVersion"
    #>
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
        Write-Log "Error reading registry value from $($Path): $($_.Exception.Message)" -Color Red
        return $null
    }
}


function Get-CreatedShortcuts {
    <#
    .SYNOPSIS
       Find .app files in Start Menu Programs locations and build a list of shortcut metadata.
    #>
    [CmdletBinding()]
    param (
        # No parameters needed, uses environment
    )

    $list = @()

    # Paths to search
    $allUsersStart = Join-Path $env:ALLUSERSPROFILE "Microsoft\Windows\Start Menu\Programs"
    $currentUserStart = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"

    $locations = @(
        @{ Root = $allUsersStart; UserType = "AllUsers" },
        @{ Root = $currentUserStart; UserType = "CurrentUser" }
    )

    foreach ($loc in $locations) {
        $rootPath = $loc.Root
        $userType = $loc.UserType

        # If the folder doesn’t exist, skip
        if (-not (Test-Path $rootPath)) {
            continue
        }

        # Recursively find all .app files under that folder
        $appFiles = Get-ChildItem -Path $rootPath -Recurse -Filter *.app -ErrorAction SilentlyContinue

        foreach ($f in $appFiles) {
            # The parent folder (folder containing this .app) is the ShortcutPath
            $shortcutFolder = Split-Path -Parent $f.FullName

            # Load the .app file JSON
            $jsonObj = Load-AppFile $f.FullName
            if ($null -eq $jsonObj) {
                # skip if JSON parse failed
                continue
            }

            # Create an object with needed properties
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

# Add Win32 API for icon extraction (only add once in your script)
if (-not ([System.Management.Automation.PSTypeName]'IconExtractor').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class IconExtractor
{
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);
    
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern int ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, int nIcons);
    
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool DestroyIcon(IntPtr handle);
}
"@
}

function Get-IconFromFile {
    <#
    .SYNOPSIS
        Extracts an icon from a file (DLL, EXE, ICO) by index
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $false)]
        [int]$IconIndex = 0,
        
        [Parameter(Mandatory = $false)]
        [switch]$Large
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            Write-Error "File not found: $FilePath"
            return $null
        }
        
        if ($Large) {
            # Extract large icon (32x32)
            $largeIcons = New-Object IntPtr[] 1
            
            $result = [IconExtractor]::ExtractIconEx($FilePath, $IconIndex, $largeIcons, $null, 1)
            
            if ($result -gt 0 -and $largeIcons[0] -ne [IntPtr]::Zero) {
                $icon = [System.Drawing.Icon]::FromHandle($largeIcons[0]).Clone()
                [IconExtractor]::DestroyIcon($largeIcons[0])
                return $icon
            }
        }
        else {
            # Extract small icon (16x16)
            $smallIcons = New-Object IntPtr[] 1
            
            $result = [IconExtractor]::ExtractIconEx($FilePath, $IconIndex, $null, $smallIcons, 1)
            
            if ($result -gt 0 -and $smallIcons[0] -ne [IntPtr]::Zero) {
                $icon = [System.Drawing.Icon]::FromHandle($smallIcons[0]).Clone()
                [IconExtractor]::DestroyIcon($smallIcons[0])
                return $icon
            }
        }
        
        # Fallback: use simple ExtractIcon
        $hIcon = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $FilePath, $IconIndex)
        if ($hIcon -ne [IntPtr]::Zero) {
            $icon = [System.Drawing.Icon]::FromHandle($hIcon).Clone()
            [IconExtractor]::DestroyIcon($hIcon)
            return $icon
        }
        
        return $null
    }
    catch {
        Write-Error "Error extracting icon: $($_.Exception.Message)"
        return $null
    }
}

# Helper: pick first valid icon from an array or single icon
function Get-FirstIcon($icons) {
    if ($null -eq $icons) {
        return $null
    }
    if ($icons -is [System.Array]) {
        foreach ($candidate in $icons) {
            if ($candidate -is [System.Drawing.Icon]) {
                return $candidate
            }
        }
        return $null
    }
    else {
        if ($icons -is [System.Drawing.Icon]) {
            return $icons
        }
        else {
            return $null
        }
    }
}

function Load-AppFile {
    param (
        [string] $appFilePath
    )

    # Read the entire JSON file as a single string (if possible)
    $jsonText = $null
    if ($PSVersionTable.PSVersion.Major -ge 3) {
        $jsonText = Get-Content -LiteralPath $appFilePath -Raw -ErrorAction SilentlyContinue
    }
    else {
        $lines = Get-Content -LiteralPath $appFilePath -ErrorAction SilentlyContinue
        if ($lines) {
            $jsonText = $lines -join "`n"
        }
    }

    if ([string]::IsNullOrEmpty($jsonText)) {
        return $null
    }

    try {
        $obj = ConvertFrom-Json -InputObject $jsonText
    }
    catch {
        Show-ErrorReport $_
        return $null
    }
    return $obj
}

function Get-ScalarInt {
    param (
        $value
    )
    if ($null -eq $value) { return 0 }
    if ($value -is [System.Array]) {
        $v = $value[0]
    }
    else {
        $v = $value
    }
    try {
        return [int]$v
    }
    catch {
        return 0
    }
}

function Show-ManageUI {
    param(
        [Parameter(Mandatory)] $appWrappers,
        [Parameter(Mandatory)] $createdShortcuts
    )

    # region ─── Helper Functions ──────────────────────────────────────────────

    function New-Form {
        param($title, $icon)
        $form = New-Object System.Windows.Forms.Form
        $form.Text = $title
        if ($icon) { $form.Icon = $icon }
        $form.Size = New-Object System.Drawing.Size(850, 500)
        $form.StartPosition = "CenterScreen"
        $form.MinimumSize = $form.Size
        $form.Padding = New-Object System.Windows.Forms.Padding(10)
        return $form
    }

    function Initialize-ImageList {
        $imgList = New-Object System.Windows.Forms.ImageList
        $imgList.ImageSize = New-Object System.Drawing.Size(16, 16)
        return $imgList
    }

    function Add-IconToImageList {
        param($imgList, $file, $index, $key, [switch]$Warn)
        $ico = Get-FirstIcon (Get-IconFromFile -FilePath $file -IconIndex $index -Large)
        if ($ico) {
            $imgList.Images.Add($key, $ico.ToBitmap()) | Out-Null
        }
        elseif ($Warn) {
            Write-Warning "Could not load icon from $file index $index"
        }
    }

    function New-Label {
        param($text)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $text
        $lbl.AutoSize = $true
        $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $lbl.Margin = New-Object System.Windows.Forms.Padding(0, 3, 5, 3)
        return $lbl
    }

    function New-ComboBox {
        param($items)
        $cmb = New-Object System.Windows.Forms.ComboBox
        $cmb.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $cmb.Items.AddRange($items)
        $cmb.SelectedIndex = 0
        $cmb.Dock = [System.Windows.Forms.DockStyle]::Fill
        $cmb.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
        return $cmb
    }

    function New-Button {
        param($text, $dock = [System.Windows.Forms.DockStyle]::Fill)
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $text
        $btn.Dock = $dock
        $btn.Margin = New-Object System.Windows.Forms.Padding(3)
        return $btn
    }

    function New-TreeView {
        param($imgList)
        $tree = New-Object System.Windows.Forms.TreeView
        $tree.Dock = [System.Windows.Forms.DockStyle]::Fill
        $tree.CheckBoxes = $true
        $tree.ImageList = $imgList
        $tree.Margin = New-Object System.Windows.Forms.Padding(0)
        return $tree
    }

    function New-DetailsTextBox {
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Multiline = $true
        $txt.ReadOnly = $true
        $txt.ScrollBars = "Vertical"
        $txt.Dock = [System.Windows.Forms.DockStyle]::Fill
        $txt.Margin = New-Object System.Windows.Forms.Padding(0)
        return $txt
    }

    function Set-AppNodeIcon {
        param(
            [System.Windows.Forms.TreeNode] $Node,
            [object] $App,
            [System.Windows.Forms.ImageList] $ImgList
        )
        if (($App.HasShortcut) -and ($App.IsBothSame -eq $false) -and $ImgList.Images.ContainsKey('changed')) {
            $Node.ImageKey = 'changed'
            $Node.SelectedImageKey = 'changed'
        }
        elseif ($App.HasShortcut -and $ImgList.Images.ContainsKey('startMenu')) {
            $Node.ImageKey = 'startMenu'
            $Node.SelectedImageKey = 'startMenu'
        }
        elseif ($App.IsInstalled -and $ImgList.Images.ContainsKey('installed')) {
            $Node.ImageKey = 'installed'
            $Node.SelectedImageKey = 'installed'
        }
        elseif ($ImgList.Images.ContainsKey('portable')) {
            $Node.ImageKey = 'portable'
            $Node.SelectedImageKey = 'portable'
        }
    }

    function Update-DetailsTextBox {
        param(
            [System.Windows.Forms.TextBox] $txt,
            [object] $w
        )
        if (-not $w) {
            $txt.Text = ""
            return
        }
        $sb = [System.Text.StringBuilder]::new()

        if ($w.HasShortcut) {
            if ($w.IsBothSame) {
                $sb.AppendLine("App Name: [On Start Menu - identical]")
            }
            else {
                $sb.AppendLine("App Name: [On Start Menu - changed]")
            }
        }
        elseif ($w.IsInstalled) {
            $sb.AppendLine("App Name: [Already Installed]")
        }
        else {
            $sb.AppendLine("App Name:")
        }
        $sb.AppendLine($w.appName)
        $sb.AppendLine()

        if ($w.PSObject.Properties.Match("appVersion").Count) {
            if ($w.HasShortcut) {
                $sb.AppendLine("StartMenu -> Portable Version:")
                $sb.AppendLine(" $($w.ShortcutAppVersion) -> $($w.appVersion)")
            }
            elseif ($w.IsInstalled) {
                $sb.AppendLine("Installed -> Portable Version:")
                $sb.AppendLine("$($w.InstalledVersion) -> $($w.appVersion)")
            }
            else {
                $sb.AppendLine("Version:")
                $sb.AppendLine($w.appVersion)
            }
        }
        $sb.AppendLine()

        $sb.AppendLine("Group: " + @(if ($w.appGroup) { $w.appGroup } else { "<UNGROUPED>" }))
        $sb.AppendLine()

        $sb.AppendLine("StartMenu Folder: " + $w.appStartMenuFolderName)
        $sb.AppendLine()

        $desc = $w.appDescription -replace '\\n', [Environment]::NewLine
        $sb.AppendLine("Description:`n$desc")

        $txt.Text = $sb.ToString().TrimEnd()
    }

    function Update-ParentCheckState {
        param([System.Windows.Forms.TreeNode] $node)
        while ($null -ne $node) {
            $parent = $node.Parent
            if ($null -eq $parent) { break }
            $children = $parent.Nodes
            $countChecked = ($children | Where-Object { $_.Checked }).Count
            if ($countChecked -eq $children.Count -and $children.Count -gt 0) {
                $parent.Checked = $true
            }
            else {
                $parent.Checked = $false
            }
            $node = $parent
        }
    }

    function Set-AllTreeNodesChecked {
        param($nodes, [bool]$checked)
        foreach ($n in $nodes) {
            $n.Checked = $checked
            if ($n.Nodes.Count -gt 0) {
                Set-AllTreeNodesChecked $n.Nodes $checked
            }
        }
    }

    function Invert-AllTreeNodesChecked {
        param($nodes)
        foreach ($n in $nodes) {
            $n.Checked = -not $n.Checked
            if ($n.Nodes.Count -gt 0) {
                Invert-AllTreeNodesChecked $n.Nodes
            }
        }
    }

    function Populate-Tree {
        param($tree, $appWrappers, $imgList, $filter)

        $tree.Nodes.Clear()

        # Choose wrappers based on filter
        switch ($filter) {
            "All" {
                $filtered = $appWrappers
            }
            "Installed" {
                $filtered = $appWrappers | Where-Object { $_.IsInstalled }
            }
            "Portable on StartMenu" {
                $filtered = $appWrappers | Where-Object { $_.HasShortcut }
            }
            "Portable not used" {
                $filtered = $appWrappers | Where-Object { (-not $_.IsInstalled) -and (-not $_.HasShortcut) }
            }
            default {
                $filtered = $appWrappers
            }
        }

        # First annotate shortcuts info
        foreach ($w in $filtered) {
            $m = $createdShortcuts | Where-Object { $_.ShortcutAppName -eq $w.appName }
            if ($m) {
                $w | Add-Member -NotePropertyName HasShortcut -NotePropertyValue $true -Force
                $w | Add-Member -NotePropertyName ShortcutUserType -NotePropertyValue $m.ShortcutUserType -Force
                $w | Add-Member -NotePropertyName ShortcutAppVersion -NotePropertyValue $m.ShortcutAppVersion -Force
                $portableHash = (Get-FileHash -Path $w.AppFilePath -Algorithm MD5).Hash
                $shortcutAppPath = Join-Path $m.ShortcutPath ".app"
                $shortcutHash = (Get-FileHash -Path $shortcutAppPath -Algorithm MD5).Hash
                $w | Add-Member -NotePropertyName IsBothSame -NotePropertyValue ($portableHash -eq $shortcutHash) -Force
            }
            else {
                $w | Add-Member -NotePropertyName HasShortcut -NotePropertyValue $false -Force
            }
        }

        # Sort & group
        $sorted = $filtered | Sort-Object @{Expression = { $_.appGroup } }, @{Expression = { $_.appName } }
        $groupMap = @{}
        $ungrouped = @()

        foreach ($w in $sorted) {
            # Determine installation
            $w | Add-Member -NotePropertyName IsInstalled -NotePropertyValue $false -Force
            $w | Add-Member -NotePropertyName InstalledVersion -NotePropertyValue "" -Force
            if ($w.appInstallRegistryData) {
                foreach ($root in @(
                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                    )) {
                    $rp = Join-Path $root $w.appInstallRegistryData
                    if (Get-ValidateRegistryExists -Path $rp) {
                        $w.IsInstalled = $true
                        $ver = Get-RegistryValue -Path $rp -Name "DisplayVersion"
                        if ($ver) { $w.InstalledVersion = $ver }
                        break
                    }
                }
            }

            if ($w.appGroup) {
                if (-not $groupMap.ContainsKey($w.appGroup)) {
                    $groupNode = New-Object System.Windows.Forms.TreeNode($w.appGroup)
                    $groupNode.ImageKey = "folder"
                    $groupNode.SelectedImageKey = "folder"
                    $groupMap[$w.appGroup] = $groupNode
                    $tree.Nodes.Add($groupNode) | Out-Null
                }
                $appNode = New-Object System.Windows.Forms.TreeNode($w.appName)
                $appNode.Tag = $w
                Set-AppNodeIcon -Node $appNode -App $w -ImgList $imgList
                $groupMap[$w.appGroup].Nodes.Add($appNode) | Out-Null
            }
            else {
                $ungrouped += $w
            }
        }

        foreach ($w in $ungrouped) {
            $node = New-Object System.Windows.Forms.TreeNode($w.appName)
            $node.Tag = $w
            Set-AppNodeIcon -Node $node -App $w -ImgList $imgList
            $tree.Nodes.Add($node) | Out-Null
        }

        $tree.ExpandAll()
    }

    # endregion ───────────────────────────────────────────────────────────

    # region ─── UI Construction with TableLayoutPanel ────────────────────
    # Convert your .ico file to Base64 using PowerShell once:
    #   [Convert]::ToBase64String([IO.File]::ReadAllBytes("C:\path\to\icon.ico")) | Out-File -Encoding ascii icon_base64.txt
    $changedStartMenuShortcutIcon = @'
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAQAQAAAAAAAAAAAAAAAAAAAAAAAD
/37IA7a1cAPzcsQH/3bAB/t2uAf7cqgH93K0B7718AXmfiQFRrLoBuLacAaKwnQFauMcBu7+oAeinUgB7q6
wA7axbAOieQD7rqld+7ataf+yrWH/sqliA7KpXgOefRIDVkj2AwYpCgK99OICwfDl/wYg+f9eWRX7fkzM+y
JNPAPzcsADsqlh++dap+/zbrvv72qz9+9mq/frZq/3st3T9oJpv/X2dlP2kl3v9mpV8/X+hmPvEtJH74Z9K
fp+ijwH/37MA7a1df/zcsvv+3rL7/t2v/f3crP383a79779//YGxmf0z2fD9XszV/WbT3v1t4u/7ws25++q
pVH+Mx8sA/uC2AeytXX/73bP8/t+0/f3esv793a7//N2x//DAgP+VsZb/Mt36/zbj+v9L5vv+Z+Hz/dLUuf
zpqVaAqtDJAf7htwHsrV6A+961/f7gtv3937P//d6x//zesv/Nwpz/XbnJ/0nZ+/844Pv/NOL8/z3g+f2Dz
9H9y6ZjgFLU6gH/5L4B7K5fgPvguf3+47z9/uG6//3guP/03rr/fsXY/2vN8v963Pv/Wtz9/zTc/P854fv9
SNnv/V2xtoEzyPUB88SHAeqlT4DywIH988OG/fPDhf/zwoP/8MGD/8qvev+vq4P/gbe1/2XQ9f9Kx+r/k7a
g/bWzif3BomOAlrOhAfPCgwHqpU6A8r99/fPBgv3zwYH/88B///PBf//srl//665g/9PBl/9iw+f/e8LQ/+
a/hf3vu3b96aFGgPK/fQH/4LYB7K9ggPvds/3+37T9/t6y//3drv/+3rH/8sB///HCg//54Ln/y9TA/9zWu
f/93K79+tmr/eyrV4D93K0B/t+zAeyvYYD73bL9/t6y/f3dr//93Kz//t2u//LAf//ywoP//uC3//zesf/8
3a///tys/fvZqf3sqleA/tyqAf7gtwHsr2J/+961/P7ftf393rP+/d2v//7esv/ywYH/8sOF//7huv/937P
//d6y/v7dr/372qz87atZf/7drgH/4rkA7bBjf/vft/v+4bj7/t+1/f7esv3/37T988GC/fPDhv3/47z9/u
C2/f7ftP3+3rL7+9uu++2rWn//3rAA/OC4AOyuYH7527H7/N+3+/vetf373bL9/N2z/fLAf/3xwIH9/N+5/
fvetf373bP9/Nyy+/nXqfvtrVt+/NyxAO2wZADooko+665gfu2wY3/sr2J/7a9hgO2vYIDqplCA6qVPgO2u
X4Dtrl+A7K1ef+2tXX/rqlh+6KBDPu2tXAD/47sA7bBkAPzguAH/4rkB/uC3Af7fswH/4LYB88KDAfPEhwH
/5L4B/uG3Af7gtQH/37MB/NyvAe2sWwD/37IAwAP//4AB//+AAP//gAH//wAA//8AAP//AAD//wAA//8AAP
//AAD//wAA//8AAP//gAH//4AB//+AAf//wAP//w==
'@

    $iconBytes = [Convert]::FromBase64String($changedStartMenuShortcutIcon)
    $stream = New-Object IO.MemoryStream(, $iconBytes)
    $changedIcon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Icon]::new($stream)).Handle)

    $iconPath = Join-Path $env:SystemRoot "System32\shell32.dll"
    $formIcon = Get-FirstIcon (Get-IconFromFile -FilePath $iconPath -IconIndex 26 -Large)

    $form = New-Form "$scriptName $scriptVersion" $formIcon
    $imgList = Initialize-ImageList

    $iconDll = Join-Path $env:SystemRoot "System32\imageres.dll"
    Add-IconToImageList $imgList $iconDll 82 "installed" -Warn
    Add-IconToImageList $imgList $iconDll 3  "folder" -Warn
    Add-IconToImageList $imgList $iconDll 10 "empty"  -Warn
    Add-IconToImageList $imgList $iconDll 248 "startMenu" -Warn
    $imgList.Images.Add("changed", $changedIcon.ToBitmap())
    if ($formIcon) {
        $imgList.Images.Add("portable", $formIcon.ToBitmap())
    }

    # Main TableLayoutPanel (3 rows)
    $mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainLayout.ColumnCount = 1
    $mainLayout.RowCount = 3
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35))) | Out-Null  # Top bar
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null  # Main content
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45))) | Out-Null  # Bottom buttons

    # Row 0: Top bar with filter and About button
    $topBarLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $topBarLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $topBarLayout.ColumnCount = 4
    $topBarLayout.RowCount = 1
    $topBarLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null  # Label
    $topBarLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 200))) | Out-Null  # ComboBox
    $topBarLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null  # Spacer
    $topBarLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90))) | Out-Null  # About button

    $lblFilter = New-Label "Filters:"
    $cmbFilter = New-ComboBox @("All", "Installed", "Portable on StartMenu", "Portable not used")
    $btnAbout = New-Button "About"

    $topBarLayout.Controls.Add($lblFilter, 0, 0)
    $topBarLayout.Controls.Add($cmbFilter, 1, 0)
    $topBarLayout.Controls.Add($btnAbout, 3, 0)

    # Row 1: Main content area (TreeView + Details + Selection buttons)
    $contentLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $contentLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $contentLayout.ColumnCount = 2
    $contentLayout.RowCount = 1
    $contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 310))) | Out-Null  # Left panel
    $contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null  # Right panel

    # Left panel: TreeView + Selection buttons
    $leftPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $leftPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $leftPanel.ColumnCount = 1
    $leftPanel.RowCount = 2
    $leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null  # TreeView
    $leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35))) | Out-Null  # Selection buttons

    $tree = New-TreeView $imgList

    # Selection buttons panel
    $selectionButtonsLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $selectionButtonsLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $selectionButtonsLayout.ColumnCount = 3
    $selectionButtonsLayout.RowCount = 1
    $selectionButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
    $selectionButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
    $selectionButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.34))) | Out-Null

    $btnSelectAll = New-Button "Select All"
    $btnUnselectAll = New-Button "Unselect All"
    $btnInvert = New-Button "Invert Selection"

    $selectionButtonsLayout.Controls.Add($btnSelectAll, 0, 0)
    $selectionButtonsLayout.Controls.Add($btnUnselectAll, 1, 0)
    $selectionButtonsLayout.Controls.Add($btnInvert, 2, 0)

    $leftPanel.Controls.Add($tree, 0, 0)
    $leftPanel.Controls.Add($selectionButtonsLayout, 0, 1)

    # Right panel: Details textbox
    $txt = New-DetailsTextBox

    $contentLayout.Controls.Add($leftPanel, 0, 0)
    $contentLayout.Controls.Add($txt, 1, 0)

    # Row 2: Bottom buttons (Add Shortcut, Remove Shortcut, Create .app file, Exit)
    $bottomButtonsLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $bottomButtonsLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $bottomButtonsLayout.ColumnCount = 5
    $bottomButtonsLayout.RowCount = 1
    $bottomButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null  # Spacer
    $bottomButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 140))) | Out-Null  # Create .app
    $bottomButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 140))) | Out-Null  # Remove Shortcut
    $bottomButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120))) | Out-Null  # Add Shortcut
    $bottomButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90))) | Out-Null  # Exit

    $btnCreateAppFile = New-Button "Create '.app' file"
    $btnRemoveShortcut = New-Button "Remove Shortcut"
    $btnAddShortcut = New-Button "Add Shortcut"
    $btnExit = New-Button "Exit"

    $bottomButtonsLayout.Controls.Add($btnCreateAppFile, 1, 0)
    $bottomButtonsLayout.Controls.Add($btnRemoveShortcut, 2, 0)
    $bottomButtonsLayout.Controls.Add($btnAddShortcut, 3, 0)
    $bottomButtonsLayout.Controls.Add($btnExit, 4, 0)

    # Add all rows to main layout
    $mainLayout.Controls.Add($topBarLayout, 0, 0)
    $mainLayout.Controls.Add($contentLayout, 0, 1)
    $mainLayout.Controls.Add($bottomButtonsLayout, 0, 2)

    $form.Controls.Add($mainLayout)

    # Wire events
    $cmbFilter.Add_SelectedIndexChanged({
            Populate-Tree $tree $appWrappers $imgList $cmbFilter.SelectedItem
        })

    $tree.Add_AfterSelect({
            param($sender, $e)
            $node = $e.Node
            if ($node -and $node.Tag) {
                Update-DetailsTextBox -txt $txt -w $node.Tag
            }
            else {
                $txt.Text = ""
            }
        })

    $tree.Add_AfterCheck({
            param($sender, $e)
            if ($e.Action -eq [System.Windows.Forms.TreeViewAction]::ByMouse) {
                foreach ($c in $e.Node.Nodes) {
                    $c.Checked = $e.Node.Checked
                }
                Update-ParentCheckState $e.Node
            }
        })

    $btnSelectAll.Add_Click({
            Set-AllTreeNodesChecked $tree.Nodes $true
            foreach ($n in $tree.Nodes) {
                Update-ParentCheckState $n
            }
        })

    $btnUnselectAll.Add_Click({
            Set-AllTreeNodesChecked $tree.Nodes $false
            foreach ($n in $tree.Nodes) {
                Update-ParentCheckState $n
            }
        })

    $btnInvert.Add_Click({
            Invert-AllTreeNodesChecked $tree.Nodes
            foreach ($n in $tree.Nodes) {
                Update-ParentCheckState $n
            }
        })

    $btnAbout.Add_Click({
            [void] [System.Windows.Forms.MessageBox]::Show(
                "This script was fully written by AIs assistant (ChatGPT + Claude.ai) and then guided, managed, and refined by me.`r`n`r`n" +
                "Manage Portable Apps is a PowerShell utility designed to help you keep your portable applications organized and accessible. " +
                "It scans a folder of portable apps (each identified by its own .app metadata file) and compares them against your Start Menu shortcuts. " +
                "You can see which apps are installed, which already have shortcuts, and which are unused. " +
                "With just a few clicks you can add or remove shortcuts and create new .app entries.`r`n`r`n" +
                "Author: MrkTheCoder`r`n" +
                "Version: $scriptVersion`r`n" +
                "© 2025 MrkTheCoder – All Rights Reserved",
                "About Manage-Portable-Apps",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        })

    $btnExit.Add_Click({ $form.Close() })

    Populate-Tree $tree $appWrappers $imgList $cmbFilter.SelectedItem

    [void]$form.ShowDialog()

    # endregion
}



# ---- Main ----

# Discover created shortcuts
$createdShortcuts = Get-CreatedShortcuts
Write-Host "Found $($createdShortcuts.Count) created shortcuts in Start Menu folders"
# You can optionally print them or debug:
# foreach ($cs in $createdShortcuts) {
#     Write-Host "[$($cs.ShortcutUserType)] $($cs.ShortcutPath) → $($cs.ShortcutAppName) v$($cs.ShortcutAppVersion)"
# }

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
# -------------------------------------
# Comment the following line on production release 
#$scriptDir = "d:\Portables"
# -------------------------------------
$appFiles = Get-ChildItem -Path $scriptDir -Recurse -Filter *.app -ErrorAction SilentlyContinue
Write-Host "Total Portable Apps: '$($appFiles.Count)'"
$appWrappers = @()
foreach ($f in $appFiles) {
    $json = Load-AppFile $f.FullName
    if ($null -eq $json) {
        continue
    }
    # wrap JSON object properties + metadata into single PSObject
    $wrapper = New-Object PSObject
    foreach ($prop in $json.PSObject.Properties) {
        $wrapper | Add-Member NoteProperty $prop.Name -Value $prop.Value
    }
    $wrapper | Add-Member NoteProperty "AppFilePath" -Value $f.FullName
    $appWrappers += $wrapper
}

if ($appWrappers.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No .app files found under $scriptDir", "Error", "OK", "Error")
    exit 0
}

Show-ManageUI -appWrappers $appWrappers -createdShortcuts $createdShortcuts
