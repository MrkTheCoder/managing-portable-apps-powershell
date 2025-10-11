# Manage-Portable-Apps.ps1

$scriptName = "Manage Portable Apps"
$scriptVersion = "v0.7.202510 Pre-Alpha"
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
    param (
        [array] $appWrappers
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
        $form.SuspendLayout()
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

    function New-Button {
        param($text, $x, $y, $width, $height)
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $text
        $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size($width, $height)
        return $btn
    }

    function Set-ButtonLocation {
        param (
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Form]$form,

            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Button]$btnExit
        )
    
        # Set the button's location to the bottom-right corner of the form
        $btnExit.Location = New-Object System.Drawing.Point(
            ($form.ClientSize.Width - $btnExit.Width - 15),
            ($form.ClientSize.Height - $btnExit.Height - 5)
        )
    }

    function New-TreeView {
        param($imgList)
        $tree = New-Object System.Windows.Forms.TreeView
        $tree.Location = New-Object System.Drawing.Point(10, 40)
        $tree.Size = New-Object System.Drawing.Size(300, 380)
        $tree.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Bottom"
        $tree.CheckBoxes = $true
        $tree.ImageList = $imgList

        return $tree
    }

    function New-DetailsTextBox {
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point(320, 40)
        $txt.Size = New-Object System.Drawing.Size(510, 380)
        $txt.Multiline = $true
        $txt.ReadOnly = $true
        $txt.ScrollBars = "Vertical"
        $txt.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Bottom,Right"
        return $txt
    }

    function Populate-Tree {
        param($tree, $appWrappers, $imgList)

        $sorted = $appWrappers | Sort-Object @{Expression = { $_.appGroup } }, @{Expression = { $_.appName } }
        $groupMap = @{}
        $ungrouped = @()

        $uninstallRoots = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        )

        foreach ($w in $sorted) {
            $w | Add-Member IsInstalled $false -Force
            $w | Add-Member InstalledVersion "" -Force

            if ($w.appInstallRegistryData) {
                foreach ($root in $uninstallRoots) {
                    $regPath = Join-Path $root $w.appInstallRegistryData
                    if (Get-ValidateRegistryExists -Path $regPath) {
                        $w.IsInstalled = $true
                        $ver = Get-RegistryValue -Path $regPath -Name "DisplayVersion"
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
                #Write-Host "Inspecting Group $($w): $($w | Format-List -Force)"
                $appNode.Tag = $w
                if ($w.IsInstalled) {
                    $appNode.ImageKey = "installed"
                    $appNode.SelectedImageKey = "installed"
                }
                else {
                    $appNode.ImageKey = "portable"
                    $appNode.SelectedImageKey = "portable"
                }
                $groupMap[$w.appGroup].Nodes.Add($appNode) | Out-Null
            }
            else {
                $ungrouped += $w
            }
        }

        # Add ungrouped apps
        foreach ($w in $ungrouped) {
            $node = New-Object System.Windows.Forms.TreeNode($w.appName)
            #Write-Host "Inspecting Ungroup $($w): $($w | Format-List -Force)"
            $node.Tag = $w
            if ($w.IsInstalled) {
                $node.ImageKey = "installed"
                $node.SelectedImageKey = "installed"
            }
            else {
                $node.ImageKey = "portable"
                $node.SelectedImageKey = "portable"
            }
            $tree.Nodes.Add($node) | Out-Null
        }

        $tree.ExpandAll()
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

    function Add-EventHandlers {
        param($form, $tree, $txt, $btnSelectAll, $btnUnselectAll, $btnInvert, $btnExit)

        $tree.Add_AfterSelect({
                param($sender, $e)
                $node = $e.Node
                if ($node.Tag) {
                    $w = $node.Tag
                    $sb = [System.Text.StringBuilder]::new()
                    $sb.AppendLine("App Name: " + $(if ($w.IsInstalled) { "[Already Installed]" } else { "" }))
                    $sb.AppendLine($w.appName)
                    $sb.AppendLine()
                    if ($w.PSObject.Properties.Match("appVersion").Count) {
                        if ($w.IsInstalled) {
                            $sb.AppendLine("Installed -> Portable Version:")
                            $sb.AppendLine("$($w.InstalledVersion) -> $($w.appVersion)")
                        }
                        else {
                            $sb.AppendLine("Version:")
                            $sb.AppendLine($w.appVersion)
                        }
                    }
                    $sb.AppendLine()
                    $sb.AppendLine("Group: " + $(if ($w.appGroup) { $w.appGroup } else { "<UNGROUPED>" }))
                    $sb.AppendLine()
                    $sb.AppendLine("StartMenu Folder: " + $w.appStartMenuFolderName)
                    $sb.AppendLine()
                    $desc = $w.appDescription -replace '\\n', [Environment]::NewLine
                    $sb.AppendLine("Description:`n$desc")
                    $txt.Text = $sb.ToString().TrimEnd()
                }
                else {
                    $txt.Text = ""
                }
            })

        $tree.Add_AfterCheck({
                if ($_.Action -eq [System.Windows.Forms.TreeViewAction]::ByMouse) {
                    foreach ($c in $_.Node.Nodes) { $c.Checked = $_.Node.Checked }
                }
            })

        $btnSelectAll.Add_Click({ Set-AllTreeNodesChecked $tree.Nodes $true })
        $btnUnselectAll.Add_Click({ Set-AllTreeNodesChecked $tree.Nodes $false })
        $btnInvert.Add_Click({ Invert-AllTreeNodesChecked $tree.Nodes })

        $btnExit.Add_Click({ $form.Close() })
    }

    # endregion

    # region ─── Main UI Construction ───────────────────────────────────────────

    $iconPath = Join-Path $env:SystemRoot "System32\shell32.dll"
    $formIcon = Get-FirstIcon (Get-IconFromFile -FilePath $iconPath -IconIndex 26 -Large)

    $form = New-Form "$scriptName $scriptVersion" $formIcon
    $imgList = Initialize-ImageList

    $iconDll = Join-Path $env:SystemRoot "System32\imageres.dll"
    Add-IconToImageList $imgList $iconDll 82 "installed" -Warn
    Add-IconToImageList $imgList $iconDll 3  "folder" -Warn
    Add-IconToImageList $imgList $iconDll 10 "empty"  -Warn
    Add-IconToImageList $imgList $iconDll 248 "startMenu"  -Warn
    if ($formIcon) { $imgList.Images.Add("portable", $formIcon.ToBitmap()) }

    $tree = New-TreeView $imgList
    $txt = New-DetailsTextBox
    $btnSelectAll = New-Button "Select All" 10 10 80 25
    $btnUnselectAll = New-Button "Unselect All" 100 10 80 25
    $btnInvert = New-Button "Invert Selection" 190 10 120 25
    $btnExit = New-Button "Exit" 0 0 80 30
    $btnExit.Anchor = [System.Windows.Forms.AnchorStyles] "Bottom,Right"

    $form.Controls.AddRange(@($tree, $txt, $btnSelectAll, $btnUnselectAll, $btnInvert, $btnExit))

    Populate-Tree $tree $appWrappers $imgList
    Add-EventHandlers $form $tree $txt $btnSelectAll $btnUnselectAll $btnInvert $btnExit

    # Exit button positioning on resize
    $form.Add_Resize({
            try {
                Set-ButtonLocation $form $btnExit
            }
            catch {}
        })

    Set-ButtonLocation $form $btnExit

    # Resume layout before showing the form
    $form.ResumeLayout($false)
    [void]$form.ShowDialog()
        
    # endregion
}


# ---- Main ----

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
# -------------------------------------
# Comment the following line on production release 
$scriptDir = "d:\Portables"
# -------------------------------------
$appFiles = Get-ChildItem -Path $scriptDir -Recurse -Filter *.app -ErrorAction SilentlyContinue
Write-Host "Total Portable Apps: '$($appFiles.Count)'"
$appObjs = @()
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
    $appObjs += $wrapper
}

if ($appObjs.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No .app files found under $scriptDir", "Error", "OK", "Error")
    exit 0
}

Show-ManageUI -appWrappers $appObjs
