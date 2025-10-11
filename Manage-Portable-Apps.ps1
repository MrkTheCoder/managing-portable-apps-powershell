# Manage-Portable-Apps.ps1

$scriptName = "Manage Portable Apps"
$scriptVersion = "v0.1.202510 Pre-Alpha"
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
    # Path to the DLL (or EXE) containing the icon
    $iconPath = [System.IO.Path]::Combine($env:SystemRoot, "System32", "shell32.dll")
    $icon = Get-IconFromFile -FilePath $iconPath -IconIndex 26 -Large

    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$scriptName $scriptVersion"
    if ($icon) {
        $form.Icon = $icon[1]
    }
    else {
        Write-Warning "Could not load icon"
    }
    $form.Size = New-Object System.Drawing.Size(830, 500)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = $form.Size

    $form.SuspendLayout()

    # TreeView
    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Location = New-Object System.Drawing.Point(10, 40)
    $tree.Size = New-Object System.Drawing.Size(300, 380)
    $tree.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Bottom"
    $tree.CheckBoxes = $true
    $form.Controls.Add($tree)

    # TextBox for details
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(320, 40)
    $txt.Size = New-Object System.Drawing.Size(490, 380)
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = "Vertical"
    $txt.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Bottom,Right"
    $form.Controls.Add($txt)

    # Buttons above tree
    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = "Select All"
    $btnSelectAll.Location = New-Object System.Drawing.Point(10, 10)
    $btnSelectAll.Size = New-Object System.Drawing.Size(80, 25)
    $form.Controls.Add($btnSelectAll)

    $btnUnselectAll = New-Object System.Windows.Forms.Button
    $btnUnselectAll.Text = "Unselect All"
    $btnUnselectAll.Location = New-Object System.Drawing.Point(100, 10)
    $btnUnselectAll.Size = New-Object System.Drawing.Size(80, 25)
    $form.Controls.Add($btnUnselectAll)

    $btnInvert = New-Object System.Windows.Forms.Button
    $btnInvert.Text = "Invert Selection"
    $btnInvert.Location = New-Object System.Drawing.Point(190, 10)
    $btnInvert.Size = New-Object System.Drawing.Size(120, 25)
    $form.Controls.Add($btnInvert)

    # Exit button
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text = "Exit"
    $btnExit.Size = New-Object System.Drawing.Size(80, 30)
    # Set the button's margin (left, top, right, bottom)
    $btnExit.Margin = New-Object System.Windows.Forms.Padding(10, 50, 10, 20)

    $btnExit.Anchor = [System.Windows.Forms.AnchorStyles] "Bottom,Right"
    $form.Controls.Add($btnExit)

    # Sort wrappers by group and name (case-sensitive by default in Sort-Object)
    $sorted = $appWrappers | Sort-Object @{ Expression = { $_.appGroup } }, @{ Expression = { $_.appName } }

    # Build tree: group → app nodes
    # Sort wrappers
    $sorted = $appWrappers | Sort-Object @{ Expression = { $_.appGroup } }, @{ Expression = { $_.appName } }

    $groupMap = @{}
    # A list for ungrouped apps
    $ungrouped = @()

    foreach ($w in $sorted) {
        if (-not [string]::IsNullOrEmpty($w.appGroup)) {
            # App has a group
            $group = $w.appGroup
            if (-not $groupMap.ContainsKey($group)) {
                $gn = New-Object System.Windows.Forms.TreeNode($group)
                $groupMap[$group] = $gn
                $null = $tree.Nodes.Add($gn)
            }
            $child = New-Object System.Windows.Forms.TreeNode($w.appName)
            $child.Tag = $w
            $null = $groupMap[$group].Nodes.Add($child)
        }
        else {
            # No group — handle later
            $ungrouped += $w
        }
    }

    # After adding grouped ones, now add ungrouped apps as root nodes
    foreach ($w in $ungrouped) {
        $node = New-Object System.Windows.Forms.TreeNode($w.appName)
        $node.Tag = $w
        $null = $tree.Nodes.Add($node)
    }

    $tree.ExpandAll()

    # Place Exit button with correct padding
    $fw = Get-ScalarInt $form.ClientSize.Width
    $fh = Get-ScalarInt $form.ClientSize.Height
    $bw = Get-ScalarInt $btnExit.Width
    $bh = Get-ScalarInt $btnExit.Height

    $btnExit.Location = New-Object System.Drawing.Point(
        ($fw - $bw - 10),
        ($fh - $bh - 10)
    )
    
    # Event: AfterSelect
    $tree.Add_AfterSelect({
            param($senderParm, $e)
            $node = $e.Node
            if ($node -and $node.Tag) {
                $w = $node.Tag
                # The wrapper object holds properties from the JSON (flat)
                $sb = New-Object System.Text.StringBuilder
                $nl = [Environment]::NewLine
                $sb.AppendLine("App Name:$nl$($w.appName)$nl")
                if ($w.PSObject.Properties.Match("appVersion").Count -gt 0) {
                    $sb.AppendLine("Version:$nl$($w.appVersion)$nl")
                }
                $sb.AppendLine("Group:$nl$($w.appGroup)$nl")
                $sb.AppendLine("StartMenu Folder:$nl$($w.appStartMenuFolderName)$nl")
                $sb.AppendLine("Description:")
                $desc = $w.appDescription -replace '\\n', $nl
                $sb.AppendLine($desc)
                $txt.Text = $sb.ToString().TrimEnd()
            }
            else {
                $txt.Text = ""
            }
        })

    # AfterCheck for cascading
    $tree.Add_AfterCheck({
            param($senderParm, $e)
            if ($e.Action -eq [System.Windows.Forms.TreeViewAction]::ByMouse) {
                $n = $e.Node
                foreach ($c in $n.Nodes) {
                    $c.Checked = $n.Checked
                }
            }
        })

    # Button actions
    $btnSelectAll.Add_Click({
            foreach ($gn in $tree.Nodes) {
                foreach ($cn in $gn.Nodes) {
                    $cn.Checked = $true
                }
            }
        })
    $btnUnselectAll.Add_Click({
            foreach ($gn in $tree.Nodes) {
                foreach ($cn in $gn.Nodes) {
                    $cn.Checked = $false
                }
            }
        })
    $btnInvert.Add_Click({
            foreach ($gn in $tree.Nodes) {
                foreach ($cn in $gn.Nodes) {
                    $cn.Checked = (-not $cn.Checked)
                }
            }
        })
    $btnExit.Add_Click({ $form.Close() })

    # Resize event for repositioning Exit button
    $form.Add_Resize({
            try {
                $fw2 = Get-ScalarInt $form.ClientSize.Width
                $fh2 = Get-ScalarInt $form.ClientSize.Height
                $bw2 = Get-ScalarInt $btnExit.Width
                $bh2 = Get-ScalarInt $btnExit.Height
                $btnExit.Location = New-Object System.Drawing.Point(
                    ($fw2 - $bw2 - 10),
                    ($fh2 - $bh2 - 10)
                )
            }
            catch {
                # silently ignore positioning errors during resize
            }
        })

    $form.ResumeLayout($false)
    [void] $form.ShowDialog()
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