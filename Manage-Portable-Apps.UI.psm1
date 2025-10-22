<##
.SYNOPSIS
    UI module for Manage Portable Apps - builds and shows the WinForms UI.

.DESCRIPTION
    Contains UI factories, panel builders, Set-FillTree, Update-MenuItemStates,
    Register-EventHandlers and exported Show-ManageUI function.

.NOTES
    - PowerShell compatibility: aims for PowerShell v2.0+ compatibility.
    - Requires: .NET System.Windows.Forms, System.Drawing assemblies (available on Windows PowerShell).
    - Icon extraction uses small Add-Type P/Invoke wrapper - also works on PowerShell v2.
    - Exported functions: Show-ManageUI

.SUGGESTED NAME
    Manage-Portable-Apps.UI.psm1
#>

# Ensure WinForms / Drawing types loaded
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------------------------------------------------
# Icon extraction Add-Type (Win32 calls). Safe to add once.
# ----------------------------------------------------------------------
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
"@ -ErrorAction Stop
}

# ----------------------------------------------------------------------
# Helper: Extract icon from file
# ----------------------------------------------------------------------
function Get-IconFromFile {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [int]$IconIndex = 0,
        [switch]$Large
    )

    try {
        if (-not (Test-Path $FilePath)) { return $null }

        if ($Large) {
            $largeIcons = New-Object IntPtr[] 1
            $result = [IconExtractor]::ExtractIconEx($FilePath, $IconIndex, $largeIcons, $null, 1)
            if ($result -gt 0 -and $largeIcons[0] -ne [IntPtr]::Zero) {
                $icon = [System.Drawing.Icon]::FromHandle($largeIcons[0]).Clone()
                [IconExtractor]::DestroyIcon($largeIcons[0])
                return $icon
            }
        }
        else {
            $smallIcons = New-Object IntPtr[] 1
            $result = [IconExtractor]::ExtractIconEx($FilePath, $IconIndex, $null, $smallIcons, 1)
            if ($result -gt 0 -and $smallIcons[0] -ne [IntPtr]::Zero) {
                $icon = [System.Drawing.Icon]::FromHandle($smallIcons[0]).Clone()
                [IconExtractor]::DestroyIcon($smallIcons[0])
                return $icon
            }
        }

        # Fallback
        $hIcon = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $FilePath, $IconIndex)
        if ($hIcon -ne [IntPtr]::Zero) {
            $icon = [System.Drawing.Icon]::FromHandle($hIcon).Clone()
            [IconExtractor]::DestroyIcon($hIcon)
            return $icon
        }
        return $null
    }
    catch {
        Write-Warning "Get-IconFromFile: $($_.Exception.Message)"
        return $null
    }
}

# ----------------------------------------------------------------------
# Helper: pick first valid icon
# ----------------------------------------------------------------------
function Get-FirstIcon {
    param($icons)
    if ($null -eq $icons) { return $null }
    if ($icons -is [System.Array]) {
        foreach ($candidate in $icons) {
            if ($candidate -is [System.Drawing.Icon]) { return $candidate }
        }
        return $null
    }
    else {
        if ($icons -is [System.Drawing.Icon]) { return $icons }
        return $null
    }
}

# ----------------------------------------------------------------------
# Small UI factories (kept small and reused)
# ----------------------------------------------------------------------
function New-Form {
    param($title, $icon)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    if ($icon) { $form.Icon = $icon }
    $form.Size = New-Object System.Drawing.Size(850, 500)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = $form.Size
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
    if ($items) { $cmb.Items.AddRange($items) }
    if ($cmb.Items.Count -gt 0) { $cmb.SelectedIndex = 0 }
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

# ----------------------------------------------------------------------
# Small helpers used by UI
# ----------------------------------------------------------------------
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
            $sb.AppendLine("App Name: [On Start Menu - identical]") | Out-Null
        }
        else {
            $sb.AppendLine("App Name: [On Start Menu - changed]") | Out-Null
        }
    }
    elseif ($w.IsInstalled) {
        $sb.AppendLine("App Name: [Already Installed]") | Out-Null
    }
    else {
        $sb.AppendLine("App Name:") | Out-Null
    }
    $sb.AppendLine($w.appName) | Out-Null
    $sb.AppendLine() | Out-Null

    if ($w.PSObject.Properties.Match("appVersion").Count) {
        if ($w.HasShortcut) {
            $sb.AppendLine("StartMenu -> Portable Version:") | Out-Null
            $sb.AppendLine(" $($w.ShortcutAppVersion) -> $($w.appVersion)") | Out-Null
        }
        elseif ($w.IsInstalled) {
            $sb.AppendLine("Installed -> Portable Version:") | Out-Null
            $sb.AppendLine("$($w.InstalledVersion) -> $($w.appVersion)") | Out-Null
        }
        else {
            $sb.AppendLine("Version:") | Out-Null
            $sb.AppendLine($w.appVersion) | Out-Null
        }
    }
    $sb.AppendLine() | Out-Null

    $grp = if ($w.appGroup) { $w.appGroup } else { "<UNGROUPED>" }
    $sb.AppendLine("Group: " + $grp) | Out-Null
    $sb.AppendLine() | Out-Null

    $sb.AppendLine("StartMenu Folder: " + $w.appStartMenuFolderName) | Out-Null
    $sb.AppendLine() | Out-Null

    $desc = $w.appDescription -replace '\\n', [Environment]::NewLine
    $sb.AppendLine("Description:`n$desc") | Out-Null

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

function Set-ToggleTreeNodeSelection {
    param($nodes)
    foreach ($n in $nodes) {
        $n.Checked = -not $n.Checked
        if ($n.Nodes.Count -gt 0) {
            Set-ToggleTreeNodeSelection $n.Nodes
        }
    }
}

# ----------------------------------------------------------------------
# Set-FillTree: builds nodes from app wrappers and created shortcuts
# - explicit parameters; no reliance on outer variables
# ----------------------------------------------------------------------
function Set-FillTree {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TreeView] $Tree,
        [Parameter(Mandatory)][object[]] $AppWrappers,
        [Parameter(Mandatory)][System.Windows.Forms.ImageList] $ImgList,
        [string] $Filter = "All"
    )

    $Tree.Nodes.Clear()

    switch ($Filter) {
        "All" { $filtered = $AppWrappers }
        "Installed" { $filtered = $AppWrappers | Where-Object { $_.IsInstalled } }
        "Portable on StartMenu" { $filtered = $AppWrappers | Where-Object { $_.HasShortcut } }
        "Portable not used" { $filtered = $AppWrappers | Where-Object { (-not $_.IsInstalled) -and (-not $_.HasShortcut) } }
        default { $filtered = $AppWrappers }
    }

    $sorted = $filtered | Sort-Object @{Expression = { $_.appGroup } }, @{Expression = { $_.appName } }
    $groupMap = @{}
    $ungrouped = @()

    foreach ($w in $sorted) {
        if ($w.appGroup) {
            if (-not $groupMap.ContainsKey($w.appGroup)) {
                $groupNode = New-Object System.Windows.Forms.TreeNode($w.appGroup)
                $groupNode.ImageKey = "folder"
                $groupNode.SelectedImageKey = "folder"
                $groupMap[$w.appGroup] = $groupNode
                $Tree.Nodes.Add($groupNode) | Out-Null
            }
            $appNode = New-Object System.Windows.Forms.TreeNode($w.appName)
            $appNode.Tag = $w
            Set-AppNodeIcon -Node $appNode -App $w -ImgList $ImgList
            $groupMap[$w.appGroup].Nodes.Add($appNode) | Out-Null
        }
        else {
            $ungrouped += $w
        }
    }

    foreach ($w in $ungrouped) {
        $node = New-Object System.Windows.Forms.TreeNode($w.appName)
        $node.Tag = $w
        Set-AppNodeIcon -Node $node -App $w -ImgList $ImgList
        $Tree.Nodes.Add($node) | Out-Null
    }

    $Tree.ExpandAll()
}

# ----------------------------------------------------------------------
# Update-MenuItemStates: controls enabling/disabling of menu items
# ----------------------------------------------------------------------
function Update-MenuItemStates {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TreeView] $Tree,
        [Parameter(Mandatory)][System.Windows.Forms.ToolStripMenuItem] $MenuAddShortcut,
        [Parameter(Mandatory)][System.Windows.Forms.ToolStripMenuItem] $MenuRemoveShortcut
    )

    # Get checked nodes with Tag
    $checkedApps = New-Object System.Collections.ArrayList

    function Get-CheckedNodes {
        param($nodes)
        foreach ($node in $nodes) {
            if ($node.Checked -and $node.Tag) { [void]$checkedApps.Add($node.Tag) }
            if ($node.Nodes.Count -gt 0) { Get-CheckedNodes $node.Nodes }
        }
    }
    Get-CheckedNodes $Tree.Nodes

    $hasItemsToAdd = $false
    foreach ($app in $checkedApps) {
        if (($app.PSObject.Properties.Match("HasShortcut").Count -gt 0 -and -not $app.HasShortcut) -or
            ($app.PSObject.Properties.Match("IsBothSame").Count -gt 0 -and -not $app.IsBothSame)) {
            $hasItemsToAdd = $true
            break
        }
    }
    $MenuAddShortcut.Enabled = $hasItemsToAdd

    $hasItemsToRemove = $false
    foreach ($app in $checkedApps) {
        if ($app.PSObject.Properties.Match("HasShortcut").Count -gt 0 -and $app.HasShortcut) {
            $hasItemsToRemove = $true
            break
        }
    }
    $MenuRemoveShortcut.Enabled = $hasItemsToRemove
}

# ----------------------------------------------------------------------
# Build panel helper functions: return small hashtable of relevant controls
# ----------------------------------------------------------------------

function Build-MenuStrip {
    $menuStrip = New-Object System.Windows.Forms.MenuStrip
    $menuStrip.Dock = [System.Windows.Forms.DockStyle]::Top

    # File Menu
    $menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuFile.Text = "&File"

    $menuAddShortcut = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuAddShortcut.Text = "&Add or Update Selected Shortcuts"
    $menuAddShortcut.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::A
    $menuAddShortcut.Enabled = $false

    $menuRemoveShortcut = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuRemoveShortcut.Text = "&Remove Selected Shortcuts"
    $menuRemoveShortcut.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::R
    $menuRemoveShortcut.Enabled = $false

    $menuSeparator1 = New-Object System.Windows.Forms.ToolStripSeparator

    $menuCreateApp = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuCreateApp.Text = "&Create '.app' file"
    $menuCreateApp.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::N

    $menuSeparator2 = New-Object System.Windows.Forms.ToolStripSeparator

    $menuExit = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuExit.Text = "E&xit"
    $menuExit.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4

    $menuFile.DropDownItems.Add($menuAddShortcut) | Out-Null
    $menuFile.DropDownItems.Add($menuRemoveShortcut) | Out-Null
    $menuFile.DropDownItems.Add($menuSeparator1) | Out-Null
    $menuFile.DropDownItems.Add($menuCreateApp) | Out-Null
    $menuFile.DropDownItems.Add($menuSeparator2) | Out-Null
    $menuFile.DropDownItems.Add($menuExit) | Out-Null

    # Help Menu
    $menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuHelp.Text = "&Help"

    $menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuAbout.Text = "&About"
    $menuAbout.ShortcutKeys = [System.Windows.Forms.Keys]::F1

    $menuHelp.DropDownItems.Add($menuAbout) | Out-Null

    $menuStrip.Items.Add($menuFile) | Out-Null
    $menuStrip.Items.Add($menuHelp) | Out-Null

    return @{
        MenuStrip          = $menuStrip
        MenuAddShortcut    = $menuAddShortcut
        MenuRemoveShortcut = $menuRemoveShortcut
        MenuCreateApp      = $menuCreateApp
        MenuExit           = $menuExit
        MenuAbout          = $menuAbout
    }
}

function Build-TopBar {
    param([System.Windows.Forms.ImageList] $ImgList)

    $lblFilter = New-Label "Filters:"
    $cmbFilter = New-ComboBox @("All", "Installed", "Portable on StartMenu", "Portable not used")
    $btnAbout = New-Button "About" ([System.Windows.Forms.DockStyle]::Right)

    return @{ Label = $lblFilter; CmbFilter = $cmbFilter; BtnAbout = $btnAbout }
}

function Build-LeftPanel {
    param([System.Windows.Forms.ImageList] $ImgList)

    $tree = New-TreeView $ImgList
    $btnSelectAll = New-Button "Select All"
    $btnUnselectAll = New-Button "Unselect All"
    $btnInvert = New-Button "Invert Selection"

    return @{ Tree = $tree; BtnSelectAll = $btnSelectAll; BtnUnselectAll = $btnUnselectAll; BtnInvert = $btnInvert }
}

function Build-BottomPanel {
    $btnExit = New-Button "Exit"
    return @{
        BtnExit = $btnExit
    }
}

# ----------------------------------------------------------------------
# Build-MainLayout: assemble TableLayoutPanels and place controls; returns a hashtable with main controls
# ----------------------------------------------------------------------
function Build-MainLayout {
    param(
        [string] $Title,
        $Icon,
        [System.Windows.Forms.ImageList] $ImgList,
        [hashtable] $MenuStripInfo,
        [hashtable] $TopBar,
        [hashtable] $LeftPanel,
        [System.Windows.Forms.TextBox] $TxtDetails,
        [hashtable] $BottomPanel
    )

    $form = New-Form -title $Title -icon $Icon
    $form.Padding = New-Object System.Windows.Forms.Padding(0)
        
    # Add MenuStrip to form
    $form.MainMenuStrip = $MenuStripInfo.MenuStrip

    $MenuStripInfo.MenuStrip.Margin = New-Object System.Windows.Forms.Padding(0)

    # Main layout (3 rows now)
    $mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainLayout.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
    $mainLayout.Margin = New-Object System.Windows.Forms.Padding(0)
    # ----------------------------------------------------------------------
    $mainLayout.ColumnCount = 1
    $mainLayout.RowCount = 4
    # ----------------------------------------------------------------------
    # FIX IS HERE: Padding (Left, Top, Right, Bottom) - set Top to 0
    #$mainLayout.Padding = New-Object System.Windows.Forms.Padding(10, 0, 10, 10) 
    # ----------------------------------------------------------------------
    # Row Styles:
    # 1. MenuStrip (height: 25-30 is typical for a menu strip)
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 27))) | Out-Null
    # 2. TopBar (Filter controls)
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35))) | Out-Null
    # 3. Content (Tree and Details)
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    # 4. BottomPanel (Exit button)
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45))) | Out-Null



    # TopBar table (filter only now)
    $topBarLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $topBarLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $topBarLayout.ColumnCount = 3
    $topBarLayout.RowCount = 1
    $topBarLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $topBarLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 200))) | Out-Null
    $topBarLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

    $topBarLayout.Controls.Add($TopBar.Label, 0, 0)
    $topBarLayout.Controls.Add($TopBar.CmbFilter, 1, 0)

    # Content layout
    $contentLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $contentLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $contentLayout.ColumnCount = 2
    $contentLayout.RowCount = 1
    $contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 310))) | Out-Null
    $contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

    # Left panel layout
    $leftPanelLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $leftPanelLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $leftPanelLayout.ColumnCount = 1
    $leftPanelLayout.RowCount = 2
    $leftPanelLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $leftPanelLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35))) | Out-Null

    # Selection buttons sublayout
    $selectionButtonsLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $selectionButtonsLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $selectionButtonsLayout.ColumnCount = 3
    $selectionButtonsLayout.RowCount = 1
    $selectionButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
    $selectionButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
    $selectionButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.34))) | Out-Null

    $selectionButtonsLayout.Controls.Add($LeftPanel.BtnSelectAll, 0, 0)
    $selectionButtonsLayout.Controls.Add($LeftPanel.BtnUnselectAll, 1, 0)
    $selectionButtonsLayout.Controls.Add($LeftPanel.BtnInvert, 2, 0)

    $leftPanelLayout.Controls.Add($LeftPanel.Tree, 0, 0)
    $leftPanelLayout.Controls.Add($selectionButtonsLayout, 0, 1)

    $contentLayout.Controls.Add($leftPanelLayout, 0, 0)
    $contentLayout.Controls.Add($TxtDetails, 1, 0)

    # Bottom buttons layout (only Exit button now)
    $bottomButtonsLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $bottomButtonsLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $bottomButtonsLayout.ColumnCount = 2
    $bottomButtonsLayout.RowCount = 1
    $bottomButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $bottomButtonsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90))) | Out-Null

    $bottomButtonsLayout.Controls.Add($BottomPanel.BtnExit, 1, 0)

    # Add rows to main
    $mainLayout.Controls.Add($MenuStripInfo.MenuStrip, 0, 0) # <--- New: MenuStrip in Row 0
    $mainLayout.Controls.Add($topBarLayout, 0, 1)          # <--- Row 1
    $mainLayout.Controls.Add($contentLayout, 0, 2)         # <--- Row 2
    $mainLayout.Controls.Add($bottomButtonsLayout, 0, 3)   # <--- Row 3

    $form.Controls.Add($mainLayout)

    # Return controls of interest
    return @{
        Form               = $form
        Tree               = $LeftPanel.Tree
        CmbFilter          = $TopBar.CmbFilter
        Txt                = $TxtDetails
        BtnSelectAll       = $LeftPanel.BtnSelectAll
        BtnUnselectAll     = $LeftPanel.BtnUnselectAll
        BtnInvert          = $LeftPanel.BtnInvert
        MenuAddShortcut    = $MenuStripInfo.MenuAddShortcut
        MenuRemoveShortcut = $MenuStripInfo.MenuRemoveShortcut
        MenuCreateApp      = $MenuStripInfo.MenuCreateApp
        MenuAbout          = $MenuStripInfo.MenuAbout
        MenuExit           = $MenuStripInfo.MenuExit
        BtnExit            = $BottomPanel.BtnExit
        ImgList            = $ImgList
    }
}

# ----------------------------------------------------------------------
# Register-EventHandlers: wire events; parameters explicit
# NOTE: .GetNewClosure() is used to ensure event handlers keep references
# to the controls after this function returns.
# ----------------------------------------------------------------------
function Register-EventHandlers {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Form] $Form,
        [Parameter(Mandatory)][System.Windows.Forms.TreeView] $Tree,
        [Parameter(Mandatory)][System.Windows.Forms.ComboBox] $CmbFilter,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox] $Txt,
        [Parameter(Mandatory)][System.Windows.Forms.Button] $BtnSelectAll,
        [Parameter(Mandatory)][System.Windows.Forms.Button] $BtnUnselectAll,
        [Parameter(Mandatory)][System.Windows.Forms.Button] $BtnInvert,
        [Parameter(Mandatory)][System.Windows.Forms.ToolStripMenuItem] $MenuAddShortcut,
        [Parameter(Mandatory)][System.Windows.Forms.ToolStripMenuItem] $MenuRemoveShortcut,
        [Parameter(Mandatory)][System.Windows.Forms.ToolStripMenuItem] $MenuCreateApp,
        [Parameter(Mandatory)][System.Windows.Forms.ToolStripMenuItem] $MenuAbout,
        [Parameter(Mandatory)][System.Windows.Forms.ToolStripMenuItem] $MenuExit,
        [Parameter(Mandatory)][System.Windows.Forms.Button] $BtnExit,
        [Parameter(Mandatory)][object[]] $AppWrappers,
        [Parameter(Mandatory)][System.Windows.Forms.ImageList] $ImgList
    )

    # Filter change: use new-closure to capture current variables
    $sbFilter = {
        Set-FillTree -Tree $Tree -AppWrappers $AppWrappers -ImgList $ImgList -Filter $CmbFilter.SelectedItem
        Update-MenuItemStates -Tree $Tree -MenuAddShortcut $MenuAddShortcut -MenuRemoveShortcut $MenuRemoveShortcut
    }.GetNewClosure()
    $CmbFilter.Add_SelectedIndexChanged($sbFilter)

    # Tree selection
    $sbSel = {
        param($button, $e)
        $node = $e.Node
        if ($node -and $node.Tag) {
            Update-DetailsTextBox -txt $Txt -w $node.Tag
        }
        else { $Txt.Text = "" }
    }.GetNewClosure()
    $Tree.Add_AfterSelect($sbSel)

    # AfterCheck
    $sbAfterCheck = {
        param($button, $e)
        if ($e.Action -eq [System.Windows.Forms.TreeViewAction]::ByMouse) {
            foreach ($c in $e.Node.Nodes) { $c.Checked = $e.Node.Checked }
            Update-ParentCheckState $e.Node
        }
        Update-MenuItemStates -Tree $Tree -MenuAddShortcut $MenuAddShortcut -MenuRemoveShortcut $MenuRemoveShortcut
    }.GetNewClosure()
    $Tree.Add_AfterCheck($sbAfterCheck)

    # Selection buttons
    $sbSelectAll = {
        Set-AllTreeNodesChecked $Tree.Nodes $true
        foreach ($n in $Tree.Nodes) { Update-ParentCheckState $n }
        Update-MenuItemStates -Tree $Tree -MenuAddShortcut $MenuAddShortcut -MenuRemoveShortcut $MenuRemoveShortcut
    }.GetNewClosure()
    $BtnSelectAll.Add_Click($sbSelectAll)

    $sbUnselectAll = {
        Set-AllTreeNodesChecked $Tree.Nodes $false
        foreach ($n in $Tree.Nodes) { Update-ParentCheckState $n }
        Update-MenuItemStates -Tree $Tree -MenuAddShortcut $MenuAddShortcut -MenuRemoveShortcut $MenuRemoveShortcut
    }.GetNewClosure()
    $BtnUnselectAll.Add_Click($sbUnselectAll)

    $sbInvert = {
        Set-ToggleTreeNodeSelection $Tree.Nodes
        foreach ($n in $Tree.Nodes) { Update-ParentCheckState $n }
        Update-MenuItemStates -Tree $Tree -MenuAddShortcut $MenuAddShortcut -MenuRemoveShortcut $MenuRemoveShortcut
    }.GetNewClosure()
    $BtnInvert.Add_Click($sbInvert)

    # Action menu items - TODO: these should call out to main script functions (dot-sourced)
    $sbAdd = {
        [System.Windows.Forms.MessageBox]::Show("Add Shortcut functionality will be implemented by main script.", "Add Shortcut", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }.GetNewClosure()
    $MenuAddShortcut.Add_Click($sbAdd)

    $sbRemove = {
        [System.Windows.Forms.MessageBox]::Show("Remove Shortcut functionality will be implemented by main script.", "Remove Shortcut", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }.GetNewClosure()
    $MenuRemoveShortcut.Add_Click($sbRemove)

    $sbCreate = {
        [System.Windows.Forms.MessageBox]::Show("Create '.app' file functionality will be implemented by main script.", "Create .app File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }.GetNewClosure()
    $MenuCreateApp.Add_Click($sbCreate)

    # About: show brief info (capturing $scriptVersion is optional; if not present will be blank)
    $sbAbout = {
        [void] [System.Windows.Forms.MessageBox]::Show(
            "This script was fully written by AI assistants (ChatGPT + Claude.ai) and then guided, managed, and refined by me.`r`n`r`n" +
            "Manage Portable Apps is a PowerShell utility designed to help you keep your portable applications organized and accessible. " +
            "It scans a folder of portable apps (each identified by its own .app metadata file) and compares them against your Start Menu shortcuts. " +
            "You can see which apps are installed, which already have shortcuts, and which are unused. " +
            "With just a few clicks you can add or remove shortcuts and create new .app entries.`r`n`r`n" +
            "Author: MrkTheCoder`r`n" +
            ("Version: $($scriptVersion)") + "`r`n" +
            "© 2025 MrkTheCoder = All Rights Reserved",
            "About Manage-Portable-Apps",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }.GetNewClosure()
    $MenuAbout.Add_Click($sbAbout)

    $sbExit = { $Form.Close() }.GetNewClosure()
    $MenuExit.Add_Click($sbExit)
    $BtnExit.Add_Click($sbExit)
}

# ----------------------------------------------------------------------
# Show-ManageUI: exported orchestrator that builds the UI and shows it
# ----------------------------------------------------------------------
function Show-ManageUI {
    param(
        [Parameter(Mandatory = $true)][object[]] $appWrappers,
        [string] $Title = "Manage Portable Apps"
    )

    # Prepare ImageList and icons
    $imgList = Initialize-ImageList

    # Add standard icons (guarded)
    $iconDll = Join-Path $env:SystemRoot "System32\imageres.dll"
    Add-IconToImageList -imgList $imgList -file $iconDll -index 82 -key "installed" -Warn
    Add-IconToImageList -imgList $imgList -file $iconDll -index 3  -key "folder" -Warn
    Add-IconToImageList -imgList $imgList -file $iconDll -index 10 -key "empty" -Warn
    Add-IconToImageList -imgList $imgList -file $iconDll -index 248 -key "startMenu" -Warn

    # small "changed" icon from embedded base64 - keep inline to avoid separate file
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
    $mem = New-Object IO.MemoryStream(, $iconBytes)
    try {
        $changedIcon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Icon]::new($mem)).Handle)
        $imgList.Images.Add("changed", $changedIcon.ToBitmap())
    }
    catch { }
    finally {
        if ($mem) { $mem.Dispose() }
        if ($changedIcon) { try { $changedIcon.Dispose() } catch {} }
    }

    # If a form icon was provided, add a portable icon
    # Build a small icon for the form (optional) - try reading shell32 default icon
    $FormIcon = $null
    try {
        $sysIco = Join-Path $env:SystemRoot "System32\shell32.dll"
        $formIcon = Get-FirstIcon (Get-IconFromFile -FilePath $sysIco -IconIndex 26 -Large)
        $imgList.Images.Add("portable", $formIcon.ToBitmap())
    }
    catch { }


    # Build panels
    $topBar = Build-TopBar -ImgList $imgList
    $leftPanel = Build-LeftPanel -ImgList $imgList
    $bottomPanel = Build-BottomPanel
    $menuStrip = Build-MenuStrip
    $txtDetails = New-DetailsTextBox

    $layoutInfo = Build-MainLayout -Title $Title -Icon $FormIcon -ImgList $imgList -MenuStripInfo $menuStrip -TopBar $topBar -LeftPanel $leftPanel -TxtDetails $txtDetails -BottomPanel $bottomPanel

    # Wire events (use GetNewClosure in Register-EventHandlers)
    Register-EventHandlers -Form $layoutInfo.Form -Tree $layoutInfo.Tree -CmbFilter $layoutInfo.CmbFilter -Txt $layoutInfo.Txt `
        -BtnSelectAll $layoutInfo.BtnSelectAll -BtnUnselectAll $layoutInfo.BtnUnselectAll -BtnInvert $layoutInfo.BtnInvert `
        -MenuAddShortcut $layoutInfo.MenuAddShortcut -MenuRemoveShortcut $layoutInfo.MenuRemoveShortcut -MenuCreateApp $layoutInfo.MenuCreateApp `
        -MenuAbout $layoutInfo.MenuAbout -MenuExit $layoutInfo.MenuExit -BtnExit $layoutInfo.BtnExit -AppWrappers $appWrappers -ImgList $imgList

    # Initial populate & button state
    Set-FillTree -Tree $layoutInfo.Tree -AppWrappers $appWrappers -ImgList $imgList -Filter $layoutInfo.CmbFilter.SelectedItem
    Update-MenuItemStates -Tree $layoutInfo.Tree -MenuAddShortcut $layoutInfo.MenuAddShortcut -MenuRemoveShortcut $layoutInfo.MenuRemoveShortcut

    # Show dialog
    [void]$layoutInfo.Form.ShowDialog()

    # Cleanup
    try { $imgList.Dispose() } catch {}
    try { $layoutInfo.Form.Dispose() } catch {}
}

# ==============================================
# Exported entry point
# ==============================================
Export-ModuleMember -Function Show-ManageUI, Register-EventHandlers, Set-FillTree, Update-MenuItemStates, Update-ParentCheckState, Update-DetailsTextBox, Get-IconFromFile, Get-FirstIcon, Set-AllTreeNodesChecked, Set-ToggleTreeNodeSelection