# Create-AppFile.ps1
# Requires: WScript.Shell COM availability

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Helper Functions

function Select-FolderDialog {
    param([string]$Title = "Select Folder")
    
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Title
    $folderBrowser.ShowNewFolderButton = $true
    
    if (($folderBrowser.ShowDialog()) -eq 'OK') {
        return $folderBrowser.SelectedPath
    }
    return $null
}

function Clean-Description {
    param([string]$desc)
    
    if ([string]::IsNullOrEmpty($desc)) { return "" }
    
    $d = $desc.Trim('"')
    
    # Check if it looks like a path
    if ($d -match '^[A-Za-z]:\\') { return "" }
    if ($d -match '^\\\\') { return "" }
    if ($d -match '[\\/].*[\\/]') { return "" }
    if ($d -match '\.(exe|lnk|bat|url|cmd|msi)$') { return "" }
    
    return $desc
}

function Convert-PathToToken {
    param(
        [string]$fullPath,
        [string]$appPath
    )
    
    if ([string]::IsNullOrEmpty($fullPath)) { return "" }
    
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
    
    # Fallback: system paths
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
    
    return $normFull
}

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

function Get-ShortcutProperties {
    param([string]$lnkPath)
    
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
    param(
        [object]$obj,
        [int]$indent = 0
    )
    
    if ($null -eq $obj) { return "null" }
    
    $pad = " " * $indent
    
    if ($obj -is [hashtable] -or $obj -is [System.Collections.IDictionary]) {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("{")
        $first = $true
        foreach ($k in $obj.Keys) {
            if (-not $first) { [void]$sb.AppendLine(",") }
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
            if (-not $first) { [void]$sb.AppendLine(",") }
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

#endregion

#region GUI Factory Functions

function New-Label {
    param(
        [string]$Text,
        [System.Windows.Forms.AnchorStyles]$Anchor = "Left"
    )
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.Anchor = $Anchor
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $label.Margin = New-Object System.Windows.Forms.Padding(3, 8, 3, 3)
    return $label
}

function New-TextBox {
    param(
        [System.Windows.Forms.AnchorStyles]$Anchor = "Left,Right",
        [bool]$ReadOnly = $false,
        [bool]$Multiline = $false,
        [bool]$Scrollbar = $true,
        [int]$Height = 23
    )
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Anchor = $Anchor
    $textBox.ReadOnly = $ReadOnly
    $textBox.Multiline = $Multiline
    if ($Multiline) {
        $textBox.Height = $Height
        if ($Scrollbar) { $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical }
        $textBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom
    }
    $textBox.Margin = New-Object System.Windows.Forms.Padding(3)
    return $textBox
}

function New-InfoIcon {
    param(
        [string]$TooltipText,
        [string]$HelpText,
        [System.Windows.Forms.TextBox]$HelpTextBox = $null
    )
    
    $icon = New-Object System.Windows.Forms.Label
    $icon.Text = "i"
    $icon.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $icon.ForeColor = [System.Drawing.Color]::DodgerBlue
    $icon.AutoSize = $true
    $icon.Cursor = [System.Windows.Forms.Cursors]::Help
    $icon.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $icon.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $icon.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
    
    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.SetToolTip($icon, $TooltipText)
    $tooltip.AutoPopDelay = 30000
    $tooltip.InitialDelay = 100
    $tooltip.ReshowDelay = 100
    $tooltip.ShowAlways = $true
    
    # Add hover events to update help textbox if provided
    if ($null -ne $HelpTextBox) {
        $icon.Add_MouseEnter({
                $HelpTextBox.Text = $HelpText
            }.GetNewClosure())
        
        $icon.Add_MouseLeave({
                $HelpTextBox.Text = ""
            }.GetNewClosure())
    }
    
    return $icon
}

function New-BrowseButton {
    param([scriptblock]$ClickAction)
    
    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Browse..."
    $button.Width = 80
    $button.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $button.Margin = New-Object System.Windows.Forms.Padding(3)
    $button.Add_Click($ClickAction)
    return $button
}

function New-ActionButton {
    param(
        [string]$Text,
        [int]$Width = 120
    )
    
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Width = $Width
    $button.Height = 30
    $button.Margin = New-Object System.Windows.Forms.Padding(5, 3, 5, 3)
    return $button
}

#endregion

#region GUI Layout Construction

function Build-InputRowLayout {
    param(
        [bool]$HasExtraControl = $false
    )
    
    $subLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $subLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $subLayout.RowCount = 1
    $subLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $subLayout.Margin = New-Object System.Windows.Forms.Padding(0)
    
    # Base columns: Input (stretches) | Icon (fixed 30px)
    $subLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null # Input Control
    $subLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 30))) | Out-Null  # Info Icon
    
    if ($HasExtraControl) {
        # Add a third column for the Browse button
        $subLayout.ColumnCount = 3
        $subLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90))) | Out-Null # Extra Control (Button)
    }
    else {
        $subLayout.ColumnCount = 2
    }
    
    return $subLayout
}

function Build-FormLayout {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Shortcut-to-.app Converter"
    $form.Size = New-Object System.Drawing.Size(700, 600)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(650, 550)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.MaximizeBox = $true
    $FormIcon = $null
    try {
        $sysIco = Join-Path $env:SystemRoot "System32\shell32.dll"
        $formIcon = Get-FirstIcon (Get-IconFromFile -FilePath $sysIco -IconIndex 146 -Large)
        $form.Icon = $formIcon
    }
    catch { }

    # Main TableLayoutPanel
    $mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainLayout.ColumnCount = 1
    $mainLayout.RowCount = 2
    $mainLayout.Padding = New-Object System.Windows.Forms.Padding(10)
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45))) | Out-Null
    
    # Content panel
    $contentLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $contentLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $contentLayout.ColumnCount = 2
    $contentLayout.RowCount = 8
    $contentLayout.AutoSize = $false
    
    # Column styles: Label | Controls
    $contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130))) | Out-Null
    $contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    
    # Row styles - using fixed heights for rows 0-5, expandable for row 6 (Description), fixed for row 7 (Help)
    for ($i = 0; $i -lt 6; $i++) {
        $contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35))) | Out-Null
    }
    # Row 6: Description row - expandable
    $contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    # Row 7: Help row - fixed height for 3 lines of text
    $contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 120))) | Out-Null
    
    return @{
        Form          = $form
        MainLayout    = $mainLayout
        ContentLayout = $contentLayout
    }
}

function Add-FormRow {
    param(
        [System.Windows.Forms.TableLayoutPanel]$Layout,
        [int]$Row,
        [string]$LabelText,
        [System.Windows.Forms.Control]$InputControl,
        [string]$TooltipText,
        [string]$HelpText,
        [System.Windows.Forms.Control]$ExtraControl = $null,
        [System.Windows.Forms.TextBox]$HelpTextBox = $null
    )
    
    # 1. Add label (Main Layout Col 0)
    $label = New-Label -Text $LabelText
    $label.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
    $Layout.Controls.Add($label, 0, $Row)
    
    # 2. Create the inner layout panel for all controls (Main Layout Col 1)
    $subLayout = Build-InputRowLayout -HasExtraControl ($null -ne $ExtraControl)
    $Layout.Controls.Add($subLayout, 1, $Row)
    
    # 3. Add Input Control (Sub-Layout Col 0)
    $InputControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $subLayout.Controls.Add($InputControl, 0, 0)
    
    # 4. Add Info Icon (Sub-Layout Col 1)
    if (-not [string]::IsNullOrEmpty($TooltipText)) {
        $icon = New-InfoIcon -TooltipText $TooltipText -HelpTextBox $HelpTextBox -HelpText $HelpText
        $subLayout.Controls.Add($icon, 1, 0)
    }
    
    # 5. Add extra control (Sub-Layout Col 2 - if it exists)
    if ($null -ne $ExtraControl) {
        $ExtraControl.Anchor = [System.Windows.Forms.AnchorStyles]::Right
        $subLayout.Controls.Add($ExtraControl, 2, 0)
    }
}

function Add-DualInputRow {
    param(
        [System.Windows.Forms.TableLayoutPanel]$Layout,
        [int]$Row,
        [string]$Label1Text,
        [System.Windows.Forms.Control]$Input1,
        [string]$Tooltip1,
        [string]$HelpText1,
        [string]$Label2Text,
        [System.Windows.Forms.Control]$Input2,
        [string]$Tooltip2,
        [string]$HelpText2,
        [System.Windows.Forms.TextBox]$HelpTextBox = $null
    )
    
    # Create sub-layout for dual inputs
    $subLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $subLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $subLayout.ColumnCount = 6
    $subLayout.RowCount = 1
    $subLayout.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
    $subLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

    # Column styles for sub-layout
    $subLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 133))) | Out-Null  # Label1
    $subLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 80))) | Out-Null  # Input1
    $subLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 20))) | Out-Null  # Icon1
    $subLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null  # Label2
    $subLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120))) | Out-Null  # Input2
    $subLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 33))) | Out-Null  # Icon2
    
    $lbl1 = New-Label -Text $Label1Text
    $lbl1.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
    $lbl1.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $subLayout.Controls.Add($lbl1, 0, 0)
    
    $Input1.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $Input1.Margin = New-Object System.Windows.Forms.Padding(0)
    $subLayout.Controls.Add($Input1, 1, 0)
    
    $icon1 = New-InfoIcon -TooltipText $Tooltip1 -HelpTextBox $HelpTextBox -HelpText $HelpText1
    $icon1.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $subLayout.Controls.Add($icon1, 2, 0)
    
    $lbl2 = New-Label -Text $Label2Text
    $lbl2.Margin = New-Object System.Windows.Forms.Padding(15, 0, 5, 0)
    $lbl2.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $subLayout.Controls.Add($lbl2, 3, 0)
    
    $Input2.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $Input2.Margin = New-Object System.Windows.Forms.Padding(0)
    $subLayout.Controls.Add($Input2, 4, 0)
    
    $icon2 = New-InfoIcon -TooltipText $Tooltip2 -HelpTextBox $HelpTextBox -HelpText $HelpText2
    $icon2.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $subLayout.Controls.Add($icon2, 5, 0)
    
    # Add sub-layout spanning remaining columns
    $Layout.Controls.Add($subLayout, 0, $Row)
    $Layout.SetColumnSpan($subLayout, 2)
}

#endregion

#region Main Form Builder

function Show-CreateAppFileForm {
    # Build layout structure
    $layoutInfo = Build-FormLayout
    $form = $layoutInfo.Form
    $mainLayout = $layoutInfo.MainLayout
    $contentLayout = $layoutInfo.ContentLayout
    
    # Create all input controls
    $txtGroup = New-TextBox
    $txtName = New-TextBox
    $txtVersion = New-TextBox
    $txtStartMenuFolder = New-TextBox
    $txtRegistryKey = New-TextBox
    $txtShortcutFolder = New-TextBox
    $txtAppPath = New-TextBox
    $txtDescription = New-TextBox -Multiline $true -Height 40
    
    # Help textbox (read-only, 3 lines height)
    $txtHelp = New-TextBox -ReadOnly $true -Multiline $true -Scrollbar $false -Height 120
    $txtHelp.BackColor = [System.Drawing.SystemColors]::Control
    $txtHelp.Dock = [System.Windows.Forms.DockStyle]::Fill
    
    # Variables to store paths
    $script:selectedShortcutFolder = $null
    $script:selectedAppPath = $null
    
    # Browse buttons
    $btnBrowseShortcuts = New-BrowseButton {
        $folder = Select-FolderDialog -Title "Select folder containing .lnk shortcuts"
        if ($folder) {
            $script:selectedShortcutFolder = $folder
            $txtShortcutFolder.Text = $folder
        }
    }
    
    $btnBrowseAppPath = New-BrowseButton {
        $folder = Select-FolderDialog -Title "Select Portable App base folder"
        if ($folder) {
            $script:selectedAppPath = $folder
            $txtAppPath.Text = $folder
        }
    }

    $nl = [System.Environment]::NewLine

    # Row 0: Group Name
    Add-FormRow -Layout $contentLayout -Row 0 -LabelText "Group Name:" -InputControl $txtGroup `
        -TooltipText "Organize apps into groups (e.g., 'Development Tools', 'Media Players'). Leave empty for ungrouped apps." `
        -HelpText "[Optional] You can assign a group name to categorize similar apps under the same group in the main app tree (e.g., Antiviruses, Text Editors, etc.).$($nl)These group names are solely for organizing similar apps in the main app view and will not create any folders in the Start Menu." `
        -HelpTextBox $txtHelp
    
    # Row 1: Name and Version (dual input)
    Add-DualInputRow -Layout $contentLayout -Row 1 `
        -Label1Text "Name*:" -Input1 $txtName -Tooltip1 "The display name of your portable application (Required)." -HelpText1 "(Required) Enter a name for your portable app.$($nl)This name will only be used to display it in the main app tree." `
        -Label2Text "Version:" -Input2 $txtVersion -Tooltip2 "Version number of the application (e.g., '1.0.0')." -HelpText2 "[Optional] Provide the app version.$($nl)This version will only be used in the app description and to compare it with the installed app, if it is already installed and you have provided its registry key." `
        -HelpTextBox $txtHelp
    
    # Row 2: StartMenu Folder
    Add-FormRow -Layout $contentLayout -Row 2 -LabelText "StartMenu Folder*:" -InputControl $txtStartMenuFolder `
        -TooltipText "The folder name that will appear in the Start Menu (Required)." `
        -HelpText "(Required) The name you specify here will be used to create a folder with the same name in the Start menu, where all shortcuts will be copied.$($nl)For portable apps, it is recommended to add '(Portable)' at the end of the name for convenience." `
        -HelpTextBox $txtHelp
  
    # Row 3: Registry Key
    Add-FormRow -Layout $contentLayout -Row 3 -LabelText "Registry Key:" -InputControl $txtRegistryKey `
        -TooltipText "Registry key to detect if app is installed (e.g., 'Google\Chrome'). Leave empty if not applicable." `
        -HelpText "[Optional] To find the name of the last key (e.g., Avira, 7-Zip), you can identify the last key path for any application installed on Windows by examining one of the following registry paths, depending on your platform (x86 or x64):$($nl)* HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall$($nl)* HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" `
        -HelpTextBox $txtHelp
    
    # Row 4: Shortcuts Folder
    Add-FormRow -Layout $contentLayout -Row 4 -LabelText "Shortcuts Folder*:" -InputControl $txtShortcutFolder `
        -TooltipText "Folder containing .lnk shortcut files that will be processed (Required)." `
        -HelpText "(Required) Provide the path where you manually created all shortcuts for your portable app.$($nl)This script will then automatically convert them into a JSON file with the `.app` file marker and place the resulting file in the root folder of the portable app." `
        -ExtraControl $btnBrowseShortcuts -HelpTextBox $txtHelp
    
    # Row 5: Portable App Folder
    Add-FormRow -Layout $contentLayout -Row 5 -LabelText "Portable App Folder*:" -InputControl $txtAppPath `
        -TooltipText "Base folder where the .app file will be saved and referenced (Required)." `
        -HelpText "(Required) Specify the path to the folder where your portable app is located.$($nl)Based on this path and the shortcuts you defined in the section above, the `.app` file marker will be created in this folder." `
        -ExtraControl $btnBrowseAppPath -HelpTextBox $txtHelp
    
    # Row 6: Description (expandable)
    Add-FormRow -Layout $contentLayout -Row 6 -LabelText "Description:" -InputControl $txtDescription `
        -TooltipText "Brief description of the application (1-2 lines recommended)." `
        -HelpText "[Optional] Provide a concise application overview in 1-2 lines, ideally within 80 characters long each line." `
        -HelpTextBox $txtHelp
    
    # Row 7: Help
    $lblHelp = New-Label -Text "Help:"
    $lblHelp.Margin = New-Object System.Windows.Forms.Padding(0, 5, 0, 0)
    $contentLayout.Controls.Add($lblHelp, 0, 7)
    $contentLayout.Controls.Add($txtHelp, 1, 7)
    
    # Bottom buttons panel
    $bottomPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $bottomPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $bottomPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $bottomPanel.WrapContents = $false
    $bottomPanel.Padding = New-Object System.Windows.Forms.Padding(0)
    
    $btnExit = New-ActionButton -Text "Exit" -Width 90
    $btnCreate = New-ActionButton -Text "Create .app File" -Width 130
    
    $bottomPanel.Controls.Add($btnExit)
    $bottomPanel.Controls.Add($btnCreate)
    
    # Add layouts to main
    $mainLayout.Controls.Add($contentLayout, 0, 0)
    $mainLayout.Controls.Add($bottomPanel, 0, 1)
    $form.Controls.Add($mainLayout)
    
    #region Event Handlers
    
    $btnExit.Add_Click({
            $form.Close()
        })
    
    $btnCreate.Add_Click({
            # Validation
            $errors = @()

            if ([string]::IsNullOrWhiteSpace($txtName.Text)) {
                $errors += "• App Name is required"
            }
        
            if ([string]::IsNullOrWhiteSpace($txtStartMenuFolder.Text)) {
                $errors += "• StartMenu Folder is required"
            }
        
            $script:selectedShortcutFolder = $txtShortcutFolder.Text
            if (-not $script:selectedShortcutFolder -or -not (Test-Path -LiteralPath $script:selectedShortcutFolder)) {
                $errors += "• Please select a valid Shortcuts Folder"
            }

            $script:selectedAppPath = $txtAppPath.Text 
            if (-not $script:selectedAppPath -or -not (Test-Path -LiteralPath $script:selectedAppPath)) {
                $errors += "• Please select a valid Portable App Folder"
            }
        
            if ($errors.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please fix the following errors:`n`n" + ($errors -join "`n"),
                    "Validation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }
        
            # Process shortcuts
            try {
                $lnkFiles = Get-ChildItem -Path $script:selectedShortcutFolder -Recurse -Filter "*.lnk" -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer }
            
                if ($lnkFiles.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "No .lnk files found in the selected Shortcuts Folder.",
                        "Warning",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    )
                    return
                }
            
                $shortcutEntries = @()
                foreach ($lnk in $lnkFiles) {
                    $props = Get-ShortcutProperties $lnk.FullName
                    if ($null -eq $props) {
                        Write-Warning "Cannot read shortcut: $($lnk.FullName)"
                        continue
                    }
                    $entry = [ordered]@{}
                    $entry["name"] = $lnk.BaseName
                    $entry["target"] = Convert-PathToToken $props.TargetPath $script:selectedAppPath
                    $entry["arguments"] = Convert-PathToToken $props.Arguments $script:selectedAppPath
                    $entry["workingDirectory"] = Convert-PathToToken $props.WorkingDirectory $script:selectedAppPath
                    $entry["icon"] = $props.IconLocation # Convert-PathToToken $props.IconLocation $script:selectedAppPath
                    $entry["windowStyle"] = $props.WindowStyle
                    $entry["description"] = Clean-Description $props.Description
                    $shortcutEntries += $entry
                }
            
                $desc = $txtDescription.Text.Trim() -replace "`r?`n", '\n'
                $desc = $desc -replace "\n\n", "\n"
                # Build app object
                $appObj = [ordered]@{
                    appName                = $txtName.Text.Trim()
                    appVersion             = $txtVersion.Text.Trim()
                    appGroup               = $txtGroup.Text.Trim()
                    appDescription         = $desc
                    appInstallRegistryData = $txtRegistryKey.Text.Trim()
                    appStartMenuFolderName = $txtStartMenuFolder.Text.Trim()
                    shortcuts              = $shortcutEntries
                }
            
                $jsonText = ConvertTo-JsonSimple $appObj 0
            
                # Save file
                $appFileName = ".app"
                $appFilePath = Join-Path -Path $script:selectedAppPath -ChildPath $appFileName
                Set-Content -LiteralPath $appFilePath -Value $jsonText -Encoding UTF8
            
                [System.Windows.Forms.MessageBox]::Show(
                    "Successfully created .app file at:`n`n$appFilePath",
                    "Success",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Error creating .app file:`n`n$($_.Exception.Message)",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        })
    
    #endregion
    
    # Show form
    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
}

#endregion

# Execute
Show-CreateAppFileForm