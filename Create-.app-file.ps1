# Create-AppFile.ps1
# Requires: WScript.Shell COM availability

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Helper Functions

function Show-TopMostDialog($dialog) {
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.StartPosition = "Manual"
    $form.Location = New-Object System.Drawing.Point(-32000, -32000)
    $form.Size = New-Object System.Drawing.Size(1, 1)
    $form.Show()
    $result = $dialog.ShowDialog($form)
    $form.Dispose()
    return $result
}

function Select-FolderDialog {
    param([string]$Title = "Select Folder")
    
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Title
    $folderBrowser.ShowNewFolderButton = $true
    
    if ((Show-TopMostDialog $folderBrowser) -eq 'OK') {
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
        [int]$Height = 23
    )
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Anchor = $Anchor
    $textBox.ReadOnly = $ReadOnly
    $textBox.Multiline = $Multiline
    if ($Multiline) {
        $textBox.Height = $Height
        $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $textBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom
    }
    $textBox.Margin = New-Object System.Windows.Forms.Padding(3)
    return $textBox
}

function New-InfoIcon {
    param([string]$TooltipText)
    
    $icon = New-Object System.Windows.Forms.Label
    $icon.Text = "ℹ"
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
    $form.Text = "Create .app File"
    $form.Size = New-Object System.Drawing.Size(700, 560)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(650, 500)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.MaximizeBox = $true
    
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
    $contentLayout.RowCount = 10
    $contentLayout.AutoSize = $false
    
    # Column styles: Label | TextBox | Icon | Button
    $contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130))) | Out-Null
    $contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    # $contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 30))) | Out-Null
    # $contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90))) | Out-Null
    
    # Row styles - using fixed heights for consistent alignment
    for ($i = 0; $i -lt 6; $i++) {
        $contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35))) | Out-Null
    }
    # Description row - expandable
    $contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    
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
        [System.Windows.Forms.Control]$ExtraControl = $null
        # NOTE: $ColumnSpan is no longer needed/used here
    )
    
    # 1. Add label (Main Layout Col 0)
    $label = New-Label -Text $LabelText
    $label.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
    $Layout.Controls.Add($label, 0, $Row)
    
    # 2. Create the inner layout panel for all controls (Main Layout Col 1)
    $subLayout = Build-InputRowLayout -HasExtraControl ($null -ne $ExtraControl)
    $Layout.Controls.Add($subLayout, 1, $Row)
    
    # 3. Add Input Control (Sub-Layout Col 0)
    # This input control now correctly stretches across the available percentage width.
    $InputControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $subLayout.Controls.Add($InputControl, 0, 0)
    
    # 4. Add Info Icon (Sub-Layout Col 1)
    if (-not [string]::IsNullOrEmpty($TooltipText)) {
        $icon = New-InfoIcon -TooltipText $TooltipText
        # Icon is placed in Column 1 of the subLayout (fixed 30px width)
        $subLayout.Controls.Add($icon, 1, 0)
    }
    
    # 5. Add extra control (Sub-Layout Col 2 - if it exists)
    if ($null -ne $ExtraControl) {
        # ExtraControl is placed in Column 2 of the subLayout (fixed 90px width)
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
        [string]$Label2Text,
        [System.Windows.Forms.Control]$Input2,
        [string]$Tooltip2
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
    
    $icon1 = New-InfoIcon -TooltipText $Tooltip1
    $icon1.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $subLayout.Controls.Add($icon1, 2, 0)
    
    $lbl2 = New-Label -Text $Label2Text
    $lbl2.Margin = New-Object System.Windows.Forms.Padding(15, 0, 5, 0)
    $lbl2.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $subLayout.Controls.Add($lbl2, 3, 0)
    
    $Input2.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $Input2.Margin = New-Object System.Windows.Forms.Padding(0)
    $subLayout.Controls.Add($Input2, 4, 0)
    
    $icon2 = New-InfoIcon -TooltipText $Tooltip2
    $icon2.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $subLayout.Controls.Add($icon2, 5, 0)
    
    # Add sub-layout spanning remaining columns
    $Layout.Controls.Add($subLayout, 0, $Row)
    $Layout.SetColumnSpan($subLayout, 4)
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
    $txtDescription = New-TextBox -Multiline $true -Height 80
    
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
    
    # Row 0: Group Name
    Add-FormRow -Layout $contentLayout -Row 0 -LabelText "Group Name:" -InputControl $txtGroup `
        -TooltipText "Organize apps into groups (e.g., 'Development Tools', 'Media Players'). Leave empty for ungrouped apps." `
    
    # Row 1: Name and Version (dual input)
    Add-DualInputRow -Layout $contentLayout -Row 1 `
        -Label1Text "Name:" -Input1 $txtName -Tooltip1 "The display name of your portable application (Required)." `
        -Label2Text "Version:" -Input2 $txtVersion -Tooltip2 "Version number of the application (e.g., '1.0.0')."
    
    # Row 2: StartMenu Folder
    Add-FormRow -Layout $contentLayout -Row 2 -LabelText "StartMenu Folder:" -InputControl $txtStartMenuFolder `
        -TooltipText "The folder name that will appear in the Start Menu (Required)." `
  
    # Row 3: Registry Key
    Add-FormRow -Layout $contentLayout -Row 3 -LabelText "Registry Key:" -InputControl $txtRegistryKey `
        -TooltipText "Registry key to detect if app is installed (e.g., 'Google\Chrome'). Leave empty if not applicable." `
    
    # Row 4: Shortcuts Folder
    Add-FormRow -Layout $contentLayout -Row 4 -LabelText "Shortcuts Folder:" -InputControl $txtShortcutFolder `
        -TooltipText "Folder containing .lnk shortcut files that will be processed." `
        -ExtraControl $btnBrowseShortcuts -ColumnSpan 1
    
    # Row 5: Portable App Folder
    Add-FormRow -Layout $contentLayout -Row 5 -LabelText "Portable App Folder:" -InputControl $txtAppPath `
        -TooltipText "Base folder where the .app file will be saved and referenced." `
        -ExtraControl $btnBrowseAppPath -ColumnSpan 1
    
    # Row 6: Description
    Add-FormRow -Layout $contentLayout -Row 6 -LabelText "Description:" -InputControl $txtDescription `
        -TooltipText "Brief description of the application (1-2 lines recommended)." `
    
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
        
            if (-not $script:selectedShortcutFolder -or -not (Test-Path -LiteralPath $script:selectedShortcutFolder)) {
                $errors += "• Please select a valid Shortcuts Folder"
            }
        
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
                    $entry["icon"] = Convert-PathToToken $props.IconLocation $script:selectedAppPath
                    $entry["windowStyle"] = $props.WindowStyle
                    $entry["description"] = Clean-Description $props.Description
                    $shortcutEntries += $entry
                }
            
                # Build app object
                $appObj = [ordered]@{
                    appName                = $txtName.Text.Trim()
                    appVersion             = $txtVersion.Text.Trim()
                    appGroup               = $txtGroup.Text.Trim()
                    appDescription         = $txtDescription.Text.Trim()
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
            
                $form.Close()
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
    $form.TopMost = $true
    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
}

#endregion

# Execute
Show-CreateAppFileForm