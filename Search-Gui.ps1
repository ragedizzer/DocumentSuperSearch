#GUI Wrapper SuperSearch Search Travis Webb V1.2026
#Travis Webb October 2026
#Document Search GUI wrapper for Search.ps1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$CoreScript = Join-Path -Path $PSScriptRoot -ChildPath "Search.ps1"
if (-not (Test-Path $CoreScript)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Missing Search.ps1 in $PSScriptRoot.",
        "Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}

. $CoreScript

$form = New-Object System.Windows.Forms.Form
$form.Text = "Document Search"
$form.ClientSize = New-Object System.Drawing.Size(680, 540)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Add_Shown({
    $form.TopMost = $true
    $form.Activate()
    $form.TopMost = $false
})

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 10000
$toolTip.InitialDelay = 400
$toolTip.ReshowDelay = 200
$toolTip.ShowAlways = $true

$documentsPath = [Environment]::GetFolderPath('MyDocuments')
if ([string]::IsNullOrWhiteSpace($documentsPath)) {
    $documentsPath = $PSScriptRoot
}

$settingsPath = Join-Path -Path $PSScriptRoot -ChildPath "Search-Gui-Settings.json"
$savedSettings = $null
if (Test-Path -LiteralPath $settingsPath) {
    try {
        $savedSettings = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
    } catch {
        $savedSettings = $null
    }
}

function Save-GuiSettings {
    param(
        [string]$SearchPath,
        [string]$OutputPath,
        [string]$OutputFormat
    )

    $settings = [PSCustomObject]@{
        SearchPath = $SearchPath
        OutputPath = $OutputPath
        OutputFormat = $OutputFormat
    }

    try {
        $settings | ConvertTo-Json -Depth 2 | Set-Content -LiteralPath $settingsPath -Encoding UTF8
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Unable to save default paths.",
            "Settings",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
}

$defaultSearchPath = if ($savedSettings -and $savedSettings.SearchPath) { $savedSettings.SearchPath } else { $documentsPath }
$defaultOutputPath = if ($savedSettings -and $savedSettings.OutputPath) { $savedSettings.OutputPath } else { $documentsPath }
$defaultOutputFormat = if ($savedSettings -and $savedSettings.OutputFormat) { $savedSettings.OutputFormat } else { "Excel" }

$pathTextX = 140
$pathTextY = 12
$pathTextWidth = 420
$iconButtonSize = New-Object System.Drawing.Size(26, 26)
$iconButtonSpacing = 6
$pathButtonX = [int]($pathTextX + $pathTextWidth + $iconButtonSpacing)

$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Search path"
$pathLabel.Location = New-Object System.Drawing.Point(12, 15)
$pathLabel.Size = New-Object System.Drawing.Size(120, 20)

$pathText = New-Object System.Windows.Forms.TextBox
$pathText.Location = New-Object System.Drawing.Point($pathTextX, $pathTextY)
$pathText.Size = New-Object System.Drawing.Size($pathTextWidth, 20)
$pathText.Text = $defaultSearchPath

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "..."
$browseButton.Location = New-Object System.Drawing.Point($pathButtonX, 10)
$browseButton.Size = $iconButtonSize
$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if (-not [string]::IsNullOrWhiteSpace($pathText.Text)) {
        $dialog.SelectedPath = $pathText.Text
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $pathText.Text = $dialog.SelectedPath
    }
})
$toolTip.SetToolTip($browseButton, "Browse for search folder")

$pathDefaultButton = New-Object System.Windows.Forms.Button
$pathDefaultButton.Text = "S"
$pathDefaultButton.Location = New-Object System.Drawing.Point(([int]($pathButtonX + $iconButtonSize.Width + $iconButtonSpacing)), 10)
$pathDefaultButton.Size = $iconButtonSize
$pathDefaultButton.Add_Click({
    $currentPath = $pathText.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($currentPath)) {
        Save-GuiSettings -SearchPath $currentPath -OutputPath $outputPathText.Text.Trim() -OutputFormat $outputFormatCombo.SelectedItem
    }
})
$toolTip.SetToolTip($pathDefaultButton, "Save search path as default")

$pathResetButton = New-Object System.Windows.Forms.Button
$pathResetButton.Text = "R"
$pathResetButton.Location = New-Object System.Drawing.Point(([int]($pathButtonX + (($iconButtonSize.Width + $iconButtonSpacing) * 2))), 10)
$pathResetButton.Size = $iconButtonSize
$pathResetButton.Add_Click({
    $pathText.Text = $documentsPath
    Save-GuiSettings -SearchPath $documentsPath -OutputPath $outputPathText.Text.Trim() -OutputFormat $outputFormatCombo.SelectedItem
})
$toolTip.SetToolTip($pathResetButton, "Reset search path to Documents")

$outputPathLabel = New-Object System.Windows.Forms.Label
$outputPathLabel.Text = "Output folder"
$outputPathLabel.Location = New-Object System.Drawing.Point(12, 45)
$outputPathLabel.Size = New-Object System.Drawing.Size(120, 20)

$outputPathText = New-Object System.Windows.Forms.TextBox
$outputPathText.Location = New-Object System.Drawing.Point($pathTextX, 42)
$outputPathText.Size = New-Object System.Drawing.Size($pathTextWidth, 20)
$outputPathText.Text = $defaultOutputPath

$outputBrowseButton = New-Object System.Windows.Forms.Button
$outputBrowseButton.Text = "..."
$outputBrowseButton.Location = New-Object System.Drawing.Point($pathButtonX, 40)
$outputBrowseButton.Size = $iconButtonSize
$outputBrowseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if (-not [string]::IsNullOrWhiteSpace($outputPathText.Text)) {
        $dialog.SelectedPath = $outputPathText.Text
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputPathText.Text = $dialog.SelectedPath
    }
})
$toolTip.SetToolTip($outputBrowseButton, "Browse for output folder")

$outputDefaultButton = New-Object System.Windows.Forms.Button
$outputDefaultButton.Text = "S"
$outputDefaultButton.Location = New-Object System.Drawing.Point(([int]($pathButtonX + $iconButtonSize.Width + $iconButtonSpacing)), 40)
$outputDefaultButton.Size = $iconButtonSize
$outputDefaultButton.Add_Click({
    $currentOutput = $outputPathText.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($currentOutput)) {
        Save-GuiSettings -SearchPath $pathText.Text.Trim() -OutputPath $currentOutput -OutputFormat $outputFormatCombo.SelectedItem
    }
})
$toolTip.SetToolTip($outputDefaultButton, "Save output path as default")

$outputResetButton = New-Object System.Windows.Forms.Button
$outputResetButton.Text = "R"
$outputResetButton.Location = New-Object System.Drawing.Point(([int]($pathButtonX + (($iconButtonSize.Width + $iconButtonSpacing) * 2))), 40)
$outputResetButton.Size = $iconButtonSize
$outputResetButton.Add_Click({
    $outputPathText.Text = $documentsPath
    Save-GuiSettings -SearchPath $pathText.Text.Trim() -OutputPath $documentsPath -OutputFormat $outputFormatCombo.SelectedItem
})
$toolTip.SetToolTip($outputResetButton, "Reset output path to Documents")

$outputFormatLabel = New-Object System.Windows.Forms.Label
$outputFormatLabel.Text = "Output format"
$outputFormatLabel.Location = New-Object System.Drawing.Point(12, 70)
$outputFormatLabel.Size = New-Object System.Drawing.Size(120, 20)

$outputFormatCombo = New-Object System.Windows.Forms.ComboBox
$outputFormatCombo.Location = New-Object System.Drawing.Point(140, 68)
$outputFormatCombo.Size = New-Object System.Drawing.Size(160, 20)
$outputFormatCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$null = $outputFormatCombo.Items.AddRange(@("Excel","ExcelTable","Csv"))
if ($outputFormatCombo.Items.Contains($defaultOutputFormat)) {
    $outputFormatCombo.SelectedItem = $defaultOutputFormat
} else {
    $outputFormatCombo.SelectedItem = "Excel"
}

$includeSubfoldersCheck = New-Object System.Windows.Forms.CheckBox
$includeSubfoldersCheck.Text = "Include subfolders"
$includeSubfoldersCheck.Location = New-Object System.Drawing.Point(140, 120)
$includeSubfoldersCheck.Size = New-Object System.Drawing.Size(160, 20)
$includeSubfoldersCheck.Checked = $true

$preventSleepCheck = New-Object System.Windows.Forms.CheckBox
$preventSleepCheck.Text = "Keep awake"
$preventSleepCheck.Location = New-Object System.Drawing.Point(320, 120)
$preventSleepCheck.Size = New-Object System.Drawing.Size(200, 20)
$preventSleepCheck.Checked = $true

$termsLabel = New-Object System.Windows.Forms.Label
$termsLabel.Text = "Search terms (comma or semicolon separated)"
$termsLabel.Location = New-Object System.Drawing.Point(12, 150)
$termsLabel.Size = New-Object System.Drawing.Size(260, 20)

$termsText = New-Object System.Windows.Forms.TextBox
$termsText.Location = New-Object System.Drawing.Point(280, 147)
$termsText.Size = New-Object System.Drawing.Size(380, 20)
$termsText.Text = "fnbm"

$matchCaseCheck = New-Object System.Windows.Forms.CheckBox
$matchCaseCheck.Text = "Match case"
$matchCaseCheck.Location = New-Object System.Drawing.Point(140, 185)
$matchCaseCheck.Size = New-Object System.Drawing.Size(140, 20)

$matchWholeCheck = New-Object System.Windows.Forms.CheckBox
$matchWholeCheck.Text = "Match whole word"
$matchWholeCheck.Location = New-Object System.Drawing.Point(320, 185)
$matchWholeCheck.Size = New-Object System.Drawing.Size(160, 20)
$matchWholeCheck.Checked = $true

$searchTextCheck = New-Object System.Windows.Forms.CheckBox
$searchTextCheck.Text = "Search content"
$searchTextCheck.Location = New-Object System.Drawing.Point(140, 210)
$searchTextCheck.Size = New-Object System.Drawing.Size(160, 20)
$searchTextCheck.Checked = $true

$searchLinksCheck = New-Object System.Windows.Forms.CheckBox
$searchLinksCheck.Text = "Search links"
$searchLinksCheck.Location = New-Object System.Drawing.Point(320, 210)
$searchLinksCheck.Size = New-Object System.Drawing.Size(160, 20)
$searchLinksCheck.Checked = $true

$linkModeLabel = New-Object System.Windows.Forms.Label
$linkModeLabel.Text = "Link mode"
$linkModeLabel.Location = New-Object System.Drawing.Point(140, 285)
$linkModeLabel.Size = New-Object System.Drawing.Size(70, 20)

$linkModeCombo = New-Object System.Windows.Forms.ComboBox
$linkModeCombo.Location = New-Object System.Drawing.Point(215, 283)
$linkModeCombo.Size = New-Object System.Drawing.Size(115, 20)
$linkModeCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$null = $linkModeCombo.Items.AddRange(@("AddressOnly","AddressAndSub","All"))
$linkModeCombo.SelectedItem = "All"

$docTimeoutLabel = New-Object System.Windows.Forms.Label
$docTimeoutLabel.Text = "Doc timeout (s)"
$docTimeoutLabel.Location = New-Object System.Drawing.Point(340, 285)
$docTimeoutLabel.Size = New-Object System.Drawing.Size(95, 20)

$docTimeoutInput = New-Object System.Windows.Forms.NumericUpDown
$docTimeoutInput.Location = New-Object System.Drawing.Point(445, 283)
$docTimeoutInput.Size = New-Object System.Drawing.Size(60, 20)
$docTimeoutInput.Minimum = 0
$docTimeoutInput.Maximum = 3600
$docTimeoutInput.Value = 120
$toolTip.SetToolTip($docTimeoutInput, "0 disables timeout")

$searchFileNameCheck = New-Object System.Windows.Forms.CheckBox
$searchFileNameCheck.Text = "Search file names"
$searchFileNameCheck.Location = New-Object System.Drawing.Point(140, 235)
$searchFileNameCheck.Size = New-Object System.Drawing.Size(160, 20)
$searchFileNameCheck.Checked = $false

$searchMetadataCheck = New-Object System.Windows.Forms.CheckBox
$searchMetadataCheck.Text = "Search metadata"
$searchMetadataCheck.Location = New-Object System.Drawing.Point(320, 235)
$searchMetadataCheck.Size = New-Object System.Drawing.Size(180, 20)
$searchMetadataCheck.Checked = $true

$includeMetadataCheck = New-Object System.Windows.Forms.CheckBox
$includeMetadataCheck.Text = "Include metadata columns"
$includeMetadataCheck.Location = New-Object System.Drawing.Point(320, 260)
$includeMetadataCheck.Size = New-Object System.Drawing.Size(200, 20)
$includeMetadataCheck.Checked = $true

$sendEmailCheck = New-Object System.Windows.Forms.CheckBox
$sendEmailCheck.Text = "Email results"
$sendEmailCheck.Location = New-Object System.Drawing.Point(140, 260)
$sendEmailCheck.Size = New-Object System.Drawing.Size(140, 20)

$emailToLabel = New-Object System.Windows.Forms.Label
$emailToLabel.Text = "Email to"
$emailToLabel.Location = New-Object System.Drawing.Point(12, 315)
$emailToLabel.Size = New-Object System.Drawing.Size(120, 20)

$emailToText = New-Object System.Windows.Forms.TextBox
$emailToText.Location = New-Object System.Drawing.Point(140, 312)
$emailToText.Size = New-Object System.Drawing.Size(520, 20)

$emailFromLabel = New-Object System.Windows.Forms.Label
$emailFromLabel.Text = "Send on behalf of (optional)"
$emailFromLabel.Location = New-Object System.Drawing.Point(12, 345)
$emailFromLabel.Size = New-Object System.Drawing.Size(180, 20)

$emailFromText = New-Object System.Windows.Forms.TextBox
$emailFromText.Location = New-Object System.Drawing.Point(200, 342)
$emailFromText.Size = New-Object System.Drawing.Size(460, 20)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(12, 385)
$statusLabel.Size = New-Object System.Drawing.Size(648, 40)
$statusLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "Search"
$searchButton.Location = New-Object System.Drawing.Point(140, 465)
$searchButton.Size = New-Object System.Drawing.Size(100, 30)

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = "Stop"
$stopButton.Location = New-Object System.Drawing.Point(250, 465)
$stopButton.Size = New-Object System.Drawing.Size(100, 30)
$stopButton.Enabled = $false
$stopButton.Add_Click({
    $script:StopRequested = $true
    $stopButton.Enabled = $false
    $statusLabel.Text = "Stopping..."
})

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Location = New-Object System.Drawing.Point(360, 465)
$closeButton.Size = New-Object System.Drawing.Size(100, 30)
$closeButton.Add_Click({ $form.Close() })

$toggleEmailFields = {
    $enabled = $sendEmailCheck.Checked
    $emailToLabel.Enabled = $enabled
    $emailToText.Enabled = $enabled
    $emailFromLabel.Enabled = $enabled
    $emailFromText.Enabled = $enabled
}
$sendEmailCheck.Add_CheckedChanged($toggleEmailFields)
& $toggleEmailFields

$toggleLinkMode = {
    $enabled = $searchLinksCheck.Checked
    $linkModeLabel.Visible = $enabled
    $linkModeCombo.Visible = $enabled
}
$searchLinksCheck.Add_CheckedChanged($toggleLinkMode)
& $toggleLinkMode

$searchButton.Add_Click({
    $statusLabel.Text = ""

    $pathValue = $pathText.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($pathValue)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a search path.","Validation",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (-not (Test-Path -Path $pathValue)) {
        [System.Windows.Forms.MessageBox]::Show("Search path not found.","Validation",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $outputPathValue = $outputPathText.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outputPathValue)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter an output folder.","Validation",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (-not (Test-Path -Path $outputPathValue)) {
        $createResult = [System.Windows.Forms.MessageBox]::Show(
            "Output folder does not exist. Create it?",
            "Create folder",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($createResult -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
        try {
            New-Item -ItemType Directory -Path $outputPathValue -Force | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Unable to create output folder.","Validation",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
    }

    $findTerms = @($termsText.Text -split "[,;]" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | Select-Object -Unique)
    if (-not $findTerms -or $findTerms.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please enter at least one search term.","Validation",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    if (-not $searchTextCheck.Checked -and -not $searchLinksCheck.Checked -and -not $searchMetadataCheck.Checked -and -not $searchFileNameCheck.Checked) {
        [System.Windows.Forms.MessageBox]::Show("Enable text content, link paths, metadata, or file name searching.","Validation",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $emailTo = @()
    $emailFrom = ""
    if ($sendEmailCheck.Checked) {
        $emailTo = @($emailToText.Text -split "[,;]" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | Select-Object -Unique)
        if (-not $emailTo -or $emailTo.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please enter at least one email address.","Validation",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        $emailFrom = $emailFromText.Text.Trim()
    }

    $searchButton.Enabled = $false
    $stopButton.Enabled = $true
    $closeButton.Enabled = $false
    $form.UseWaitCursor = $true
    $statusLabel.Text = "Searching..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $script:StopRequested = $false
        Save-GuiSettings -SearchPath $pathValue -OutputPath $outputPathValue -OutputFormat $outputFormatCombo.SelectedItem
        $emailSubject = Get-SearchResultsSubject -SearchTerms $findTerms
        $outputFormat = $outputFormatCombo.SelectedItem
        $outputExtension = if ($outputFormat -eq "Csv") { ".csv" } else { ".xlsx" }
        $outputFileName = Get-SearchResultsFileName -Subject $emailSubject -Extension $outputExtension
        $outputPath = Join-Path -Path $outputPathValue -ChildPath $outputFileName

        $shouldStop = { [System.Windows.Forms.Application]::DoEvents(); return $script:StopRequested }

        $searchParams = @{
            Path = $pathValue
            IncludeSubfolders = $includeSubfoldersCheck.Checked
            PreventSleep = $preventSleepCheck.Checked
            FindTerms = $findTerms
            MatchCase = $matchCaseCheck.Checked
            MatchWholeWord = $matchWholeCheck.Checked
            SearchTextContent = $searchTextCheck.Checked
            SearchLinkPaths = $searchLinksCheck.Checked
            LinkSearchMode = $linkModeCombo.SelectedItem
            SearchFileName = $searchFileNameCheck.Checked
            SearchMetadata = $searchMetadataCheck.Checked
            IncludeMetadataColumns = $includeMetadataCheck.Checked
            SendEmailResults = $sendEmailCheck.Checked
            EmailTo = $emailTo
            EmailFrom = $emailFrom
            OutputDirectory = $outputPathValue
            DocumentTimeoutSeconds = [int]$docTimeoutInput.Value
            ShouldStop = $shouldStop
            OutputFormat = $outputFormat
        }

        $results = Invoke-DocumentSearch @searchParams
        $matchCount = 0
        if ($results -ne $null) {
            $matchCount = $results.Count
        }
        $statusLabel.Text = "Complete. Matches: $matchCount"
        [System.Windows.Forms.MessageBox]::Show(
            "Search complete.`nMatches: $matchCount`nOutput: $outputPath",
            "Search complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } catch {
        $statusLabel.Text = "Search failed."
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $form.UseWaitCursor = $false
        $searchButton.Enabled = $true
        $stopButton.Enabled = $false
        $closeButton.Enabled = $true
        $script:StopRequested = $false
    }
})

$form.Controls.AddRange(@(
    $pathLabel,
    $pathText,
    $browseButton,
    $pathDefaultButton,
    $pathResetButton,
    $outputPathLabel,
    $outputPathText,
    $outputBrowseButton,
    $outputDefaultButton,
    $outputResetButton,
    $outputFormatLabel,
    $outputFormatCombo,
    $includeSubfoldersCheck,
    $preventSleepCheck,
    $termsLabel,
    $termsText,
    $matchCaseCheck,
    $matchWholeCheck,
    $searchTextCheck,
    $searchLinksCheck,
    $linkModeLabel,
    $linkModeCombo,
    $docTimeoutLabel,
    $docTimeoutInput,
    $searchFileNameCheck,
    $searchMetadataCheck,
    $includeMetadataCheck,
    $sendEmailCheck,
    $emailToLabel,
    $emailToText,
    $emailFromLabel,
    $emailFromText,
    $statusLabel,
    $searchButton,
    $stopButton,
    $closeButton
))

[void]$form.ShowDialog()

