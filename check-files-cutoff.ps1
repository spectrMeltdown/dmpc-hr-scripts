#Requires -Version 5.1

param(
    [string]$EnvPath = (Join-Path $PSScriptRoot ".env"),
    [switch]$NoPopup
)

$script:log_path = $null
$script:branch_payroll_start_day = 'Wednesday'
$script:branch_payroll_end_day = 'Tuesday'
$script:payroll_target_period = 'next'
$script:ref_run_date = $null
$script:branch_paths = @()
$script:use_branch_year_month = $false
$script:branch_year = $null
$script:branch_month = $null
$script:refresh_interval_seconds = 0
$script:cutoff_file_extensions = @('.xlsx', '.xls')
$script:recursive = $false
$script:dir_level_search = 2

function Import-DotEnv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Error "Env file not found: $Path. Copy .env.example to .env and edit values."
        exit 1
    }

    $envMap = @{}

    foreach ($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $line = $line.Trim()
        if (-not $line -or $line.StartsWith('#')) {
            continue
        }

        $eqIndex = $line.IndexOf('=')
        if ($eqIndex -lt 1) {
            continue
        }

        $key = $line.Substring(0, $eqIndex).Trim()
        $value = $line.Substring($eqIndex + 1).Trim()

        if ($value.Length -ge 2 -and $value.StartsWith('"') -and $value.EndsWith('"')) {
            $value = $value.Substring(1, $value.Length - 2)
        }
        elseif ($value.Length -ge 2 -and $value.StartsWith("'") -and $value.EndsWith("'")) {
            $value = $value.Substring(1, $value.Length - 2)
        }

        # Repeatable keys (e.g. BRANCH_PATHS on multiple lines) are joined with ';'
        if ($key -eq 'BRANCH_PATHS' -and $envMap.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace($envMap[$key])) {
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $envMap[$key] = $envMap[$key] + ';' + $value
            }
        }
        else {
            $envMap[$key] = $value
        }
    }

    return $envMap
}

function Resolve-ConfigPath {
    param(
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $null
    }

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }

    return Join-Path $PSScriptRoot $PathValue
}

function Test-EnvBool {
    param(
        [string]$Value,
        [bool]$Default = $false
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Default
    }

    return $Value.Trim().ToLowerInvariant() -in @('true', '1', 'yes')
}

function Assert-BranchYearMonthConfig {
    param(
        [bool]$UseBranchYearMonth,
        [string]$BranchYear,
        [string]$BranchMonth
    )

    if (-not $UseBranchYearMonth) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($BranchYear)) {
        throw 'BRANCH_YEAR is required when USE_BRANCH_YEAR_MONTH is true'
    }

    if ([string]::IsNullOrWhiteSpace($BranchMonth)) {
        throw 'BRANCH_MONTH is required when USE_BRANCH_YEAR_MONTH is true'
    }
}

function Resolve-BranchFolderPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [bool]$UseBranchYearMonth = $false,

        [string]$BranchYear,

        [string]$BranchMonth
    )

    if (-not $UseBranchYearMonth) {
        return $BasePath
    }

    Assert-BranchYearMonthConfig -UseBranchYearMonth $true -BranchYear $BranchYear -BranchMonth $BranchMonth
    return (Join-Path (Join-Path $BasePath $BranchYear.Trim()) $BranchMonth.Trim())
}

function ConvertTo-DayOfWeekEnum {
    param([string]$Name)

    if (-not $Name) {
        throw 'Day name is required.'
    }

    $n = $Name.Trim().ToLowerInvariant()
    $map = @{
        sunday    = 0
        monday    = 1
        tuesday   = 2
        wednesday = 3
        thursday  = 4
        friday    = 5
        saturday  = 6
    }

    if ($map.ContainsKey($n)) {
        return [DayOfWeek]$map[$n]
    }

    throw "Invalid day of week: $Name"
}

function Get-PayrollPeriodSpanDays {
    param(
        [DayOfWeek]$StartDow,
        [DayOfWeek]$EndDow
    )

    $diff = ([int]$EndDow - [int]$StartDow + 7) % 7
    if ($diff -eq 0) {
        return 8
    }

    return $diff + 1
}

function Get-BranchPayrollPeriod {
    param(
        [string]$StartDay,
        [string]$EndDay,
        [string]$TargetPeriod = 'next',
        [DateTime]$ReferenceDate = [DateTime]::Today
    )

    $startDow = ConvertTo-DayOfWeekEnum $StartDay
    $endDow = ConvertTo-DayOfWeekEnum $EndDay
    $spanDays = Get-PayrollPeriodSpanDays -StartDow $startDow -EndDow $endDow
    $ref = $ReferenceDate.Date
    $period = $TargetPeriod.Trim().ToLowerInvariant()

    if ($period -eq 'current') {
        $end = $ref
        while ($end.DayOfWeek -ne $endDow) {
            $end = $end.AddDays(-1)
        }
        $start = $end.AddDays(-($spanDays - 1))
    }
    elseif ($period -eq 'previous') {
        $end = $ref
        while ($end.DayOfWeek -ne $endDow) {
            $end = $end.AddDays(-1)
        }
        $end = $end.AddDays(-$spanDays)
        $start = $end.AddDays(-($spanDays - 1))
    }
    elseif ($period -eq 'next') {
        $start = $ref
        while ($start.DayOfWeek -ne $startDow) {
            $start = $start.AddDays(1)
        }
        if ($start -le $ref) {
            $start = $start.AddDays(7)
        }
        $end = $start.AddDays($spanDays - 1)
    }
    else {
        throw "PAYROLL_TARGET_PERIOD must be 'current', 'next', or 'previous', got: $TargetPeriod"
    }

    return @{
        Start = $start
        End   = $end
    }
}

function Parse-BranchPaths {
    param(
        [string]$Value
    )

    $list = [System.Collections.Generic.List[object]]::new()

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @($list.ToArray())
    }

    $segments = $Value -split ';'
    foreach ($segment in $segments) {
        $part = $segment.Trim()
        if (-not $part) {
            continue
        }

        $eqIndex = $part.IndexOf('=')
        if ($eqIndex -lt 1) {
            throw "Invalid BRANCH_PATHS segment (expected Label=Path): $part"
        }

        $label = $part.Substring(0, $eqIndex).Trim()
        $pathValue = $part.Substring($eqIndex + 1).Trim()

        if ([string]::IsNullOrWhiteSpace($label) -or [string]::IsNullOrWhiteSpace($pathValue)) {
            throw "Invalid BRANCH_PATHS segment (empty label or path): $part"
        }

        $resolved = Resolve-ConfigPath -PathValue $pathValue
        $list.Add([pscustomobject]@{
                Label = $label
                Path  = $resolved
            })
    }

    return @($list.ToArray())
}

function Get-BranchGroupLabel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    return ($Label -replace '-\d+$', '')
}

function Group-BranchPaths {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$BranchPaths
    )

    $ordered = [System.Collections.Generic.List[string]]::new()
    $groups = @{}

    foreach ($branch in $BranchPaths) {
        $groupLabel = Get-BranchGroupLabel -Label $branch.Label
        if (-not $groups.ContainsKey($groupLabel)) {
            $groups[$groupLabel] = [System.Collections.Generic.List[string]]::new()
            $ordered.Add($groupLabel)
        }
        $groups[$groupLabel].Add($branch.Path)
    }

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($groupLabel in $ordered) {
        $paths = @($groups[$groupLabel].ToArray())
        $list.Add([pscustomobject]@{
                Label = $groupLabel
                Paths = $paths
                Path  = $paths[0]
            })
    }

    return @($list.ToArray())
}

function Parse-CutoffFileExtensions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $extensions = [System.Collections.Generic.List[string]]::new()

    foreach ($segment in ($Value -split ';')) {
        $part = $segment.Trim()
        if (-not $part) {
            continue
        }

        if (-not $part.StartsWith('.')) {
            $part = '.' + $part
        }

        $extensions.Add($part.ToLowerInvariant())
    }

    if ($extensions.Count -eq 0) {
        throw 'CUTOFF_FILE_EXTENSIONS must contain at least one extension'
    }

    return @($extensions.ToArray())
}

function Get-DayRangeMatchPattern {
    param(
        [Parameter(Mandatory = $true)]
        [int]$StartDay,

        [Parameter(Mandatory = $true)]
        [int]$EndDay
    )

    return "(?<!\d)0?$StartDay\s{0,2}-\s{0,2}0?$EndDay(?!\d)"
}

function Test-FilenameContainsDayRange {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [Parameter(Mandatory = $true)]
        [int]$StartDay,

        [Parameter(Mandatory = $true)]
        [int]$EndDay
    )

    $pattern = Get-DayRangeMatchPattern -StartDay $StartDay -EndDay $EndDay
    return [bool]($FileName -match $pattern)
}

function Test-BranchHasCutoffFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath,

        [Parameter(Mandatory = $true)]
        [int]$StartDay,

        [Parameter(Mandatory = $true)]
        [int]$EndDay
    )

    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        return $false
    }

    $allowedExtensions = $script:cutoff_file_extensions
    $childParams = @{
        LiteralPath = $FolderPath
        File        = $true
        ErrorAction = 'Stop'
    }
    if ($script:recursive) {
        $childParams['Recurse'] = $true
        $childParams['Depth'] = $script:dir_level_search
    }

    $files = @(Get-ChildItem @childParams |
        Where-Object {
            -not $_.Name.StartsWith('~$') -and
            ($allowedExtensions -icontains $_.Extension)
        })

    foreach ($file in $files) {
        if (Test-FilenameContainsDayRange -FileName $file.Name -StartDay $StartDay -EndDay $EndDay) {
            return $true
        }
    }

    return $false
}

function Get-BranchesWithCutoffFiles {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$BranchPaths,

        [Parameter(Mandatory = $true)]
        [int]$StartDay,

        [Parameter(Mandatory = $true)]
        [int]$EndDay,

        [scriptblock]$OnError
    )

    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($branch in $BranchPaths) {
        $paths = if ($branch.PSObject.Properties['Paths'] -and $branch.Paths) {
            @($branch.Paths)
        }
        else {
            @($branch.Path)
        }

        $hasFile = $false
        $matchedPath = $null
        $pathResults = [System.Collections.Generic.List[object]]::new()

        foreach ($folderPath in $paths) {
            $pathHasFile = $false
            try {
                $pathHasFile = Test-BranchHasCutoffFile -FolderPath $folderPath -StartDay $StartDay -EndDay $EndDay
            }
            catch {
                $pathHasFile = $false
                if ($OnError) {
                    & $OnError ([pscustomobject]@{
                            Label = $branch.Label
                            Path  = $folderPath
                        }) $_.Exception.Message
                }
            }

            $pathResults.Add([pscustomobject]@{
                    Path    = $folderPath
                    HasFile = $pathHasFile
                })

            if ($pathHasFile -and -not $hasFile) {
                $hasFile = $true
                $matchedPath = $folderPath
            }
        }

        $results.Add([pscustomobject]@{
                Label       = $branch.Label
                Path        = if ($matchedPath) { $matchedPath } else { $paths[0] }
                Paths       = $paths
                PathResults = @($pathResults.ToArray())
                HasFile     = $hasFile
            })
    }

    return @($results.ToArray())
}

function Format-CutoffFilesChecklist {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results
    )

    $lines = foreach ($item in $Results) {
        if ($item.HasFile) {
            "[OK] $($item.Label)"
        }
        else {
            "[X] $($item.Label)"
        }
    }

    return ($lines -join [Environment]::NewLine)
}

function Test-CutoffFilesDisplayChanged {
    param(
        [hashtable]$Previous,
        [hashtable]$Current
    )

    if (-not $Previous) {
        return $false
    }

    return $Previous.Body -ne $Current.Body
}

function Invoke-CutoffFilesChangeAlert {
    try {
        Add-Type -AssemblyName System.Media -ErrorAction Stop
        [System.Media.SystemSounds]::Exclamation.Play()
    }
    catch {
        Write-Host "Could not play notification sound: $($_.Exception.Message)"
    }
}

function Open-CutoffFolderInExplorer {
    param(
        [string]$FolderPath
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            'No folder path is configured for this branch.',
            'Open folder',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $FolderPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Folder not found:`n$FolderPath",
            'Open folder',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    Start-Process -FilePath 'explorer.exe' -ArgumentList @($FolderPath)
}

function Show-CutoffFilesPopup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Results,

        [string]$Body,

        [int]$RefreshIntervalSeconds = 0,
        [scriptblock]$OnRefresh
    )

    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing | Out-Null

    if ([string]::IsNullOrWhiteSpace($Body)) {
        $Body = Format-CutoffFilesChecklist -Results $Results
    }

    $margin = 16
    $formWidth = 520
    $contentWidth = $formWidth - (2 * $margin)
    $buttonHeight = 28
    $buttonWidth = 86
    $buttonGap = 12
    $rowHeight = 30
    $rowGap = 2
    $openButtonWidth = 56
    $openButtonHeight = 26
    $openButtonGap = 8
    $minContentHeight = 80

    $font = New-Object System.Drawing.Font('Consolas', 11)
    $rowCount = @($Results).Count
    $desiredContentHeight = [Math]::Max(
        $minContentHeight,
        ($rowCount * ($rowHeight + $rowGap)) + 4
    )

    $workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $maxFormHeight = [Math]::Floor($workingArea.Height * 0.7)
    $maxContentHeight = [Math]::Max(
        $minContentHeight,
        $maxFormHeight - (2 * $margin) - $buttonHeight - $buttonGap
    )
    $needsScroll = $desiredContentHeight -gt $maxContentHeight
    $contentHeight = if ($needsScroll) { $maxContentHeight } else { $desiredContentHeight }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ShowInTaskbar = $true
    $form.ClientSize = New-Object System.Drawing.Size(
        $formWidth,
        ((2 * $margin) + $contentHeight + $buttonGap + $buttonHeight)
    )

    $listPanel = New-Object System.Windows.Forms.Panel
    $listPanel.Location = New-Object System.Drawing.Point($margin, $margin)
    $listPanel.Size = New-Object System.Drawing.Size($contentWidth, $contentHeight)
    $listPanel.AutoScroll = $true
    $listPanel.BorderStyle = 'None'

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $okButton.Location = New-Object System.Drawing.Point(
        ($formWidth - $margin - $buttonWidth),
        ($margin + $contentHeight + $buttonGap)
    )

    $rebuildRows = {
        param([object[]]$RowResults)

        $listPanel.SuspendLayout()
        $listPanel.Controls.Clear()

        $scrollPad = 0
        if ($needsScroll -or (@($RowResults).Count * ($rowHeight + $rowGap) -gt $contentHeight)) {
            $scrollPad = [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
        }
        $rowInnerWidth = [Math]::Max(120, $listPanel.ClientSize.Width - $scrollPad)
        $labelWidth = [Math]::Max(
            40,
            $rowInnerWidth - $openButtonWidth - $openButtonGap
        )
        $labelHeight = [Math]::Max($font.Height + 2, 18)

        $y = 0
        foreach ($item in @($RowResults)) {
            $status = if ($item.HasFile) { '[OK]' } else { '[X]' }
            $rowPanel = New-Object System.Windows.Forms.Panel
            $rowPanel.Location = New-Object System.Drawing.Point(0, $y)
            $rowPanel.Size = New-Object System.Drawing.Size($rowInnerWidth, $rowHeight)

            $statusLabel = New-Object System.Windows.Forms.Label
            $statusLabel.Text = ("$status $($item.Label)" -replace '[\r\n]+', ' ')
            $statusLabel.Font = $font
            $statusLabel.AutoSize = $false
            $statusLabel.AutoEllipsis = $true
            $statusLabel.Location = New-Object System.Drawing.Point(0, [Math]::Floor(($rowHeight - $labelHeight) / 2))
            $statusLabel.Size = New-Object System.Drawing.Size($labelWidth, $labelHeight)
            $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

            $openButton = New-Object System.Windows.Forms.Button
            $openButton.Text = 'Open'
            $openButton.Size = New-Object System.Drawing.Size($openButtonWidth, $openButtonHeight)
            $openButton.Location = New-Object System.Drawing.Point(
                ($rowInnerWidth - $openButtonWidth),
                [Math]::Floor(($rowHeight - $openButtonHeight) / 2)
            )
            $openButton.Tag = $item.Path
            $openButton.Add_Click({
                param($sender, $eventArgs)
                Open-CutoffFolderInExplorer -FolderPath ([string]$sender.Tag)
            })

            $rowPanel.Controls.Add($statusLabel)
            $rowPanel.Controls.Add($openButton)
            $listPanel.Controls.Add($rowPanel)
            $y += ($rowHeight + $rowGap)
        }

        $listPanel.ResumeLayout()
    }

    & $rebuildRows $Results

    $form.Controls.Add($listPanel)
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $timer = $null
    if ($RefreshIntervalSeconds -gt 0 -and $OnRefresh) {
        $refreshCallback = $OnRefresh
        $previousDisplay = @{ Body = $Body }
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = $RefreshIntervalSeconds * 1000
        $timer.Add_Tick({
            try {
                $display = & $refreshCallback
                if ($display) {
                    if (Test-CutoffFilesDisplayChanged -Previous $previousDisplay -Current $display) {
                        Invoke-CutoffFilesChangeAlert
                    }
                    $previousDisplay = $display
                    & $rebuildRows -RowResults ([object[]]@($display.Results))
                    $form.Text = if ($display.Title -match '\(updated ') {
                        $display.Title
                    }
                    else {
                        "$($display.Title) (updated $(Get-Date -Format 'HH:mm:ss'))"
                    }
                }
            }
            catch {
                Write-Host "Refresh failed: $($_.Exception.Message)"
            }
        })
        $timer.Start()
    }

    [void]$form.ShowDialog()
    if ($timer) {
        $timer.Stop()
        $timer.Dispose()
    }
    $form.Dispose()
    $font.Dispose()
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] $Level  $Message"
    Write-Host $line

    if ([string]::IsNullOrWhiteSpace($script:log_path)) {
        return
    }

    $logDir = Split-Path -Parent $script:log_path
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    Add-Content -LiteralPath $script:log_path -Value $line -Encoding UTF8
}

function Initialize-Config {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $envMap = Import-DotEnv -Path $Path

    if (-not $envMap.ContainsKey('BRANCH_PATHS') -or [string]::IsNullOrWhiteSpace($envMap['BRANCH_PATHS'])) {
        Write-Error 'Missing or empty required env key: BRANCH_PATHS'
        exit 1
    }

    if (-not $envMap.ContainsKey('LOG_PATH') -or [string]::IsNullOrWhiteSpace($envMap['LOG_PATH'])) {
        Write-Error 'Missing or empty required env key: LOG_PATH'
        exit 1
    }

    $script:log_path = Resolve-ConfigPath -PathValue $envMap['LOG_PATH']

    if ($envMap.ContainsKey('BRANCH_PAYROLL_START_DAY') -and -not [string]::IsNullOrWhiteSpace($envMap['BRANCH_PAYROLL_START_DAY'])) {
        $script:branch_payroll_start_day = $envMap['BRANCH_PAYROLL_START_DAY'].Trim()
    }

    if ($envMap.ContainsKey('BRANCH_PAYROLL_END_DAY') -and -not [string]::IsNullOrWhiteSpace($envMap['BRANCH_PAYROLL_END_DAY'])) {
        $script:branch_payroll_end_day = $envMap['BRANCH_PAYROLL_END_DAY'].Trim()
    }

    if ($envMap.ContainsKey('PAYROLL_TARGET_PERIOD') -and -not [string]::IsNullOrWhiteSpace($envMap['PAYROLL_TARGET_PERIOD'])) {
        $script:payroll_target_period = $envMap['PAYROLL_TARGET_PERIOD'].Trim()
    }

    $script:ref_run_date = $null
    if ($envMap.ContainsKey('REF_RUN_DATE') -and -not [string]::IsNullOrWhiteSpace($envMap['REF_RUN_DATE'])) {
        $script:ref_run_date = [DateTime]::ParseExact(
            $envMap['REF_RUN_DATE'].Trim(),
            'yyyy-MM-dd',
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    }

    $useYearMonthValue = if ($envMap.ContainsKey('USE_BRANCH_YEAR_MONTH')) { $envMap['USE_BRANCH_YEAR_MONTH'] } else { $null }
    $script:use_branch_year_month = Test-EnvBool -Value $useYearMonthValue -Default $false
    $script:branch_year = if ($envMap.ContainsKey('BRANCH_YEAR')) { $envMap['BRANCH_YEAR'] } else { $null }
    $script:branch_month = if ($envMap.ContainsKey('BRANCH_MONTH')) { $envMap['BRANCH_MONTH'] } else { $null }

    try {
        Assert-BranchYearMonthConfig `
            -UseBranchYearMonth $script:use_branch_year_month `
            -BranchYear $script:branch_year `
            -BranchMonth $script:branch_month
    }
    catch {
        Write-Error $_.Exception.Message
        exit 1
    }

    try {
        $script:branch_paths = @(Parse-BranchPaths -Value $envMap['BRANCH_PATHS'])
    }
    catch {
        Write-Error $_.Exception.Message
        exit 1
    }

    if ($script:branch_paths.Count -eq 0) {
        Write-Error 'BRANCH_PATHS must contain at least one Label=Path entry'
        exit 1
    }

    if ($script:use_branch_year_month) {
        $script:branch_paths = @(
            foreach ($branch in $script:branch_paths) {
                [pscustomobject]@{
                    Label = $branch.Label
                    Path  = Resolve-BranchFolderPath `
                        -BasePath $branch.Path `
                        -UseBranchYearMonth $true `
                        -BranchYear $script:branch_year `
                        -BranchMonth $script:branch_month
                }
            }
        )
    }

    $script:branch_paths = @(Group-BranchPaths -BranchPaths $script:branch_paths)

    $script:refresh_interval_seconds = 0
    if ($envMap.ContainsKey('REFRESH_INTERVAL_SECONDS') -and -not [string]::IsNullOrWhiteSpace($envMap['REFRESH_INTERVAL_SECONDS'])) {
        $parsedRefresh = 0
        if (-not [int]::TryParse($envMap['REFRESH_INTERVAL_SECONDS'].Trim(), [ref]$parsedRefresh) -or $parsedRefresh -lt 0) {
            Write-Error 'REFRESH_INTERVAL_SECONDS must be a non-negative integer'
            exit 1
        }
        $script:refresh_interval_seconds = $parsedRefresh
    }

    if ($envMap.ContainsKey('CUTOFF_FILE_EXTENSIONS') -and -not [string]::IsNullOrWhiteSpace($envMap['CUTOFF_FILE_EXTENSIONS'])) {
        try {
            $script:cutoff_file_extensions = @(Parse-CutoffFileExtensions -Value $envMap['CUTOFF_FILE_EXTENSIONS'])
        }
        catch {
            Write-Error $_.Exception.Message
            exit 1
        }
    }

    $recursiveValue = if ($envMap.ContainsKey('RECURSIVE')) { $envMap['RECURSIVE'] } else { $null }
    $script:recursive = Test-EnvBool -Value $recursiveValue -Default $false

    $script:dir_level_search = 2
    if ($envMap.ContainsKey('DIR_LEVEL_SEARCH') -and -not [string]::IsNullOrWhiteSpace($envMap['DIR_LEVEL_SEARCH'])) {
        $parsedDepth = 0
        if (-not [int]::TryParse($envMap['DIR_LEVEL_SEARCH'].Trim(), [ref]$parsedDepth) -or $parsedDepth -lt 0) {
            Write-Error 'DIR_LEVEL_SEARCH must be a non-negative integer'
            exit 1
        }
        $script:dir_level_search = $parsedDepth
    }
}

function Get-CutoffFilesCheckDisplay {
    $referenceDate = if ($script:ref_run_date) { $script:ref_run_date } else { [DateTime]::Today }

    $period = Get-BranchPayrollPeriod `
        -StartDay $script:branch_payroll_start_day `
        -EndDay $script:branch_payroll_end_day `
        -TargetPeriod $script:payroll_target_period `
        -ReferenceDate $referenceDate

    $startDay = $period.Start.Day
    $endDay = $period.End.Day
    $dayRangeLabel = '{0}-{1}' -f $startDay, $endDay

    $branches_with_cutoff_files = Get-BranchesWithCutoffFiles `
        -BranchPaths $script:branch_paths `
        -StartDay $startDay `
        -EndDay $endDay `
        -OnError {
            param($branch, $message)
            Write-Log "Cannot scan '$($branch.Label)' ($($branch.Path)): $message" -Level ERROR
        }

    $checklist = Format-CutoffFilesChecklist -Results $branches_with_cutoff_files
    $title = 'Cutoff files check - {0} ({1:yyyy-MM-dd} to {2:yyyy-MM-dd})' -f `
        $dayRangeLabel, $period.Start, $period.End

    return @{
        Title         = $title
        Body          = $checklist
        Results       = $branches_with_cutoff_files
        Period        = $period
        DayRangeLabel = $dayRangeLabel
    }
}

function Write-CutoffFilesCheckLog {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Display
    )

    Write-Log ("Payroll period ({0}): {1:yyyy-MM-dd} to {2:yyyy-MM-dd} (day range {3})" -f `
            $script:payroll_target_period, $Display.Period.Start, $Display.Period.End, $Display.DayRangeLabel)
    Write-Log "Cutoff files checklist:`n$($Display.Body)"

    foreach ($item in $Display.Results) {
        if ($item.PSObject.Properties['PathResults'] -and $item.PathResults) {
            foreach ($pathResult in $item.PathResults) {
                $status = if ($pathResult.HasFile) { 'found' } else { 'missing' }
                Write-Log ("{0}: {1} ({2})" -f $item.Label, $status, $pathResult.Path)
            }
        }
        else {
            $status = if ($item.HasFile) { 'found' } else { 'missing' }
            Write-Log ("{0}: {1} ({2})" -f $item.Label, $status, $item.Path)
        }
    }
}

function Invoke-CheckFilesCutoff {
    param(
        [switch]$SkipPopup
    )

    $display = Get-CutoffFilesCheckDisplay
    Write-CutoffFilesCheckLog -Display $display

    if (-not $SkipPopup) {
        Show-CutoffFilesPopup -Title $display.Title -Results $display.Results -Body $display.Body
    }

    return $display.Results
}

# True only when dot-sourced (`. .\script.ps1`). Do NOT match `.\script.ps1` via Line.
$isDotSourced = $MyInvocation.InvocationName -eq '.'
if (-not $isDotSourced) {
    Initialize-Config -Path $EnvPath

    if ($script:refresh_interval_seconds -le 0) {
        [void](Invoke-CheckFilesCutoff -SkipPopup:$NoPopup)
    }
    elseif ($NoPopup) {
        while ($true) {
            Initialize-Config -Path $EnvPath
            $display = Get-CutoffFilesCheckDisplay
            Write-CutoffFilesCheckLog -Display $display
            Start-Sleep -Seconds $script:refresh_interval_seconds
        }
    }
    else {
        $display = Get-CutoffFilesCheckDisplay
        Write-CutoffFilesCheckLog -Display $display

        $envPathForRefresh = $EnvPath
        Show-CutoffFilesPopup -Title $display.Title -Results $display.Results -Body $display.Body `
            -RefreshIntervalSeconds $script:refresh_interval_seconds `
            -OnRefresh {
                Initialize-Config -Path $envPathForRefresh
                $refreshed = Get-CutoffFilesCheckDisplay
                Write-CutoffFilesCheckLog -Display $refreshed
                return $refreshed
            }
    }
}
