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
        throw "PAYROLL_TARGET_PERIOD must be 'current' or 'next', got: $TargetPeriod"
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

function Test-BranchHasIncentiveFile {
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

    $files = @(Get-ChildItem -LiteralPath $FolderPath -File -ErrorAction Stop |
        Where-Object {
            -not $_.Name.StartsWith('~$') -and
            ($_.Extension -ieq '.xlsx' -or $_.Extension -ieq '.xls')
        })

    foreach ($file in $files) {
        if (Test-FilenameContainsDayRange -FileName $file.Name -StartDay $StartDay -EndDay $EndDay) {
            return $true
        }
    }

    return $false
}

function Get-BranchesWithIncentives {
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
        $hasFile = $false
        try {
            $hasFile = Test-BranchHasIncentiveFile -FolderPath $branch.Path -StartDay $StartDay -EndDay $EndDay
        }
        catch {
            $hasFile = $false
            if ($OnError) {
                & $OnError $branch $_.Exception.Message
            }
        }

        $results.Add([pscustomobject]@{
                Label   = $branch.Label
                Path    = $branch.Path
                HasFile = $hasFile
            })
    }

    return @($results.ToArray())
}

function Format-IncentivesChecklist {
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

function Show-IncentivesPopup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Body
    )

    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing | Out-Null

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ShowInTaskbar = $true
    $form.ClientSize = New-Object System.Drawing.Size(420, 320)

    $label = New-Object System.Windows.Forms.Label
    $label.AutoSize = $false
    $label.Location = New-Object System.Drawing.Point(16, 16)
    $label.Size = New-Object System.Drawing.Size(388, 240)
    $label.Font = New-Object System.Drawing.Font('Consolas', 11)
    $label.Text = $Body

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(318, 270)
    $okButton.Size = New-Object System.Drawing.Size(86, 28)

    $form.Controls.Add($label)
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    [void]$form.ShowDialog()
    $form.Dispose()
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
}

function Invoke-CheckIncentives {
    param(
        [switch]$SkipPopup
    )

    $referenceDate = if ($script:ref_run_date) { $script:ref_run_date } else { [DateTime]::Today }

    $period = Get-BranchPayrollPeriod `
        -StartDay $script:branch_payroll_start_day `
        -EndDay $script:branch_payroll_end_day `
        -TargetPeriod $script:payroll_target_period `
        -ReferenceDate $referenceDate

    $startDay = $period.Start.Day
    $endDay = $period.End.Day
    $dayRangeLabel = '{0}-{1}' -f $startDay, $endDay

    Write-Log ("Payroll period ({0}): {1:yyyy-MM-dd} to {2:yyyy-MM-dd} (day range {3})" -f `
            $script:payroll_target_period, $period.Start, $period.End, $dayRangeLabel)

    $branches_with_incentives = Get-BranchesWithIncentives `
        -BranchPaths $script:branch_paths `
        -StartDay $startDay `
        -EndDay $endDay `
        -OnError {
            param($branch, $message)
            Write-Log "Cannot scan '$($branch.Label)' ($($branch.Path)): $message" -Level ERROR
        }

    $checklist = Format-IncentivesChecklist -Results $branches_with_incentives
    Write-Log "Incentives checklist:`n$checklist"

    foreach ($item in $branches_with_incentives) {
        $status = if ($item.HasFile) { 'found' } else { 'missing' }
        Write-Log ("{0}: {1} ({2})" -f $item.Label, $status, $item.Path)
    }

    if (-not $SkipPopup) {
        $title = 'Incentives check - {0} ({1:yyyy-MM-dd} to {2:yyyy-MM-dd})' -f `
            $dayRangeLabel, $period.Start, $period.End
        Show-IncentivesPopup -Title $title -Body $checklist
    }

    return $branches_with_incentives
}

# True only when dot-sourced (`. .\script.ps1`). Do NOT match `.\script.ps1` via Line.
$isDotSourced = $MyInvocation.InvocationName -eq '.'
if (-not $isDotSourced) {
    Initialize-Config -Path $EnvPath
    [void](Invoke-CheckIncentives -SkipPopup:$NoPopup)
}
