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
$script:date_subpath = $null
$script:use_branch_year_month = $false
$script:branch_year = $null
$script:branch_month = $null
$script:refresh_interval_seconds = 0
$script:cutoff_file_extensions = @('.xlsx', '.xls')
$script:recursive = $false
$script:dir_level_search = 2
$script:log_level = 'INFO'
$script:sr_paths = @()
$script:sr_file_extensions = @('.pdf', '.png', '.jpg', '.jpeg', '.xlsx', '.xls')
$script:sr_open_max = 10
$script:show_sr_button = $true
$script:final_sc_incentives_path = $null
$script:cutoffPopupForm = $null

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

        # Repeatable keys (e.g. BRANCH_PATHS / SR_PATHS on multiple lines) are joined with ';'
        if ($key -in @('BRANCH_PATHS', 'SR_PATHS') -and $envMap.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace($envMap[$key])) {
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

function Get-LogLevelRank {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Level
    )

    switch ($Level.ToUpperInvariant()) {
        'TRACE' { return 0 }
        'DEBUG' { return 1 }
        'INFO' { return 2 }
        'WARNING' { return 3 }
        'ERROR' { return 4 }
        default { return -1 }
    }
}

function Resolve-LogLevel {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return 'INFO'
    }

    $normalized = $Value.Trim().ToUpperInvariant()
    if ((Get-LogLevelRank -Level $normalized) -lt 0) {
        Write-Error "LOG_LEVEL must be one of: TRACE, DEBUG, INFO, WARNING, ERROR. Got: $Value"
        exit 1
    }

    return $normalized
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

function Expand-DateSubpathPattern {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Pattern,

        [Parameter(Mandatory = $true)]
        [DateTime]$Date
    )

    if ([string]::IsNullOrEmpty($Pattern)) {
        return $Pattern
    }

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    # Longest tokens first so MMMM wins over MMM/MM/M.
    $tokens = @(
        @{ Name = 'yyyy'; Value = $Date.ToString('yyyy', $culture) }
        @{ Name = 'MMMM'; Value = $Date.ToString('MMMM', $culture) }
        @{ Name = 'MMM'; Value = $Date.ToString('MMM', $culture) }
        @{ Name = 'MM'; Value = $Date.ToString('MM', $culture) }
        @{ Name = 'M'; Value = $Date.ToString('%M', $culture) }
    )

    $sb = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $Pattern.Length) {
        $matched = $false
        foreach ($token in $tokens) {
            $name = $token.Name
            $len = $name.Length
            if (($i + $len) -le $Pattern.Length -and $Pattern.Substring($i, $len) -eq $name) {
                [void]$sb.Append($token.Value)
                $i += $len
                $matched = $true
                break
            }
        }
        if (-not $matched) {
            [void]$sb.Append($Pattern[$i])
            $i++
        }
    }

    return $sb.ToString()
}

function Resolve-DateSubpath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Pattern,

        [Parameter(Mandatory = $true)]
        [DateTime]$Date
    )

    if ([string]::IsNullOrWhiteSpace($Pattern)) {
        return $BasePath
    }

    $expanded = Expand-DateSubpathPattern -Pattern $Pattern.Trim() -Date $Date
    $segments = @(
        $expanded -split '[\\/]+' |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    $current = $BasePath
    foreach ($segment in $segments) {
        if ($segment -match '[*?]') {
            if (-not (Test-Path -LiteralPath $current -PathType Container)) {
                throw ("Cannot resolve DATE_SUBPATH segment '{0}': parent folder missing: {1}" -f $segment, $current)
            }

            # "*Aug" should match "8. August"; standard wildcards are whole-string, so
            # a leading-* segment with no trailing wildcard gets an implicit trailing "*".
            $matchPattern = $segment
            if ($matchPattern.StartsWith('*') -and $matchPattern -notmatch '[*?]$') {
                $matchPattern = $matchPattern + '*'
            }

            $wildcard = [System.Management.Automation.WildcardPattern]::new(
                $matchPattern,
                [System.Management.Automation.WildcardOptions]::IgnoreCase
            )
            $matched = @(
                Get-ChildItem -LiteralPath $current -Directory -ErrorAction SilentlyContinue |
                Where-Object { $wildcard.IsMatch($_.Name) }
            )

            if ($matched.Count -eq 0) {
                throw ("DATE_SUBPATH segment '{0}' matched no folders under '{1}'" -f $segment, $current)
            }

            if ($matched.Count -gt 1) {
                $names = ($matched | ForEach-Object { $_.Name }) -join ', '
                throw ("DATE_SUBPATH segment '{0}' matched multiple folders under '{1}': {2}" -f `
                        $segment, $current, $names)
            }

            $current = $matched[0].FullName
        }
        else {
            $current = Join-Path $current $segment
        }
    }

    return $current
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

function Get-BranchMonthNumber {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BranchMonth
    )

    if ($BranchMonth -match '^\s*0*(\d{1,2})') {
        return [int]$Matches[1]
    }

    return $null
}

function Resolve-SrBranchFolderPath {
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

    $yearPath = Join-Path $BasePath $BranchYear.Trim()
    $fallbackPath = Join-Path $yearPath $BranchMonth.Trim()
    $monthNumber = Get-BranchMonthNumber -BranchMonth $BranchMonth

    if ($null -eq $monthNumber) {
        Write-Log ("[open-sr] Could not parse month number from BRANCH_MONTH='{0}'; using literal folder '{1}'" -f `
                $BranchMonth, $fallbackPath) -Level WARNING
        return $fallbackPath
    }

    if (-not (Test-Path -LiteralPath $yearPath -PathType Container)) {
        Write-Log ("[open-sr] SR year folder missing: Path='{0}'; falling back to '{1}'" -f `
                $yearPath, $fallbackPath) -Level WARNING
        return $fallbackPath
    }

    $monthPattern = '^0?' + $monthNumber + '(?:\D|$)'
    $matchedFolder = Get-ChildItem -LiteralPath $yearPath -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match $monthPattern } |
    Select-Object -First 1

    if ($matchedFolder) {
        Write-Log ("[open-sr] Resolved SR month folder: BranchMonth='{0}', MonthNumber={1}, Path='{2}'" -f `
                $BranchMonth, $monthNumber, $matchedFolder.FullName) -Level DEBUG
        return $matchedFolder.FullName
    }

    Write-Log ("[open-sr] No SR month folder matching month {0} under '{1}'; falling back to '{2}'" -f `
            $monthNumber, $yearPath, $fallbackPath) -Level WARNING
    return $fallbackPath
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
        $start = $end.AddDays( - ($spanDays - 1))
    }
    elseif ($period -eq 'previous') {
        $end = $ref
        while ($end.DayOfWeek -ne $endDow) {
            $end = $end.AddDays(-1)
        }
        $end = $end.AddDays(-$spanDays)
        $start = $end.AddDays( - ($spanDays - 1))
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

function Parse-LabeledPaths {
    param(
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$SettingName
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
            throw "Invalid $SettingName segment (expected Label=Path): $part"
        }

        $label = $part.Substring(0, $eqIndex).Trim()
        $pathValue = $part.Substring($eqIndex + 1).Trim()

        if ([string]::IsNullOrWhiteSpace($label) -or [string]::IsNullOrWhiteSpace($pathValue)) {
            throw "Invalid $SettingName segment (empty label or path): $part"
        }

        $resolved = Resolve-ConfigPath -PathValue $pathValue
        $list.Add([pscustomobject]@{
                Label = $label
                Path  = $resolved
            })
    }

    return @($list.ToArray())
}

function Parse-BranchPaths {
    param(
        [string]$Value
    )

    return @(Parse-LabeledPaths -Value $Value -SettingName 'BRANCH_PATHS')
}

function Parse-SrPaths {
    param(
        [string]$Value
    )

    return @(Parse-LabeledPaths -Value $Value -SettingName 'SR_PATHS')
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
            $groups[$groupLabel] = @{
                Paths     = [System.Collections.Generic.List[string]]::new()
                BasePaths = [System.Collections.Generic.List[string]]::new()
            }
            $ordered.Add($groupLabel)
        }
        $groups[$groupLabel].Paths.Add($branch.Path)
        if ($branch.PSObject.Properties['BasePath'] -and -not [string]::IsNullOrWhiteSpace([string]$branch.BasePath)) {
            $groups[$groupLabel].BasePaths.Add([string]$branch.BasePath)
        }
    }

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($groupLabel in $ordered) {
        $group = $groups[$groupLabel]
        [string[]]$paths = @($group.Paths.ToArray())
        if ($group.BasePaths.Count -gt 0 -and $group.BasePaths.Count -eq $group.Paths.Count) {
            [string[]]$basePaths = @($group.BasePaths.ToArray())
            $list.Add([pscustomobject]@{
                    Label     = $groupLabel
                    Paths     = [string[]]@($paths)
                    Path      = $paths[0]
                    BasePaths = [string[]]@($basePaths)
                    BasePath  = $basePaths[0]
                })
        }
        else {
            $list.Add([pscustomobject]@{
                    Label = $groupLabel
                    Paths = [string[]]@($paths)
                    Path  = $paths[0]
                })
        }
    }

    return @($list.ToArray())
}

function Parse-FileExtensions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$SettingName
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
        throw "$SettingName must contain at least one extension"
    }

    return @($extensions.ToArray())
}

function Parse-CutoffFileExtensions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return @(Parse-FileExtensions -Value $Value -SettingName 'CUTOFF_FILE_EXTENSIONS')
}

function Parse-SrFileExtensions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return @(Parse-FileExtensions -Value $Value -SettingName 'SR_FILE_EXTENSIONS')
}

function Get-DayRangeMatchPattern {
    param(
        [Parameter(Mandatory = $true)]
        [int]$StartDay,

        [Parameter(Mandatory = $true)]
        [int]$EndDay
    )

    return "(?<!\d)0?$StartDay\D{0,20}-\D{0,20}0?$EndDay(?!\d)"
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

function Get-PayrollPeriodDayNumbers {
    param(
        [Parameter(Mandatory = $true)]
        [DateTime]$PeriodStart,

        [Parameter(Mandatory = $true)]
        [DateTime]$PeriodEnd
    )

    $days = [System.Collections.Generic.HashSet[int]]::new()
    $cursor = $PeriodStart.Date
    $end = $PeriodEnd.Date

    while ($cursor -le $end) {
        [void]$days.Add($cursor.Day)
        $cursor = $cursor.AddDays(1)
    }

    # PowerShell 5.1 HashSet[T] has no LINQ ToArray(); enumerate instead.
    return @($days | Sort-Object)
}

function Get-PayrollPeriodMonthDayGroups {
    param(
        [Parameter(Mandatory = $true)]
        [DateTime]$PeriodStart,

        [Parameter(Mandatory = $true)]
        [DateTime]$PeriodEnd
    )

    $groups = [System.Collections.Generic.List[object]]::new()
    $cursor = $PeriodStart.Date
    $end = $PeriodEnd.Date

    $currentAnchor = $null
    $currentYear = 0
    $currentMonth = 0
    $currentDays = $null

    while ($cursor -le $end) {
        if ($null -eq $currentDays -or $cursor.Year -ne $currentYear -or $cursor.Month -ne $currentMonth) {
            if ($null -ne $currentDays) {
                $groups.Add([pscustomobject]@{
                        AnchorDate = $currentAnchor
                        Days       = [int[]]@($currentDays.ToArray())
                    })
            }
            $currentAnchor = $cursor
            $currentYear = $cursor.Year
            $currentMonth = $cursor.Month
            $currentDays = [System.Collections.Generic.List[int]]::new()
        }

        $currentDays.Add($cursor.Day)
        $cursor = $cursor.AddDays(1)
    }

    if ($null -ne $currentDays) {
        $groups.Add([pscustomobject]@{
                AnchorDate = $currentAnchor
                Days       = [int[]]@($currentDays.ToArray())
            })
    }

    return @($groups.ToArray())
}

function Test-FilenameContainsSalesReportDay {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [Parameter(Mandatory = $true)]
        [int[]]$PeriodDays
    )

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    if ([string]::IsNullOrWhiteSpace($stem)) {
        return $false
    }

    foreach ($day in $PeriodDays) {
        if ($day -lt 10) {
            $dayToken = "0?$day"
        }
        else {
            $dayToken = "$day"
        }

        $pattern = "(^|[-_\s.,])$dayToken($|[-_\s.,])"
        if ($stem -match $pattern) {
            return $true
        }
    }

    return $false
}

function Get-SalesReportFilesInFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath,

        [Parameter(Mandatory = $true)]
        [int[]]$PeriodDays
    )

    $matches = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        Write-Log "[open-sr] Scan path missing or not a folder: Path='$FolderPath'" -Level DEBUG
        return @()
    }

    $allowedExtensions = $script:sr_file_extensions
    $childParams = @{
        LiteralPath = $FolderPath
        File        = $true
        ErrorAction = 'Stop'
    }
    if ($script:recursive) {
        $childParams['Recurse'] = $true
        $childParams['Depth'] = $script:dir_level_search
    }

    Write-Log ("[open-sr] Scanning path: Path='{0}', PeriodDays='{1}', Recursive={2}, Depth={3}, Extensions='{4}'" -f `
            $FolderPath, ($PeriodDays -join ','), $script:recursive, $script:dir_level_search, ($allowedExtensions -join ',')) -Level DEBUG

    $files = @(Get-ChildItem @childParams |
        Where-Object {
            -not $_.Name.StartsWith('~$') -and
            ($allowedExtensions -icontains $_.Extension)
        })

    Write-Log ("[open-sr] Candidate files found: Path='{0}', Count={1}" -f $FolderPath, $files.Count) -Level DEBUG

    foreach ($file in $files) {
        if (Test-FilenameContainsSalesReportDay -FileName $file.Name -PeriodDays $PeriodDays) {
            Write-Log ("[open-sr] Matched sales report file: Path='{0}', File='{1}'" -f $FolderPath, $file.FullName) -Level DEBUG
            $matches.Add($file)
        }
    }

    return @($matches.ToArray())
}

function Resolve-SrScanFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [DateTime]$AnchorDate
    )

    if (-not [string]::IsNullOrWhiteSpace($script:date_subpath)) {
        return Resolve-DateSubpath `
            -BasePath $BasePath `
            -Pattern $script:date_subpath `
            -Date $AnchorDate
    }

    if ($script:use_branch_year_month) {
        return Resolve-SrBranchFolderPath `
            -BasePath $BasePath `
            -UseBranchYearMonth $true `
            -BranchYear $AnchorDate.ToString('yyyy') `
            -BranchMonth $AnchorDate.Month.ToString()
    }

    return $BasePath
}

function Get-SrBranchGroupByLabel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BranchLabel
    )

    foreach ($branch in $script:sr_paths) {
        if ($branch.Label -eq $BranchLabel) {
            return $branch
        }
    }

    return $null
}

function Find-BranchSalesReportFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BranchLabel,

        [Parameter(Mandatory = $true)]
        [DateTime]$PeriodStart,

        [Parameter(Mandatory = $true)]
        [DateTime]$PeriodEnd
    )

    $branch = Get-SrBranchGroupByLabel -BranchLabel $BranchLabel
    if (-not $branch) {
        Write-Log ("[open-sr] No SR paths configured for branch: Label='{0}'" -f $BranchLabel) -Level DEBUG
        return @()
    }

    $monthGroups = @(Get-PayrollPeriodMonthDayGroups -PeriodStart $PeriodStart -PeriodEnd $PeriodEnd)
    if ($monthGroups.Count -eq 0) {
        return @()
    }

    $hasBasePaths = $branch.PSObject.Properties['BasePaths'] -and $branch.BasePaths -and @($branch.BasePaths).Count -gt 0
    [string[]]$scanBases = if ($hasBasePaths) {
        @($branch.BasePaths)
    }
    elseif ($branch.PSObject.Properties['BasePath'] -and -not [string]::IsNullOrWhiteSpace([string]$branch.BasePath)) {
        @([string]$branch.BasePath)
    }
    else {
        @()
    }

    [string[]]$fallbackPaths = if ($branch.PSObject.Properties['Paths'] -and $branch.Paths) {
        @($branch.Paths)
    }
    else {
        @($branch.Path)
    }

    $resolvePerMonth = $scanBases.Count -gt 0 -and (
        -not [string]::IsNullOrWhiteSpace($script:date_subpath) -or $script:use_branch_year_month
    )

    [string[]]$roots = if ($resolvePerMonth) { $scanBases } else { $fallbackPaths }

    Write-Log ("[open-sr] Finding sales report files: Label='{0}', RootCount={1}, MonthGroups={2}, ResolvePerMonth={3}, Roots='{4}'" -f `
            $BranchLabel, $roots.Count, $monthGroups.Count, $resolvePerMonth, ($roots -join "'; '")) -Level DEBUG

    $seen = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $allMatches = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

    foreach ($root in $roots) {
        foreach ($group in $monthGroups) {
            $folderPath = $null
            try {
                if ($resolvePerMonth) {
                    $folderPath = Resolve-SrScanFolder -BasePath $root -AnchorDate $group.AnchorDate
                }
                else {
                    $folderPath = $root
                }
            }
            catch {
                Write-Log ("[open-sr] Could not resolve SR folder for month: Label='{0}', Base='{1}', Anchor='{2:yyyy-MM-dd}', Message='{3}'" -f `
                        $BranchLabel, $root, $group.AnchorDate, $_.Exception.Message) -Level DEBUG
                continue
            }

            Write-Log ("[open-sr] Month segment scan: Label='{0}', Anchor='{1:yyyy-MM-dd}', Days='{2}', Path='{3}'" -f `
                    $BranchLabel, $group.AnchorDate, ($group.Days -join ','), $folderPath) -Level DEBUG

            try {
                $folderMatches = @(Get-SalesReportFilesInFolder -FolderPath $folderPath -PeriodDays $group.Days)
                foreach ($match in $folderMatches) {
                    if ($seen.Add($match.FullName)) {
                        $allMatches.Add($match)
                    }
                }
            }
            catch {
                Write-Log ("[open-sr] Scan error: Label='{0}', Path='{1}', Message='{2}'" -f `
                        $BranchLabel, $folderPath, $_.Exception.Message) -Level ERROR
            }
        }
    }

    return @($allMatches.ToArray())
}

function Get-CurrentPayrollPeriod {
    $referenceDate = if ($script:ref_run_date) { $script:ref_run_date } else { [DateTime]::Today }

    return Get-BranchPayrollPeriod `
        -StartDay $script:branch_payroll_start_day `
        -EndDay $script:branch_payroll_end_day `
        -TargetPeriod $script:payroll_target_period `
        -ReferenceDate $referenceDate
}

function Open-BranchSalesReports {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BranchLabel
    )

    Write-Log ("[open-sr] Open requested: Label='{0}'" -f $BranchLabel) -Level DEBUG

    if ($script:sr_paths.Count -eq 0) {
        Write-Log "[open-sr] No SR paths configured in environment" -Level WARNING
        Show-CutoffMessageBox `
            -Text 'No SR paths are configured. Add SR_PATHS entries to your .env file.' `
            -Caption 'Open SR'
        return
    }

    $branch = Get-SrBranchGroupByLabel -BranchLabel $BranchLabel
    if (-not $branch) {
        Write-Log ("[open-sr] No SR path mapping for branch: Label='{0}'" -f $BranchLabel) -Level WARNING
        Show-CutoffMessageBox `
            -Text "No SR path is configured for branch '$BranchLabel'.`nEnsure SR_PATHS uses the same label as BRANCH_PATHS." `
            -Caption 'Open SR'
        return
    }

    $period = Get-CurrentPayrollPeriod
    try {
        $monthGroups = @(Get-PayrollPeriodMonthDayGroups -PeriodStart $period.Start -PeriodEnd $period.End)
        $periodDays = @(Get-PayrollPeriodDayNumbers -PeriodStart $period.Start -PeriodEnd $period.End)
    }
    catch {
        Write-Log ("[open-sr] Failed to build period day list: Label='{0}', Period='{1:yyyy-MM-dd}' to '{2:yyyy-MM-dd}', Message='{3}'" -f `
                $BranchLabel, $period.Start, $period.End, $_.Exception.Message) -Level ERROR
        Show-CutoffMessageBox `
            -Text "Could not determine payroll days for '$BranchLabel'.`n$($_.Exception.Message)" `
            -Caption 'Open SR'
        return
    }

    if ($monthGroups.Count -eq 0 -or $periodDays.Count -eq 0) {
        Write-Log ("[open-sr] Period day list unexpectedly empty: Label='{0}', Period='{1:yyyy-MM-dd}' to '{2:yyyy-MM-dd}'" -f `
                $BranchLabel, $period.Start, $period.End) -Level ERROR
        Show-CutoffMessageBox `
            -Text "Could not determine payroll days for '$BranchLabel' (empty day list)." `
            -Caption 'Open SR'
        return
    }

    Write-Log ("[open-sr] Period days for matching: Label='{0}', Days='{1}', MonthGroups={2}, Period='{3:yyyy-MM-dd}' to '{4:yyyy-MM-dd}'" -f `
            $BranchLabel, ($periodDays -join ','), $monthGroups.Count, $period.Start, $period.End) -Level DEBUG

    $matchedFiles = @(Find-BranchSalesReportFiles `
            -BranchLabel $BranchLabel `
            -PeriodStart $period.Start `
            -PeriodEnd $period.End)
    if ($matchedFiles.Count -eq 0) {
        Write-Log ("[open-sr] No matching sales report files: Label='{0}', Days='{1}'" -f `
                $BranchLabel, ($periodDays -join ',')) -Level WARNING
        Show-CutoffMessageBox `
            -Text "No sales report files found for '$BranchLabel' matching payroll days $($periodDays -join ', ')." `
            -Caption 'Open SR'
        return
    }

    $filesToOpen = @($matchedFiles | Select-Object -First $script:sr_open_max)
    $truncated = $matchedFiles.Count -gt $filesToOpen.Count

    if ($truncated) {
        Write-Log ("[open-sr] Match count exceeds SR_OPEN_MAX: Label='{0}', Total={1}, Opening={2}, Max={3}" -f `
                $BranchLabel, $matchedFiles.Count, $filesToOpen.Count, $script:sr_open_max) -Level WARNING
        Show-CutoffMessageBox `
            -Text ("Opened {0} of {1} matching files for '{2}'.`nIncrease SR_OPEN_MAX to open more." -f `
                $filesToOpen.Count, $matchedFiles.Count, $BranchLabel) `
            -Caption 'Open SR' `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Information)
    }

    $openedCount = 0
    foreach ($file in $filesToOpen) {
        try {
            Write-Log ("[open-sr] Launching file: Label='{0}', File='{1}'" -f $BranchLabel, $file.FullName) -Level DEBUG
            Start-Process -FilePath $file.FullName -ErrorAction Stop
            $openedCount++
        }
        catch {
            $message = $_.Exception.Message
            Write-Log ("[open-sr] Could not open file (missing handler or shell launch error): Label='{0}', File='{1}', Message='{2}'" -f `
                    $BranchLabel, $file.FullName, $message) -Level ERROR
        }
    }

    Write-Log ("[open-sr] Open completed: Label='{0}', Opened={1}, TotalMatches={2}, Truncated={3}" -f `
            $BranchLabel, $openedCount, $matchedFiles.Count, $truncated) -Level INFO
}

function Open-BranchSalesReportFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BranchLabel
    )

    Write-Log ("[open-sr-folder] Open requested: Label='{0}'" -f $BranchLabel) -Level DEBUG

    if ($script:sr_paths.Count -eq 0) {
        Write-Log "[open-sr-folder] No SR paths configured in environment" -Level WARNING
        Show-CutoffMessageBox `
            -Text 'No SR paths are configured. Add SR_PATHS entries to your .env file.' `
            -Caption 'SR Folder'
        return
    }

    $branch = Get-SrBranchGroupByLabel -BranchLabel $BranchLabel
    if (-not $branch) {
        Write-Log ("[open-sr-folder] No SR path mapping for branch: Label='{0}'" -f $BranchLabel) -Level WARNING
        Show-CutoffMessageBox `
            -Text "No SR path is configured for branch '$BranchLabel'.`nEnsure SR_PATHS uses the same label as BRANCH_PATHS." `
            -Caption 'SR Folder'
        return
    }

    $period = Get-CurrentPayrollPeriod
    try {
        $monthGroups = @(Get-PayrollPeriodMonthDayGroups -PeriodStart $period.Start -PeriodEnd $period.End)
    }
    catch {
        Write-Log ("[open-sr-folder] Failed to build period month groups: Label='{0}', Period='{1:yyyy-MM-dd}' to '{2:yyyy-MM-dd}', Message='{3}'" -f `
                $BranchLabel, $period.Start, $period.End, $_.Exception.Message) -Level ERROR
        Show-CutoffMessageBox `
            -Text "Could not determine payroll period folders for '$BranchLabel'.`n$($_.Exception.Message)" `
            -Caption 'SR Folder'
        return
    }

    if ($monthGroups.Count -eq 0) {
        Write-Log ("[open-sr-folder] Period month groups unexpectedly empty: Label='{0}', Period='{1:yyyy-MM-dd}' to '{2:yyyy-MM-dd}'" -f `
                $BranchLabel, $period.Start, $period.End) -Level ERROR
        Show-CutoffMessageBox `
            -Text "Could not determine payroll period folders for '$BranchLabel' (empty month list)." `
            -Caption 'SR Folder'
        return
    }

    $hasBasePaths = $branch.PSObject.Properties['BasePaths'] -and $branch.BasePaths -and @($branch.BasePaths).Count -gt 0
    [string[]]$scanBases = if ($hasBasePaths) {
        @($branch.BasePaths)
    }
    elseif ($branch.PSObject.Properties['BasePath'] -and -not [string]::IsNullOrWhiteSpace([string]$branch.BasePath)) {
        @([string]$branch.BasePath)
    }
    else {
        @()
    }

    [string[]]$fallbackPaths = if ($branch.PSObject.Properties['Paths'] -and $branch.Paths) {
        @($branch.Paths)
    }
    else {
        @($branch.Path)
    }

    $resolvePerMonth = $scanBases.Count -gt 0 -and (
        -not [string]::IsNullOrWhiteSpace($script:date_subpath) -or $script:use_branch_year_month
    )

    [string[]]$roots = if ($resolvePerMonth) { $scanBases } else { $fallbackPaths }

    $seen = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $resolvedFolders = [System.Collections.Generic.List[string]]::new()

    foreach ($root in $roots) {
        foreach ($group in $monthGroups) {
            $folderPath = $null
            try {
                if ($resolvePerMonth) {
                    $folderPath = Resolve-SrScanFolder -BasePath $root -AnchorDate $group.AnchorDate
                }
                else {
                    $folderPath = $root
                }
            }
            catch {
                Write-Log ("[open-sr-folder] Could not resolve SR folder for month: Label='{0}', Base='{1}', Anchor='{2:yyyy-MM-dd}', Message='{3}'" -f `
                        $BranchLabel, $root, $group.AnchorDate, $_.Exception.Message) -Level DEBUG
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($folderPath) -and $seen.Add($folderPath)) {
                $resolvedFolders.Add($folderPath)
                Write-Log ("[open-sr-folder] Resolved SR folder: Label='{0}', Anchor='{1:yyyy-MM-dd}', Path='{2}'" -f `
                        $BranchLabel, $group.AnchorDate, $folderPath) -Level DEBUG
            }
        }
    }

    if ($resolvedFolders.Count -eq 0) {
        Write-Log ("[open-sr-folder] No SR folders resolved: Label='{0}'" -f $BranchLabel) -Level WARNING
        Show-CutoffMessageBox `
            -Text "No SR folder could be resolved for '$BranchLabel'." `
            -Caption 'SR Folder'
        return
    }

    $folderToOpen = $resolvedFolders[0]
    Write-Log ("[open-sr-folder] Opening SR folder: Label='{0}', Path='{1}', ResolvedCount={2}" -f `
            $BranchLabel, $folderToOpen, $resolvedFolders.Count) -Level INFO
    Open-CutoffFolderInExplorer -FolderPath $folderToOpen -BranchLabel $BranchLabel
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
        Write-Log "[open-flow] Scan path missing or not a folder: Path='$FolderPath'" -Level DEBUG
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

    Write-Log ("[open-flow] Scanning path: Path='{0}', StartDay={1}, EndDay={2}, Recursive={3}, Depth={4}, Extensions={5}" -f `
            $FolderPath, $StartDay, $EndDay, $script:recursive, $script:dir_level_search, ($allowedExtensions -join ',')) -Level DEBUG

    $files = @(Get-ChildItem @childParams |
        Where-Object {
            -not $_.Name.StartsWith('~$') -and
            ($allowedExtensions -icontains $_.Extension)
        })

    Write-Log ("[open-flow] Candidate files found: Path='{0}', Count={1}" -f $FolderPath, $files.Count) -Level DEBUG

    foreach ($file in $files) {
        if (Test-FilenameContainsDayRange -FileName $file.Name -StartDay $StartDay -EndDay $EndDay) {
            Write-Log ("[open-flow] Matched cutoff file: Path='{0}', File='{1}'" -f $FolderPath, $file.FullName) -Level DEBUG
            return $true
        }
    }

    Write-Log "[open-flow] No matching cutoff file in path: Path='$FolderPath'" -Level DEBUG
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

        [DateTime]$PeriodStart,

        [DateTime]$PeriodEnd,

        [scriptblock]$OnError
    )

    $hasPeriod = $PSBoundParameters.ContainsKey('PeriodStart') -and $PSBoundParameters.ContainsKey('PeriodEnd')
    $monthGroups = @()
    if ($hasPeriod) {
        $monthGroups = @(Get-PayrollPeriodMonthDayGroups -PeriodStart $PeriodStart -PeriodEnd $PeriodEnd)
    }

    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($branch in $BranchPaths) {
        [string[]]$fallbackPaths = if ($branch.PSObject.Properties['Paths'] -and $branch.Paths) {
            @($branch.Paths)
        }
        else {
            @($branch.Path)
        }

        $hasBasePaths = $branch.PSObject.Properties['BasePaths'] -and $branch.BasePaths -and @($branch.BasePaths).Count -gt 0
        [string[]]$scanBases = if ($hasBasePaths) {
            @($branch.BasePaths)
        }
        elseif ($branch.PSObject.Properties['BasePath'] -and -not [string]::IsNullOrWhiteSpace([string]$branch.BasePath)) {
            @([string]$branch.BasePath)
        }
        else {
            @()
        }

        $resolvePerMonth = $hasPeriod -and $monthGroups.Count -gt 0 -and $scanBases.Count -gt 0 -and (
            -not [string]::IsNullOrWhiteSpace($script:date_subpath) -or $script:use_branch_year_month
        )

        $scanFolders = [System.Collections.Generic.List[string]]::new()
        if ($resolvePerMonth) {
            foreach ($root in $scanBases) {
                foreach ($group in $monthGroups) {
                    try {
                        $folderPath = Resolve-SrScanFolder -BasePath $root -AnchorDate $group.AnchorDate
                        if (-not [string]::IsNullOrWhiteSpace($folderPath) -and -not ($scanFolders -contains $folderPath)) {
                            $scanFolders.Add($folderPath)
                        }
                        Write-Log ("[open-flow] Month segment folder: Label='{0}', Anchor='{1:yyyy-MM-dd}', Path='{2}'" -f `
                                $branch.Label, $group.AnchorDate, $folderPath) -Level DEBUG
                    }
                    catch {
                        Write-Log ("[open-flow] Could not resolve incentives folder for month: Label='{0}', Base='{1}', Anchor='{2:yyyy-MM-dd}', Message='{3}'" -f `
                                $branch.Label, $root, $group.AnchorDate, $_.Exception.Message) -Level DEBUG
                    }
                }
            }
        }

        [string[]]$paths = if ($scanFolders.Count -gt 0) {
            @($scanFolders.ToArray())
        }
        else {
            $fallbackPaths
        }

        Write-Log ("[open-flow] Evaluating branch: Label='{0}', PathCount={1}, ResolvePerMonth={2}, Paths='{3}'" -f `
                $branch.Label, $paths.Count, $resolvePerMonth, ($paths -join "'; '")) -Level DEBUG

        $hasFile = $false
        $matchedPath = $null
        $pathResults = [System.Collections.Generic.List[object]]::new()

        foreach ($folderPath in $paths) {
            $pathHasFile = $false
            try {
                Write-Log ("[open-flow] Checking branch path: Label='{0}', Path='{1}'" -f $branch.Label, $folderPath) -Level DEBUG
                $pathHasFile = Test-BranchHasCutoffFile -FolderPath $folderPath -StartDay $StartDay -EndDay $EndDay
                Write-Log ("[open-flow] Branch path result: Label='{0}', Path='{1}', HasFile={2}" -f `
                        $branch.Label, $folderPath, $pathHasFile) -Level DEBUG
            }
            catch {
                $pathHasFile = $false
                Write-Log ("[open-flow] Branch path scan error: Label='{0}', Path='{1}', Message='{2}'" -f `
                        $branch.Label, $folderPath, $_.Exception.Message) -Level DEBUG
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

        $selectedPath = if ($matchedPath) { $matchedPath } elseif ($paths.Count -gt 0) { $paths[0] } else { $fallbackPaths[0] }
        Write-Log ("[open-flow] Branch final result: Label='{0}', HasFile={1}, SelectedPath='{2}', MatchedPath='{3}'" -f `
                $branch.Label, $hasFile, $selectedPath, $matchedPath) -Level DEBUG

        $results.Add([pscustomobject]@{
                Label       = $branch.Label
                Path        = $selectedPath
                Paths       = [string[]]@($paths)
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
        [System.Media.SystemSounds]::Exclamation.Play()
    }
    catch {
        Write-Host "Could not play notification sound: $($_.Exception.Message)"
    }
}

function Get-ExplorerFolderPathArgument {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    $resolvedPath = (Resolve-Path -LiteralPath $FolderPath).ProviderPath
    return '"' + $resolvedPath + '"'
}

function Show-CutoffMessageBox {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Caption,

        $Buttons = $null,

        $Icon = $null
    )

    if ($null -eq $Buttons) {
        $Buttons = [System.Windows.Forms.MessageBoxButtons]::OK
    }
    if ($null -eq $Icon) {
        $Icon = [System.Windows.Forms.MessageBoxIcon]::Warning
    }

    $owner = $script:cutoffPopupForm
    if ($owner -and -not $owner.IsDisposed) {
        [void][System.Windows.Forms.MessageBox]::Show($owner, $Text, $Caption, $Buttons, $Icon)
    }
    else {
        [void][System.Windows.Forms.MessageBox]::Show($Text, $Caption, $Buttons, $Icon)
    }
}

function Open-CutoffFolderInExplorer {
    param(
        [string]$FolderPath,

        [string]$BranchLabel
    )

    Write-Log ("[open-flow] Open requested: Label='{0}', InputPath='{1}'" -f $BranchLabel, $FolderPath) -Level DEBUG

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        Write-Log ("[open-flow] Open blocked: Label='{0}', Reason='blank path'" -f $BranchLabel) -Level DEBUG
        Show-CutoffMessageBox `
            -Text 'No folder path is configured for this branch.' `
            -Caption 'Open folder'
        return
    }

    $pathExists = Test-Path -LiteralPath $FolderPath -PathType Container
    Write-Log ("[open-flow] Open path test: Label='{0}', InputPath='{1}', Exists={2}" -f $BranchLabel, $FolderPath, $pathExists) -Level DEBUG

    if (-not $pathExists) {
        Write-Log ("[open-flow] Open blocked: Label='{0}', Reason='folder not found', InputPath='{1}'" -f $BranchLabel, $FolderPath) -Level DEBUG
        Show-CutoffMessageBox `
            -Text "Folder not found:`n$FolderPath" `
            -Caption 'Open folder'
        return
    }

    try {
        $explorerPathArgument = Get-ExplorerFolderPathArgument -FolderPath $FolderPath
        Write-Log ("[open-flow] Explorer launch: Label='{0}', InputPath='{1}', Argument={2}" -f `
                $BranchLabel, $FolderPath, $explorerPathArgument) -Level DEBUG
        Start-Process -FilePath 'explorer.exe' -ArgumentList $explorerPathArgument -ErrorAction Stop
        Write-Log ("[open-flow] Explorer launch requested successfully: Label='{0}', InputPath='{1}'" -f $BranchLabel, $FolderPath) -Level DEBUG
    }
    catch {
        Write-Log ("[open-flow] Explorer launch failed: Label='{0}', InputPath='{1}', Message='{2}'" -f `
                $BranchLabel, $FolderPath, $_.Exception.Message) -Level DEBUG
        Show-CutoffMessageBox `
            -Text "Could not open folder:`n$FolderPath`n`n$($_.Exception.Message)" `
            -Caption 'Open folder'
    }
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
    $showSrButton = [bool]$script:show_sr_button
    $formWidth = if ($showSrButton) { 780 } else { 520 }
    $contentWidth = $formWidth - (2 * $margin)
    $buttonHeight = 28
    $buttonWidth = 86
    $buttonGap = 12
    $rowHeight = 30
    $rowGap = 2
    $incentivesButtonWidth = 100
    $srButtonWidth = 70
    $srFolderButtonWidth = 90
    $openButtonHeight = 26
    $openButtonGap = 8
    $rowButtonsWidth = if ($showSrButton) {
        $incentivesButtonWidth + $openButtonGap + $srButtonWidth + $openButtonGap + $srFolderButtonWidth
    }
    else {
        $incentivesButtonWidth
    }
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
    $form.FormBorderStyle = 'FixedSingle'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.ShowInTaskbar = $true
    $form.TopMost = $false
    $form.ClientSize = New-Object System.Drawing.Size(
        $formWidth,
        ((2 * $margin) + $contentHeight + $buttonGap + $buttonHeight)
    )
    $script:cutoffPopupForm = $form

    $listPanel = New-Object System.Windows.Forms.Panel
    $listPanel.Location = New-Object System.Drawing.Point($margin, $margin)
    $listPanel.Size = New-Object System.Drawing.Size($contentWidth, $contentHeight)
    $listPanel.AutoScroll = $true
    $listPanel.BorderStyle = 'None'

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = 'Close'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $okButton.Location = New-Object System.Drawing.Point(
        ($formWidth - $margin - $buttonWidth),
        ($margin + $contentHeight + $buttonGap)
    )

    $showFinalScButton = -not [string]::IsNullOrWhiteSpace($script:final_sc_incentives_path)
    $finalScButton = $null
    if ($showFinalScButton) {
        $finalScButtonWidth = 120
        $finalScButton = New-Object System.Windows.Forms.Button
        $finalScButton.Text = 'Final SC'
        $finalScButton.Size = New-Object System.Drawing.Size($finalScButtonWidth, $buttonHeight)
        $finalScButton.Location = New-Object System.Drawing.Point(
            ($formWidth - $margin - $buttonWidth - $buttonGap - $finalScButtonWidth),
            ($margin + $contentHeight + $buttonGap)
        )
        $finalScButton.Add_Click({
                param($sender, $eventArgs)
                Write-Log ("[open-final-sc] Open Final SC clicked: Path='{0}'" -f $script:final_sc_incentives_path) -Level DEBUG
                Open-CutoffFolderInExplorer -FolderPath ([string]$script:final_sc_incentives_path) -BranchLabel 'Final SC'
            })
    }

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
            $rowInnerWidth - $rowButtonsWidth - $openButtonGap
        )
        $labelHeight = [Math]::Max($font.Height + 2, 18)

        $y = 0
        foreach ($item in @($RowResults)) {
            $status = if ($item.HasFile) { '[OK]' } else { '[X]' }
            $pathResultsForLog = if ($item.PSObject.Properties['PathResults'] -and $item.PathResults) {
                @(foreach ($pathResult in $item.PathResults) {
                        '{0}={1}' -f $pathResult.Path, $pathResult.HasFile
                    }) -join '; '
            }
            else {
                ''
            }
            Write-Log ("[open-flow] Popup row bind: Label='{0}', HasFile={1}, ButtonPath='{2}', PathResults='{3}'" -f `
                    $item.Label, $item.HasFile, $item.Path, $pathResultsForLog) -Level DEBUG

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

            $incentivesButton = New-Object System.Windows.Forms.Button
            $incentivesButton.Text = 'Open'
            $incentivesButton.Size = New-Object System.Drawing.Size($incentivesButtonWidth, $openButtonHeight)
            $incentivesButton.Location = New-Object System.Drawing.Point(
                ($rowInnerWidth - $rowButtonsWidth),
                [Math]::Floor(($rowHeight - $openButtonHeight) / 2)
            )
            $incentivesButton.Tag = [pscustomobject]@{
                Label       = $item.Label
                Path        = $item.Path
                HasFile     = $item.HasFile
                PathResults = $item.PathResults
            }
            $incentivesButton.Add_Click({
                    param($sender, $eventArgs)
                    $tag = $sender.Tag
                    Write-Log ("[open-flow] Open Incentives clicked: Label='{0}', Path='{1}', HasFile={2}" -f `
                            $tag.Label, $tag.Path, $tag.HasFile) -Level DEBUG
                    Open-CutoffFolderInExplorer -FolderPath ([string]$tag.Path) -BranchLabel ([string]$tag.Label)
                })

            $rowPanel.Controls.Add($statusLabel)
            $rowPanel.Controls.Add($incentivesButton)

            if ($showSrButton) {
                $srFolderButton = New-Object System.Windows.Forms.Button
                $srFolderButton.Text = 'SR Folder'
                $srFolderButton.Size = New-Object System.Drawing.Size($srFolderButtonWidth, $openButtonHeight)
                $srFolderButton.Location = New-Object System.Drawing.Point(
                    ($rowInnerWidth - $srFolderButtonWidth),
                    [Math]::Floor(($rowHeight - $openButtonHeight) / 2)
                )
                $srFolderButton.Tag = [pscustomobject]@{
                    Label = $item.Label
                }
                $srFolderButton.Add_Click({
                        param($sender, $eventArgs)
                        $tag = $sender.Tag
                        Write-Log ("[open-sr-folder] SR Folder clicked: Label='{0}'" -f $tag.Label) -Level DEBUG
                        Open-BranchSalesReportFolder -BranchLabel ([string]$tag.Label)
                    })

                $srButton = New-Object System.Windows.Forms.Button
                $srButton.Text = 'SR'
                $srButton.Size = New-Object System.Drawing.Size($srButtonWidth, $openButtonHeight)
                $srButton.Location = New-Object System.Drawing.Point(
                    ($rowInnerWidth - $srFolderButtonWidth - $openButtonGap - $srButtonWidth),
                    [Math]::Floor(($rowHeight - $openButtonHeight) / 2)
                )
                $srButton.Tag = [pscustomobject]@{
                    Label = $item.Label
                }
                $srButton.Add_Click({
                        param($sender, $eventArgs)
                        $tag = $sender.Tag
                        Write-Log ("[open-sr] Open SR clicked: Label='{0}'" -f $tag.Label) -Level DEBUG
                        Open-BranchSalesReports -BranchLabel ([string]$tag.Label)
                    })

                $rowPanel.Controls.Add($srButton)
                $rowPanel.Controls.Add($srFolderButton)
            }

            $listPanel.Controls.Add($rowPanel)
            $y += ($rowHeight + $rowGap)
        }

        $listPanel.ResumeLayout()
    }

    & $rebuildRows $Results

    $form.Controls.Add($listPanel)
    $form.Controls.Add($okButton)
    if ($finalScButton) {
        $form.Controls.Add($finalScButton)
    }
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
    $script:cutoffPopupForm = $null
    $form.Dispose()
    $font.Dispose()
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('TRACE', 'DEBUG', 'INFO', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $configuredLevel = if ([string]::IsNullOrWhiteSpace($script:log_level)) { 'INFO' } else { $script:log_level }
    if ((Get-LogLevelRank -Level $Level) -lt (Get-LogLevelRank -Level $configuredLevel)) {
        return
    }

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

    $logLevelValue = if ($envMap.ContainsKey('LOG_LEVEL')) { $envMap['LOG_LEVEL'] } else { $null }
    $script:log_level = Resolve-LogLevel -Value $logLevelValue
    Write-Log ("[open-flow] Config loaded: EnvPath='{0}', LogPath='{1}', LogLevel={2}" -f `
            $Path, $script:log_path, $script:log_level) -Level DEBUG

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

    $script:date_subpath = $null
    if ($envMap.ContainsKey('DATE_SUBPATH') -and -not [string]::IsNullOrWhiteSpace($envMap['DATE_SUBPATH'])) {
        $script:date_subpath = $envMap['DATE_SUBPATH'].Trim()
    }

    $useYearMonthValue = if ($envMap.ContainsKey('USE_BRANCH_YEAR_MONTH')) { $envMap['USE_BRANCH_YEAR_MONTH'] } else { $null }
    $script:use_branch_year_month = Test-EnvBool -Value $useYearMonthValue -Default $false
    $script:branch_year = if ($envMap.ContainsKey('BRANCH_YEAR')) { $envMap['BRANCH_YEAR'] } else { $null }
    $script:branch_month = if ($envMap.ContainsKey('BRANCH_MONTH')) { $envMap['BRANCH_MONTH'] } else { $null }

    # DATE_SUBPATH takes precedence; legacy year/month trio only applies when DATE_SUBPATH is empty.
    $useLegacyYearMonth = [string]::IsNullOrWhiteSpace($script:date_subpath) -and $script:use_branch_year_month

    try {
        Assert-BranchYearMonthConfig `
            -UseBranchYearMonth $useLegacyYearMonth `
            -BranchYear $script:branch_year `
            -BranchMonth $script:branch_month
    }
    catch {
        Write-Error $_.Exception.Message
        exit 1
    }

    $subpathReferenceDate = if ($script:ref_run_date) { $script:ref_run_date } else { [DateTime]::Today }
    $subpathPeriod = $null
    if (-not [string]::IsNullOrWhiteSpace($script:date_subpath)) {
        try {
            $subpathPeriod = Get-BranchPayrollPeriod `
                -StartDay $script:branch_payroll_start_day `
                -EndDay $script:branch_payroll_end_day `
                -TargetPeriod $script:payroll_target_period `
                -ReferenceDate $subpathReferenceDate
            Write-Log ("[open-flow] DATE_SUBPATH period: Pattern='{0}', PeriodStart='{1:yyyy-MM-dd}', PeriodEnd='{2:yyyy-MM-dd}'" -f `
                    $script:date_subpath, $subpathPeriod.Start, $subpathPeriod.End) -Level DEBUG
        }
        catch {
            Write-Error $_.Exception.Message
            exit 1
        }
    }

    try {
        $script:branch_paths = @(Parse-BranchPaths -Value $envMap['BRANCH_PATHS'])
        Write-Log ("[open-flow] Parsed branch paths: Count={0}" -f $script:branch_paths.Count) -Level DEBUG
    }
    catch {
        Write-Error $_.Exception.Message
        exit 1
    }

    if ($script:branch_paths.Count -eq 0) {
        Write-Error 'BRANCH_PATHS must contain at least one Label=Path entry'
        exit 1
    }

    if (-not [string]::IsNullOrWhiteSpace($script:date_subpath)) {
        try {
            $script:branch_paths = @(
                foreach ($branch in $script:branch_paths) {
                    $resolved = Resolve-DateSubpath `
                        -BasePath $branch.Path `
                        -Pattern $script:date_subpath `
                        -Date $subpathPeriod.Start
                    Write-Log ("[open-flow] DATE_SUBPATH resolved: Label='{0}', Base='{1}', Path='{2}'" -f `
                            $branch.Label, $branch.Path, $resolved) -Level DEBUG
                    [pscustomobject]@{
                        Label    = $branch.Label
                        Path     = $resolved
                        BasePath = $branch.Path
                    }
                }
            )
            Write-Log ("[open-flow] Applied DATE_SUBPATH: Pattern='{0}', PeriodStart='{1:yyyy-MM-dd}', Count={2}" -f `
                    $script:date_subpath, $subpathPeriod.Start, $script:branch_paths.Count) -Level DEBUG
        }
        catch {
            Write-Error $_.Exception.Message
            exit 1
        }
    }
    elseif ($useLegacyYearMonth) {
        $script:branch_paths = @(
            foreach ($branch in $script:branch_paths) {
                [pscustomobject]@{
                    Label    = $branch.Label
                    Path     = Resolve-BranchFolderPath `
                        -BasePath $branch.Path `
                        -UseBranchYearMonth $true `
                        -BranchYear $script:branch_year `
                        -BranchMonth $script:branch_month
                    BasePath = $branch.Path
                }
            }
        )
        Write-Log ("[open-flow] Applied legacy branch year/month suffix: Year='{0}', Month='{1}', Count={2}" -f `
                $script:branch_year, $script:branch_month, $script:branch_paths.Count) -Level DEBUG
    }

    $script:branch_paths = @(Group-BranchPaths -BranchPaths $script:branch_paths)
    Write-Log ("[open-flow] Grouped branch paths: Count={0}" -f $script:branch_paths.Count) -Level DEBUG
    foreach ($branch in $script:branch_paths) {
        [string[]]$paths = if ($branch.PSObject.Properties['Paths'] -and $branch.Paths) {
            @($branch.Paths)
        }
        else {
            @($branch.Path)
        }
        Write-Log ("[open-flow] Grouped branch: Label='{0}', PrimaryPath='{1}', Paths='{2}'" -f `
                $branch.Label, $branch.Path, ($paths -join "'; '")) -Level DEBUG
    }

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

    $showSrButtonValue = if ($envMap.ContainsKey('SHOW_SR_BUTTON')) { $envMap['SHOW_SR_BUTTON'] } else { $null }
    $script:show_sr_button = Test-EnvBool -Value $showSrButtonValue -Default $true

    $script:final_sc_incentives_path = $null
    if ($envMap.ContainsKey('FINAL_SC_INCENTIVES_PATH') -and -not [string]::IsNullOrWhiteSpace($envMap['FINAL_SC_INCENTIVES_PATH'])) {
        $script:final_sc_incentives_path = Resolve-ConfigPath -PathValue $envMap['FINAL_SC_INCENTIVES_PATH'].Trim()
        Write-Log ("[open-final-sc] Configured path: Path='{0}'" -f $script:final_sc_incentives_path) -Level DEBUG
    }

    $script:sr_open_max = 10
    if ($envMap.ContainsKey('SR_OPEN_MAX') -and -not [string]::IsNullOrWhiteSpace($envMap['SR_OPEN_MAX'])) {
        $parsedSrOpenMax = 0
        if (-not [int]::TryParse($envMap['SR_OPEN_MAX'].Trim(), [ref]$parsedSrOpenMax) -or $parsedSrOpenMax -lt 1) {
            Write-Error 'SR_OPEN_MAX must be a positive integer'
            exit 1
        }
        $script:sr_open_max = $parsedSrOpenMax
    }

    if ($envMap.ContainsKey('SR_FILE_EXTENSIONS') -and -not [string]::IsNullOrWhiteSpace($envMap['SR_FILE_EXTENSIONS'])) {
        try {
            $script:sr_file_extensions = @(Parse-SrFileExtensions -Value $envMap['SR_FILE_EXTENSIONS'])
        }
        catch {
            Write-Error $_.Exception.Message
            exit 1
        }
    }

    $script:sr_paths = @()
    if ($envMap.ContainsKey('SR_PATHS') -and -not [string]::IsNullOrWhiteSpace($envMap['SR_PATHS'])) {
        try {
            $script:sr_paths = @(Parse-SrPaths -Value $envMap['SR_PATHS'])
            Write-Log ("[open-sr] Parsed SR paths: Count={0}" -f $script:sr_paths.Count) -Level DEBUG
        }
        catch {
            Write-Error $_.Exception.Message
            exit 1
        }

        if (-not [string]::IsNullOrWhiteSpace($script:date_subpath)) {
            try {
                $script:sr_paths = @(
                    foreach ($branch in $script:sr_paths) {
                        $resolved = Resolve-DateSubpath `
                            -BasePath $branch.Path `
                            -Pattern $script:date_subpath `
                            -Date $subpathPeriod.Start
                        Write-Log ("[open-sr] DATE_SUBPATH resolved: Label='{0}', Base='{1}', Path='{2}'" -f `
                                $branch.Label, $branch.Path, $resolved) -Level DEBUG
                        [pscustomobject]@{
                            Label    = $branch.Label
                            Path     = $resolved
                            BasePath = $branch.Path
                        }
                    }
                )
                Write-Log ("[open-sr] Applied DATE_SUBPATH: Pattern='{0}', PeriodStart='{1:yyyy-MM-dd}', Count={2}" -f `
                        $script:date_subpath, $subpathPeriod.Start, $script:sr_paths.Count) -Level DEBUG
            }
            catch {
                Write-Error $_.Exception.Message
                exit 1
            }
        }
        elseif ($useLegacyYearMonth) {
            $script:sr_paths = @(
                foreach ($branch in $script:sr_paths) {
                    [pscustomobject]@{
                        Label    = $branch.Label
                        Path     = Resolve-SrBranchFolderPath `
                            -BasePath $branch.Path `
                            -UseBranchYearMonth $true `
                            -BranchYear $script:branch_year `
                            -BranchMonth $script:branch_month
                        BasePath = $branch.Path
                    }
                }
            )
            Write-Log ("[open-sr] Applied legacy SR year/month resolution: Year='{0}', Month='{1}', Count={2}" -f `
                    $script:branch_year, $script:branch_month, $script:sr_paths.Count) -Level DEBUG
            foreach ($branch in $script:sr_paths) {
                Write-Log ("[open-sr] Resolved SR path: Label='{0}', Path='{1}'" -f $branch.Label, $branch.Path) -Level DEBUG
            }
        }

        $script:sr_paths = @(Group-BranchPaths -BranchPaths $script:sr_paths)
        Write-Log ("[open-sr] Grouped SR paths: Count={0}" -f $script:sr_paths.Count) -Level DEBUG
    }

    Write-Log ("[open-flow] Config scan options: RefreshIntervalSeconds={0}, Recursive={1}, Depth={2}, Extensions='{3}', ShowSrButton={4}, SrOpenMax={5}, SrExtensions='{6}', SrPathCount={7}, FinalScPath='{8}'" -f `
            $script:refresh_interval_seconds, $script:recursive, $script:dir_level_search, ($script:cutoff_file_extensions -join ','), `
            $script:show_sr_button, $script:sr_open_max, ($script:sr_file_extensions -join ','), $script:sr_paths.Count, $script:final_sc_incentives_path) -Level DEBUG
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
    Write-Log ("[open-flow] Display build: ReferenceDate='{0:yyyy-MM-dd}', TargetPeriod='{1}', PeriodStart='{2:yyyy-MM-dd}', PeriodEnd='{3:yyyy-MM-dd}', DayRange='{4}'" -f `
            $referenceDate, $script:payroll_target_period, $period.Start, $period.End, $dayRangeLabel) -Level DEBUG

    $branches_with_cutoff_files = Get-BranchesWithCutoffFiles `
        -BranchPaths $script:branch_paths `
        -StartDay $startDay `
        -EndDay $endDay `
        -PeriodStart $period.Start `
        -PeriodEnd $period.End `
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
