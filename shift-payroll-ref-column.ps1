#Requires -Version 5.1

param(
    [string]$EnvPath = (Join-Path $PSScriptRoot ".env"),

    [switch]$DryRun
)

$file_path = $null
$folder_path = $null
$recursive = $false
$log_path = $null
$log_level = 'INFO'
$TargetCells = @()
$onlyNumSheets = $false
$dryRun = $false
$incrementMode = $true
$columnDelta = 1

# Excel quoted sheet names escape apostrophes as '' inside the quotes
$ExternalRefPattern = "('(?:''|[^'])*'!)(`$?)([A-Z]{1,3})(`$?)(\d+)"

function Import-DotEnv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Error "Env file not found: $Path. Copy .env.example to .env and edit values."
        exit 1
    }

    $env = @{}

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

        $env[$key] = $value
    }

    return $env
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

function Import-TargetCellGroups {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Env
    )

    $groups = @()
    $n = 1

    while ($Env.ContainsKey("TARGET_CELLS_$n")) {
        $cellsKey = "TARGET_CELLS_$n"
        $cellsValue = $Env[$cellsKey]
        if ([string]::IsNullOrWhiteSpace($cellsValue)) {
            Write-Error "Missing or empty required env key: $cellsKey"
            exit 1
        }

        $cells = ($cellsValue -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($cells.Count -eq 0) {
            Write-Error "$cellsKey must contain at least one cell address"
            exit 1
        }

        $groups += @{ Cells = $cells }
        $n++
    }

    if ($groups.Count -eq 0) {
        Write-Error "At least one TARGET_CELLS_N group required"
        exit 1
    }

    return $groups
}

function Initialize-Config {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$DryRunSwitch
    )

    $env = Import-DotEnv -Path $Path
    $requiredKeys = @('LOG_PATH')

    foreach ($key in $requiredKeys) {
        if (-not $env.ContainsKey($key) -or [string]::IsNullOrWhiteSpace($env[$key])) {
            Write-Error "Missing or empty required env key: $key"
            exit 1
        }
    }

    $cellGroups = Import-TargetCellGroups -Env $env
    $script:TargetCells = @()
    foreach ($group in $cellGroups) {
        $script:TargetCells += $group.Cells
    }

    if ($script:TargetCells.Count -eq 0) {
        Write-Error "At least one TARGET_CELLS_N group must contain a cell address"
        exit 1
    }

    $folderPathValue = if ($env.ContainsKey('FOLDER_PATH')) { $env['FOLDER_PATH'] } else { $null }
    $filePathValue = if ($env.ContainsKey('FILE_PATH')) { $env['FILE_PATH'] } else { $null }

    $script:folder_path = Resolve-ConfigPath -PathValue $folderPathValue
    $script:file_path = Resolve-ConfigPath -PathValue $filePathValue

    if ([string]::IsNullOrWhiteSpace($script:folder_path) -and [string]::IsNullOrWhiteSpace($script:file_path)) {
        Write-Error "At least one of FOLDER_PATH or FILE_PATH must be set"
        exit 1
    }

    $recursiveValue = if ($env.ContainsKey('RECURSIVE')) { $env['RECURSIVE'] } else { $null }
    $script:recursive = Test-EnvBool -Value $recursiveValue

    $onlyNumSheetsValue = if ($env.ContainsKey('ONLY_NUM_SHEETS')) { $env['ONLY_NUM_SHEETS'] } else { $null }
    $script:onlyNumSheets = Test-EnvBool -Value $onlyNumSheetsValue

    $dryRunValue = if ($env.ContainsKey('DRY_RUN')) { $env['DRY_RUN'] } else { $null }
    $script:dryRun = Test-EnvBool -Value $dryRunValue
    if ($DryRunSwitch) {
        $script:dryRun = $true
    }

    $incrementModeValue = if ($env.ContainsKey('INCREMENT_MODE')) { $env['INCREMENT_MODE'] } else { 'true' }
    $script:incrementMode = Test-EnvBool -Value $incrementModeValue -Default $true
    $script:columnDelta = if ($script:incrementMode) { 1 } else { -1 }

    $script:log_path = Resolve-ConfigPath -PathValue $env['LOG_PATH']

    $logLevelValue = if ($env.ContainsKey('LOG_LEVEL')) { $env['LOG_LEVEL'] } else { $null }
    $script:log_level = Resolve-LogLevel -Value $logLevelValue
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

    $logDir = Split-Path -Parent $log_path
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    Add-Content -LiteralPath $log_path -Value $line -Encoding UTF8
}

function Test-IntegerSheetName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    return $Name -match '^\d+$'
}

function Get-TargetWorksheets {
    param(
        [Parameter(Mandatory = $true)]
        $Workbook
    )

    $sheets = @()
    foreach ($sheet in @($Workbook.Worksheets)) {
        if ($sheet.Visible -ne -1) {
            continue
        }

        if (-not $onlyNumSheets -or (Test-IntegerSheetName -Name $sheet.Name)) {
            $sheets += $sheet
        }
    }

    return $sheets
}

function Convert-ColumnLettersToNumber {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Letters
    )

    $num = 0
    foreach ($char in $Letters.ToUpperInvariant().ToCharArray()) {
        $num = $num * 26 + ([int][char]$char - [int][char]'A' + 1)
    }

    return $num
}

function Convert-NumberToColumnLetters {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Number
    )

    if ($Number -lt 1) {
        return $null
    }

    $result = ''
    $n = $Number

    while ($n -gt 0) {
        $remainder = ($n - 1) % 26
        $result = [char]([int][char]'A' + $remainder) + $result
        $n = [math]::Floor(($n - 1) / 26)
    }

    return $result
}

function Shift-ExternalRefColumns {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Formula,

        [Parameter(Mandatory = $true)]
        [int]$Delta
    )

    if ($Delta -eq 0) {
        return $Formula
    }

    if ($Formula -notmatch $ExternalRefPattern) {
        return $Formula
    }

    return [regex]::Replace($Formula, $ExternalRefPattern, {
        param($match)

        $prefix = $match.Groups[1].Value
        $colDollar = $match.Groups[2].Value
        $column = $match.Groups[3].Value
        $rowDollar = $match.Groups[4].Value
        $row = $match.Groups[5].Value

        $colNum = Convert-ColumnLettersToNumber -Letters $column
        $newColNum = $colNum + $Delta

        if ($newColNum -lt 1) {
            return $match.Value
        }

        $newColumn = Convert-NumberToColumnLetters -Number $newColNum
        return "$prefix$colDollar$newColumn$rowDollar$row"
    })
}

function Update-TargetCellFormulas {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $true)]
        [string]$SheetLabel
    )

    $updated = 0
    $skipped = 0

    foreach ($address in $TargetCells) {
        $cell = $Worksheet.Range($address)
        $formula = $cell.Formula

        if (-not $formula -or $formula -notlike '=*') {
            Write-Log "Sheet $SheetLabel cell ${address}: not a formula, skipped" -Level DEBUG
            $skipped++
            continue
        }

        if ($formula -notmatch $ExternalRefPattern) {
            Write-Log "Sheet $SheetLabel cell ${address}: no external sheet reference in formula, skipped" -Level DEBUG
            $skipped++
            continue
        }

        $newFormula = Shift-ExternalRefColumns -Formula $formula -Delta $columnDelta

        if ($newFormula -eq $formula) {
            Write-Log "Sheet $SheetLabel cell ${address}: column shift produced no change, skipped" -Level DEBUG
            $skipped++
            continue
        }

        if ($dryRun) {
            Write-Log "Sheet $SheetLabel cell ${address}: WOULD UPDATE old=$formula new=$newFormula" -Level DEBUG
        }
        else {
            $cell.Formula = $newFormula
            Write-Log "Sheet $SheetLabel cell ${address}: updated old=$formula new=$newFormula" -Level DEBUG
        }

        $updated++
    }

    return @{
        Updated = $updated
        Skipped = $skipped
    }
}

function Get-PayrollWorkbookPaths {
    if (-not [string]::IsNullOrWhiteSpace($folder_path)) {
        if (-not (Test-Path -LiteralPath $folder_path -PathType Container)) {
            throw "Folder not found: $folder_path"
        }

        $childParams = @{
            LiteralPath = $folder_path
            Filter      = '*.xlsx'
            File        = $true
        }
        if ($recursive) {
            $childParams['Recurse'] = $true
        }

        $paths = @(Get-ChildItem @childParams |
            Where-Object { -not $_.Name.StartsWith('~$') } |
            ForEach-Object { $_.FullName } |
            Sort-Object)

        if ($paths.Count -eq 0) {
            throw "No .xlsx files found in folder: $folder_path"
        }

        return $paths
    }

    if (-not (Test-Path -LiteralPath $file_path -PathType Leaf)) {
        throw "File not found: $file_path"
    }

    return @($file_path)
}

function Invoke-ShiftPayrollRefColumn {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookPath
    )

    Write-Log "Starting shift payroll reference column letters"
    Write-Log "workbook_path=$WorkbookPath target_cells=$($TargetCells.Count) increment_mode=$incrementMode column_delta=$columnDelta only_num_sheets=$onlyNumSheets dry_run=$dryRun" -Level DEBUG

    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($WorkbookPath)

        $targetSheets = Get-TargetWorksheets -Workbook $workbook
        if ($targetSheets.Count -eq 0) {
            Write-Log "No worksheets matched ONLY_NUM_SHEETS=$onlyNumSheets filter" -Level WARNING
        }

        $totalUpdated = 0
        $totalSkipped = 0

        foreach ($sheet in $targetSheets) {
            $sheetLabel = $sheet.Name
            Write-Log "Processing sheet '$sheetLabel'" -Level DEBUG

            $result = Update-TargetCellFormulas -Worksheet $sheet -SheetLabel $sheetLabel
            $totalUpdated += $result.Updated
            $totalSkipped += $result.Skipped
            Write-Log "Sheet '$sheetLabel': $($result.Updated) updated, $($result.Skipped) skipped"
        }

        if ($dryRun) {
            Write-Log "Dry run: workbook not saved ($totalUpdated would-be updates, $totalSkipped skipped)"
        }
        else {
            $workbook.Save()
            Write-Log "Workbook saved successfully ($totalUpdated updated, $totalSkipped skipped)"
        }

        Write-Log "Shift payroll reference column letters completed for $WorkbookPath"
    }
    catch {
        Write-Log "Unexpected error: $($_.Exception.Message)" -Level ERROR
        throw
    }
    finally {
        if ($workbook) {
            $workbook.Close(-not $dryRun) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }

        if ($excel) {
            $excel.Quit() | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Initialize-Config -Path $EnvPath -DryRunSwitch:$DryRun

try {
    $workbookPaths = Get-PayrollWorkbookPaths
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}

Write-Log "Processing $($workbookPaths.Count) workbook(s)"
if (-not [string]::IsNullOrWhiteSpace($folder_path)) {
    Write-Log "folder_path=$folder_path recursive=$recursive" -Level DEBUG
}

$failures = @()
foreach ($path in $workbookPaths) {
    Write-Log "Processing workbook: $path"
    try {
        Invoke-ShiftPayrollRefColumn -WorkbookPath $path
    }
    catch {
        Write-Log "Failed workbook ${path}: $($_.Exception.Message)" -Level ERROR
        $failures += $path
    }
}

if ($failures.Count -gt 0) {
    Write-Log "Completed with $($failures.Count) failure(s)" -Level ERROR
    exit 1
}

Write-Log "All workbook column reference shifts completed successfully"
