#Requires -Version 5.1

param(
    [string]$EnvPath = (Join-Path $PSScriptRoot ".env"),

    [switch]$DryRun
)

$file_path = $null
$folder_path = $null
$recursive = $false
$log_path = $null
$TargetCells = @()
$sheetRefIndex = 1
$onlyNumSheets = $false
$dryRun = $false

$BracketedFileRefPattern = "'[^']*\[[^\]]+\][^']+'!"

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

function Initialize-Config {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$DryRunSwitch
    )

    $env = Import-DotEnv -Path $Path
    $requiredKeys = @('LOG_PATH', 'TARGET_CELLS')

    foreach ($key in $requiredKeys) {
        if (-not $env.ContainsKey($key) -or [string]::IsNullOrWhiteSpace($env[$key])) {
            Write-Error "Missing or empty required env key: $key"
            exit 1
        }
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

    $refIndexValue = if ($env.ContainsKey('SHEET_REF_INDEX')) { $env['SHEET_REF_INDEX'] } else { '1' }
    try {
        $script:sheetRefIndex = [int]$refIndexValue
    }
    catch {
        Write-Error "SHEET_REF_INDEX must be an integer: $refIndexValue"
        exit 1
    }

    if ($script:sheetRefIndex -lt 1) {
        Write-Error "SHEET_REF_INDEX must be >= 1"
        exit 1
    }

    $onlyNumSheetsValue = if ($env.ContainsKey('ONLY_NUM_SHEETS')) { $env['ONLY_NUM_SHEETS'] } else { $null }
    $script:onlyNumSheets = Test-EnvBool -Value $onlyNumSheetsValue

    $dryRunValue = if ($env.ContainsKey('DRY_RUN')) { $env['DRY_RUN'] } else { $null }
    $script:dryRun = Test-EnvBool -Value $dryRunValue
    if ($DryRunSwitch) {
        $script:dryRun = $true
    }

    $script:log_path = Resolve-ConfigPath -PathValue $env['LOG_PATH']

    $script:TargetCells = ($env['TARGET_CELLS'] -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    if ($script:TargetCells.Count -eq 0) {
        Write-Error "TARGET_CELLS must contain at least one cell address"
        exit 1
    }
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

function Get-RefSheetName {
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,

        [Parameter(Mandatory = $true)]
        [int]$Index
    )

    $visibleSheets = @()
    foreach ($sheet in @($Workbook.Worksheets)) {
        if ($sheet.Visible -eq -1) {
            $visibleSheets += $sheet
        }
    }

    $visibleCount = $visibleSheets.Count
    if ($Index -lt 1 -or $Index -gt $visibleCount) {
        throw "SHEET_REF_INDEX ($Index) is out of range; workbook has $visibleCount visible sheet(s)"
    }

    return [string]$visibleSheets[$Index - 1].Name
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

function Fix-BracketedFileReferences {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Formula,

        [Parameter(Mandatory = $true)]
        [string]$RefSheetName
    )

    if ($Formula -notmatch $BracketedFileRefPattern) {
        return $Formula
    }

    $escapedName = $RefSheetName -replace "'", "''"
    $replacement = "'$escapedName'!"

    return [regex]::Replace($Formula, $BracketedFileRefPattern, $replacement)
}

function Update-TargetCellFormulas {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $true)]
        [string]$RefSheetName,

        [Parameter(Mandatory = $true)]
        [string]$SheetLabel
    )

    $updated = 0
    $skipped = 0

    foreach ($address in $TargetCells) {
        $cell = $Worksheet.Range($address)
        $formula = $cell.Formula

        if (-not $formula -or $formula -notlike '=*') {
            Write-Log "Sheet $SheetLabel cell ${address}: not a formula, skipped" -Level WARN
            $skipped++
            continue
        }

        $newFormula = Fix-BracketedFileReferences -Formula $formula -RefSheetName $RefSheetName

        if ($newFormula -eq $formula) {
            Write-Log "Sheet $SheetLabel cell ${address}: no bracketed file reference, skipped" -Level WARN
            $skipped++
            continue
        }

        if ($dryRun) {
            Write-Log "Sheet $SheetLabel cell ${address}: WOULD UPDATE old=$formula new=$newFormula"
        }
        else {
            $cell.Formula = $newFormula
            Write-Log "Sheet $SheetLabel cell ${address}: updated old=$formula new=$newFormula"
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

function Invoke-FixCopiedPayrollSheet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookPath
    )

    Write-Log "Starting fix copied payroll sheet references"
    Write-Log "workbook_path=$WorkbookPath sheet_ref_index=$sheetRefIndex only_num_sheets=$onlyNumSheets dry_run=$dryRun"

    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $refSheetName = Get-RefSheetName -Workbook $workbook -Index $sheetRefIndex
        Write-Log "Reference sheet index $sheetRefIndex name='$refSheetName'"

        $targetSheets = Get-TargetWorksheets -Workbook $workbook
        if ($targetSheets.Count -eq 0) {
            Write-Log "No worksheets matched ONLY_NUM_SHEETS=$onlyNumSheets filter" -Level WARN
        }

        $totalUpdated = 0
        $totalSkipped = 0

        foreach ($sheet in $targetSheets) {
            $sheetLabel = $sheet.Name
            Write-Log "Processing sheet '$sheetLabel'"
            $result = Update-TargetCellFormulas -Worksheet $sheet -RefSheetName $refSheetName -SheetLabel $sheetLabel
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

        Write-Log "Fix copied payroll sheet references completed for $WorkbookPath"
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
    Write-Log "folder_path=$folder_path recursive=$recursive"
}

$failures = @()
foreach ($path in $workbookPaths) {
    Write-Log "Processing workbook: $path"
    try {
        Invoke-FixCopiedPayrollSheet -WorkbookPath $path
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

Write-Log "All workbook reference fixes completed successfully"
