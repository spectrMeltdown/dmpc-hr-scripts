#Requires -Version 5.1

param(
    [string]$EnvPath = (Join-Path $PSScriptRoot ".env")
)

$file_path = $null
$folder_path = $null
$recursive = $false
$sheet_name = $null
$sheet_range = @()
$log_path = $null
$TargetCells = @()

$ExternalRefPattern = "('[^']+'![A-Z]{1,3})(\d+)"

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
        [string]$Path
    )

    $env = Import-DotEnv -Path $Path
    $requiredKeys = @('SHEET_NAME', 'SHEET_RANGE', 'LOG_PATH')

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
    if ([string]::IsNullOrWhiteSpace($recursiveValue)) {
        $script:recursive = $false
    }
    else {
        $script:recursive = $recursiveValue.Trim().ToLowerInvariant() -in @('true', '1', 'yes')
    }

    $script:sheet_name = $env['SHEET_NAME']

    $rangeParts = ($env['SHEET_RANGE'] -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    if ($rangeParts.Count -lt 2) {
        Write-Error "SHEET_RANGE must contain two comma-separated integers, e.g. 3,5"
        exit 1
    }

    try {
        $script:sheet_range = @([int]$rangeParts[0], [int]$rangeParts[1])
    }
    catch {
        Write-Error "SHEET_RANGE values must be integers: $($env['SHEET_RANGE'])"
        exit 1
    }

    $script:log_path = Resolve-ConfigPath -PathValue $env['LOG_PATH']

    $cellGroups = Import-TargetCellGroups -Env $env
    $script:TargetCells = @($cellGroups | ForEach-Object { $_.Cells })
    if ($script:TargetCells.Count -eq 0) {
        Write-Error "At least one TARGET_CELLS_N group must contain a cell address"
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

function Increment-ExternalRefRows {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Formula,

        [Parameter(Mandatory = $true)]
        [int]$Increment
    )

    if ($Increment -eq 0) {
        return $Formula
    }

    if ($Formula -notmatch $ExternalRefPattern) {
        return $Formula
    }

    return [regex]::Replace($Formula, $ExternalRefPattern, {
        param($match)
        $prefix = $match.Groups[1].Value
        $row = [int]$match.Groups[2].Value
        "$prefix$($row + $Increment)"
    })
}

function Update-TargetCellFormulas {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $true)]
        [int]$Increment,

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

        $newFormula = Increment-ExternalRefRows -Formula $formula -Increment $Increment

        if ($newFormula -eq $formula) {
            Write-Log "Sheet $SheetLabel cell ${address}: no external ref row to increment, skipped" -Level WARN
            $skipped++
            continue
        }

        $cell.Formula = $newFormula
        $updated++
    }

    return @{
        Updated = $updated
        Skipped = $skipped
    }
}

function Get-WorksheetByName {
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    foreach ($sheet in @($Workbook.Worksheets)) {
        if ($sheet.Name -eq $Name) {
            return $sheet
        }
    }

    return $null
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

function Invoke-PayrollSheetUpdate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookPath
    )

    Write-Log "Starting payroll sheet update"
    Write-Log "workbook_path=$WorkbookPath sheet_name=$sheet_name sheet_range=[$($sheet_range[0]), $($sheet_range[1])]"

    if ($sheet_range.Count -lt 2) {
        Write-Log "sheet_range must contain start and stop values, e.g. @(3, 5)" -Level ERROR
        throw "sheet_range must contain start and stop values"
    }

    $rangeStart = [int]$sheet_range[0]
    $rangeStop = [int]$sheet_range[1]

    if ($rangeStart -gt $rangeStop) {
        Write-Log "sheet_range start ($rangeStart) must be <= stop ($rangeStop)" -Level ERROR
        throw "sheet_range start ($rangeStart) must be <= stop ($rangeStop)"
    }

    $templateName = [string]$sheet_name
    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $templateSheet = Get-WorksheetByName -Workbook $workbook -Name $templateName

        if (-not $templateSheet) {
            Write-Log "Template sheet '$templateName' not found in workbook" -Level ERROR
            throw "Template sheet '$templateName' not found in workbook"
        }

        for ($n = $rangeStart; $n -le $rangeStop; $n++) {
            $targetName = [string]$n
            $existingSheet = Get-WorksheetByName -Workbook $workbook -Name $targetName

            if ($existingSheet) {
                $result = Update-TargetCellFormulas -Worksheet $existingSheet -Increment 1 -SheetLabel $targetName
                Write-Log "Updated sheet $targetName (+1 row offset on $($result.Updated) cells, $($result.Skipped) skipped)"
            }
            else {
                $insertAfterSheet = $templateSheet
                $predecessorSheet = Get-WorksheetByName -Workbook $workbook -Name ([string]($n - 1))
                if ($predecessorSheet) {
                    $insertAfterSheet = $predecessorSheet
                }

                $insertAfterSheet.Copy([Type]::Missing, $insertAfterSheet)
                $newSheet = $insertAfterSheet.Next
                $newSheet.Name = $targetName

                $offset = $n - [int]$sheet_name
                $result = Update-TargetCellFormulas -Worksheet $newSheet -Increment $offset -SheetLabel $targetName
                Write-Log "Created sheet $targetName from template $templateName (+$offset row offset on $($result.Updated) cells, $($result.Skipped) skipped)"
            }
        }

        $workbook.Save()
        Write-Log "Workbook saved successfully"
        Write-Log "Payroll sheet update completed for $WorkbookPath"
    }
    catch {
        Write-Log "Unexpected error: $($_.Exception.Message)" -Level ERROR
        throw
    }
    finally {
        if ($workbook) {
            $workbook.Close($true) | Out-Null
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

Initialize-Config -Path $EnvPath

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
        Invoke-PayrollSheetUpdate -WorkbookPath $path
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

Write-Log "All workbook updates completed successfully"
