#Requires -Version 5.1

param(
    [string]$EnvPath = (Join-Path $PSScriptRoot ".env")
)

$file_path = $null
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

function Initialize-Config {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $env = Import-DotEnv -Path $Path
    $requiredKeys = @('FILE_PATH', 'SHEET_NAME', 'SHEET_RANGE', 'LOG_PATH', 'TARGET_CELLS')

    foreach ($key in $requiredKeys) {
        if (-not $env.ContainsKey($key) -or [string]::IsNullOrWhiteSpace($env[$key])) {
            Write-Error "Missing or empty required env key: $key"
            exit 1
        }
    }

    $script:file_path = $env['FILE_PATH']
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

    $logPathValue = $env['LOG_PATH']
    if ([System.IO.Path]::IsPathRooted($logPathValue)) {
        $script:log_path = $logPathValue
    }
    else {
        $script:log_path = Join-Path $PSScriptRoot $logPathValue
    }

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

function Invoke-PayrollSheetUpdate {
    Write-Log "Starting payroll sheet update"
    Write-Log "file_path=$file_path sheet_name=$sheet_name sheet_range=[$($sheet_range[0]), $($sheet_range[1])]"

    if (-not (Test-Path -LiteralPath $file_path)) {
        Write-Log "File not found: $file_path" -Level ERROR
        exit 1
    }

    if ($sheet_range.Count -lt 2) {
        Write-Log "sheet_range must contain start and stop values, e.g. @(3, 5)" -Level ERROR
        exit 1
    }

    $rangeStart = [int]$sheet_range[0]
    $rangeStop = [int]$sheet_range[1]

    if ($rangeStart -gt $rangeStop) {
        Write-Log "sheet_range start ($rangeStart) must be <= stop ($rangeStop)" -Level ERROR
        exit 1
    }

    $templateName = [string]$sheet_name
    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($file_path)
        $templateSheet = Get-WorksheetByName -Workbook $workbook -Name $templateName

        if (-not $templateSheet) {
            Write-Log "Template sheet '$templateName' not found in workbook" -Level ERROR
            exit 1
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
        Write-Log "Payroll sheet update completed"
    }
    catch {
        Write-Log "Unexpected error: $($_.Exception.Message)" -Level ERROR
        exit 1
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
Invoke-PayrollSheetUpdate
