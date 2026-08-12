#Requires -Version 5.1

param(
    [string]$EnvPath = (Join-Path $PSScriptRoot ".env")
)

$file_path = $null
$folder_path = $null
$recursive = $false
$sheet_range = @()
$sheet_password = $null
$log_path = $null
$log_level = 'INFO'

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
        [string]$Path
    )

    $env = Import-DotEnv -Path $Path
    $requiredKeys = @('SHEET_RANGE', 'SHEET_PASSWORD', 'LOG_PATH')

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

    $script:sheet_password = $env['SHEET_PASSWORD']
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

function Invoke-UnlockSheets {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookPath
    )

    Write-Log "Starting sheet unlock"
    Write-Log "workbook_path=$WorkbookPath sheet_range=[$($sheet_range[0]), $($sheet_range[1])]" -Level DEBUG

    $rangeStart = [int]$sheet_range[0]
    $rangeStop = [int]$sheet_range[1]

    if ($rangeStart -gt $rangeStop) {
        Write-Log "sheet_range start ($rangeStart) must be <= stop ($rangeStop)" -Level ERROR
        throw "sheet_range start ($rangeStart) must be <= stop ($rangeStop)"
    }

    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        # UpdateLinks=0, ReadOnly=$false, IgnoreReadOnlyRecommended=$true
        $workbook = $excel.Workbooks.Open(
            $WorkbookPath,
            0,
            $false,
            [Type]::Missing,
            [Type]::Missing,
            [Type]::Missing,
            $true
        )

        if ($workbook.ReadOnly) {
            throw "Workbook opened as read-only; cannot unlock and save: $WorkbookPath"
        }

        $unlocked = 0
        $skipped = 0

        for ($n = $rangeStart; $n -le $rangeStop; $n++) {
            $sheetName = [string]$n
            $sheet = Get-WorksheetByName -Workbook $workbook -Name $sheetName

            if (-not $sheet) {
                Write-Log "Sheet '$sheetName' not found, skipped" -Level WARNING
                $skipped++
                continue
            }

            if (-not $sheet.ProtectContents) {
                Write-Log "Sheet '$sheetName' not protected, skipped" -Level WARNING
                $skipped++
                continue
            }

            $sheet.Unprotect($sheet_password) | Out-Null
            if ($sheet.ProtectContents) {
                $msg = "Sheet '$sheetName' still protected after Unprotect (wrong SHEET_PASSWORD?)"
                Write-Log $msg -Level ERROR
                throw $msg
            }

            Write-Log "Sheet '$sheetName' unlocked"
            $unlocked++
        }

        $workbook.Save()
        Write-Log "Workbook saved successfully ($unlocked unlocked, $skipped skipped)"
        Write-Log "Sheet unlock completed for $WorkbookPath"
    }
    catch {
        Write-Log "Unexpected error: $($_.Exception.Message)" -Level ERROR
        throw
    }
    finally {
        if ($workbook) {
            $workbook.Close($false) | Out-Null
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
    Write-Log "folder_path=$folder_path recursive=$recursive" -Level DEBUG
}

$failures = @()
foreach ($path in $workbookPaths) {
    Write-Log "Processing workbook: $path"
    try {
        Invoke-UnlockSheets -WorkbookPath $path
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

Write-Log "All sheet unlocks completed successfully"
