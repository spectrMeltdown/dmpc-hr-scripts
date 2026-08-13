#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'check-files-cutoff.ps1')

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-PeriodRange {
    param(
        [hashtable]$Period,
        [DateTime]$ExpectedStart,
        [DateTime]$ExpectedEnd
    )
    if ($Period.Start.Date -ne $ExpectedStart.Date -or $Period.End.Date -ne $ExpectedEnd.Date) {
        throw ("Period mismatch: got {0:yyyy-MM-dd}..{1:yyyy-MM-dd}, expected {2:yyyy-MM-dd}..{3:yyyy-MM-dd}" -f `
                $Period.Start, $Period.End, $ExpectedStart, $ExpectedEnd)
    }
}

# --- Period math (aligned with dmpc-payroll) ---

$wedSpan = Get-PayrollPeriodSpanDays -StartDow ([DayOfWeek]::Wednesday) -EndDow ([DayOfWeek]::Tuesday)
Assert-True ($wedSpan -eq 7) "Wed/Tue span expected 7, got $wedSpan"

$refMay10 = [DateTime]::ParseExact('2026-05-10', 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
$wedNext = Get-BranchPayrollPeriod -StartDay Wednesday -EndDay Tuesday -TargetPeriod next -ReferenceDate $refMay10
Assert-PeriodRange $wedNext ([DateTime]'2026-05-13') ([DateTime]'2026-05-19')

$refMay12 = [DateTime]::ParseExact('2026-05-12', 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
$wedCurrent = Get-BranchPayrollPeriod -StartDay Wednesday -EndDay Tuesday -TargetPeriod current -ReferenceDate $refMay12
Assert-PeriodRange $wedCurrent ([DateTime]'2026-05-06') ([DateTime]'2026-05-12')

$refJul12 = [DateTime]::ParseExact('2026-07-12', 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
$julyNext = Get-BranchPayrollPeriod -StartDay Wednesday -EndDay Tuesday -TargetPeriod next -ReferenceDate $refJul12
Assert-PeriodRange $julyNext ([DateTime]'2026-07-15') ([DateTime]'2026-07-21')

# --- Day-range filename match ---

Assert-True (Test-FilenameContainsDayRange -FileName 'JULY 8-14, 2026.xlsx' -StartDay 8 -EndDay 14) `
    'Expected match for JULY 8-14, 2026.xlsx'
Assert-True (Test-FilenameContainsDayRange -FileName 'JULY.08-14,2026.xlsx' -StartDay 8 -EndDay 14) `
    'Expected match for JULY.08-14,2026.xlsx'
Assert-True (-not (Test-FilenameContainsDayRange -FileName 'JULY 15-21, 2026.xlsx' -StartDay 8 -EndDay 14)) `
    'Did not expect match for JULY 15-21 against 8-14'
Assert-True (-not (Test-FilenameContainsDayRange -FileName 'JULY 18-140, 2026.xlsx' -StartDay 8 -EndDay 14)) `
    'Did not expect false-positive on 18-140'

# --- Sales report day-token match ---

$crossMonthDays = @(Get-PayrollPeriodDayNumbers -PeriodStart ([DateTime]'2026-07-29') -PeriodEnd ([DateTime]'2026-08-05'))
Assert-True ($crossMonthDays.Count -eq 8) "Expected 8 cross-month days, got $($crossMonthDays.Count)"
Assert-True ($crossMonthDays -contains 29 -and $crossMonthDays -contains 5) 'Cross-month day set should include 29 and 5'

Assert-True (Test-FilenameContainsSalesReportDay -FileName 'SALES REPORT AUG 2 2026.pdf' -PeriodDays @(2)) `
    'Expected match for space-delimited day 2'
Assert-True (Test-FilenameContainsSalesReportDay -FileName 'SR-AUG-31-2026.pdf' -PeriodDays @(31)) `
    'Expected match for dash-delimited day 31'
Assert-True (Test-FilenameContainsSalesReportDay -FileName 'SR_8_2026.xlsx' -PeriodDays @(8)) `
    'Expected match for underscore-delimited day 8'
Assert-True (-not (Test-FilenameContainsSalesReportDay -FileName 'SR_18_2026.pdf' -PeriodDays @(8))) `
    'Did not expect embedded day 8 in 18'
Assert-True (-not (Test-FilenameContainsSalesReportDay -FileName 'version12.pdf' -PeriodDays @(2))) `
    'Did not expect day 2 embedded in version12'

Assert-True (Test-EnvBool -Value 'true' -Default $false) 'SHOW_SR_BUTTON=true should parse as true'
Assert-True (-not (Test-EnvBool -Value 'false' -Default $true)) 'SHOW_SR_BUTTON=false should parse as false'

# --- BRANCH_PATHS parse ---

$parsed = @(Parse-BranchPaths -Value 'Pandesalan=C:\branch\a;Mindoro=C:\branch\b')
Assert-True ($parsed.Count -eq 2) "Expected 2 branches, got $($parsed.Count)"
Assert-True ($parsed[0].Label -eq 'Pandesalan') "Expected label Pandesalan, got $($parsed[0].Label)"
Assert-True ($parsed[0].Path -eq 'C:\branch\a') "Expected path C:\branch\a, got $($parsed[0].Path)"
Assert-True ($parsed[1].Label -eq 'Mindoro') "Expected label Mindoro, got $($parsed[1].Label)"
Assert-True ($parsed[1].Path -eq 'C:\branch\b') "Expected path C:\branch\b, got $($parsed[1].Path)"

$empty = @(Parse-BranchPaths -Value '')
Assert-True ($empty.Count -eq 0) 'Empty BRANCH_PATHS should parse to zero entries'

$threw = $false
try {
    [void](Parse-BranchPaths -Value 'NoEqualsHere;Mindoro=C:\ok')
}
catch {
    $threw = $true
}
Assert-True $threw 'Expected Parse-BranchPaths to throw on invalid segment'

# --- Branch year/month suffix ---

$baseUnc = '\\172.10.0.11\AST-AKLAN Branch\Payroll\3. Sales Clerk Incentives'
$offPath = Resolve-BranchFolderPath -BasePath $baseUnc -UseBranchYearMonth $false -BranchYear '2026' -BranchMonth '7. July'
Assert-True ($offPath -eq $baseUnc) "Toggle off should leave path unchanged, got $offPath"

$onPath = Resolve-BranchFolderPath -BasePath $baseUnc -UseBranchYearMonth $true -BranchYear '2026' -BranchMonth '7. July'
$expectedOn = Join-Path (Join-Path $baseUnc '2026') '7. July'
Assert-True ($onPath -eq $expectedOn) "Toggle on path mismatch: got $onPath, expected $expectedOn"

$threwYear = $false
try {
    Assert-BranchYearMonthConfig -UseBranchYearMonth $true -BranchYear '' -BranchMonth '7. July'
}
catch {
    $threwYear = $true
}
Assert-True $threwYear 'Expected Assert-BranchYearMonthConfig to throw when BRANCH_YEAR is empty'

$threwMonth = $false
try {
    Assert-BranchYearMonthConfig -UseBranchYearMonth $true -BranchYear '2026' -BranchMonth ''
}
catch {
    $threwMonth = $true
}
Assert-True $threwMonth 'Expected Assert-BranchYearMonthConfig to throw when BRANCH_MONTH is empty'

Assert-BranchYearMonthConfig -UseBranchYearMonth $false -BranchYear '' -BranchMonth ''

# Grouped single-path branches must keep the whole string, not the first character.
$singleGroupedPath = '\\server\AST-AKLAN Branch\Payroll\3. Sales Clerk Incentives\2026\7. July'
$singleGrouped = @(Group-BranchPaths -BranchPaths @(
        [pscustomobject]@{ Label = 'Aklan'; Path = $singleGroupedPath }
    ))
Assert-True ($singleGrouped.Count -eq 1) "Expected 1 grouped branch, got $($singleGrouped.Count)"
Assert-True ($singleGrouped[0].Path -eq $singleGroupedPath) `
    "Grouped primary path mismatch: got $($singleGrouped[0].Path), expected $singleGroupedPath"
Assert-True ($singleGrouped[0].Paths[0] -eq $singleGroupedPath) `
    "Grouped Paths[0] mismatch: got $($singleGrouped[0].Paths[0]), expected $singleGroupedPath"

$singleGroupedResults = @(Get-BranchesWithCutoffFiles -BranchPaths $singleGrouped -StartDay 8 -EndDay 14)
Assert-True ($singleGroupedResults[0].Path -eq $singleGroupedPath) `
    "No-match selected path mismatch: got $($singleGroupedResults[0].Path), expected $singleGroupedPath"

# --- Explorer path argument ---

$openTempRoot = Join-Path ([IO.Path]::GetTempPath()) ("check-files-open path-" + [guid]::NewGuid().ToString('N'))
try {
    $openFolder = Join-Path $openTempRoot 'Branch Folder'
    New-Item -ItemType Directory -Path $openFolder -Force | Out-Null

    $slashPath = $openFolder.Replace('\', '/')
    $openArgument = Get-ExplorerFolderPathArgument -FolderPath $slashPath
    $expectedOpenArgument = '"' + (Resolve-Path -LiteralPath $openFolder).ProviderPath + '"'
    Assert-True ($openArgument -eq $expectedOpenArgument) `
        "Explorer argument mismatch: got $openArgument, expected $expectedOpenArgument"
}
finally {
    if (Test-Path -LiteralPath $openTempRoot) {
        Remove-Item -LiteralPath $openTempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$debugLogRoot = Join-Path ([IO.Path]::GetTempPath()) ("check-files-debug-log-" + [guid]::NewGuid().ToString('N'))
$debugLogPath = Join-Path $debugLogRoot 'open-flow.log'
try {
    $script:log_path = $debugLogPath
    $script:log_level = 'INFO'
    Write-Log '[open-flow] suppressed message' -Level DEBUG
    Assert-True (-not (Test-Path -LiteralPath $debugLogPath)) 'DEBUG lines should be suppressed when LOG_LEVEL is INFO'

    $script:log_level = 'DEBUG'
    Write-Log '[open-flow] enabled message' -Level DEBUG
    Assert-True (Test-Path -LiteralPath $debugLogPath) 'DEBUG lines should create a log file when LOG_LEVEL is DEBUG'
    $debugLogText = Get-Content -LiteralPath $debugLogPath -Raw -Encoding UTF8
    Assert-True ($debugLogText -match '\[open-flow\] enabled message') 'DEBUG logging should write the open-flow marker'

    Assert-True ((Resolve-LogLevel -Value $null) -eq 'INFO') 'Missing LOG_LEVEL should default to INFO'
    Assert-True ((Resolve-LogLevel -Value 'debug') -eq 'DEBUG') 'LOG_LEVEL=debug should normalize to DEBUG'
    Assert-True ((Resolve-LogLevel -Value ' WARNING ') -eq 'WARNING') 'LOG_LEVEL=WARNING should trim and normalize'
}
finally {
    $script:log_level = 'INFO'
    $script:log_path = $null
    if (Test-Path -LiteralPath $debugLogRoot) {
        Remove-Item -LiteralPath $debugLogRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Multi-line BRANCH_PATHS= lines are joined by Import-DotEnv
$envFile = Join-Path ([IO.Path]::GetTempPath()) ("branch-paths-env-" + [guid]::NewGuid().ToString('N') + '.env')
try {
    @(
        'LOG_PATH=logs\check-files-cutoff.log'
        'BRANCH_PATHS=Pandesalan=C:\branch\a'
        'BRANCH_PATHS=Mindoro=C:\branch\b'
    ) | Set-Content -LiteralPath $envFile -Encoding UTF8

    $envMap = Import-DotEnv -Path $envFile
    Assert-True ($envMap['BRANCH_PATHS'] -eq 'Pandesalan=C:\branch\a;Mindoro=C:\branch\b') `
        "Expected joined BRANCH_PATHS, got $($envMap['BRANCH_PATHS'])"

    $fromEnv = @(Parse-BranchPaths -Value $envMap['BRANCH_PATHS'])
    Assert-True ($fromEnv.Count -eq 2) "Expected 2 branches from multi-line env, got $($fromEnv.Count)"
    Assert-True ($fromEnv[0].Label -eq 'Pandesalan') 'Multi-line env first label should be Pandesalan'
    Assert-True ($fromEnv[1].Label -eq 'Mindoro') 'Multi-line env second label should be Mindoro'
}
finally {
    if (Test-Path -LiteralPath $envFile) {
        Remove-Item -LiteralPath $envFile -Force -ErrorAction SilentlyContinue
    }
}

# --- Folder scan ---

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("check-files-cutoff-tests-" + [guid]::NewGuid().ToString('N'))
$branchA = Join-Path $tempRoot 'A'
$branchB = Join-Path $tempRoot 'B'
$missing = Join-Path $tempRoot 'Missing'

try {
    New-Item -ItemType Directory -Path $branchA -Force | Out-Null
    New-Item -ItemType Directory -Path $branchB -Force | Out-Null

    Set-Content -LiteralPath (Join-Path $branchA 'Incentives JULY 8-14, 2026.xlsx') -Value 'dummy' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $branchA '~$Incentives JULY 8-14, 2026.xlsx') -Value 'lock' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $branchB 'Other JULY 15-21, 2026.xlsx') -Value 'dummy' -Encoding UTF8

    $lockOnly = Join-Path $tempRoot 'LockOnly'
    New-Item -ItemType Directory -Path $lockOnly -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $lockOnly '~$Incentives JULY 8-14, 2026.xlsx') -Value 'lock' -Encoding UTF8

    Assert-True (Test-BranchHasCutoffFile -FolderPath $branchA -StartDay 8 -EndDay 14) `
        'Branch A should have cutoff files for 8-14'
    Assert-True (-not (Test-BranchHasCutoffFile -FolderPath $branchB -StartDay 8 -EndDay 14)) `
        'Branch B should not have cutoff files for 8-14'
    Assert-True (-not (Test-BranchHasCutoffFile -FolderPath $missing -StartDay 8 -EndDay 14)) `
        'Missing folder should report no cutoff files'
    Assert-True (-not (Test-BranchHasCutoffFile -FolderPath $lockOnly -StartDay 8 -EndDay 14)) `
        'Folder with only a ~$ lock file should report no cutoff files'

    $branches = @(
        [pscustomobject]@{ Label = 'Alpha'; Path = $branchA }
        [pscustomobject]@{ Label = 'Beta'; Path = $branchB }
        [pscustomobject]@{ Label = 'Gone'; Path = $missing }
    )

    $results = @(Get-BranchesWithCutoffFiles -BranchPaths $branches -StartDay 8 -EndDay 14)
    Assert-True ($results.Count -eq 3) "Expected 3 results, got $($results.Count)"
    Assert-True ($results[0].HasFile -eq $true) 'Alpha should be true'
    Assert-True ($results[1].HasFile -eq $false) 'Beta should be false'
    Assert-True ($results[2].HasFile -eq $false) 'Gone should be false'

    $checklist = Format-CutoffFilesChecklist -Results $results
    Assert-True ($checklist -match [regex]::Escape('[OK] Alpha')) 'Checklist should mark Alpha present'
    Assert-True ($checklist -match [regex]::Escape('[X] Beta')) 'Checklist should mark Beta missing'
    Assert-True ($checklist -match [regex]::Escape('[X] Gone')) 'Checklist should mark Gone missing'
}
finally {
    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# --- SR folder scan ---

$srTempRoot = Join-Path ([IO.Path]::GetTempPath()) ("check-files-sr-tests-" + [guid]::NewGuid().ToString('N'))
$srBranch = Join-Path $srTempRoot 'SRBranch'
try {
    New-Item -ItemType Directory -Path $srBranch -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $srBranch 'SALES REPORT AUG 8 2026.pdf') -Value 'dummy' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $srBranch 'SR_9_2026.jpg') -Value 'dummy' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $srBranch 'SR_18_2026.pdf') -Value 'dummy' -Encoding UTF8

    $script:sr_file_extensions = @('.pdf', '.jpg')
    $script:recursive = $false
    $periodDays = @(8, 9)

    $srMatches = @(Get-SalesReportFilesInFolder -FolderPath $srBranch -PeriodDays $periodDays)
    Assert-True ($srMatches.Count -eq 2) "Expected 2 SR matches, got $($srMatches.Count)"
    Assert-True (@($srMatches | ForEach-Object { $_.Name }) -contains 'SALES REPORT AUG 8 2026.pdf') `
        'Expected day 8 PDF match'
    Assert-True (@($srMatches | ForEach-Object { $_.Name }) -contains 'SR_9_2026.jpg') `
        'Expected day 9 JPG match'
}
finally {
    if (Test-Path -LiteralPath $srTempRoot) {
        Remove-Item -LiteralPath $srTempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host 'Check-Files-Cutoff.Tests.ps1: all assertions passed.'
