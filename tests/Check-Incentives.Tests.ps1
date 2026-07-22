#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'check-incentives.ps1')

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

# Multi-line BRANCH_PATHS= lines are joined by Import-DotEnv
$envFile = Join-Path ([IO.Path]::GetTempPath()) ("branch-paths-env-" + [guid]::NewGuid().ToString('N') + '.env')
try {
    @(
        'LOG_PATH=logs\check-incentives.log'
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

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("check-incentives-tests-" + [guid]::NewGuid().ToString('N'))
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

    Assert-True (Test-BranchHasIncentiveFile -FolderPath $branchA -StartDay 8 -EndDay 14) `
        'Branch A should have incentives for 8-14'
    Assert-True (-not (Test-BranchHasIncentiveFile -FolderPath $branchB -StartDay 8 -EndDay 14)) `
        'Branch B should not have incentives for 8-14'
    Assert-True (-not (Test-BranchHasIncentiveFile -FolderPath $missing -StartDay 8 -EndDay 14)) `
        'Missing folder should report no incentives'
    Assert-True (-not (Test-BranchHasIncentiveFile -FolderPath $lockOnly -StartDay 8 -EndDay 14)) `
        'Folder with only a ~$ lock file should report no incentives'

    $branches = @(
        [pscustomobject]@{ Label = 'Alpha'; Path = $branchA }
        [pscustomobject]@{ Label = 'Beta'; Path = $branchB }
        [pscustomobject]@{ Label = 'Gone'; Path = $missing }
    )

    $results = @(Get-BranchesWithIncentives -BranchPaths $branches -StartDay 8 -EndDay 14)
    Assert-True ($results.Count -eq 3) "Expected 3 results, got $($results.Count)"
    Assert-True ($results[0].HasFile -eq $true) 'Alpha should be true'
    Assert-True ($results[1].HasFile -eq $false) 'Beta should be false'
    Assert-True ($results[2].HasFile -eq $false) 'Gone should be false'

    $checklist = Format-IncentivesChecklist -Results $results
    Assert-True ($checklist -match [regex]::Escape('✓ Alpha')) 'Checklist should mark Alpha present'
    Assert-True ($checklist -match [regex]::Escape('✗ Beta')) 'Checklist should mark Beta missing'
    Assert-True ($checklist -match [regex]::Escape('✗ Gone')) 'Checklist should mark Gone missing'
}
finally {
    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host 'Check-Incentives.Tests.ps1: all assertions passed.'
