$ErrorActionPreference = "Stop"

$ProjectRoot = "D:\Workplace\SCS"
$MasterDo = Join-Path $ProjectRoot "00_master.do"

$preferredNames = @(
    "StataMP-64.exe",
    "StataSE-64.exe",
    "StataBE-64.exe",
    "Stata-64.exe"
)

function Add-Candidate {
    param(
        [System.Collections.Generic.List[string]] $List,
        [string] $Path
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if ((Test-Path -LiteralPath $resolved) -and -not $List.Contains($resolved)) {
        [void] $List.Add($resolved)
    }
}

$candidateDirs = [System.Collections.Generic.List[string]]::new()

$registryRoots = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($root in $registryRoots) {
    Get-ItemProperty $root -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match "Stata" -or $_.InstallLocation -match "Stata" } |
        ForEach-Object {
            Add-Candidate $candidateDirs $_.InstallLocation
            if ($_.DisplayIcon) {
                $iconPath = ($_.DisplayIcon -replace ',.*$', '').Trim('"')
                if (Test-Path -LiteralPath $iconPath) {
                    Add-Candidate $candidateDirs (Split-Path -Parent $iconPath)
                }
            }
        }
}

$commonDirs = @(
    "D:\Staata18",
    "D:\Stata18",
    "C:\Program Files\Stata18",
    "C:\Program Files\Stata17",
    "C:\Program Files\Stata16",
    "C:\Program Files (x86)\Stata18",
    "C:\Program Files (x86)\Stata17",
    "C:\Program Files (x86)\Stata16"
)

foreach ($dir in $commonDirs) {
    Add-Candidate $candidateDirs $dir
}

$stataExe = $null
foreach ($name in $preferredNames) {
    foreach ($dir in $candidateDirs) {
        $candidate = Join-Path $dir $name
        if (Test-Path -LiteralPath $candidate) {
            $stataExe = $candidate
            break
        }
    }
    if ($stataExe) { break }
}

if (-not $stataExe) {
    Write-Error ("No Stata executable found. Looked for: {0}. Checked install locations: {1}" -f ($preferredNames -join ", "), ($candidateDirs -join "; "))
    exit 1
}

if (-not (Test-Path -LiteralPath $MasterDo)) {
    Write-Error "Master do-file not found: $MasterDo"
    exit 1
}

Set-Location -LiteralPath $ProjectRoot
Write-Host "Using Stata executable: $stataExe"
Write-Host "Running master do-file: $MasterDo"

& $stataExe /e do $MasterDo
$exitCode = $LASTEXITCODE
$MasterLog = Join-Path $ProjectRoot "outputs\logs\00_master.log"
$logCompleted = $false
$stataErrors = @()

if (Test-Path -LiteralPath $MasterLog) {
    $masterLogText = Get-Content -LiteralPath $MasterLog -Raw
    $logCompleted = $masterLogText -match "Completed SCS workflow"
    $stataErrors = Select-String -Path (Join-Path $ProjectRoot "outputs\logs\*.log") -Pattern "^r\([0-9]+\);" -ErrorAction SilentlyContinue
}

if ($exitCode -ne 0) {
    if ($logCompleted -and (($stataErrors | Measure-Object).Count -eq 0)) {
        Write-Warning "Stata returned exit code $exitCode, but the master log reports completion and no Stata r() errors were found."
        exit 0
    }
    Write-Error "Stata batch run failed with exit code $exitCode. Check log: $MasterLog"
    exit $exitCode
}

Write-Host "Stata batch run completed successfully."
