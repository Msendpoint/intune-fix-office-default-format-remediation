<#
.SYNOPSIS
    One-click installer and launcher for Fix Office Apps Defaulting to ODS/ODT Format with Intune Proactive Remediation

.DESCRIPTION
    This script checks prerequisites, sets execution policy, and runs the main
    automation script. Safe to run multiple times.

.NOTES
    Author:      Souhaiel Morhag
    Company:     MSEndpoint.com
    Blog:        https://msendpoint.com
    LinkedIn:    https://linkedin.com/in/souhaiel-morhag
    Created:     2026-03-28
    License:     MIT

.EXAMPLE
    .\Install.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "    Fix Office Apps Defaulting to ODS/ODT Format with Intune Proactive Remediation" -ForegroundColor Cyan
Write-Host "    MSEndpoint.com — Souhaiel Morhag" -ForegroundColor DarkGray
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""

# ── 1. Check PowerShell version ──────────────────────────────
$minVersion = [Version]"5.1"
if ($PSVersionTable.PSVersion -lt $minVersion) {
    Write-Error "PowerShell $minVersion or later is required. Current: $($PSVersionTable.PSVersion)"
    exit 1
}
Write-Host "  [OK] PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Green

# ── 2. Check execution policy ────────────────────────────────
$policy = Get-ExecutionPolicy -Scope CurrentUser
if ($policy -eq 'Restricted') {
    Write-Host "  [..] Setting execution policy to RemoteSigned..." -ForegroundColor Yellow
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    Write-Host "  [OK] Execution policy updated" -ForegroundColor Green
} else {
    Write-Host "  [OK] Execution policy: $policy" -ForegroundColor Green
}

# ── 3. Check required modules ────────────────────────────────
$requiredModules = @('Microsoft.Graph')
foreach ($mod in $requiredModules) {
    if (Get-Module -ListAvailable -Name $mod) {
        Write-Host "  [OK] Module: $mod" -ForegroundColor Green
    } else {
        Write-Host "  [..] Installing module: $mod ..." -ForegroundColor Yellow
        try {
            Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
            Write-Host "  [OK] Module: $mod installed" -ForegroundColor Green
        } catch {
            Write-Warning "  [!!] Could not install $mod — install manually: Install-Module $mod -Scope CurrentUser"
        }
    }
}

# ── 4. Create output directory ────────────────────────────────
$outputDir = Join-Path $scriptDir "output"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# ── 5. Launch main script ────────────────────────────────────
$mainScript = Join-Path $scriptDir "scripts/intune-fix-office-default-format-remediation.ps1"
if (Test-Path $mainScript) {
    Write-Host ""
    Write-Host "  Launching main script..." -ForegroundColor Cyan
    Write-Host "  ────────────────────────────────────────────" -ForegroundColor DarkGray
    & $mainScript @PSBoundParameters
} else {
    Write-Error "Main script not found: $mainScript"
    exit 1
}