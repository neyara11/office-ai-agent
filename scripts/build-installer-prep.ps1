# Prepare Release code outputs and audit OfficeAgent.vdproj SourcePath inputs.
# This does NOT build the MSI (.vdproj requires Visual Studio Installer Projects UI/devenv).
#
# Usage:
#   powershell -File .\scripts\build-installer-prep.ps1
#   powershell -File .\scripts\build-installer-prep.ps1 -SkipBuild
#   powershell -File .\scripts\build-installer-prep.ps1 -IncludePlainFiles

param(
    [switch]$SkipBuild,
    [switch]$IncludePlainFiles,
    [string]$Configuration = "Release",
    [string]$Platform = "AnyCPU"
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$buildCode = Join-Path $repoRoot "scripts\build-code-projects.ps1"
$auditInputs = Join-Path $repoRoot "build\AuditInstallerInputs.ps1"
$vdproj = Join-Path $repoRoot "OfficeAgent\OfficeAgent.vdproj"

Write-Host "============================================================"
Write-Host " Installer prep (code Release + vdproj input audit)"
Write-Host " Repo: $repoRoot"
Write-Host "============================================================"
Write-Host ""
Write-Host "NOTE: This script does NOT build OfficeAgent.vdproj / MSI."
Write-Host "      After PASS, open OfficeAgent.vdproj in Visual Studio"
Write-Host "      (with 'Microsoft Visual Studio Installer Projects' installed)"
Write-Host "      and Build the installer project."
Write-Host ""

if (-not (Test-Path -LiteralPath $vdproj)) {
    throw "Installer project not found: $vdproj"
}

if (-not $SkipBuild) {
    if ($Configuration -ne "Release") {
        Write-Host "WARNING: OfficeAgent.vdproj SourcePath entries point at bin\Release."
        Write-Host "         Prefer -Configuration Release (current: $Configuration)."
        Write-Host ""
    }

    Write-Host "==> Step 1/2: Build code projects [$Configuration|$Platform]"
    & $buildCode -Configuration $Configuration -Platform $Platform
    if ($LASTEXITCODE -ne 0) {
        throw "Code project build failed (exit $LASTEXITCODE)."
    }
}
else {
    Write-Host "==> Step 1/2: SkipBuild - using existing $Configuration outputs"
}

Write-Host ""
Write-Host "==> Step 2/2: Audit installer SourcePath inputs"
$auditArgs = @{}
if ($IncludePlainFiles) {
    $auditArgs["IncludePlainFiles"] = $true
}

# Capture exit code explicitly; PowerShell may leave $LASTEXITCODE unset on some hosts.
$global:LASTEXITCODE = 0
& $auditInputs @auditArgs
$auditExit = 0
if ($null -ne $LASTEXITCODE) {
    $auditExit = [int]$LASTEXITCODE
}
if ($auditExit -ne 0) {
    Write-Host ""
    Write-Host "Installer input audit FAILED (exit=$auditExit)."
    Write-Host "Fix missing Release outputs, then re-run this script."
    Write-Host "Typical fix: remove -SkipBuild, or build Release code first:"
    Write-Host "  powershell -File .\scripts\build-installer-prep.ps1"
    exit $auditExit
}

Write-Host ""
Write-Host "============================================================"
Write-Host " Installer prep PASSED"
Write-Host "============================================================"

# Building VSTO projects inside Visual Studio recreates HKCU add-in registrations pointing at
# ..\bin\Release\...vsto. If they stay, Office loads both the dev and the MSI-installed add-in
# and shows duplicate ribbon tabs. Remove them so only the installed registration remains.
$cleanupScript = Join-Path $repoRoot "scripts\clean-vsto-debug-registrations.ps1"
if (Test-Path -LiteralPath $cleanupScript) {
    Write-Host ""
    Write-Host "==> Step 3/3: Remove stale VSTO dev registrations"
    & $cleanupScript
}

Write-Host "Next steps:"
Write-Host "  1. Open Visual Studio with Installer Projects extension."
Write-Host "  2. Open OfficeAgent\OfficeAgent.vdproj (or load it from AiHelper.sln)."
Write-Host "  3. Set configuration to Release and Build the installer project."
Write-Host "  4. Optional full gate: powershell -File .\build\RunReleaseChecks.ps1"
Write-Host ""
Write-Host "Do NOT treat a failed full-solution Debug Rebuild (vdproj missing"
Write-Host "Release paths) as a code compile failure. Use build-code for code."
Write-Host ""
