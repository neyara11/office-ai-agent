# Sets one release version consistently across all assemblies and the MSI project.
#
# Updates, in one pass:
#   - AssemblyFileVersion / AssemblyInformationalVersion in the four AssemblyInfo.vb files
#   - ApplicationVersion in the three VSTO host .vbproj files (ClickOnce manifest version
#     packaged into the MSI; AuditVersion.ps1 fails if it drifts from the assemblies)
#   - ProductVersion in OfficeAgent.vdproj
#   - ProductCode / PackageCode (new GUIDs so the MSI is a proper upgrade)
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\set-version.ps1 -Version 2.8.12
#   powershell -ExecutionPolicy Bypass -File .\scripts\set-version.ps1 -Version 2.8.12 -SkipCodes

param(
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [switch]$SkipCodes
)

$ErrorActionPreference = "Stop"

if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    throw "Version must have the form major.minor.patch, e.g. 2.8.12 (got: $Version)"
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$fileVersion = "$Version.0"
$latin1 = [System.Text.Encoding]::GetEncoding(28591)

$assemblyInfoFiles = @(
    "WordAi\My Project\AssemblyInfo.vb",
    "ExcelAi\My Project\AssemblyInfo.vb",
    "PowerPointAi\My Project\AssemblyInfo.vb",
    "ShareRibbon\My Project\AssemblyInfo.vb"
)

foreach ($relative in $assemblyInfoFiles) {
    $path = Join-Path $repoRoot $relative
    if (-not (Test-Path -LiteralPath $path)) { throw "Not found: $path" }

    # Latin-1 round-trip keeps every non-ASCII byte untouched (these files are not UTF-8).
    $text = $latin1.GetString([System.IO.File]::ReadAllBytes($path))
    $text = [regex]::Replace($text, 'AssemblyFileVersion\("[^"]*"\)', "AssemblyFileVersion(`"$fileVersion`")")
    $text = [regex]::Replace($text, 'AssemblyInformationalVersion\("[^"]*"\)', "AssemblyInformationalVersion(`"$Version`")")
    [System.IO.File]::WriteAllBytes($path, $latin1.GetBytes($text))
    Write-Host "Updated $relative -> file $fileVersion / informational $Version"
}

# VSTO host projects: ApplicationVersion feeds the ClickOnce manifest that ends up in the MSI.
$applicationProjectFiles = @(
    "WordAi\WordAi.vbproj",
    "ExcelAi\ExcelAi.vbproj",
    "PowerPointAi\PowerPointAi.vbproj"
)

foreach ($relative in $applicationProjectFiles) {
    $path = Join-Path $repoRoot $relative
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Host "Skipped optional missing file $relative"
        continue
    }

    $text = $latin1.GetString([System.IO.File]::ReadAllBytes($path))
    $updated = [regex]::Replace($text, '<ApplicationVersion>[^<]*</ApplicationVersion>', "<ApplicationVersion>$fileVersion</ApplicationVersion>")
    if ($updated -eq $text) {
        Write-Host "Unchanged $relative (no ApplicationVersion)"
        continue
    }

    [System.IO.File]::WriteAllBytes($path, $latin1.GetBytes($updated))
    Write-Host "Updated $relative -> ApplicationVersion $fileVersion"
}

$vdproj = Join-Path $repoRoot "OfficeAgent\OfficeAgent.vdproj"
if (-not (Test-Path -LiteralPath $vdproj)) { throw "Not found: $vdproj" }

# OfficeAgent.vdproj is GBK-encoded. Read/write it through Latin-1 so every non-ASCII byte
# survives untouched; only ASCII version fields are replaced.
$vdprojText = $latin1.GetString([System.IO.File]::ReadAllBytes($vdproj))
$vdprojText = [regex]::Replace($vdprojText, '"ProductVersion"\s*=\s*"8:[^"]*"', "`"ProductVersion`" = `"8:$Version`"")

if (-not $SkipCodes) {
    $productCode = "{" + [guid]::NewGuid().ToString().ToUpperInvariant() + "}"
    $packageCode = "{" + [guid]::NewGuid().ToString().ToUpperInvariant() + "}"
    # Only GUID-form values: the .NET Framework prerequisite rows also use a "ProductCode" key
    # but their value is ".NETFramework,Version=v4.7.2" and must not be touched.
    $vdprojText = [regex]::Replace($vdprojText, '"ProductCode"\s*=\s*"8:\{[0-9A-Fa-f\-]+\}"', "`"ProductCode`" = `"8:$productCode`"")
    $vdprojText = [regex]::Replace($vdprojText, '"PackageCode"\s*=\s*"8:\{[0-9A-Fa-f\-]+\}"', "`"PackageCode`" = `"8:$packageCode`"")
    Write-Host "Updated OfficeAgent.vdproj -> ProductVersion $Version, ProductCode $productCode, PackageCode $packageCode"
}
else {
    Write-Host "Updated OfficeAgent.vdproj -> ProductVersion $Version (codes kept)"
}

[System.IO.File]::WriteAllBytes($vdproj, $latin1.GetBytes($vdprojText))
Write-Host "Done. Rebuild code, then build OfficeAgent.vdproj to produce the MSI."
