# Removes stale VSTO add-in registrations created by building inside Visual Studio.
# Such keys point at a developer path like ...\bin\Release\...vsto and make Office load
# the add-in twice (installed via MSI + dev build), which shows up as duplicate ribbon tabs.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\clean-vsto-debug-registrations.ps1
#   powershell -ExecutionPolicy Bypass -File .\scripts\clean-vsto-debug-registrations.ps1 -WhatIfOnly

param(
    [switch]$WhatIfOnly
)

$ErrorActionPreference = "Stop"

$apps = @('Excel', 'Word', 'PowerPoint')
$removed = 0

foreach ($app in $apps) {
    $addinsRoot = "HKCU:\Software\Microsoft\Office\$app\Addins"
    if (-not (Test-Path $addinsRoot)) { continue }

    foreach ($key in Get-ChildItem $addinsRoot -ErrorAction SilentlyContinue) {
        $props = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
        $manifest = [string]$props.Manifest
        if ([string]::IsNullOrWhiteSpace($manifest)) { continue }

        $isDevPath = ($manifest -match '[\\/]bin[\\/]Debug[\\/]') -or ($manifest -match '[\\/]bin[\\/]Release[\\/]')
        if (-not $isDevPath) { continue }

        Write-Host "Stale dev registration: $($key.PSPath)"
        Write-Host "    Manifest: $manifest"
        if (-not $WhatIfOnly) {
            Remove-Item $key.PSPath -Recurse -Force
        }
        $removed += 1
    }
}

if ($WhatIfOnly) {
    Write-Host "WhatIf: found $removed stale registration(s), nothing removed."
}
else {
    Write-Host "Removed $removed stale dev registration(s). Restart Office to apply."
}
