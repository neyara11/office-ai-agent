# Ensure per-project VSTO TemporaryKey.pfx exist for local Debug signing.
# Does NOT use or create release/code-signing certificates.
#
#   powershell -File .\scripts\ensure-vsto-temp-keys.ps1
#   powershell -File .\scripts\ensure-vsto-temp-keys.ps1 -Force
#   powershell -File .\scripts\ensure-vsto-temp-keys.ps1 -WhatIf

param(
    [switch]$Force,
    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")

$projects = @(
    @{ Name = "ExcelAi"; Dir = "ExcelAi"; Pfx = "ExcelAi_TemporaryKey.pfx"; Subject = "CN=OfficeAI-ExcelAi-Dev" },
    @{ Name = "WordAi"; Dir = "WordAi"; Pfx = "WordAi_TemporaryKey.pfx"; Subject = "CN=OfficeAI-WordAi-Dev" },
    @{ Name = "PowerPointAi"; Dir = "PowerPointAi"; Pfx = "PowerPointAi_TemporaryKey.pfx"; Subject = "CN=OfficeAI-PowerPointAi-Dev" }
)

function Test-PfxUsable {
    param([string]$PfxPath)
    if (-not (Test-Path -LiteralPath $PfxPath)) { return $false }
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PfxPath, "")
        if ($cert.NotAfter -lt (Get-Date).AddDays(7)) {
            Write-Host "  Expired or expiring soon: $PfxPath (NotAfter=$($cert.NotAfter))"
            return $false
        }
        if (-not $cert.HasPrivateKey) {
            Write-Host "  No private key: $PfxPath"
            return $false
        }
        return $true
    }
    catch {
        Write-Host "  Cannot load pfx: $PfxPath ($($_.Exception.Message))"
        return $false
    }
}

function Update-ManifestThumbprint {
    param(
        [string]$VbprojPath,
        [string]$Thumbprint
    )
    if (-not (Test-Path -LiteralPath $VbprojPath)) {
        Write-Warning "vbproj not found: $VbprojPath"
        return
    }
    $content = Get-Content -LiteralPath $VbprojPath -Raw
    $pattern = '(?s)(<ManifestCertificateThumbprint>)([^<]*)(</ManifestCertificateThumbprint>)'
    if ($content -notmatch $pattern) {
        Write-Warning "ManifestCertificateThumbprint not found in $VbprojPath"
        return
    }
    $newContent = [regex]::Replace($content, $pattern, "`${1}$Thumbprint`${3}")
    if ($newContent -eq $content) {
        Write-Host "  Thumbprint already up to date in $(Split-Path $VbprojPath -Leaf)"
        return
    }
    if ($WhatIf) {
        Write-Host "  WhatIf: would update ManifestCertificateThumbprint to $Thumbprint"
        return
    }
    Set-Content -LiteralPath $VbprojPath -Value $newContent -Encoding UTF8 -NoNewline
    Write-Host "  Updated ManifestCertificateThumbprint in $(Split-Path $VbprojPath -Leaf)"
}

function New-DevPfx {
    param(
        [string]$PfxPath,
        [string]$Subject
    )
    if ($WhatIf) {
        Write-Host "  WhatIf: would create $PfxPath ($Subject)"
        return $null
    }

    $cert = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $Subject `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -KeyLength 2048 `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears(2) `
        -CertStoreLocation "Cert:\CurrentUser\My"

    $pwd = New-Object System.Security.SecureString
    # Empty password: matches typical VSTO TemporaryKey usage in this repo.
    Export-PfxCertificate -Cert $cert -FilePath $PfxPath -Password $pwd | Out-Null

    # Remove from store copy optional - keep for local trust if desired
    Write-Host "  Created $PfxPath"
    Write-Host "  Thumbprint: $($cert.Thumbprint)"
    return $cert.Thumbprint
}

Write-Host "VSTO TemporaryKey ensure (local Debug only)"
Write-Host "Repo: $repoRoot"
Write-Host ""

foreach ($p in $projects) {
    $pfxPath = Join-Path $repoRoot (Join-Path $p.Dir $p.Pfx)
    $vbproj = Join-Path $repoRoot (Join-Path $p.Dir ($p.Name + ".vbproj"))
    Write-Host "==> $($p.Name)"

    $needNew = $Force -or -not (Test-PfxUsable -PfxPath $pfxPath)
    if (-not $needNew) {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxPath, "")
        Write-Host "  OK existing pfx (NotAfter=$($cert.NotAfter), Thumbprint=$($cert.Thumbprint))"
        Update-ManifestThumbprint -VbprojPath $vbproj -Thumbprint $cert.Thumbprint
        continue
    }

    if ((Test-Path -LiteralPath $pfxPath) -and $Force -and -not $WhatIf) {
        Remove-Item -LiteralPath $pfxPath -Force
    }

    $thumb = New-DevPfx -PfxPath $pfxPath -Subject $p.Subject
    if ($thumb) {
        Update-ManifestThumbprint -VbprojPath $vbproj -Thumbprint $thumb
    }
}

Write-Host ""
Write-Host "Done. TemporaryKey files are for local Debug only - do not commit *.pfx."
Write-Host "Release signing: build\SignArtifacts.ps1 + OFFICE_AI_SIGN_* env vars."
Write-Host "See docs\signing-and-certificates.md"
