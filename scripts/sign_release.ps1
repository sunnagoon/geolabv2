param(
    [Parameter(Mandatory = $true)]
    [string]$TimestampUrl,

    [string]$Thumbprint,
    [string]$PfxPath,
    [string]$PfxPassword,
    [string]$Name = "GeoLab",
    [string]$MsiPath = ""
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

function Require-Command($cmd) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        throw "Required command not found: $cmd"
    }
}

Require-Command signtool

$exePath = "dist\$Name\$Name.exe"
if (-not (Test-Path $exePath)) {
    throw "Executable not found: $exePath"
}

if (-not $MsiPath) {
    $msiCandidate = Get-ChildItem "release" -Filter "*.msi" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($msiCandidate) {
        $MsiPath = $msiCandidate.FullName
    }
}

$baseArgs = @("sign", "/fd", "SHA256", "/tr", $TimestampUrl, "/td", "SHA256")

if ($Thumbprint) {
    $signArgs = $baseArgs + @("/sha1", $Thumbprint)
} elseif ($PfxPath) {
    $signArgs = $baseArgs + @("/f", $PfxPath)
    if ($PfxPassword) {
        $signArgs += @("/p", $PfxPassword)
    }
} else {
    throw "Provide either -Thumbprint or -PfxPath for signing."
}

Write-Host "Signing EXE: $exePath"
signtool @signArgs $exePath

if ($MsiPath -and (Test-Path $MsiPath)) {
    Write-Host "Signing MSI: $MsiPath"
    signtool @signArgs $MsiPath
} else {
    Write-Host "No MSI found to sign. Skipping MSI signing."
}

Write-Host "Signing complete."
