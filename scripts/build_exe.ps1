param(
    [string]$Name = "GeoLab",
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if ($Clean) {
    if (Test-Path "build") { Remove-Item "build" -Recurse -Force }
    if (Test-Path "dist") { Remove-Item "dist" -Recurse -Force }
    if (Test-Path "$Name.spec") { Remove-Item "$Name.spec" -Force }
}

$args = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--windowed",
    "--name", $Name,
    "--add-data", "templates;templates"
)

if (Test-Path "assets") {
    $args += @("--add-data", "assets;assets")
}

$args += "app/main.py"

Write-Host "Building executable..."
python @args

Write-Host "Build complete:"
Write-Host "  $root\dist\$Name\$Name.exe"
