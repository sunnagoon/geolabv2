param(
    [string]$ProductVersion = "1.0.0",
    [string]$Name = "GeoLab"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

function Require-Command($cmd) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        throw "Required command not found: $cmd"
    }
}

if (-not (Test-Path "dist\$Name\$Name.exe")) {
    Write-Host "Executable missing. Building first..."
    & "$PSScriptRoot\build_exe.ps1" -Name $Name
}

Require-Command heat
Require-Command candle
Require-Command light

New-Item -ItemType Directory -Path "installer\obj" -Force | Out-Null
New-Item -ItemType Directory -Path "release" -Force | Out-Null

$filesWxs = "installer\GeoLabFiles.wxs"
$productWxs = "installer\Product.wxs"
$objFiles = "installer\obj\GeoLabFiles.wixobj"
$objProduct = "installer\obj\Product.wixobj"
$msiOut = "release\$Name-$ProductVersion.msi"

Write-Host "Harvesting files from dist\$Name ..."
heat dir "dist\$Name" `
  -gg `
  -srd `
  -cg GeoLabFiles `
  -dr INSTALLFOLDER `
  -var var.SourceDir `
  -out $filesWxs

Write-Host "Compiling WiX sources..."
candle `
  -dProductVersion=$ProductVersion `
  -dSourceDir="dist\$Name" `
  -out "installer\obj\" `
  $productWxs `
  $filesWxs

Write-Host "Linking MSI..."
light `
  -out $msiOut `
  $objProduct `
  $objFiles

Write-Host "MSI complete:"
Write-Host "  $root\$msiOut"
