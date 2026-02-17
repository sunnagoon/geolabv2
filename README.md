# GeoLab v2

Soils lab management app (desktop Tkinter + SQLite) for Terrapacific Consultants.

## Quick start
1. Create a virtualenv
2. Install deps:
   pip install -r requirements.txt
3. Run app:
   python -m app.main

## Data
- SQLite DB is created at data/geolab.db on first run.

## Templates
- Place Excel templates in templates/ (billing and results). The app currently exports a basic billing sheet; template integration is stubbed.

## Release workflow (EXE + MSI + signing)
### Build EXE
- Command:
  `powershell -ExecutionPolicy Bypass -File scripts/build_exe.ps1`
- Output:
  `dist/GeoLab/GeoLab.exe`

### Build MSI (WiX v3)
Prerequisites:
- Install WiX Toolset v3 and ensure `heat`, `candle`, and `light` are on PATH.

Command:
- `powershell -ExecutionPolicy Bypass -File scripts/build_msi.ps1 -ProductVersion 1.0.0`

Output:
- `release/GeoLab-1.0.0.msi`

### Code-sign EXE/MSI
Prerequisites:
- Windows SDK `signtool` available on PATH.
- Code signing certificate available in cert store (`-Thumbprint`) or as PFX (`-PfxPath`).

Examples:
- Cert store:
  `powershell -ExecutionPolicy Bypass -File scripts/sign_release.ps1 -TimestampUrl http://timestamp.digicert.com -Thumbprint YOUR_CERT_THUMBPRINT`
- PFX:
  `powershell -ExecutionPolicy Bypass -File scripts/sign_release.ps1 -TimestampUrl http://timestamp.digicert.com -PfxPath C:\certs\code_sign.pfx -PfxPassword YOUR_PASSWORD`

## Backup configuration
- Open app -> `Settings` tab.
- Set `Backup Folder`.
- Click `Save Backup Folder`.
- Click `Backup Now` to create timestamped DB backup.
