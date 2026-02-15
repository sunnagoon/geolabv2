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