import json
import os
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

from app.db import get_connection


class MapTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.points = []
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        wrapper = ttk.Frame(self)
        wrapper.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        header = ttk.Frame(wrapper)
        header.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(
            header,
            text="Interactive map covers California + Arizona by default. Add Lat/Lon in Projects.",
        ).pack(side=tk.LEFT)
        ttk.Button(header, text="Refresh", command=self.refresh).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(header, text="Open Interactive Map", command=self.open_interactive_map).pack(side=tk.RIGHT)

        self.listbox = tk.Listbox(wrapper, height=28)
        self.listbox.pack(fill=tk.BOTH, expand=True)

    def refresh(self):
        conn = get_connection()
        rows = conn.execute(
            """
            SELECT file_number, job_name, location_text, latitude, longitude
            FROM projects
            ORDER BY created_at DESC
            """
        ).fetchall()
        conn.close()

        self.points = [dict(r) for r in rows]
        self._fill_list(self.points)

    def _fill_list(self, points):
        self.listbox.delete(0, tk.END)
        mapped = 0
        for p in points:
            has_coords = p["latitude"] is not None and p["longitude"] is not None
            if has_coords:
                mapped += 1
            location = p["location_text"] or "No location text"
            coord_text = (
                f"({p['latitude']:.5f}, {p['longitude']:.5f})"
                if has_coords
                else "(no coordinates)"
            )
            self.listbox.insert(
                tk.END,
                f"{p['file_number']} | {p['job_name']} | {location} {coord_text}",
            )
        self.listbox.insert(tk.END, "")
        self.listbox.insert(tk.END, f"Mapped points: {mapped} / {len(points)}")

    def open_interactive_map(self):
        points_with_coords = [
            p for p in self.points if p["latitude"] is not None and p["longitude"] is not None
        ]
        if not points_with_coords:
            messagebox.showerror("No Coordinates", "No projects have latitude/longitude yet.")
            return

        html = self._build_leaflet_html(points_with_coords)
        out_path = Path(__file__).resolve().parent.parent.parent / "data" / "project_map.html"
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(html, encoding="utf-8")

        try:
            os.startfile(str(out_path))  # type: ignore[attr-defined]
        except Exception:
            webbrowser.open(out_path.as_uri())

    def _build_leaflet_html(self, points):
        marker_data = [
            {
                "file_number": p["file_number"],
                "job_name": p["job_name"],
                "location_text": p["location_text"] or "",
                "lat": float(p["latitude"]),
                "lon": float(p["longitude"]),
            }
            for p in points
        ]
        json_data = json.dumps(marker_data)

        return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8"/>
  <title>GeoLab Project Map</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
  <style>
    html, body, #map {{ height: 100%; margin: 0; padding: 0; }}
    .legend {{
      position: absolute; z-index: 9999; right: 10px; top: 10px;
      background: rgba(255,255,255,0.95); padding: 10px 12px;
      border: 1px solid #ccc; border-radius: 8px; font-family: Segoe UI, sans-serif; font-size: 13px;
    }}
  </style>
</head>
<body>
  <div id="map"></div>
  <div class="legend">
    <div><strong>GeoLab Jobs</strong></div>
    <div>Default extent: California + Arizona</div>
  </div>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <script>
    const points = {json_data};
    const map = L.map('map').setView([34.2, -115.5], 6);

    L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap contributors'
    }}).addTo(map);

    const bounds = [];
    for (const p of points) {{
      const marker = L.circleMarker([p.lat, p.lon], {{
        radius: 6,
        color: '#2f6e52',
        fillColor: '#4aa97a',
        fillOpacity: 0.85
      }}).addTo(map);
      marker.bindPopup(
        `<strong>${{p.file_number}}</strong><br/>${{p.job_name}}<br/>${{p.location_text || 'No location text'}}<br/>${{p.lat.toFixed(5)}}, ${{p.lon.toFixed(5)}}`
      );
      bounds.push([p.lat, p.lon]);
    }}
    if (bounds.length > 0) {{
      map.fitBounds(bounds, {{ padding: [40, 40], maxZoom: 12 }});
    }}
  </script>
</body>
</html>
"""
