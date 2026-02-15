import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app.db import get_connection, now_iso
from app.services.calculations_pti import default_payload, compute_pti, export_pti_pdf


class CalculationsTab(ttk.Frame):
    def __init__(self, parent, get_project_id):
        super().__init__(parent)
        self.get_project_id = get_project_id
        self.input_vars = {}
        self.output_vars = {}
        self._last_computed = {}
        self._build_ui()

    def _build_ui(self):
        wrap = ttk.Frame(self)
        wrap.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left = ttk.LabelFrame(wrap, text="PTI Shrink/Swell Inputs")
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        right = ttk.LabelFrame(wrap, text="Computed Outputs")
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        left_scroll = ttk.Frame(left)
        left_scroll.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.left_canvas = tk.Canvas(left_scroll, highlightthickness=0, bg="#ecf4fb")
        self.left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll = ttk.Scrollbar(left_scroll, orient=tk.VERTICAL, command=self.left_canvas.yview)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.left_canvas.configure(yscrollcommand=yscroll.set)
        self.left_body = ttk.Frame(self.left_canvas)
        self.left_window_id = self.left_canvas.create_window((0, 0), window=self.left_body, anchor="nw")
        self.left_body.bind("<Configure>", self._on_left_configure)
        self.left_canvas.bind("<Configure>", self._on_left_canvas_configure)

        fields = [
            ("project_title", "Project Title", ""),
            ("project_engineer", "Project Engineer", ""),
            ("project_number", "Project Number", ""),
            ("project_date", "Project Date", ""),
            ("layer_description", "Layer Description", ""),
            ("layer_thickness", "Layer Thickness", "ft"),
            ("ll", "Liquid Limit", ""),
            ("pl", "Plastic Limit", ""),
            ("passing_200", "% Passing #200", "%"),
            ("finer_2um", "% Finer than 2 um", "%"),
            ("dry_density", "Dry Density", "pcf"),
            ("fabric_factor", "Fabric Factor", ""),
            ("ko_drying", "Ko Drying", ""),
            ("ko_wetting", "Ko Wetting", ""),
            ("suction_initial_surface", "Wet Surface Suction", "pF"),
            ("suction_final_surface", "Dry Surface Suction", "pF"),
            ("constant_suction", "Constant Suction", "pF"),
            ("depth_constant_suction", "Depth to Constant Suction", "ft"),
            ("vertical_barrier_depth", "Vertical Barrier Depth", "ft"),
            ("horizontal_barrier_length", "Horizontal Barrier Length", "ft"),
            ("moisture_index", "Thornthwaite Moisture Index", ""),
            ("em_distance", "Em Distance (optional override)", "ft"),
        ]
        defaults = default_payload()
        row = 0
        for key, label, unit in fields:
            ttk.Label(self.left_body, text=label).grid(row=row, column=0, sticky=tk.W, padx=4, pady=3)
            var = tk.StringVar(value=str(defaults.get(key, "")))
            self.input_vars[key] = var
            ent = ttk.Entry(self.left_body, textvariable=var, width=24)
            ent.grid(row=row, column=1, sticky=tk.W, padx=4, pady=3)
            ent.bind("<KeyRelease>", lambda _e: self._compute_preview())
            ttk.Label(self.left_body, text=unit).grid(row=row, column=2, sticky=tk.W, padx=4, pady=3)
            row += 1

        btns = ttk.Frame(left)
        btns.pack(fill=tk.X, padx=6, pady=(2, 8))
        ttk.Button(btns, text="Compute + Save", command=self._compute_and_save).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btns, text="Export PDF", command=self._export_pdf).pack(side=tk.RIGHT, padx=(6, 0))

        self.summary_var = tk.StringVar(value="Select a project and enter PTI inputs.")
        ttk.Label(right, textvariable=self.summary_var).pack(anchor=tk.W, padx=8, pady=(8, 4))

        ttk.Label(right, text="Suction Profiles").pack(anchor=tk.W, padx=8, pady=(2, 2))
        self.suction_canvas = tk.Canvas(right, height=210, bg="#f7fbff", highlightthickness=0)
        self.suction_canvas.pack(fill=tk.X, padx=8, pady=(0, 6))
        self.suction_canvas.bind("<Configure>", lambda _e: self._draw_suction_graph(self._last_computed))

        out_grid = ttk.Frame(right)
        out_grid.pack(fill=tk.X, padx=8, pady=4)
        outputs = [
            ("pi", "PI"),
            ("gamma0_mean", "Gamma0 Mean"),
            ("gamma_h_shrink", "GammaH Shrink"),
            ("gamma_h_swell", "GammaH Swell"),
            ("alpha_mean", "Alpha Mean"),
            ("ym_center_shrink_in", "Ym Center (in)"),
            ("em_center_ft", "Em Center (ft)"),
            ("active_depth_ft", "Active Depth (ft)"),
        ]
        r = 0
        for key, label in outputs:
            ttk.Label(out_grid, text=label).grid(row=r, column=0, sticky=tk.W, padx=4, pady=2)
            v = tk.StringVar(value="")
            self.output_vars[key] = v
            ttk.Entry(out_grid, textvariable=v, width=16, state="readonly").grid(row=r, column=1, sticky=tk.W, padx=4, pady=2)
            r += 1

        ttk.Label(right, text="Shrink at Distance X from Edge").pack(anchor=tk.W, padx=8, pady=(8, 2))
        self.table = ttk.Treeview(right, columns=("x_ft", "x_cm", "ym_in", "ym_cm"), show="headings", height=12)
        for col, txt, w in [
            ("x_ft", "Distance (ft)", 100),
            ("x_cm", "Distance (cm)", 100),
            ("ym_in", "Shrink (in)", 100),
            ("ym_cm", "Shrink (cm)", 100),
        ]:
            self.table.heading(col, text=txt)
            self.table.column(col, width=w, anchor=tk.W)
        self.table.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

    def _on_left_configure(self, _event):
        self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))

    def _on_left_canvas_configure(self, event):
        self.left_canvas.itemconfigure(self.left_window_id, width=event.width)

    def _collect_payload(self):
        return {k: v.get().strip() for k, v in self.input_vars.items()}

    def _set_outputs(self, computed):
        self._last_computed = computed or {}
        for key, var in self.output_vars.items():
            val = computed.get(key) if computed else None
            var.set("" if val is None else f"{val}")
        for i in self.table.get_children():
            self.table.delete(i)
        if not computed:
            return
        x_ft = computed.get("distances_ft") or []
        x_cm = computed.get("distances_cm") or []
        ym_in = computed.get("ym_profile_in") or []
        ym_cm = computed.get("ym_profile_cm") or []
        n = min(len(x_ft), len(x_cm), len(ym_in), len(ym_cm))
        for i in range(n):
            self.table.insert("", tk.END, values=(x_ft[i], x_cm[i], ym_in[i], ym_cm[i]))
        self._draw_suction_graph(computed)

    def _draw_suction_graph(self, computed):
        c = self.suction_canvas
        if not c or not c.winfo_exists():
            return
        c.delete("all")
        data = computed or {}
        depth_ft = data.get("suction_depth_ft") or []
        wet_pf = data.get("suction_wet_pf") or []
        dry_pf = data.get("suction_dry_pf") or []
        const_pf = data.get("suction_const_pf") or []
        if not depth_ft:
            c.create_text(12, 12, anchor=tk.NW, text="No suction profile data yet.", fill="#35587a")
            return

        w = max(360, c.winfo_width())
        h = max(160, c.winfo_height())
        left, right, top, bottom = 56, 20, 18, 34
        gx = left
        gy = top
        gw = max(120, w - left - right)
        gh = max(90, h - top - bottom)
        c.create_rectangle(gx, gy, gx + gw, gy + gh, outline="#6b8aa8")

        z_min = min(depth_ft)
        z_max = max(depth_ft)
        if abs(z_max - z_min) <= 1e-9:
            return
        all_s = [v for v in (wet_pf + dry_pf + const_pf) if isinstance(v, (int, float))]
        if not all_s:
            return
        s_min = min(all_s)
        s_max = max(all_s)
        s_pad = max(0.25, 0.1 * (s_max - s_min if s_max > s_min else 1.0))
        x_lo = s_min - s_pad
        x_hi = s_max + s_pad

        def pxy(suction, depth):
            px = gx + ((suction - x_lo) / (x_hi - x_lo)) * gw
            py = gy + gh - ((depth - z_min) / (z_max - z_min)) * gh
            return px, py

        for i in range(7):
            sv = x_lo + (i / 6.0) * (x_hi - x_lo)
            px, _ = pxy(sv, z_min)
            c.create_line(px, gy, px, gy + gh, fill="#d4e1ee")
            c.create_text(px, gy + gh + 11, text=f"{sv:.2f}".rstrip("0").rstrip("."), anchor=tk.N, fill="#23496f", font=("Segoe UI", 8))
        for i in range(6):
            zv = z_min + (i / 5.0) * (z_max - z_min)
            _, py = pxy(x_lo, zv)
            c.create_line(gx, py, gx + gw, py, fill="#d4e1ee")
            c.create_text(gx - 6, py, text=f"{zv:.1f}".rstrip("0").rstrip("."), anchor=tk.E, fill="#23496f", font=("Segoe UI", 8))

        def draw_profile(values, color):
            pts = []
            for z, s in zip(depth_ft, values):
                if not isinstance(s, (int, float)):
                    continue
                px, py = pxy(float(s), float(z))
                pts.extend([px, py])
            if len(pts) >= 4:
                c.create_line(*pts, fill=color, width=2)
            for i in range(0, len(pts), 2):
                c.create_oval(pts[i] - 2, pts[i + 1] - 2, pts[i] + 2, pts[i + 1] + 2, fill=color, outline="")

        draw_profile(wet_pf, "#1f73b8")
        draw_profile(dry_pf, "#ba3f38")
        draw_profile(const_pf, "#2c8a46")

        c.create_text(gx + gw / 2, gy + gh + 24, text="Suction (pF)", fill="#8a1f1f", font=("Segoe UI", 9, "bold"))
        c.create_text(18, gy + gh / 2, text="Depth (ft)", angle=90, fill="#8a1f1f", font=("Segoe UI", 9, "bold"))

        legend = [("Initial suction", "#1f73b8"), ("Final suction", "#ba3f38"), ("Constant suction", "#2c8a46")]
        lx = gx + 4
        ly = gy + 8
        for label, color in legend:
            c.create_line(lx, ly, lx + 16, ly, fill=color, width=2)
            c.create_text(lx + 21, ly, text=label, anchor=tk.W, fill="#3a5773", font=("Segoe UI", 8))
            lx += 150

    def refresh(self):
        project_id = self.get_project_id()
        if not project_id:
            self.summary_var.set("Select a project and enter PTI inputs.")
            self._set_outputs({})
            return
        conn = get_connection()
        row = conn.execute(
            """
            SELECT payload_json, computed_json
            FROM calculations_runs
            WHERE project_id = ? AND calc_key = 'pti_shrink'
            """,
            (project_id,),
        ).fetchone()
        proj = conn.execute(
            "SELECT file_number, job_name FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()
        conn.close()

        data = default_payload()
        computed = {}
        if row:
            try:
                data.update(json.loads(row["payload_json"] or "{}"))
            except Exception:
                pass
            try:
                computed = json.loads(row["computed_json"] or "{}")
            except Exception:
                computed = {}

        data["project_title"] = data.get("project_title") or (proj["job_name"] if proj else "")
        data["project_number"] = data.get("project_number") or (proj["file_number"] if proj else "")
        for k, v in self.input_vars.items():
            v.set(str(data.get(k, "")))

        self.summary_var.set("PTI-style shrink/swell distortion estimate from lab + suction profile inputs.")
        self._set_outputs(computed)
        self._compute_preview()

    def _compute_preview(self):
        project_id = self.get_project_id()
        if not project_id:
            return
        payload = self._collect_payload()
        computed = compute_pti(payload)
        self._set_outputs(computed)

    def _compute_and_save(self):
        project_id = self.get_project_id()
        if not project_id:
            messagebox.showerror("No Project", "Select a project first.")
            return
        payload = self._collect_payload()
        computed = compute_pti(payload)
        conn = get_connection()
        conn.execute(
            """
            INSERT INTO calculations_runs (project_id, calc_key, payload_json, computed_json, updated_at)
            VALUES (?, 'pti_shrink', ?, ?, ?)
            ON CONFLICT(project_id) DO UPDATE SET
                calc_key = excluded.calc_key,
                payload_json = excluded.payload_json,
                computed_json = excluded.computed_json,
                updated_at = excluded.updated_at
            """,
            (project_id, json.dumps(payload), json.dumps(computed), now_iso()),
        )
        conn.commit()
        conn.close()
        self._set_outputs(computed)
        messagebox.showinfo("Saved", "PTI shrink/swell calculation saved.")

    def _export_pdf(self):
        project_id = self.get_project_id()
        if not project_id:
            messagebox.showerror("No Project", "Select a project first.")
            return
        payload = self._collect_payload()
        computed = compute_pti(payload)
        conn = get_connection()
        project = conn.execute(
            "SELECT file_number, job_name FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()
        conn.close()
        if not project:
            messagebox.showerror("Missing", "Project not found.")
            return
        default_name = f"PTI_Shrink_{project['file_number']}.pdf"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_name)
        if not path:
            return
        try:
            export_pti_pdf(path, dict(project), payload, computed)
        except Exception as exc:
            messagebox.showerror("Export Failed", f"Failed to export PDF: {exc}")
            return
        messagebox.showinfo("Exported", f"Calculation exported to:\n{path}")
