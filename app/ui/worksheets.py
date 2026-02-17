import json
import math
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app.db import get_connection, now_iso
from app.services.worksheet_d1557 import (
    calculate_d1557,
    compute_d1557_rows,
    export_d1557_pdf,
    extract_points,
)
from app.services.worksheet_generic import (
    compute_values as compute_generic_values,
    dumps_payload,
    export_generic_pdf,
    export_grain_size_pdf,
    grain_size_default_payload,
    grain_size_section_flags,
    compute_grain_size,
    grain_curve_points,
    get_spec,
    loads_payload,
    map_results,
)

GRAIN_TESTS = {"-200 Washed Sieve", "Sieve Part. Analysis", "Hydrometer"}
D1557_LIKE_TESTS = {"Max Density", "698 Max", "C Max"}


class WorksheetsTab(ttk.Frame):
    def __init__(self, parent, get_project_id, on_saved=None):
        super().__init__(parent)
        self.get_project_id = get_project_id
        self.on_saved = on_saved
        self.test_cols = 6
        self.current_test_name = None
        self.current_sample_id = None
        self.current_spec = None
        self.current_mode = "none"
        self.grain_include_wash = False
        self.grain_include_dry_sieve = False
        self.grain_include_hydrometer = False
        self.grain_hydrometer_var = tk.StringVar(value="no")
        self.grain_graph_window = None
        self.grain_graph_canvas = None
        self.grain_graph_info_var = tk.StringVar(value="")
        self.generic_input_vars = {}
        self.generic_comp_vars = {}
        self._build_ui()

    def _build_ui(self):
        wrap = ttk.Frame(self)
        wrap.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left = ttk.LabelFrame(wrap, text="Scheduled Worksheets")
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        right = ttk.LabelFrame(wrap, text="Worksheet Editor")
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        right.pack_propagate(False)

        self.tree = ttk.Treeview(
            left,
            columns=("test", "sample", "depth", "status", "metric1", "metric2"),
            show="headings",
        )
        for col, label, width in [
            ("test", "Worksheet/Test", 170),
            ("sample", "Sample", 120),
            ("depth", "Depth", 120),
            ("status", "Status", 100),
            ("metric1", "Result 1", 130),
            ("metric2", "Result 2", 130),
        ]:
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width, anchor=tk.W)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.tag_configure("completed", foreground="#6b7b8a")

        self.mode_var = tk.StringVar(value="Select a scheduled worksheet item.")
        ttk.Label(right, textvariable=self.mode_var).pack(anchor=tk.W, padx=8, pady=(6, 4))

        scroll_wrap = ttk.Frame(right)
        scroll_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        self.editor_canvas = tk.Canvas(scroll_wrap, highlightthickness=0, bg="#ecf4fb")
        self.editor_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.editor_scroll = ttk.Scrollbar(scroll_wrap, orient=tk.VERTICAL, command=self.editor_canvas.yview)
        self.editor_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.editor_canvas.configure(yscrollcommand=self.editor_scroll.set)

        self.editor_host = ttk.Frame(self.editor_canvas)
        self.editor_host_id = self.editor_canvas.create_window((0, 0), window=self.editor_host, anchor="nw")
        self.editor_host.pack_propagate(True)
        self.editor_host.bind("<Configure>", self._on_editor_configure)
        self.editor_canvas.bind("<Configure>", self._on_canvas_configure)
        self.editor_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.d1557_frame = ttk.Frame(self.editor_host)
        self._build_d1557_editor(self.d1557_frame)

        self.generic_frame = ttk.Frame(self.editor_host)
        self._build_generic_editor(self.generic_frame)

        btns = ttk.Frame(right)
        btns.pack(fill=tk.X, padx=8, pady=(2, 8))
        action_bg = "#d6e8f9"
        action_fg = "#0f2e4a"
        self.group_export_btn = tk.Button(
            btns,
            text="Export Project (Same Test)",
            command=self._export_project_grouped,
            bg=action_bg,
            fg=action_fg,
            relief=tk.FLAT,
            padx=10,
            pady=6,
        )
        self.group_export_btn.pack(side=tk.RIGHT, padx=(8, 0))
        self.export_btn = tk.Button(
            btns,
            text="Export Worksheet PDF",
            command=self._export_pdf,
            bg=action_bg,
            fg=action_fg,
            relief=tk.FLAT,
            padx=10,
            pady=6,
        )
        self.export_btn.pack(side=tk.RIGHT, padx=(8, 0))
        self.compute_btn = tk.Button(
            btns,
            text="Compute + Save",
            command=self._compute_and_save,
            bg=action_bg,
            fg=action_fg,
            relief=tk.FLAT,
            padx=10,
            pady=6,
        )
        self.compute_btn.pack(side=tk.RIGHT, padx=(8, 0))
        self.graph_btn = tk.Button(
            btns,
            text="Open Sieve Graph",
            command=self._open_grain_graph,
            bg=action_bg,
            fg=action_fg,
            relief=tk.FLAT,
            padx=10,
            pady=6,
            state=tk.DISABLED,
        )
        self.graph_btn.pack(side=tk.RIGHT, padx=(8, 0))

        self._set_editor_mode("none")

    def _on_editor_configure(self, _event):
        self.editor_canvas.configure(scrollregion=self.editor_canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.editor_canvas.itemconfigure(self.editor_host_id, width=event.width)

    def _on_mousewheel(self, event):
        if self.editor_canvas.winfo_height() <= 0:
            return
        self.editor_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _build_d1557_editor(self, parent):
        table_frame = ttk.Frame(parent)
        table_frame.pack(fill=tk.X, pady=4)
        self._build_d1557_grid(table_frame)

        g_frame = ttk.Frame(parent)
        g_frame.pack(fill=tk.X, pady=4)
        ttk.Label(g_frame, text="Zero-Air-Void G values (comma separated):").pack(side=tk.LEFT)
        self.g_values_var = tk.StringVar(value="2.65,2.70,2.75")
        ttk.Entry(g_frame, textvariable=self.g_values_var, width=20).pack(side=tk.LEFT, padx=6)

        self.calc_var = tk.StringVar(value="Computed: -")
        ttk.Label(parent, textvariable=self.calc_var).pack(anchor=tk.W, pady=4)

    def _build_generic_editor(self, parent):
        self.generic_spec_var = tk.StringVar(value="")
        ttk.Label(parent, textvariable=self.generic_spec_var).pack(anchor=tk.W, pady=(2, 6))

        self.generic_fields_container = ttk.Frame(parent)
        self.generic_fields_container.pack(fill=tk.BOTH, expand=True)

        self.generic_calc_var = tk.StringVar(value="Computed: -")
        ttk.Label(parent, textvariable=self.generic_calc_var).pack(anchor=tk.W, pady=4)

    def _build_d1557_grid(self, parent):
        self.raw_vars = {k: [] for k in ("A", "B", "D", "E", "F")}
        self.calc_vars = {k: [] for k in ("C", "G", "H", "I")}

        labels = {
            "A": "A Wt. Comp. Soil + Mold (gm)",
            "B": "B Wt. of Mold (gm)",
            "C": "C Net Wt. of Soil (gm)",
            "D": "D Wet Wt. Soil + Cont. (gm)",
            "E": "E Dry Wt. Soil + Cont. (gm)",
            "F": "F Wt. of Container (gm)",
            "G": "G Moisture Content (%)",
            "H": "H Wet Density (pcf)",
            "I": "I Dry Density (pcf)",
        }

        ttk.Label(parent, text="Field").grid(row=0, column=0, sticky=tk.W, padx=2, pady=2)
        for i in range(self.test_cols):
            ttk.Label(parent, text=f"Test {i+1}").grid(row=0, column=i + 1, padx=2, pady=2)

        row_keys = ["A", "B", "C", "D", "E", "F", "G", "H", "I"]
        for r, key in enumerate(row_keys, start=1):
            ttk.Label(parent, text=labels[key]).grid(row=r, column=0, sticky=tk.W, padx=2, pady=2)
            for c in range(self.test_cols):
                if key in self.raw_vars:
                    var = tk.StringVar()
                    self.raw_vars[key].append(var)
                    e = ttk.Entry(parent, textvariable=var, width=10)
                    e.grid(row=r, column=c + 1, padx=2, pady=2)
                    e.bind("<KeyRelease>", lambda _e: self._recompute_d1557())
                else:
                    var = tk.StringVar()
                    self.calc_vars[key].append(var)
                    ttk.Entry(parent, textvariable=var, width=10, state="readonly").grid(row=r, column=c + 1, padx=2, pady=2)

    def _set_editor_mode(self, mode):
        self.current_mode = mode
        for child in (self.d1557_frame, self.generic_frame):
            child.pack_forget()
        if mode == "d1557":
            self.d1557_frame.pack(fill=tk.BOTH, expand=True)
            self.export_btn.config(text="Export D1557 PDF")
            self.graph_btn.config(state=tk.DISABLED)
        elif mode in ("generic", "grain"):
            self.generic_frame.pack(fill=tk.BOTH, expand=True)
            self.export_btn.config(text="Export Worksheet PDF")
            self.graph_btn.config(state=tk.NORMAL if mode == "grain" else tk.DISABLED)
        else:
            self.export_btn.config(text="Export Worksheet PDF")
            self.graph_btn.config(state=tk.DISABLED)
        # Grouped export only applies to simple calculation tests.
        if self.current_test_name in self._groupable_tests():
            self.group_export_btn.config(state=tk.NORMAL)
        else:
            self.group_export_btn.config(state=tk.DISABLED)

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        project_id = self.get_project_id()
        if not project_id:
            return
        conn = get_connection()
        rows = conn.execute(
            """
            SELECT st.id, t.name AS test_name, s.sample_name, s.depth_raw, st.status,
                   w.max_dry_density, w.opt_moisture,
                   st.result_value, st.result_value2, st.result_unit2
            FROM sample_tests st
            JOIN samples s ON s.id = st.sample_id
            JOIN tests t ON t.id = st.test_id
            LEFT JOIN astm1557_runs w ON w.sample_test_id = st.id
            WHERE s.project_id = ?
              AND (st.status IS NULL OR st.status IN ('scheduled', 'in progress', 'completed'))
            ORDER BY s.sample_name
            """,
            (project_id,),
        ).fetchall()
        conn.close()
        for r in rows:
            tags = ("completed",) if (r["status"] or "").lower() == "completed" else ()
            metric1 = r["max_dry_density"] if r["max_dry_density"] is not None else r["result_value"]
            metric2 = r["opt_moisture"] if r["opt_moisture"] is not None else r["result_value2"]
            if metric2 is None and (r["result_unit2"] or "").strip():
                metric2 = r["result_unit2"]
            self.tree.insert(
                "",
                tk.END,
                iid=r["id"],
                tags=tags,
                values=(
                    r["test_name"],
                    r["sample_name"],
                    r["depth_raw"] or "",
                    r["status"],
                    "" if metric1 is None else f"{metric1}",
                    "" if metric2 is None else f"{metric2}",
                ),
            )

    def _selected_sample_test(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return int(sel[0])

    def _on_select(self, _event):
        sid = self._selected_sample_test()
        if not sid:
            return
        conn = get_connection()
        test_row = conn.execute(
            """
            SELECT t.name AS test_name, st.sample_id
            FROM sample_tests st JOIN tests t ON t.id = st.test_id
            WHERE st.id = ?
            """,
            (sid,),
        ).fetchone()
        d1557_row = conn.execute(
            "SELECT points_json FROM astm1557_runs WHERE sample_test_id = ?",
            (sid,),
        ).fetchone()
        generic_row = conn.execute(
            "SELECT worksheet_key, payload_json FROM worksheet_runs WHERE sample_test_id = ?",
            (sid,),
        ).fetchone()
        grain_row = None
        sample_test_names = []
        if test_row:
            sample_test_names = [
                r["name"]
                for r in conn.execute(
                    """
                    SELECT t.name
                    FROM sample_tests st
                    JOIN tests t ON t.id = st.test_id
                    WHERE st.sample_id = ?
                    """,
                    (test_row["sample_id"],),
                ).fetchall()
            ]
            grain_row = conn.execute(
                "SELECT payload_json FROM grain_size_runs WHERE sample_id = ?",
                (test_row["sample_id"],),
            ).fetchone()
        conn.close()

        self.current_test_name = test_row["test_name"] if test_row else None
        self.current_sample_id = int(test_row["sample_id"]) if test_row and test_row["sample_id"] is not None else None
        self.current_spec = get_spec(self.current_test_name) if self.current_test_name else None

        self._clear_d1557_grid()
        self._clear_generic_grid()
        if not test_row:
            self._set_editor_mode("none")
            return

        if self.current_test_name in D1557_LIKE_TESTS:
            meta = self._d1557_meta(self.current_test_name)
            self._set_editor_mode("d1557")
            self.mode_var.set(f"{meta['astm']} worksheet mode (A/B/D/E/F -> C/G/H/I).")
            if d1557_row and d1557_row["points_json"]:
                self._load_saved_d1557_json(d1557_row["points_json"])
                self._recompute_d1557()
            else:
                self.calc_var.set("Computed: -")
            return

        if self.current_test_name in GRAIN_TESTS:
            has_wash = "-200 Washed Sieve" in sample_test_names
            has_sieve = "Sieve Part. Analysis" in sample_test_names
            has_hydrometer = "Hydrometer" in sample_test_names
            include_wash, include_dry, include_hydro = grain_size_section_flags(has_wash, has_sieve, has_hydrometer)
            payload = None
            if grain_row and grain_row["payload_json"]:
                payload = loads_payload(grain_row["payload_json"])
                if self._grain_hydrometer_enabled(payload):
                    include_hydro = True
            self.grain_include_wash = include_wash
            self.grain_include_dry_sieve = include_dry
            self.grain_include_hydrometer = include_hydro
            self._set_editor_mode("grain")
            self.mode_var.set("Combined Grain Size worksheet mode.")
            self._render_generic_fields(self._grain_spec(include_wash, include_dry, include_hydro))
            if payload is None:
                payload = grain_size_default_payload()
            if "hydro_enabled" not in payload:
                payload["hydro_enabled"] = "yes" if include_hydro else "no"
            self._load_generic_payload(payload)
            self._recompute_generic()
            return

        if self.current_spec:
            self._set_editor_mode("generic")
            self.mode_var.set(f"{self.current_test_name} worksheet mode.")
            self._render_generic_fields(self.current_spec)
            if generic_row and generic_row["payload_json"]:
                self._load_generic_payload(loads_payload(generic_row["payload_json"]))
            self._recompute_generic()
            return

        self._set_editor_mode("none")
        self.mode_var.set(f"{self.current_test_name}: worksheet form not yet implemented.")

    def _clear_d1557_grid(self):
        for key in self.raw_vars:
            for v in self.raw_vars[key]:
                v.set("")
        for key in self.calc_vars:
            for v in self.calc_vars[key]:
                v.set("")

    def _load_saved_d1557_json(self, text):
        try:
            data = json.loads(text)
        except Exception:
            return
        if isinstance(data, dict) and "tests" in data:
            tests = data.get("tests", [])
            gvals = data.get("g_values")
            if gvals:
                self.g_values_var.set(",".join(str(x) for x in gvals))
            for idx, row in enumerate(tests[: self.test_cols]):
                for key in ("A", "B", "D", "E", "F"):
                    val = row.get(key)
                    if val is not None:
                        self.raw_vars[key][idx].set(f"{val}")

    def _collect_d1557_raw_rows(self):
        rows = []
        for i in range(self.test_cols):
            rows.append(
                {
                    "A": self.raw_vars["A"][i].get().strip(),
                    "B": self.raw_vars["B"][i].get().strip(),
                    "D": self.raw_vars["D"][i].get().strip(),
                    "E": self.raw_vars["E"][i].get().strip(),
                    "F": self.raw_vars["F"][i].get().strip(),
                }
            )
        return rows

    def _recompute_d1557(self):
        try:
            rows = compute_d1557_rows(self._collect_d1557_raw_rows())
        except Exception:
            return
        for i, r in enumerate(rows):
            self.calc_vars["C"][i].set("" if r["C"] is None else f"{r['C']:.2f}")
            self.calc_vars["G"][i].set("" if r["G"] is None else f"{r['G']:.2f}")
            self.calc_vars["H"][i].set("" if r["H"] is None else f"{r['H']:.2f}")
            self.calc_vars["I"][i].set("" if r["I"] is None else f"{r['I']:.2f}")
        points = extract_points(rows)
        calc = calculate_d1557(points)
        self.calc_var.set(
            f"Computed: Max Dry Density={calc.get('max_dry_density') or '-'} pcf, "
            f"Opt Moisture={calc.get('opt_moisture') or '-'} %"
        )

    def _render_generic_fields(self, spec):
        for child in self.generic_fields_container.winfo_children():
            child.destroy()
        self.generic_input_vars = {}
        self.generic_comp_vars = {}
        self.generic_spec_var.set(f"{spec['title']} ({spec['astm']})")
        row = 0
        ttk.Label(self.generic_fields_container, text="Field").grid(row=row, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Label(self.generic_fields_container, text="Value").grid(row=row, column=1, sticky=tk.W, padx=4, pady=2)
        ttk.Label(self.generic_fields_container, text="Unit").grid(row=row, column=2, sticky=tk.W, padx=4, pady=2)
        row += 1

        for raw_field in spec.get("fields", []):
            key, label, unit = raw_field[0], raw_field[1], raw_field[2]
            default_val = raw_field[3] if len(raw_field) > 3 else ""
            readonly = bool(raw_field[4]) if len(raw_field) > 4 else False
            emphasis = bool(raw_field[5]) if len(raw_field) > 5 else False
            ttk.Label(self.generic_fields_container, text=label).grid(row=row, column=0, sticky=tk.W, padx=4, pady=2)
            var = tk.StringVar()
            self.generic_input_vars[key] = var
            if default_val not in (None, ""):
                var.set(str(default_val))
            if key == "hydro_enabled":
                self.grain_hydrometer_var = var
                cb = ttk.Combobox(
                    self.generic_fields_container,
                    textvariable=var,
                    values=["no", "yes"],
                    width=20,
                    state="readonly",
                )
                cb.grid(row=row, column=1, sticky=tk.W, padx=4, pady=2)
                cb.bind("<<ComboboxSelected>>", self._on_hydrometer_toggle)
            else:
                entry_font = ("Segoe UI", 9, "bold italic") if emphasis else ("Segoe UI", 9)
                ent = tk.Entry(
                    self.generic_fields_container,
                    textvariable=var,
                    width=22,
                    font=entry_font,
                    fg="#15385b" if not emphasis else "#1c4f7f",
                    bg="#f8fbff",
                    relief=tk.SOLID,
                    bd=1,
                    disabledforeground="#1c4f7f",
                )
                ent.grid(row=row, column=1, sticky=tk.W, padx=4, pady=2)
                if readonly:
                    ent.configure(state="readonly")
                else:
                    ent.bind("<KeyRelease>", lambda _e: self._recompute_generic())
            ttk.Label(self.generic_fields_container, text=unit or "").grid(row=row, column=2, sticky=tk.W, padx=4, pady=2)
            row += 1

        if spec.get("computed"):
            ttk.Separator(self.generic_fields_container, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=3, sticky=tk.EW, padx=2, pady=5)
            row += 1
            for key, label, unit in spec.get("computed", []):
                ttk.Label(self.generic_fields_container, text=label).grid(row=row, column=0, sticky=tk.W, padx=4, pady=2)
                var = tk.StringVar()
                self.generic_comp_vars[key] = var
                ttk.Entry(self.generic_fields_container, textvariable=var, width=20, state="readonly").grid(
                    row=row, column=1, sticky=tk.W, padx=4, pady=2
                )
                ttk.Label(self.generic_fields_container, text=unit or "").grid(row=row, column=2, sticky=tk.W, padx=4, pady=2)
                row += 1

    def _clear_generic_grid(self):
        for var in self.generic_input_vars.values():
            var.set("")
        for var in self.generic_comp_vars.values():
            var.set("")
        self.generic_calc_var.set("Computed: -")

    def _collect_generic_payload(self):
        payload = {}
        for k, var in self.generic_input_vars.items():
            payload[k] = var.get().strip()
        return payload

    def _load_generic_payload(self, payload):
        for k, var in self.generic_input_vars.items():
            if k in payload:
                var.set(str(payload.get(k, "")))

    def _recompute_generic(self):
        if self.current_mode == "grain":
            payload = self._collect_generic_payload()
            include_hydro = self._grain_hydrometer_enabled(payload)
            if include_hydro != self.grain_include_hydrometer:
                if include_hydro and not self.grain_include_dry_sieve:
                    self.grain_include_hydrometer = False
                    if "hydro_enabled" in self.generic_input_vars:
                        self.generic_input_vars["hydro_enabled"].set("no")
                    self.generic_calc_var.set("Computed: Hydrometer requires dry sieve data.")
                    return
                self.grain_include_hydrometer = include_hydro
                self._render_generic_fields(
                    self._grain_spec(
                        self.grain_include_wash,
                        self.grain_include_dry_sieve,
                        self.grain_include_hydrometer,
                    )
                )
                self._load_generic_payload(payload)
                payload = self._collect_generic_payload()
            computed = compute_grain_size(payload)
            for key, var in self.generic_comp_vars.items():
                val = computed.get(key)
                if isinstance(val, float):
                    var.set(f"{val:.3f}".rstrip("0").rstrip("."))
                else:
                    var.set("" if val is None else str(val))
            parts = []
            for key, label, _unit in self._grain_spec(
                self.grain_include_wash,
                self.grain_include_dry_sieve,
                self.grain_include_hydrometer,
            ).get("computed", []):
                val = computed.get(key)
                if val is None or val == "":
                    continue
                txt = f"{val:.3f}".rstrip("0").rstrip(".") if isinstance(val, float) else str(val)
                parts.append(f"{label}: {txt}")
            self.generic_calc_var.set("Computed: " + (" | ".join(parts) if parts else "-"))
            self._refresh_grain_graph()
            return

        if not self.current_test_name or not self.current_spec:
            self.generic_calc_var.set("Computed: -")
            return
        payload = self._collect_generic_payload()
        computed = compute_generic_values(self.current_test_name, payload)
        for key, var in self.generic_comp_vars.items():
            val = computed.get(key)
            if isinstance(val, float):
                var.set(f"{val:.3f}".rstrip("0").rstrip("."))
            else:
                var.set("" if val is None else str(val))
        if computed:
            parts = []
            for key, label, _unit in self.current_spec.get("computed", []):
                val = computed.get(key)
                if val is None or val == "":
                    continue
                if isinstance(val, float):
                    txt = f"{val:.3f}".rstrip("0").rstrip(".")
                else:
                    txt = str(val)
                parts.append(f"{label}: {txt}")
            self.generic_calc_var.set("Computed: " + (" | ".join(parts) if parts else "-"))
        else:
            self.generic_calc_var.set("Computed: -")

    def _current_grain_points(self):
        if self.current_mode != "grain":
            return []
        payload = self._collect_generic_payload()
        computed = compute_grain_size(payload)
        include_hydro = self._grain_hydrometer_enabled(payload) and self.grain_include_dry_sieve
        return grain_curve_points(payload, self.grain_include_dry_sieve, include_hydro, computed)

    def _grain_x_bounds(self, points):
        if not points:
            return 0.001, 100.0
        minx = max(min(p[0] for p in points), 0.001)
        maxx = max(max(p[0] for p in points), minx * 10.0)
        return minx, maxx

    def _grain_log_to_px(self, x_mm, left, width, log_min, log_max):
        lx = math.log10(max(x_mm, 0.001))
        return left + ((log_max - lx) / (log_max - log_min)) * width

    def _grain_px_to_x(self, px, left, width, log_min, log_max):
        frac = (px - left) / width
        frac = max(0.0, min(1.0, frac))
        lx = log_max - frac * (log_max - log_min)
        return 10 ** lx

    def _grain_interpolate_y(self, x_mm, points):
        ordered = sorted(points, key=lambda p: p[0], reverse=True)
        if not ordered:
            return None
        if x_mm >= ordered[0][0]:
            return ordered[0][1]
        if x_mm <= ordered[-1][0]:
            return ordered[-1][1]
        for i in range(len(ordered) - 1):
            x1, y1 = ordered[i]
            x2, y2 = ordered[i + 1]
            if x1 >= x_mm >= x2 and x1 > 0 and x2 > 0:
                lx1 = math.log10(x1)
                lx2 = math.log10(x2)
                lxm = math.log10(x_mm)
                if abs(lx2 - lx1) <= 1e-12:
                    return y1
                t = (lxm - lx1) / (lx2 - lx1)
                return y1 + t * (y2 - y1)
        return None

    def _open_grain_graph(self):
        if self.current_mode != "grain":
            messagebox.showerror("Unavailable", "Open a grain-size worksheet first.")
            return
        if self.grain_graph_window and self.grain_graph_window.winfo_exists():
            self.grain_graph_window.deiconify()
            self.grain_graph_window.lift()
            self._refresh_grain_graph()
            return
        win = tk.Toplevel(self)
        win.title("Interactive Sieve Graph")
        win.geometry("920x580")
        self.grain_graph_window = win
        self.grain_graph_info_var.set("Hover over the curve to read particle size and % finer.")
        info = ttk.Label(win, textvariable=self.grain_graph_info_var)
        info.pack(fill=tk.X, padx=10, pady=(8, 2))
        canvas = tk.Canvas(win, bg="#f7fbff", highlightthickness=0)
        canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=(2, 10))
        self.grain_graph_canvas = canvas
        canvas.bind("<Configure>", lambda _e: self._refresh_grain_graph())
        canvas.bind("<Motion>", self._on_grain_graph_hover)
        canvas.bind("<Leave>", self._on_grain_graph_leave)
        win.protocol("WM_DELETE_WINDOW", self._close_grain_graph)
        self._refresh_grain_graph()

    def _close_grain_graph(self):
        if self.grain_graph_window and self.grain_graph_window.winfo_exists():
            self.grain_graph_window.destroy()
        self.grain_graph_window = None
        self.grain_graph_canvas = None

    def _refresh_grain_graph(self):
        c = self.grain_graph_canvas
        if not c or not c.winfo_exists():
            return
        c.delete("all")
        points = self._current_grain_points()
        if not points:
            c.create_text(20, 20, anchor=tk.NW, text="No sieve/hydrometer curve points yet.", fill="#1b3d63")
            return
        w = max(300, c.winfo_width())
        h = max(220, c.winfo_height())
        left, right, top, bottom = 78, 24, 26, 54
        gx = left
        gy = top
        gw = max(100, w - left - right)
        gh = max(100, h - top - bottom)
        minx, maxx = self._grain_x_bounds(points)
        log_min = math.log10(minx)
        log_max = math.log10(maxx)
        c.create_rectangle(gx, gy, gx + gw, gy + gh, outline="#6b8aa8")
        for yp in range(0, 101, 10):
            py = gy + gh - (yp / 100.0) * gh
            c.create_line(gx, py, gx + gw, py, fill="#d4e1ee")
            c.create_text(gx - 8, py, text=f"{yp}", anchor=tk.E, fill="#23496f", font=("Segoe UI", 8))
        tick_sizes = [100, 50, 25, 10, 4.75, 2.0, 0.85, 0.425, 0.25, 0.15, 0.075, 0.02, 0.005]
        for s in tick_sizes:
            if s < minx or s > maxx:
                continue
            px = self._grain_log_to_px(s, gx, gw, log_min, log_max)
            c.create_line(px, gy, px, gy + gh, fill="#d4e1ee")
            c.create_text(px, gy + gh + 12, text=f"{s:g}", anchor=tk.N, fill="#23496f", font=("Segoe UI", 8))
        ordered = sorted(points, key=lambda p: p[0], reverse=True)
        poly = []
        for x_mm, y_pf in ordered:
            px = self._grain_log_to_px(x_mm, gx, gw, log_min, log_max)
            py = gy + gh - (max(0.0, min(100.0, y_pf)) / 100.0) * gh
            poly.extend([px, py])
        c.create_line(*poly, fill="#0f4f8a", width=2, smooth=False, tags="curve")
        for i in range(0, len(poly), 2):
            c.create_oval(poly[i] - 2, poly[i + 1] - 2, poly[i] + 2, poly[i + 1] + 2, fill="#0f4f8a", outline="")
        c.create_text(gx + gw / 2, gy + gh + 30, text="Particle Size (mm, log scale)", fill="#8a1f1f", font=("Segoe UI", 9, "bold"))
        c.create_text(24, gy + gh / 2, text="Percent Finer (%)", angle=90, fill="#8a1f1f", font=("Segoe UI", 9, "bold"))
        c.create_text(gx + 6, gy + 8, text="Hover curve for exact values", anchor=tk.NW, fill="#4f6f8f", font=("Segoe UI", 8))
        c._grain_meta = {"gx": gx, "gy": gy, "gw": gw, "gh": gh, "log_min": log_min, "log_max": log_max, "points": ordered}

    def _on_grain_graph_hover(self, event):
        c = self.grain_graph_canvas
        if not c or not c.winfo_exists() or not hasattr(c, "_grain_meta"):
            return
        meta = c._grain_meta
        gx, gy, gw, gh = meta["gx"], meta["gy"], meta["gw"], meta["gh"]
        x = event.x
        y = event.y
        c.delete("hover")
        if x < gx or x > gx + gw or y < gy or y > gy + gh:
            self.grain_graph_info_var.set("Hover over the curve to read particle size and % finer.")
            return
        x_mm = self._grain_px_to_x(x, gx, gw, meta["log_min"], meta["log_max"])
        y_interp = self._grain_interpolate_y(x_mm, meta["points"])
        if y_interp is None:
            return
        py_line = gy + gh - (max(0.0, min(100.0, y_interp)) / 100.0) * gh
        if abs(y - py_line) > 14:
            self.grain_graph_info_var.set("Move closer to the blue sieve line for a reading.")
            return
        c.create_line(x, gy, x, gy + gh, fill="#b54a4a", dash=(3, 3), tags="hover")
        c.create_line(gx, py_line, gx + gw, py_line, fill="#b54a4a", dash=(3, 3), tags="hover")
        c.create_oval(x - 4, py_line - 4, x + 4, py_line + 4, fill="#b54a4a", outline="", tags="hover")
        self.grain_graph_info_var.set(f"Particle Size: {x_mm:.4g} mm   |   Percent Finer: {y_interp:.2f}%")

    def _on_grain_graph_leave(self, _event):
        c = self.grain_graph_canvas
        if c and c.winfo_exists():
            c.delete("hover")
        self.grain_graph_info_var.set("Hover over the curve to read particle size and % finer.")

    def _parse_g_values(self):
        raw = self.g_values_var.get().strip()
        if not raw:
            return [2.65]
        values = []
        for part in raw.split(","):
            part = part.strip()
            if not part:
                continue
            values.append(float(part))
        return values or [2.65]

    def _selection_is_d1557(self, sample_test_id):
        conn = get_connection()
        row = conn.execute(
            """
            SELECT t.name AS test_name
            FROM sample_tests st
            JOIN tests t ON t.id = st.test_id
            WHERE st.id = ?
            """,
            (sample_test_id,),
        ).fetchone()
        conn.close()
        return bool(row and row["test_name"] in D1557_LIKE_TESTS)

    def _d1557_meta(self, test_name):
        if test_name == "698 Max":
            return {"astm": "ASTM D698", "file_prefix": "D698"}
        if test_name == "C Max":
            return {"astm": "ASTM D1557", "file_prefix": "CMAX_D1557"}
        return {"astm": "ASTM D1557", "file_prefix": "D1557"}

    def _compute_and_save(self):
        sid = self._selected_sample_test()
        if not sid:
            messagebox.showerror("No Selection", "Select a worksheet assignment first.")
            return
        if self._selection_is_d1557(sid):
            self._compute_save_d1557(sid)
            return
        if self.current_mode == "grain":
            self._compute_save_grain(sid)
            return
        if self.current_spec:
            self._compute_save_generic(sid)
            return
        messagebox.showerror("Not Implemented", "Worksheet form is not implemented for this test yet.")

    def _compute_save_d1557(self, sid):
        rows = compute_d1557_rows(self._collect_d1557_raw_rows())
        points = extract_points(rows)
        if len(points) < 2:
            messagebox.showerror("Need Data", "Enter enough A/B/D/E/F data to compute at least 2 points.")
            return
        calc = calculate_d1557(points)
        try:
            g_values = self._parse_g_values()
        except Exception:
            messagebox.showerror("Invalid G Values", "G values must be comma-separated numbers.")
            return

        payload = {
            "tests": [{"A": r["A"], "B": r["B"], "D": r["D"], "E": r["E"], "F": r["F"]} for r in rows],
            "g_values": g_values,
        }
        conn = get_connection()
        conn.execute(
            """
            INSERT INTO astm1557_runs (sample_test_id, points_json, max_dry_density, opt_moisture, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(sample_test_id) DO UPDATE SET
                points_json = excluded.points_json,
                max_dry_density = excluded.max_dry_density,
                opt_moisture = excluded.opt_moisture,
                updated_at = excluded.updated_at
            """,
            (sid, json.dumps(payload), calc.get("max_dry_density"), calc.get("opt_moisture"), now_iso()),
        )
        conn.execute(
            """
            UPDATE sample_tests
            SET result_value = ?, result_unit = 'pcf',
                result_value2 = ?, result_unit2 = '%',
                status = 'completed'
            WHERE id = ?
            """,
            (calc.get("max_dry_density"), calc.get("opt_moisture"), sid),
        )
        conn.commit()
        conn.close()
        self._recompute_d1557()
        self.refresh()
        self._reselect_and_notify(sid)

    def _compute_save_generic(self, sid):
        payload = self._collect_generic_payload()
        computed = compute_generic_values(self.current_test_name, payload)
        mapped = map_results(self.current_test_name, payload, computed)
        conn = get_connection()
        conn.execute(
            """
            INSERT INTO worksheet_runs (sample_test_id, worksheet_key, payload_json, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(sample_test_id) DO UPDATE SET
                worksheet_key = excluded.worksheet_key,
                payload_json = excluded.payload_json,
                updated_at = excluded.updated_at
            """,
            (sid, self.current_spec["key"], dumps_payload(payload), now_iso()),
        )
        conn.execute(
            """
            UPDATE sample_tests
            SET result_value = ?, result_unit = ?, result_value2 = ?, result_unit2 = ?,
                result_value3 = ?, result_unit3 = ?, result_value4 = ?, result_unit4 = ?,
                result_notes = ?, status = 'completed'
            WHERE id = ?
            """,
            (
                mapped["result_value"],
                mapped["result_unit"],
                mapped["result_value2"],
                mapped["result_unit2"],
                mapped["result_value3"],
                mapped["result_unit3"],
                mapped["result_value4"],
                mapped["result_unit4"],
                mapped["result_notes"],
                sid,
            ),
        )
        conn.commit()
        conn.close()
        self._recompute_generic()
        self.refresh()
        self._reselect_and_notify(sid)

    def _reselect_and_notify(self, sid):
        if self.tree.exists(str(sid)):
            self.tree.selection_set(str(sid))
            self.tree.focus(str(sid))
            self._on_select(None)
        if self.on_saved:
            self.on_saved()

    def _grain_hydrometer_enabled(self, payload):
        raw = str(payload.get("hydro_enabled", "")).strip().lower()
        return raw in {"yes", "y", "true", "1"}

    def _on_hydrometer_toggle(self, _event=None):
        if self.current_mode != "grain":
            return
        if self.grain_hydrometer_var.get().strip().lower() == "yes" and not self.grain_include_dry_sieve:
            self.grain_hydrometer_var.set("no")
            messagebox.showerror("Not Allowed", "Hydrometer requires dry sieve data.")
        self._recompute_generic()

    def _grain_test_cost(self, conn, sample_id, test_name):
        row = conn.execute(
            """
            SELECT COALESCE(tr.price, t.default_cost) AS cost
            FROM samples s
            JOIN projects p ON p.id = s.project_id
            JOIN tests t ON t.name = ?
            LEFT JOIN test_rates tr ON tr.rate_id = p.billing_rate_id AND tr.test_id = t.id
            WHERE s.id = ?
            """,
            (test_name, sample_id),
        ).fetchone()
        return float(row["cost"]) if row and row["cost"] is not None else 0.0

    def _compute_save_grain(self, sid):
        if not self.current_sample_id:
            messagebox.showerror("Missing", "Could not resolve sample for this worksheet.")
            return
        payload = self._collect_generic_payload()
        hydro_enabled = self._grain_hydrometer_enabled(payload)
        if hydro_enabled and not self.grain_include_dry_sieve:
            hydro_enabled = False
            payload["hydro_enabled"] = "no"
        computed = compute_grain_size(payload)
        conn = get_connection()
        conn.execute(
            """
            INSERT INTO grain_size_runs (sample_id, payload_json, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(sample_id) DO UPDATE SET
                payload_json = excluded.payload_json,
                updated_at = excluded.updated_at
            """,
            (self.current_sample_id, dumps_payload(payload), now_iso()),
        )
        hydro_test_row = conn.execute("SELECT id FROM tests WHERE name = 'Hydrometer'").fetchone()
        if hydro_test_row:
            hydro_test_id = int(hydro_test_row["id"])
            existing_hydro = conn.execute(
                "SELECT id FROM sample_tests WHERE sample_id = ? AND test_id = ?",
                (self.current_sample_id, hydro_test_id),
            ).fetchone()
            if hydro_enabled and not existing_hydro:
                hydro_cost = self._grain_test_cost(conn, self.current_sample_id, "Hydrometer")
                conn.execute(
                    """
                    INSERT INTO sample_tests (sample_id, test_id, cost, status)
                    VALUES (?, ?, ?, ?)
                    """,
                    (self.current_sample_id, hydro_test_id, hydro_cost, "scheduled"),
                )
            if not hydro_enabled and existing_hydro:
                conn.execute("DELETE FROM sample_tests WHERE id = ?", (existing_hydro["id"],))
        rows = conn.execute(
            """
            SELECT st.id, t.name
            FROM sample_tests st
            JOIN tests t ON t.id = st.test_id
            WHERE st.sample_id = ?
            """,
            (self.current_sample_id,),
        ).fetchall()
        for r in rows:
            test_name = r["name"]
            st_id = r["id"]
            if test_name == "-200 Washed Sieve":
                conn.execute(
                    """
                    UPDATE sample_tests
                    SET result_value = ?, result_unit = '%', result_notes = ?, status = 'completed'
                    WHERE id = ?
                    """,
                    (
                        computed.get("wash_o_passing200"),
                        None,
                        st_id,
                    ),
                )
            elif test_name == "Sieve Part. Analysis":
                pass_no200 = computed.get("sieve_pct_pass_no200")
                if pass_no200 is None:
                    pass_no200 = computed.get("wash_o_passing200")
                conn.execute(
                    """
                    UPDATE sample_tests
                    SET result_value = NULL, result_unit = ?, result_value2 = ?, result_unit2 = '%',
                        result_notes = ?, status = 'completed'
                    WHERE id = ?
                    """,
                    (
                        (payload.get("sieve_uscs_class") or "").strip(),
                        pass_no200,
                        "Combined grain-size worksheet",
                        st_id,
                    ),
                )
            elif test_name == "Hydrometer":
                if not hydro_enabled:
                    continue
                summary = computed.get("hydro_total_1440")
                if summary is None:
                    summary = computed.get("hydro_total_250")
                conn.execute(
                    """
                    UPDATE sample_tests
                    SET result_value = ?, result_unit = '%', result_notes = ?, status = 'completed'
                    WHERE id = ?
                    """,
                    (
                        summary,
                        "Hydrometer from combined grain-size worksheet",
                        st_id,
                    ),
                )
        conn.commit()
        conn.close()
        self._recompute_generic()
        self.refresh()
        self._reselect_and_notify(sid)

    def _export_pdf(self):
        sid = self._selected_sample_test()
        if not sid:
            messagebox.showerror("No Selection", "Select a worksheet assignment first.")
            return
        if self._selection_is_d1557(sid):
            self._export_d1557_pdf(sid)
            return
        if self.current_mode == "grain":
            self._export_grain_pdf(sid)
            return
        if self.current_spec:
            self._export_generic_pdf(sid)
            return
        messagebox.showerror("Not Implemented", "Worksheet PDF export is not implemented for this test.")

    def _worksheet_project_sample_row(self, sid):
        conn = get_connection()
        row = conn.execute(
            """
            SELECT p.file_number, p.job_name, s.sample_name, s.depth_raw
            FROM sample_tests st
            JOIN samples s ON s.id = st.sample_id
            JOIN projects p ON p.id = s.project_id
            WHERE st.id = ?
            """,
            (sid,),
        ).fetchone()
        conn.close()
        return row

    def _export_d1557_pdf(self, sid):
        rows = compute_d1557_rows(self._collect_d1557_raw_rows())
        points = extract_points(rows)
        if len(points) < 2:
            messagebox.showerror("Need Data", "Enter enough A/B/D/E/F data to compute at least 2 points.")
            return
        calc = calculate_d1557(points)
        try:
            g_values = self._parse_g_values()
        except Exception:
            messagebox.showerror("Invalid G Values", "G values must be comma-separated numbers.")
            return

        row = self._worksheet_project_sample_row(sid)
        if not row:
            messagebox.showerror("Missing", "Could not find selected worksheet record.")
            return
        meta = self._d1557_meta(self.current_test_name)
        sample_label = row["sample_name"] if not row["depth_raw"] else f"{row['sample_name']} @ {row['depth_raw']}"
        default_name = f"{meta['file_prefix']}_{row['file_number']}_{row['sample_name']}.pdf"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_name)
        if not path:
            return
        try:
            export_d1557_pdf(path, dict(row), sample_label, rows, calc, g_values, meta["astm"])
        except Exception as exc:
            messagebox.showerror("Export Failed", f"Failed to export PDF: {exc}")
            return
        messagebox.showinfo("Exported", f"Worksheet exported to:\n{path}")

    def _export_generic_pdf(self, sid):
        row = self._worksheet_project_sample_row(sid)
        if not row:
            messagebox.showerror("Missing", "Could not find selected worksheet record.")
            return
        sample_label = row["sample_name"] if not row["depth_raw"] else f"{row['sample_name']} @ {row['depth_raw']}"
        payload = self._collect_generic_payload()
        computed = compute_generic_values(self.current_test_name, payload)
        default_name = f"{self.current_spec['key']}_{row['file_number']}_{row['sample_name']}.pdf"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_name)
        if not path:
            return
        try:
            export_generic_pdf(path, dict(row), sample_label, self.current_test_name, payload, computed)
        except Exception as exc:
            messagebox.showerror("Export Failed", f"Failed to export PDF: {exc}")
            return
        messagebox.showinfo("Exported", f"Worksheet exported to:\n{path}")

    def _export_grain_pdf(self, sid):
        row = self._worksheet_project_sample_row(sid)
        if not row:
            messagebox.showerror("Missing", "Could not find selected worksheet record.")
            return
        payload = self._collect_generic_payload()
        computed = compute_grain_size(payload)
        include_hydro = self._grain_hydrometer_enabled(payload) and self.grain_include_dry_sieve
        sample_label = row["sample_name"] if not row["depth_raw"] else f"{row['sample_name']} @ {row['depth_raw']}"
        default_name = f"GrainSize_{row['file_number']}_{row['sample_name']}.pdf"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_name)
        if not path:
            return
        try:
            export_grain_size_pdf(
                path,
                dict(row),
                sample_label,
                payload,
                computed,
                self.grain_include_wash,
                self.grain_include_dry_sieve,
                include_hydro,
            )
        except Exception as exc:
            messagebox.showerror("Export Failed", f"Failed to export PDF: {exc}")
            return
        messagebox.showinfo("Exported", f"Worksheet exported to:\n{path}")

    def _groupable_tests(self):
        return {"Field Density/Moisture", "Expansion Index", "Sand Cone", "Moisture Content"}

    def _export_project_grouped(self):
        project_id = self.get_project_id()
        if not project_id:
            messagebox.showerror("No Project", "Select a project first.")
            return
        test_name = self.current_test_name
        if not test_name:
            messagebox.showerror("No Selection", "Select a worksheet row first.")
            return
        if test_name not in self._groupable_tests():
            messagebox.showerror("Not Supported", "Grouped export is only supported for simple calculation tests.")
            return
        from app.services.worksheet_generic import export_grouped_results_pdf

        conn = get_connection()
        project = conn.execute(
            "SELECT file_number, job_name FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()
        rows = conn.execute(
            """
            SELECT s.sample_name, s.depth_raw, st.result_value, st.result_unit,
                   st.result_value2, st.result_unit2, st.result_value3, st.result_unit3,
                   st.result_unit4
            FROM sample_tests st
            JOIN samples s ON s.id = st.sample_id
            JOIN tests t ON t.id = st.test_id
            WHERE s.project_id = ? AND t.name = ?
            ORDER BY s.sample_name, s.depth_raw
            """,
            (project_id, test_name),
        ).fetchall()
        conn.close()
        if not project:
            messagebox.showerror("Missing Project", "Project record not found.")
            return
        if not rows:
            messagebox.showerror("No Data", "No results found for this test in the current project.")
            return

        safe_test = self._safe_filename(test_name)
        default_name = f"{safe_test}_{project['file_number']}.pdf"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_name)
        if not path:
            return
        try:
            export_grouped_results_pdf(
                path=path,
                project=dict(project),
                test_name=test_name,
                rows=[dict(r) for r in rows],
            )
        except Exception as exc:
            messagebox.showerror("Export Failed", f"Grouped export failed: {exc}")
            return
        messagebox.showinfo("Exported", f"Grouped worksheet exported to:\n{path}")

    def _safe_filename(self, name):
        cleaned = "".join(ch for ch in name if ch.isalnum() or ch in (" ", "-", "_"))
        cleaned = cleaned.strip().replace(" ", "")
        return cleaned or "Export"

    def _grain_spec(self, include_wash, include_dry, include_hydro):
        fields = []
        computed = []
        fields.append(
            (
                "hydro_enabled",
                "Hydrometer (include in this worksheet)",
                "",
                "yes" if include_hydro else "no",
                False,
                False,
            )
        )
        if include_wash:
            fields.extend(
                [
                    ("wash_a_wet_tare", "A Wet Sample + Tare", "g"),
                    ("wash_b_dry_tare", "B Dry Sample + Tare", "g"),
                    ("wash_c_tare", "C Tare", "g"),
                    ("wash_e_moist_soil", "E Weight of Moist Soil", "g"),
                    ("wash_h_dry40_tare", "H Dry + Tare (#40)", "g"),
                    ("wash_i_tare40", "I Tare (#40)", "g"),
                    ("wash_k_dry200_tare", "K Dry + Tare (#200)", "g"),
                    ("wash_l_tare200", "L Tare (#200)", "g"),
                ]
            )
            computed.extend(
                [
                    ("wash_d_moisture", "D Moisture Content", "%"),
                    ("wash_f_dry_sample", "F Dry Sample Weight", "g"),
                    ("wash_m_weight", "M Weight", "g"),
                    ("wash_n_minus200", "N -200 Sieve Weight", "g"),
                    ("wash_o_passing200", "O % Passing #200", "%"),
                ]
            )
        if include_dry:
            fields.extend(
                [
                    ("sieve_uscs_class", "Dry Sieve USCS Classification", ""),
                    ("sieve_pre_3_4in", "Pre Test Sieve Wt 3/4\"", "g"),
                    ("sieve_post_3_4in", "Soil+Sieve Wt 3/4\"", "g"),
                    ("sieve_pre_3_8in", "Pre Test Sieve Wt 3/8\"", "g"),
                    ("sieve_post_3_8in", "Soil+Sieve Wt 3/8\"", "g"),
                    ("sieve_pre_no4", "Pre Test Sieve Wt #4", "g"),
                    ("sieve_post_no4", "Soil+Sieve Wt #4", "g"),
                    ("sieve_pre_no10", "Pre Test Sieve Wt #10", "g"),
                    ("sieve_post_no10", "Soil+Sieve Wt #10", "g"),
                    ("sieve_pre_no16", "Pre Test Sieve Wt #16", "g"),
                    ("sieve_post_no16", "Soil+Sieve Wt #16", "g"),
                    ("sieve_pre_no40", "Pre Test Sieve Wt #40", "g"),
                    ("sieve_post_no40", "Soil+Sieve Wt #40", "g"),
                    ("sieve_pre_no50", "Pre Test Sieve Wt #50", "g"),
                    ("sieve_post_no50", "Soil+Sieve Wt #50", "g"),
                    ("sieve_pre_no100", "Pre Test Sieve Wt #100", "g"),
                    ("sieve_post_no100", "Soil+Sieve Wt #100", "g"),
                    ("sieve_pre_no200", "Pre Test Sieve Wt #200", "g"),
                    ("sieve_post_no200", "Soil+Sieve Wt #200", "g"),
                ]
            )
            computed.extend(
                [
                    ("sieve_a_prewash_dry_weight", "A Prewash Dry Weight", "g"),
                    ("sieve_ret_3_4in", "Retained 3/4\"", "g"),
                    ("sieve_pct_pass_3_4in", "Percent Passing 3/4\"", "%"),
                    ("sieve_ret_3_8in", "Retained 3/8\"", "g"),
                    ("sieve_pct_pass_3_8in", "Percent Passing 3/8\"", "%"),
                    ("sieve_ret_no4", "Retained #4", "g"),
                    ("sieve_pct_pass_no4", "Percent Passing #4", "%"),
                    ("sieve_ret_no10", "Retained #10", "g"),
                    ("sieve_pct_pass_no10", "Percent Passing #10", "%"),
                    ("sieve_ret_no16", "Retained #16", "g"),
                    ("sieve_pct_pass_no16", "Percent Passing #16", "%"),
                    ("sieve_ret_no40", "Retained #40", "g"),
                    ("sieve_pct_pass_no40", "Percent Passing #40", "%"),
                    ("sieve_ret_no50", "Retained #50", "g"),
                    ("sieve_pct_pass_no50", "Percent Passing #50", "%"),
                    ("sieve_ret_no100", "Retained #100", "g"),
                    ("sieve_pct_pass_no100", "Percent Passing #100", "%"),
                    ("sieve_ret_no200", "Retained #200", "g"),
                    ("sieve_pct_pass_no200", "Percent Passing #200", "%"),
                ]
            )
        if include_hydro:
            fields.extend(
                [
                    ("hydro_gs", "Specific Gravity Gs", "", "2.67", False, False),
                    ("hydro_moist_sample_mass", "Hydrometer Moist Sample Mass", "g"),
                    ("hydro_hydrostatic_moisture", "Hydrostatic Moisture Content", "%"),
                    ("hydro_w", "Oven Dry Weight W (direct override)", "g", "", False, False),
                    ("hydro_pct_finer_no10", "Fraction Finer than #10", "", "1.0", False, False),
                    ("hydro_cal_t1", "Cal Temp 1", "C", "20.5", False, False),
                    ("hydro_cal_c1", "Cal Corr 1", "", "0.0025", False, False),
                    ("hydro_cal_t2", "Cal Temp 2", "C", "21.5", False, False),
                    ("hydro_cal_c2", "Cal Corr 2", "", "0.0020", False, False),
                    ("hydro_cal_t3", "Cal Temp 3", "C", "19.6", False, False),
                    ("hydro_cal_c3", "Cal Corr 3", "", "0.0040", False, False),
                    ("hydro_cal_t4", "Cal Temp 4", "C", "17.2", False, False),
                    ("hydro_cal_c4", "Cal Corr 4", "", "0.0045", False, False),
                    ("hydro_temp_1", "Temp @1 min", "C"),
                    ("hydro_ra_1", "Actual Reading Ra @1 min", ""),
                    ("hydro_temp_2", "Temp @2 min", "C"),
                    ("hydro_ra_2", "Actual Reading Ra @2 min", ""),
                    ("hydro_temp_5", "Temp @5 min", "C"),
                    ("hydro_ra_5", "Actual Reading Ra @5 min", ""),
                    ("hydro_temp_10", "Temp @10 min", "C"),
                    ("hydro_ra_10", "Actual Reading Ra @10 min", ""),
                    ("hydro_temp_15", "Temp @15 min", "C"),
                    ("hydro_ra_15", "Actual Reading Ra @15 min", ""),
                    ("hydro_temp_30", "Temp @30 min", "C"),
                    ("hydro_ra_30", "Actual Reading Ra @30 min", ""),
                    ("hydro_temp_60", "Temp @60 min", "C"),
                    ("hydro_ra_60", "Actual Reading Ra @60 min", ""),
                    ("hydro_temp_250", "Temp @250 min", "C"),
                    ("hydro_ra_250", "Actual Reading Ra @250 min", ""),
                    ("hydro_temp_1440", "Temp @1440 min", "C"),
                    ("hydro_ra_1440", "Actual Reading Ra @1440 min", ""),
                ]
            )
            computed.extend(
                [
                    ("hydro_w_dry_from_moist", "Computed W from Moisture", "g"),
                    ("hydro_w_dry_used", "W Used in Calculations", "g"),
                    ("hydro_cal_slope", "Calibration Slope m", ""),
                    ("hydro_cal_intercept", "Calibration Intercept b", ""),
                    ("hydro_rc_1", "Rc @1 min", ""),
                    ("hydro_d_1", "D @1 min", "mm"),
                    ("hydro_total_1", "Total %Finer @1 min", "%"),
                    ("hydro_rc_2", "Rc @2 min", ""),
                    ("hydro_d_2", "D @2 min", "mm"),
                    ("hydro_total_2", "Total %Finer @2 min", "%"),
                    ("hydro_rc_5", "Rc @5 min", ""),
                    ("hydro_d_5", "D @5 min", "mm"),
                    ("hydro_total_5", "Total %Finer @5 min", "%"),
                    ("hydro_rc_10", "Rc @10 min", ""),
                    ("hydro_d_10", "D @10 min", "mm"),
                    ("hydro_total_10", "Total %Finer @10 min", "%"),
                    ("hydro_rc_15", "Rc @15 min", ""),
                    ("hydro_d_15", "D @15 min", "mm"),
                    ("hydro_total_15", "Total %Finer @15 min", "%"),
                    ("hydro_rc_30", "Rc @30 min", ""),
                    ("hydro_d_30", "D @30 min", "mm"),
                    ("hydro_total_30", "Total %Finer @30 min", "%"),
                    ("hydro_rc_60", "Rc @60 min", ""),
                    ("hydro_d_60", "D @60 min", "mm"),
                    ("hydro_total_60", "Total %Finer @60 min", "%"),
                    ("hydro_rc_250", "Rc @250 min", ""),
                    ("hydro_d_250", "D @250 min", "mm"),
                    ("hydro_total_250", "Total %Finer @250 min", "%"),
                    ("hydro_rc_1440", "Rc @1440 min", ""),
                    ("hydro_d_1440", "D @1440 min", "mm"),
                    ("hydro_total_1440", "Total %Finer @1440 min", "%"),
                ]
            )
        return {
            "key": "grain_size",
            "title": "Grain Size Analysis",
            "astm": "D422 / D1140",
            "fields": fields,
            "computed": computed,
        }
