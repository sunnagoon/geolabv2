import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from app.db import get_connection
from app.services.results_export import export_results_matrix_xlsx, export_results_matrix_pdf


class ResultsTab(ttk.Frame):
    def __init__(self, parent, get_project_id):
        super().__init__(parent)
        self.get_project_id = get_project_id
        self._build_ui()
        self._test_labels = {}

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        actions = ttk.Frame(top)
        actions.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(actions, text="Export Results Table (PDF)", command=self._export_results_table_pdf).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(actions, text="Export Results Table (Excel)", command=self._export_results_table).pack(side=tk.RIGHT)

        self.tree = ttk.Treeview(
            top,
            columns=(
                "sample",
                "depth",
                "code",
                "name",
                "status",
                "value",
                "unit",
                "value2",
                "unit2",
                "value3",
                "unit3",
                "value4",
                "unit4",
            ),
            show="headings",
        )
        for col, text, width in [
            ("sample", "Sample", 120),
            ("depth", "Depth", 120),
            ("code", "Test Code", 90),
            ("name", "Test Name", 220),
            ("status", "Status", 90),
            ("value", "Value 1", 80),
            ("unit", "Unit 1", 70),
            ("value2", "Value 2", 80),
            ("unit2", "Unit 2", 70),
            ("value3", "Value 3", 80),
            ("unit3", "Unit 3", 70),
            ("value4", "Value 4", 80),
            ("unit4", "Unit 4", 70),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=tk.W)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        form = ttk.LabelFrame(top, text="Enter Test Result")
        form.pack(fill=tk.X, pady=10)

        self.value_var = tk.StringVar()
        self.unit_var = tk.StringVar()
        self.value2_var = tk.StringVar()
        self.unit2_var = tk.StringVar()
        self.value3_var = tk.StringVar()
        self.unit3_var = tk.StringVar()
        self.value4_var = tk.StringVar()
        self.unit4_var = tk.StringVar()
        self.notes_var = tk.StringVar()
        self.status_var = tk.StringVar(value="completed")
        self.chem_mode_var = tk.StringVar(value="Chlorides, Sulfates")

        self.value1_label = ttk.Label(form, text="Value 1")
        self.value1_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.value_var, width=15).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form, text="Unit 1").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.unit_combo = ttk.Combobox(
            form,
            textvariable=self.unit_var,
            values=["PSF", "%", "PHI", "Cohesion", ""],
            width=12,
        )
        self.unit_combo.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

        self.value2_label = ttk.Label(form, text="Value 2")
        self.value2_label.grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.value2_var, width=15).grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form, text="Unit 2").grid(row=0, column=6, sticky=tk.W, padx=5, pady=5)
        self.unit2_combo = ttk.Combobox(
            form,
            textvariable=self.unit2_var,
            values=["PSF", "%", "PHI", "Cohesion", ""],
            width=12,
        )
        self.unit2_combo.grid(row=0, column=7, sticky=tk.W, padx=5, pady=5)

        self.value3_label = ttk.Label(form, text="Value 3")
        self.value3_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.value3_entry = ttk.Entry(form, textvariable=self.value3_var, width=15)
        self.value3_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form, text="Unit 3").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        self.unit3_combo = ttk.Combobox(
            form,
            textvariable=self.unit3_var,
            values=["PSF", "%", "PHI", "Cohesion", ""],
            width=12,
        )
        self.unit3_combo.grid(row=1, column=3, sticky=tk.W, padx=5, pady=5)

        self.value4_label = ttk.Label(form, text="Value 4")
        self.value4_label.grid(row=1, column=4, sticky=tk.W, padx=5, pady=5)
        self.value4_entry = ttk.Entry(form, textvariable=self.value4_var, width=15)
        self.value4_entry.grid(row=1, column=5, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form, text="Unit 4").grid(row=1, column=6, sticky=tk.W, padx=5, pady=5)
        self.unit4_combo = ttk.Combobox(
            form,
            textvariable=self.unit4_var,
            values=["PSF", "%", "PHI", "Cohesion", ""],
            width=12,
        )
        self.unit4_combo.grid(row=1, column=7, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form, text="Status").grid(row=1, column=8, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.status_var, width=12).grid(row=1, column=9, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form, text="Notes").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.notes_var, width=60).grid(row=2, column=1, columnspan=5, sticky=tk.W, padx=5, pady=5)

        self.ei_label = ttk.Label(form, text="")
        self.ei_label.grid(row=2, column=6, columnspan=2, sticky=tk.W, padx=5, pady=5)

        self.chem_label = ttk.Label(form, text="Chem Fields")
        self.chem_label.grid(row=2, column=3, sticky=tk.E, padx=5, pady=5)
        self.chem_combo = ttk.Combobox(
            form,
            textvariable=self.chem_mode_var,
            values=["Chlorides, Sulfates", "All Four"],
            width=20,
        )
        self.chem_combo.grid(row=2, column=4, columnspan=2, sticky=tk.W, padx=5, pady=5)
        self.chem_combo.bind("<<ComboboxSelected>>", self._on_chem_mode)

        ttk.Button(form, text="Save Result", command=self._save).grid(row=0, column=10, rowspan=3, padx=5, pady=5)

        self.selected_id = None

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        project_id = self.get_project_id()
        if not project_id:
            return

        conn = get_connection()
        rows = conn.execute(
            """
            SELECT st.id, s.sample_name, s.depth_raw, t.code, t.name, st.status,
                   st.result_value, st.result_unit, st.result_value2, st.result_unit2,
                   st.result_value3, st.result_unit3, st.result_value4, st.result_unit4
            FROM sample_tests st
            JOIN samples s ON s.id = st.sample_id
            JOIN tests t ON t.id = st.test_id
            WHERE s.project_id = ?
            ORDER BY s.sample_name, t.code
            """,
            (project_id,),
        ).fetchall()
        conn.close()

        for row in rows:
            self.tree.insert(
                "",
                tk.END,
                iid=row["id"],
                values=(
                    row["sample_name"],
                    row["depth_raw"] or "",
                    row["code"],
                    row["name"],
                    row["status"] or "",
                    self._fmt1(row["result_value"]),
                    row["result_unit"] or "",
                    self._fmt1(row["result_value2"]),
                    row["result_unit2"] or "",
                    self._fmt1(row["result_value3"]),
                    row["result_unit3"] or "",
                    self._fmt1(row["result_value4"]),
                    row["result_unit4"] or "",
                ),
            )

    def _on_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        self.selected_id = int(sel[0])
        conn = get_connection()
        row = conn.execute(
            """
            SELECT st.result_value, st.result_unit, st.result_value2, st.result_unit2,
                   st.result_value3, st.result_unit3, st.result_value4, st.result_unit4,
                   st.result_notes, st.status, t.name
            FROM sample_tests st
            JOIN tests t ON t.id = st.test_id
            WHERE st.id = ?
            """,
            (self.selected_id,),
        ).fetchone()
        conn.close()
        if not row:
            return
        self.value_var.set(self._fmt1(row["result_value"]))
        self.unit_var.set(row["result_unit"] or "")
        self.value2_var.set(self._fmt1(row["result_value2"]))
        self.unit2_var.set(row["result_unit2"] or "")
        self.value3_var.set(self._fmt1(row["result_value3"]))
        self.unit3_var.set(row["result_unit3"] or "")
        self.value4_var.set(self._fmt1(row["result_value4"]))
        self.unit4_var.set(row["result_unit4"] or "")
        self.notes_var.set(row["result_notes"] or "")
        self.status_var.set(row["status"] or "completed")
        self._apply_labels(row["name"])

    def _save(self):
        if not self.selected_id:
            messagebox.showerror("No Selection", "Select a test row to enter results.")
            return

        def parse_number(label, raw):
            if not raw:
                return None
            if label == "USCS Classif.":
                return raw
            try:
                return float(raw)
            except ValueError:
                messagebox.showerror("Invalid", f"{label} must be a number.")
                return "error"

        value = parse_number(self.value1_label.cget("text"), self.value_var.get().strip())
        if value == "error":
            return
        value2 = parse_number(self.value2_label.cget("text"), self.value2_var.get().strip())
        if value2 == "error":
            return
        value3 = parse_number(self.value3_label.cget("text"), self.value3_var.get().strip())
        if value3 == "error":
            return
        value4 = parse_number(self.value4_label.cget("text"), self.value4_var.get().strip())
        if value4 == "error":
            return

        unit = self.unit_var.get().strip()
        unit2 = self.unit2_var.get().strip()
        unit3 = self.unit3_var.get().strip()
        unit4 = self.unit4_var.get().strip()
        status = self.status_var.get().strip() or "completed"
        notes = self.notes_var.get().strip() or None

        if self.value1_label.cget("text") == "Expansion Index":
            potential = self._expansion_potential(value)
            value2 = None
            unit2 = potential or ""
            self.value2_var.set("")
            self.unit2_var.set(unit2)

        if self.value1_label.cget("text") == "Dry Density" and self.value2_label.cget("text") == "Moisture Content":
            sat = self._calc_saturation(value, value2)
            if sat is not None:
                value3 = sat
                self.value3_var.set(f"{sat:.1f}")
                if not unit3:
                    unit3 = "%"
                    self.unit3_var.set(unit3)

        conn = get_connection()
        conn.execute(
            """
            UPDATE sample_tests
            SET result_value = ?, result_unit = ?, result_value2 = ?, result_unit2 = ?,
                result_value3 = ?, result_unit3 = ?, result_value4 = ?, result_unit4 = ?,
                result_notes = ?, status = ?
            WHERE id = ?
            """,
            (value, unit, value2, unit2, value3, unit3, value4, unit4, notes, status, self.selected_id),
        )
        conn.commit()
        conn.close()
        self.refresh()

    def _apply_labels(self, test_name):
        label_map = {
            "Max Density": ("Maximum Density", "Optimum Moisture", None, None),
            "698 Max": ("Maximum Density", "Optimum Moisture", None, None),
            "C Max": ("Maximum Density", "Optimum Moisture", None, None),
            "Direct Shear": ("Peak Phi", "Peak Cohesion", "Ultimate Phi", "Ultimate Cohesion"),
            "Field Density/Moisture": ("Dry Density", "Moisture Content", "Saturation", None),
            "Sand Cone": ("Dry Density", "Moisture Content", None, None),
            "Sieve Part. Analysis": ("USCS Classif.", "Passing No. 200", None, None),
            "Expansion Index": ("Expansion Index", "Expansion Potential", None, None),
            "Chem": ("Resistivity", "Sulfates", "Chlorides", "pH"),
            "Atterberg Limits": ("Liquid Limit", "Plasticity Index", None, None),
            "Hydro Response": ("Hydro Response", "Normal Stress", None, None),
            "R-Value": ("R-Value by Equilibration", None, None, None),
        }
        labels = label_map.get(test_name, ("Value 1", "Value 2", "Value 3", None))
        self.value1_label.config(text=labels[0] or "Value 1")
        self.value2_label.config(text=labels[1] or "Value 2")
        self.value3_label.config(text=labels[2] or "Value 3")
        self.value4_label.config(text=labels[3] or "Value 4")

        show_third = labels[2] is not None
        if show_third:
            self.value3_label.grid()
            self.value3_entry.grid()
            self.unit3_combo.grid()
        else:
            self.value3_label.grid_remove()
            self.value3_entry.grid_remove()
            self.unit3_combo.grid_remove()

        show_fourth = labels[3] is not None
        if show_fourth:
            self.value4_label.grid()
            self.value4_entry.grid()
            self.unit4_combo.grid()
        else:
            self.value4_label.grid_remove()
            self.value4_entry.grid_remove()
            self.unit4_combo.grid_remove()

        if labels[0] == "Expansion Index":
            self.ei_label.config(text="Expansion Potential is auto-calculated from EI")
        elif labels[0] == "Dry Density" and labels[1] == "Moisture Content":
            self.ei_label.config(text="Saturation (%) auto-calculates from dry density and moisture.")
        else:
            self.ei_label.config(text="")

        show_chem = test_name == "Chem"
        if show_chem:
            self.chem_label.grid()
            self.chem_combo.grid()
            self._apply_chem_mode(self.chem_mode_var.get())
        else:
            self.chem_label.grid_remove()
            self.chem_combo.grid_remove()

    def _on_chem_mode(self, _event):
        self._apply_chem_mode(self.chem_mode_var.get())

    def _apply_chem_mode(self, mode):
        if mode == "Chlorides, Sulfates":
            self.value1_label.config(text="Chlorides")
            self.value2_label.config(text="Sulfates")
            self.value3_label.config(text="")
            self.value4_label.config(text="")
            self._toggle_value3(False)
            self._toggle_value4(False)
        else:
            self.value1_label.config(text="Resistivity")
            self.value2_label.config(text="Sulfates")
            self.value3_label.config(text="Chlorides")
            self.value4_label.config(text="pH")
            self._toggle_value3(True)
            self._toggle_value4(True)

    def _toggle_value3(self, show):
        if show:
            self.value3_label.grid()
            self.value3_entry.grid()
            self.unit3_combo.grid()
        else:
            self.value3_label.grid_remove()
            self.value3_entry.grid_remove()
            self.unit3_combo.grid_remove()

    def _toggle_value4(self, show):
        if show:
            self.value4_label.grid()
            self.value4_entry.grid()
            self.unit4_combo.grid()
        else:
            self.value4_label.grid_remove()
            self.value4_entry.grid_remove()
            self.unit4_combo.grid_remove()

    def _expansion_potential(self, value):
        if value is None:
            return None
        if value <= 19:
            return "Very Low"
        if value <= 49:
            return "Low"
        if value <= 89:
            return "Medium"
        if value <= 130:
            return "High"
        return "Very High"

    def _calc_saturation(self, dry_density_pcf, moisture_percent):
        if dry_density_pcf is None or moisture_percent is None:
            return None
        try:
            gd = float(dry_density_pcf)
            w = float(moisture_percent) / 100.0
            gamma_w = 62.4
            gs = 2.65  # assumed specific gravity for auto estimate
            e = (gs * gamma_w / gd) - 1.0
            if e <= 0:
                return None
            sat = (w * gs / e) * 100.0
            return round(max(0.0, min(100.0, sat)), 1)
        except Exception:
            return None

    def _fmt1(self, value):
        if value is None:
            return ""
        try:
            return f"{float(value):.1f}"
        except Exception:
            return str(value)

    def _export_results_table(self):
        project_id = self.get_project_id()
        if not project_id:
            messagebox.showerror("No Project", "Select a project first.")
            return

        conn = get_connection()
        project = conn.execute(
            "SELECT file_number, job_name FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()
        samples = conn.execute(
            """
            SELECT id, sample_name, sample_type, depth_raw
            FROM samples
            WHERE project_id = ?
            ORDER BY sample_name, id
            """,
            (project_id,),
        ).fetchall()
        tests = conn.execute(
            """
            SELECT DISTINCT t.id, t.name, t.code
            FROM sample_tests st
            JOIN samples s ON s.id = st.sample_id
            JOIN tests t ON t.id = st.test_id
            WHERE s.project_id = ?
            ORDER BY t.code, t.name
            """,
            (project_id,),
        ).fetchall()
        results = conn.execute(
            """
            SELECT st.sample_id, st.test_id, t.name AS test_name,
                   st.result_value, st.result_unit, st.result_value2, st.result_unit2,
                   st.result_value3, st.result_unit3, st.result_value4, st.result_unit4
            FROM sample_tests st
            JOIN samples s ON s.id = st.sample_id
            JOIN tests t ON t.id = st.test_id
            WHERE s.project_id = ?
            """,
            (project_id,),
        ).fetchall()
        conn.close()

        if not project:
            messagebox.showerror("Missing Project", "Project record not found.")
            return
        if not samples:
            messagebox.showerror("No Samples", "No samples found for this project.")
            return
        if not tests:
            messagebox.showerror("No Tests", "No tests assigned for this project.")
            return
        has_data = False
        for r in results:
            for key in ("result_value", "result_value2", "result_value3", "result_value4"):
                if r[key] is not None:
                    has_data = True
                    break
            if has_data:
                break
            for key in ("result_unit", "result_unit2", "result_unit3", "result_unit4"):
                if isinstance(r[key], str) and r[key].strip():
                    has_data = True
                    break
            if has_data:
                break
        if not has_data:
            messagebox.showerror("No Entered Data", "Enter at least one result value before export.")
            return

        default_name = f"Results_{project['file_number']}.xlsx"
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name)
        if not path:
            return

        export_results_matrix_xlsx(
            path=path,
            project=dict(project),
            samples=[dict(r) for r in samples],
            tests=[dict(r) for r in tests],
            results=[dict(r) for r in results],
        )
        messagebox.showinfo("Exported", f"Results table exported to:\n{path}")

    def _export_results_table_pdf(self):
        project_id = self.get_project_id()
        if not project_id:
            messagebox.showerror("No Project", "Select a project first.")
            return

        conn = get_connection()
        project = conn.execute(
            "SELECT file_number, job_name FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()
        samples = conn.execute(
            """
            SELECT id, sample_name, sample_type, depth_raw
            FROM samples
            WHERE project_id = ?
            ORDER BY sample_name, id
            """,
            (project_id,),
        ).fetchall()
        tests = conn.execute(
            """
            SELECT DISTINCT t.id, t.name, t.code
            FROM sample_tests st
            JOIN samples s ON s.id = st.sample_id
            JOIN tests t ON t.id = st.test_id
            WHERE s.project_id = ?
            ORDER BY t.code, t.name
            """,
            (project_id,),
        ).fetchall()
        results = conn.execute(
            """
            SELECT st.sample_id, st.test_id, t.name AS test_name,
                   st.result_value, st.result_unit, st.result_value2, st.result_unit2,
                   st.result_value3, st.result_unit3, st.result_value4, st.result_unit4
            FROM sample_tests st
            JOIN samples s ON s.id = st.sample_id
            JOIN tests t ON t.id = st.test_id
            WHERE s.project_id = ?
            """,
            (project_id,),
        ).fetchall()
        conn.close()

        if not project:
            messagebox.showerror("Missing Project", "Project record not found.")
            return
        if not samples:
            messagebox.showerror("No Samples", "No samples found for this project.")
            return
        if not tests:
            messagebox.showerror("No Tests", "No tests assigned for this project.")
            return
        has_data = False
        for r in results:
            for key in ("result_value", "result_value2", "result_value3", "result_value4"):
                if r[key] is not None:
                    has_data = True
                    break
            if has_data:
                break
            for key in ("result_unit", "result_unit2", "result_unit3", "result_unit4"):
                if isinstance(r[key], str) and r[key].strip():
                    has_data = True
                    break
            if has_data:
                break
        if not has_data:
            messagebox.showerror("No Entered Data", "Enter at least one result value before export.")
            return

        default_name = f"Results_{project['file_number']}.pdf"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_name)
        if not path:
            return

        try:
            export_results_matrix_pdf(
                path=path,
                project=dict(project),
                samples=[dict(r) for r in samples],
                tests=[dict(r) for r in tests],
                results=[dict(r) for r in results],
            )
        except Exception as exc:
            messagebox.showerror("Export Failed", f"PDF export failed: {exc}")
            return
        messagebox.showinfo("Exported", f"Results table exported to:\n{path}")
