import tkinter as tk
from tkinter import ttk, messagebox

from app.db import get_connection, now_iso
from app.services.validators import is_valid_file_number


class ProjectsTab(ttk.Frame):
    def __init__(self, parent, on_project_selected):
        super().__init__(parent)
        self.on_project_selected = on_project_selected
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(
            top,
            columns=("file", "job", "client", "rate", "location", "status", "remaining"),
            show="headings",
        )
        for col, text, width in [
            ("file", "File #", 100),
            ("job", "Job Name", 260),
            ("client", "Client", 120),
            ("rate", "Billing Rate", 120),
            ("location", "Location", 220),
            ("status", "Status", 140),
            ("remaining", "Tests Remaining", 140),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=tk.W)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.pack(fill=tk.BOTH, expand=True)

        form = ttk.LabelFrame(self, text="Add Project")
        form.pack(fill=tk.X, padx=10, pady=10)

        self.file_number = tk.StringVar()
        self.job_name = tk.StringVar()
        self.client_type = tk.StringVar(value="TCI")
        self.billing_rate = tk.StringVar()
        self.billing_year = tk.StringVar()
        self.billing_kind = tk.StringVar()
        self.location_text = tk.StringVar()
        self.latitude = tk.StringVar()
        self.longitude = tk.StringVar()
        self.status_var = tk.StringVar(value="Not Scheduled")
        self.rate_choices = []
        self.editing_project_id = None

        row = 0
        ttk.Label(form, text="File # (NN-NNN)").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.file_number, width=15).grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(form, text="Job Name").grid(row=row, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.job_name, width=40).grid(row=row, column=3, sticky=tk.W, padx=5, pady=5)

        row += 1
        ttk.Label(form, text="Client Type").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Combobox(form, textvariable=self.client_type, values=["TCI", "GCI", "SBCI", "CUSTOM"], width=12).grid(
            row=row, column=1, sticky=tk.W, padx=5, pady=5
        )
        ttk.Label(form, text="").grid(row=row, column=2, sticky=tk.W, padx=5, pady=5)

        row += 1
        ttk.Label(form, text="Billing Rate ID").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        self.billing_rate_combo = ttk.Combobox(form, textvariable=self.billing_rate, width=20)
        self.billing_rate_combo.grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(form, text="Billing Year").grid(row=row, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.billing_year, width=10).grid(row=row, column=3, sticky=tk.W, padx=5, pady=5)

        row += 1
        ttk.Label(form, text="Billing Kind").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.billing_kind, width=20).grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)

        row += 1
        ttk.Label(form, text="Location (optional)").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.location_text, width=40).grid(row=row, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form, text="Lat").grid(row=row, column=3, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.latitude, width=10).grid(row=row, column=4, sticky=tk.W, padx=5, pady=5)
        ttk.Label(form, text="Lon").grid(row=row, column=5, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.longitude, width=10).grid(row=row, column=6, sticky=tk.W, padx=5, pady=5)

        row += 1
        ttk.Label(form, text="Status").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Combobox(
            form,
            textvariable=self.status_var,
            values=["Not Scheduled", "Testing in Progress", "Testing Complete"],
            width=20,
        ).grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)

        row += 1
        btns = ttk.Frame(form)
        btns.grid(row=row, column=0, columnspan=7, sticky=tk.E, padx=5, pady=5)
        ttk.Button(btns, text="Save Project", command=self._save_project).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btns, text="Update Selected", command=self._update_selected).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btns, text="Delete Selected", command=self._delete_selected).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btns, text="Reload Rates", command=self.refresh_rates).pack(side=tk.RIGHT, padx=5)

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        conn = get_connection()
        rows = conn.execute(
            """
            SELECT p.id, p.file_number, p.job_name, p.client_type, p.billing_rate_id, p.location_text, p.status,
                   SUM(CASE WHEN st.status IS NULL OR st.status != 'completed' THEN 1 ELSE 0 END) AS remaining
            FROM projects p
            LEFT JOIN samples s ON s.project_id = p.id
            LEFT JOIN sample_tests st ON st.sample_id = s.id
            GROUP BY p.id
            ORDER BY p.created_at DESC
            """
        ).fetchall()
        conn.close()
        for row in rows:
            remaining = row["remaining"] if row["remaining"] is not None else 0
            self.tree.insert(
                "",
                tk.END,
                iid=row["id"],
                values=(
                    row["file_number"],
                    row["job_name"],
                    row["client_type"],
                    row["billing_rate_id"],
                    row["location_text"] or "",
                    row["status"] or "Not Scheduled",
                    remaining,
                ),
            )
        self.refresh_rates()

    def refresh_rates(self):
        conn = get_connection()
        rows = conn.execute("SELECT rate_id FROM billing_rates ORDER BY rate_id").fetchall()
        conn.close()
        self.rate_choices = [r["rate_id"] for r in rows]
        self.billing_rate_combo["values"] = self.rate_choices

    def _on_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        project_id = int(sel[0])
        self._load_project(project_id)
        self.on_project_selected(project_id)

    def _save_project(self):
        file_number = self.file_number.get().strip()
        if not is_valid_file_number(file_number):
            messagebox.showerror("Invalid", "File number must be NN-NNN (e.g., 26-112).")
            return
        job_name = self.job_name.get().strip()
        client_type = self.client_type.get().strip() or "CUSTOM"
        client_name = client_type
        billing_rate_id = self.billing_rate.get().strip()
        status = self.status_var.get().strip() or "Not Scheduled"
        location_text = self.location_text.get().strip() or None
        latitude = self._parse_optional_float(self.latitude.get().strip(), "Latitude")
        if latitude == "error":
            return
        longitude = self._parse_optional_float(self.longitude.get().strip(), "Longitude")
        if longitude == "error":
            return

        if not job_name or not billing_rate_id:
            messagebox.showerror("Missing", "Job name and billing rate are required.")
            return

        conn = get_connection()
        try:
            conn.execute(
                """
                INSERT INTO projects (
                    file_number, job_name, client_type, client_name, billing_rate_id, billing_year, billing_kind,
                    location_text, latitude, longitude, status, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    file_number,
                    job_name,
                    client_type,
                    client_name,
                    billing_rate_id,
                    self.billing_year.get().strip() or None,
                    self.billing_kind.get().strip() or None,
                    location_text,
                    latitude,
                    longitude,
                    status,
                    now_iso(),
                ),
            )
            conn.commit()
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to save project: {exc}")
        finally:
            conn.close()

        self.refresh()
        self._clear_form()

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("No Selection", "Select a project to delete.")
            return
        project_id = int(sel[0])
        project_label = self.tree.item(sel[0], "values")
        file_no = project_label[0] if project_label else ""
        job_name = project_label[1] if project_label else ""
        prompt = (
            f"Are you sure you want to delete project {file_no} - {job_name}?\n\n"
            "This will permanently delete all related samples, tests, and worksheet data."
        )
        if not messagebox.askyesno("Confirm Project Deletion", prompt):
            return
        conn = get_connection()
        try:
            conn.execute("DELETE FROM projects WHERE id = ?", (project_id,))
            conn.commit()
        finally:
            conn.close()
        self.refresh()
        self.on_project_selected(None)
        self._clear_form()

    def _load_project(self, project_id):
        conn = get_connection()
        row = conn.execute(
            """
            SELECT id, file_number, job_name, client_type, client_name, billing_rate_id, billing_year, billing_kind,
                   location_text, latitude, longitude, status
            FROM projects WHERE id = ?
            """,
            (project_id,),
        ).fetchone()
        conn.close()
        if not row:
            return
        self.editing_project_id = row["id"]
        self.file_number.set(row["file_number"])
        self.job_name.set(row["job_name"])
        self.client_type.set(row["client_type"])
        # client_name removed from UI; keep stored as client_type
        self.billing_rate.set(row["billing_rate_id"])
        self.billing_year.set(row["billing_year"] or "")
        self.billing_kind.set(row["billing_kind"] or "")
        self.location_text.set(row["location_text"] or "")
        self.latitude.set("" if row["latitude"] is None else str(row["latitude"]))
        self.longitude.set("" if row["longitude"] is None else str(row["longitude"]))
        self.status_var.set(row["status"] or "Not Scheduled")

    def _update_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("No Selection", "Select a project to update.")
            return
        project_id = int(sel[0])

        file_number = self.file_number.get().strip()
        if not is_valid_file_number(file_number):
            messagebox.showerror("Invalid", "File number must be NN-NNN (e.g., 26-112).")
            return
        job_name = self.job_name.get().strip()
        client_type = self.client_type.get().strip() or "CUSTOM"
        client_name = client_type
        billing_rate_id = self.billing_rate.get().strip()
        status = self.status_var.get().strip() or "Not Scheduled"
        location_text = self.location_text.get().strip() or None
        latitude = self._parse_optional_float(self.latitude.get().strip(), "Latitude")
        if latitude == "error":
            return
        longitude = self._parse_optional_float(self.longitude.get().strip(), "Longitude")
        if longitude == "error":
            return

        if not job_name or not billing_rate_id:
            messagebox.showerror("Missing", "Job name and billing rate are required.")
            return

        conn = get_connection()
        try:
            conn.execute(
                """
                UPDATE projects
                SET file_number = ?, job_name = ?, client_type = ?, client_name = ?,
                    billing_rate_id = ?, billing_year = ?, billing_kind = ?,
                    location_text = ?, latitude = ?, longitude = ?, status = ?
                WHERE id = ?
                """,
                (
                    file_number,
                    job_name,
                    client_type,
                    client_name,
                    billing_rate_id,
                    self.billing_year.get().strip() or None,
                    self.billing_kind.get().strip() or None,
                    location_text,
                    latitude,
                    longitude,
                    status,
                    project_id,
                ),
            )
            conn.commit()
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to update project: {exc}")
        finally:
            conn.close()

        self.refresh()
        self._clear_form()

    def _clear_form(self):
        self.editing_project_id = None
        self.file_number.set("")
        self.job_name.set("")
        self.client_type.set("TCI")
        self.billing_rate.set("")
        self.billing_year.set("")
        self.billing_kind.set("")
        self.location_text.set("")
        self.latitude.set("")
        self.longitude.set("")
        self.status_var.set("Not Scheduled")

    def _parse_optional_float(self, raw_value, field_name):
        if not raw_value:
            return None
        try:
            return float(raw_value)
        except ValueError:
            messagebox.showerror("Invalid", f"{field_name} must be a number.")
            return "error"
