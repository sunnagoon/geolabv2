import tkinter as tk
from tkinter import ttk, messagebox

from app.db import get_connection
from app.services.validators import is_valid_sample_name, parse_depth


class SamplesTab(ttk.Frame):
    def __init__(self, parent, get_project_id, on_samples_changed=None):
        super().__init__(parent)
        self.get_project_id = get_project_id
        self.on_samples_changed = on_samples_changed
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(
            top,
            columns=("name", "type", "depth", "received", "location", "disposal", "status"),
            show="headings",
        )
        for col, text, width in [
            ("name", "Sample", 120),
            ("type", "Type", 80),
            ("depth", "Depth", 120),
            ("received", "Received", 110),
            ("location", "Storage", 140),
            ("disposal", "Disposal", 110),
            ("status", "Status", 90),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=tk.W)

        self.tree.pack(fill=tk.BOTH, expand=True)

        form = ttk.LabelFrame(self, text="Add Sample")
        form.pack(fill=tk.X, padx=10, pady=10)

        self.sample_name = tk.StringVar()
        self.depth_raw = tk.StringVar()
        self.sample_type = tk.StringVar(value="SB")
        self.received_date = tk.StringVar()
        self.storage_location = tk.StringVar()
        self.disposal_date = tk.StringVar()
        self.status = tk.StringVar(value="Inventory")
        self.editing_sample_id = None

        row = 0
        ttk.Label(form, text="Sample Name").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.sample_name, width=15).grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(form, text="Depth (e.g., 1.0'-2.0')").grid(row=row, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.depth_raw, width=20).grid(row=row, column=3, sticky=tk.W, padx=5, pady=5)
        ttk.Label(form, text="Sample Type").grid(row=row, column=4, sticky=tk.W, padx=5, pady=5)
        ttk.Combobox(
            form,
            textvariable=self.sample_type,
            values=["SB", "MB", "LB", "SPT", "Ring"],
            state="readonly",
            width=8,
        ).grid(row=row, column=5, sticky=tk.W, padx=5, pady=5)

        row += 1
        ttk.Label(form, text="Received Date").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.received_date, width=15).grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(form, text="Storage Location").grid(row=row, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.storage_location, width=20).grid(row=row, column=3, sticky=tk.W, padx=5, pady=5)

        row += 1
        ttk.Label(form, text="Disposal Date").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.disposal_date, width=15).grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(form, text="Status").grid(row=row, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(form, textvariable=self.status, width=15).grid(row=row, column=3, sticky=tk.W, padx=5, pady=5)

        ttk.Button(form, text="Save Sample", command=self._save_sample).grid(row=row, column=4, sticky=tk.E, padx=5, pady=5)
        ttk.Button(form, text="Update Selected", command=self._update_selected).grid(row=row, column=5, sticky=tk.E, padx=5, pady=5)
        ttk.Button(form, text="Delete Selected", command=self._delete_selected).grid(row=row, column=6, sticky=tk.E, padx=5, pady=5)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        project_id = self.get_project_id()
        if not project_id:
            return

        conn = get_connection()
        rows = conn.execute(
            """
            SELECT id, sample_name, sample_type, depth_raw, received_date, storage_location, disposal_date, status
            FROM samples WHERE project_id = ? ORDER BY id DESC
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
                    row["sample_type"] or "",
                    row["depth_raw"] or "",
                    row["received_date"] or "",
                    row["storage_location"] or "",
                    row["disposal_date"] or "",
                    row["status"] or "",
                ),
            )

    def _save_sample(self):
        project_id = self.get_project_id()
        if not project_id:
            messagebox.showerror("No Project", "Select a project first.")
            return

        sample_name = self.sample_name.get().strip()
        if not is_valid_sample_name(sample_name):
            messagebox.showerror("Invalid", "Sample name must be B-#, T-#, HA-#, or C-#.")
            return

        depth_raw = self.depth_raw.get().strip()
        depth_from, depth_to, depth_unit = parse_depth(depth_raw)

        conn = get_connection()
        conn.execute(
            """
            INSERT INTO samples (project_id, sample_name, sample_type, depth_raw, depth_from, depth_to, depth_unit, received_date, storage_location, disposal_date, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                project_id,
                sample_name,
                (self.sample_type.get().strip() or "SB").upper(),
                depth_raw or None,
                depth_from,
                depth_to,
                depth_unit,
                self.received_date.get().strip() or None,
                self.storage_location.get().strip() or None,
                self.disposal_date.get().strip() or None,
                self.status.get().strip() or None,
            ),
        )
        conn.commit()
        conn.close()
        self.refresh()
        self._clear_form()
        if self.on_samples_changed:
            self.on_samples_changed()

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        sample_id = int(sel[0])
        if not messagebox.askyesno("Confirm", "Delete this sample and its tests?"):
            return
        conn = get_connection()
        conn.execute("DELETE FROM samples WHERE id = ?", (sample_id,))
        conn.commit()
        conn.close()
        self.refresh()
        if self.on_samples_changed:
            self.on_samples_changed()

    def _on_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        sample_id = int(sel[0])
        self._load_sample(sample_id)

    def _load_sample(self, sample_id):
        conn = get_connection()
        row = conn.execute(
            """
            SELECT id, sample_name, sample_type, depth_raw, received_date, storage_location, disposal_date, status
            FROM samples WHERE id = ?
            """,
            (sample_id,),
        ).fetchone()
        conn.close()
        if not row:
            return
        self.editing_sample_id = row["id"]
        self.sample_name.set(row["sample_name"])
        self.sample_type.set((row["sample_type"] or "SB").upper())
        self.depth_raw.set(row["depth_raw"] or "")
        self.received_date.set(row["received_date"] or "")
        self.storage_location.set(row["storage_location"] or "")
        self.disposal_date.set(row["disposal_date"] or "")
        self.status.set(row["status"] or "")

    def _update_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("No Selection", "Select a sample to update.")
            return
        sample_id = int(sel[0])

        sample_name = self.sample_name.get().strip()
        if not is_valid_sample_name(sample_name):
            messagebox.showerror("Invalid", "Sample name must be B-#, T-#, HA-#, or C-#.")
            return

        depth_raw = self.depth_raw.get().strip()
        depth_from, depth_to, depth_unit = parse_depth(depth_raw)

        conn = get_connection()
        conn.execute(
            """
            UPDATE samples
            SET sample_name = ?, sample_type = ?, depth_raw = ?, depth_from = ?, depth_to = ?, depth_unit = ?,
                received_date = ?, storage_location = ?, disposal_date = ?, status = ?
            WHERE id = ?
            """,
            (
                sample_name,
                (self.sample_type.get().strip() or "SB").upper(),
                depth_raw or None,
                depth_from,
                depth_to,
                depth_unit,
                self.received_date.get().strip() or None,
                self.storage_location.get().strip() or None,
                self.disposal_date.get().strip() or None,
                self.status.get().strip() or None,
                sample_id,
            ),
        )
        conn.commit()
        conn.close()
        self.refresh()
        self._clear_form()
        if self.on_samples_changed:
            self.on_samples_changed()

    def _clear_form(self):
        self.editing_sample_id = None
        self.sample_name.set("")
        self.sample_type.set("SB")
        self.depth_raw.set("")
        self.received_date.set("")
        self.storage_location.set("")
        self.disposal_date.set("")
        self.status.set("Inventory")
