import tkinter as tk
from tkinter import ttk, messagebox

from app.db import get_connection


class TestsTab(ttk.Frame):
    def __init__(self, parent, get_project_id, on_tests_changed=None):
        super().__init__(parent)
        self.get_project_id = get_project_id
        self.on_tests_changed = on_tests_changed
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        assign = ttk.LabelFrame(top, text="Assign Test to Sample")
        assign.pack(fill=tk.X, pady=5)

        self.sample_choice = tk.StringVar()
        self.assign_cost = tk.StringVar()
        self.assign_status = tk.StringVar(value="scheduled")

        ttk.Label(assign, text="Sample").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.sample_combo = ttk.Combobox(assign, textvariable=self.sample_choice, width=20)
        self.sample_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(assign, text="Tests").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.test_list = tk.Listbox(assign, selectmode=tk.MULTIPLE, height=5, exportselection=False)
        self.test_list.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        self.test_scroll = ttk.Scrollbar(assign, orient=tk.VERTICAL, command=self.test_list.yview)
        self.test_scroll.grid(row=0, column=4, sticky=tk.NS, padx=0, pady=5)
        self.test_list.configure(yscrollcommand=self.test_scroll.set)

        ttk.Label(assign, text="Override Cost (optional)").grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(assign, textvariable=self.assign_cost, width=12).grid(row=0, column=6, sticky=tk.W, padx=5, pady=5)

        ttk.Label(assign, text="Status").grid(row=0, column=7, sticky=tk.W, padx=5, pady=5)
        ttk.Combobox(
            assign,
            textvariable=self.assign_status,
            values=["scheduled", "in progress", "completed"],
            width=12,
        ).grid(row=0, column=8, sticky=tk.W, padx=5, pady=5)

        ttk.Button(assign, text="Assign Selected", command=self._assign_test).grid(row=0, column=9, sticky=tk.E, padx=5, pady=5)

        self.tree = ttk.Treeview(
            top,
            columns=("sample", "depth", "code", "name", "cost", "status"),
            show="headings",
        )
        for col, text, width in [
            ("sample", "Sample", 120),
            ("depth", "Depth", 120),
            ("code", "Test Code", 100),
            ("name", "Test Name", 280),
            ("cost", "Cost", 80),
            ("status", "Status", 100),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=tk.W)

        self.tree.pack(fill=tk.BOTH, expand=True)
        ttk.Button(top, text="Delete Selected", command=self._delete_selected).pack(anchor=tk.E, pady=5)

    def refresh(self):
        project_id = self.get_project_id()
        if project_id:
            self._sync_rate_prices(project_id)
        self._refresh_combos()
        self._refresh_assignments()

    def _refresh_combos(self):
        project_id = self.get_project_id()
        if not project_id:
            self.sample_combo["values"] = []
            self.test_list.delete(0, tk.END)
            return

        conn = get_connection()
        samples = conn.execute(
            "SELECT id, sample_name, depth_raw FROM samples WHERE project_id = ? ORDER BY sample_name",
            (project_id,),
        ).fetchall()
        tests = conn.execute("SELECT id, code, name, default_cost FROM tests ORDER BY code").fetchall()
        conn.close()

        self.sample_map = {}
        for s in samples:
            depth = (s["depth_raw"] or "").strip()
            label = f"{s['sample_name']} @ {depth}" if depth else s["sample_name"]
            key = f"{label} (#{s['id']})"
            self.sample_map[key] = s["id"]
        self.test_map = {
            f"{t['code']} - {t['name']}": (t["id"], t["default_cost"])
            for t in tests
            if t["name"] != "Moisture and Density"
        }

        self.sample_combo["values"] = list(self.sample_map.keys())
        self.test_list.delete(0, tk.END)
        for key in self.test_map.keys():
            self.test_list.insert(tk.END, key)

    def _refresh_assignments(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        project_id = self.get_project_id()
        if not project_id:
            return

        conn = get_connection()
        rows = conn.execute(
            """
            SELECT st.id, s.sample_name, s.depth_raw, t.code, t.name, st.cost, st.status
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
                    row["cost"],
                    row["status"],
                ),
            )

    def _assign_test(self):
        project_id = self.get_project_id()
        if not project_id:
            messagebox.showerror("No Project", "Select a project first.")
            return

        sample_key = self.sample_choice.get()
        if sample_key not in self.sample_map:
            messagebox.showerror("Missing", "Select a sample.")
            return

        sample_id = self.sample_map[sample_key]
        selected_indices = list(self.test_list.curselection())
        if not selected_indices:
            messagebox.showerror("Missing", "Select one or more tests.")
            return

        override_cost = self.assign_cost.get().strip()
        if override_cost:
            try:
                override_cost = float(override_cost)
            except ValueError:
                messagebox.showerror("Invalid", "Override cost must be a number.")
                return
        else:
            override_cost = None

        status = self.assign_status.get().strip() or "scheduled"

        conn = get_connection()
        existing = conn.execute(
            "SELECT test_id FROM sample_tests WHERE sample_id = ?",
            (sample_id,),
        ).fetchall()
        existing_ids = {r["test_id"] for r in existing}

        for idx in selected_indices:
            test_key = self.test_list.get(idx)
            if test_key not in self.test_map:
                continue
            test_id, default_cost = self.test_map[test_key]
            if test_id in existing_ids:
                continue
            if override_cost is not None:
                cost = override_cost
            else:
                cost = float(self._get_rate_price(test_id, default_cost))
            conn.execute(
                "INSERT INTO sample_tests (sample_id, test_id, cost, status) VALUES (?, ?, ?, ?)",
                (sample_id, test_id, cost, status),
            )
        conn.commit()
        conn.close()
        self._refresh_assignments()
        if self.on_tests_changed:
            self.on_tests_changed()

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        assignment_id = int(sel[0])
        if not messagebox.askyesno("Confirm", "Delete this test assignment?"):
            return
        conn = get_connection()
        conn.execute("DELETE FROM sample_tests WHERE id = ?", (assignment_id,))
        conn.commit()
        conn.close()
        self._refresh_assignments()
        if self.on_tests_changed:
            self.on_tests_changed()

    def _get_rate_price(self, test_id, fallback):
        project_id = self.get_project_id()
        if not project_id:
            return fallback
        conn = get_connection()
        row = conn.execute(
            "SELECT billing_rate_id FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()
        if not row or not row["billing_rate_id"]:
            conn.close()
            return fallback
        rate_id = row["billing_rate_id"]
        price_row = conn.execute(
            "SELECT price FROM test_rates WHERE rate_id = ? AND test_id = ?",
            (rate_id, test_id),
        ).fetchone()
        conn.close()
        if price_row:
            return price_row["price"]
        return fallback

    def _sync_rate_prices(self, project_id):
        conn = get_connection()
        conn.execute(
            """
            UPDATE sample_tests
            SET cost = (
                SELECT tr.price
                FROM test_rates tr
                JOIN projects p ON p.billing_rate_id = tr.rate_id
                WHERE p.id = ?
                  AND tr.test_id = sample_tests.test_id
            )
            WHERE sample_id IN (SELECT id FROM samples WHERE project_id = ?)
              AND EXISTS (
                SELECT 1
                FROM test_rates tr
                JOIN projects p ON p.billing_rate_id = tr.rate_id
                WHERE p.id = ?
                  AND tr.test_id = sample_tests.test_id
              )
            """,
            (project_id, project_id, project_id),
        )
        conn.commit()
        conn.close()

    def _on_test_selected(self, _event):
        return
