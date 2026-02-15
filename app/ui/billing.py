import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from app.db import get_connection
from app.services.billing_export import export_billing_xlsx, export_billing_pdf


class BillingTab(ttk.Frame):
    def __init__(self, parent, get_project_id):
        super().__init__(parent)
        self.get_project_id = get_project_id
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(
            top,
            columns=("sample", "depth", "code", "name", "cost"),
            show="headings",
        )
        for col, text, width in [
            ("sample", "Sample", 120),
            ("depth", "Depth", 120),
            ("code", "Test Code", 100),
            ("name", "Test Name", 300),
            ("cost", "Cost", 80),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=tk.W)

        self.tree.pack(fill=tk.BOTH, expand=True)

        footer = ttk.Frame(top)
        footer.pack(fill=tk.X, pady=10)
        self.total_var = tk.StringVar(value="Total: $0.00")
        ttk.Label(footer, textvariable=self.total_var).pack(side=tk.LEFT)
        ttk.Button(footer, text="Export Excel", command=self._export_xlsx).pack(side=tk.RIGHT)
        ttk.Button(footer, text="Export PDF", command=self._export_pdf).pack(side=tk.RIGHT, padx=8)

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        project_id = self.get_project_id()
        if not project_id:
            self.total_var.set("Total: $0.00")
            return

        self._sync_rate_prices(project_id)

        conn = get_connection()
        rows = conn.execute(
            """
            SELECT s.sample_name, s.depth_raw, t.code, t.name, st.cost
            FROM sample_tests st
            JOIN samples s ON s.id = st.sample_id
            JOIN tests t ON t.id = st.test_id
            WHERE s.project_id = ?
            ORDER BY s.sample_name, t.code
            """,
            (project_id,),
        ).fetchall()
        conn.close()

        total = 0.0
        for row in rows:
            cost = float(row["cost"])
            total += cost
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row["sample_name"],
                    row["depth_raw"] or "",
                    row["code"],
                    row["name"],
                    f"{cost:.2f}",
                ),
            )

        self.total_var.set(f"Total: ${total:.2f}")

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

    def _fetch_export_data(self):
        project_id = self.get_project_id()
        if not project_id:
            messagebox.showerror("No Project", "Select a project first.")
            return None

        conn = get_connection()
        project = conn.execute(
            "SELECT file_number, job_name, client_type, billing_rate_id FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()

        line_items = conn.execute(
            """
            SELECT s.sample_name, s.depth_raw, t.code, t.name, st.cost
            FROM sample_tests st
            JOIN samples s ON s.id = st.sample_id
            JOIN tests t ON t.id = st.test_id
            WHERE s.project_id = ?
            ORDER BY s.sample_name, t.code
            """,
            (project_id,),
        ).fetchall()
        conn.close()

        if not line_items:
            messagebox.showerror("Empty", "No tests assigned for this project.")
            return None

        return dict(project), [
            {
                "sample_name": r["sample_name"],
                "depth_raw": r["depth_raw"] or "",
                "test_code": r["code"],
                "test_name": r["name"],
                "cost": r["cost"],
            }
            for r in line_items
        ]

    def _export_xlsx(self):
        data = self._fetch_export_data()
        if not data:
            return
        project, line_items = data
        default_name = f"Billing_{project['file_number']}.xlsx"
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name)
        if not path:
            return

        export_billing_xlsx(path, project=project, line_items=line_items)
        messagebox.showinfo("Exported", f"Billing exported to:\n{path}")

    def _export_pdf(self):
        data = self._fetch_export_data()
        if not data:
            return
        project, line_items = data
        default_name = f"Billing_{project['file_number']}.pdf"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_name)
        if not path:
            return
        try:
            export_billing_pdf(path, project=project, line_items=line_items)
        except Exception as exc:
            messagebox.showerror("PDF Export", f"Failed to export PDF: {exc}")
            return
        messagebox.showinfo("Exported", f"Billing exported to:\n{path}")
