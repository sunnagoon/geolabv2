import tkinter as tk
from tkinter import ttk, messagebox

from app.db import get_connection


class RatesTab(ttk.Frame):
    def __init__(self, parent, on_rates_changed):
        super().__init__(parent)
        self.on_rates_changed = on_rates_changed
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left = ttk.Frame(top)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        right = ttk.Frame(top)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.rate_tree = ttk.Treeview(left, columns=("rate", "client", "year", "kind"), show="headings")
        for col, text, width in [
            ("rate", "Rate ID", 160),
            ("client", "Client", 80),
            ("year", "Year", 70),
            ("kind", "Kind", 90),
        ]:
            self.rate_tree.heading(col, text=text)
            self.rate_tree.column(col, width=width, anchor=tk.W)
        self.rate_tree.pack(fill=tk.BOTH, expand=True)
        self.rate_tree.bind("<<TreeviewSelect>>", self._on_rate_select)

        rate_form = ttk.LabelFrame(left, text="Add Billing Rate")
        rate_form.pack(fill=tk.X, pady=8)

        self.rate_id = tk.StringVar()
        self.client_type = tk.StringVar(value="TCI")
        self.rate_year = tk.StringVar()
        self.rate_kind = tk.StringVar()
        self.rate_notes = tk.StringVar()

        ttk.Label(rate_form, text="Rate ID").grid(row=0, column=0, sticky=tk.W, padx=5, pady=4)
        ttk.Entry(rate_form, textvariable=self.rate_id, width=20).grid(row=0, column=1, sticky=tk.W, padx=5, pady=4)
        ttk.Label(rate_form, text="Client").grid(row=0, column=2, sticky=tk.W, padx=5, pady=4)
        ttk.Combobox(rate_form, textvariable=self.client_type, values=["TCI", "GCI", "SBCI", "CUSTOM"], width=10).grid(
            row=0, column=3, sticky=tk.W, padx=5, pady=4
        )

        ttk.Label(rate_form, text="Year").grid(row=1, column=0, sticky=tk.W, padx=5, pady=4)
        ttk.Entry(rate_form, textvariable=self.rate_year, width=10).grid(row=1, column=1, sticky=tk.W, padx=5, pady=4)
        ttk.Label(rate_form, text="Kind").grid(row=1, column=2, sticky=tk.W, padx=5, pady=4)
        ttk.Entry(rate_form, textvariable=self.rate_kind, width=12).grid(row=1, column=3, sticky=tk.W, padx=5, pady=4)

        ttk.Label(rate_form, text="Notes").grid(row=2, column=0, sticky=tk.W, padx=5, pady=4)
        ttk.Entry(rate_form, textvariable=self.rate_notes, width=40).grid(row=2, column=1, columnspan=3, sticky=tk.W, padx=5, pady=4)

        ttk.Button(rate_form, text="Save Rate", command=self._save_rate).grid(row=0, column=4, rowspan=2, padx=5, pady=4)
        ttk.Button(rate_form, text="Delete Selected", command=self._delete_rate).grid(row=2, column=4, padx=5, pady=4)

        price_form = ttk.LabelFrame(right, text="Set Test Prices for Selected Rate")
        price_form.pack(fill=tk.X, pady=8)

        self.selected_rate_var = tk.StringVar(value="Selected Rate: None")
        ttk.Label(price_form, textvariable=self.selected_rate_var).grid(row=0, column=0, columnspan=4, sticky=tk.W, padx=5, pady=4)

        self.test_choice = tk.StringVar()
        self.test_price = tk.StringVar()

        ttk.Label(price_form, text="Test").grid(row=1, column=0, sticky=tk.W, padx=5, pady=4)
        self.test_combo = ttk.Combobox(price_form, textvariable=self.test_choice, width=35)
        self.test_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=4)

        ttk.Label(price_form, text="Price").grid(row=1, column=2, sticky=tk.W, padx=5, pady=4)
        ttk.Entry(price_form, textvariable=self.test_price, width=12).grid(row=1, column=3, sticky=tk.W, padx=5, pady=4)

        ttk.Button(price_form, text="Save Price", command=self._save_price).grid(row=1, column=4, padx=5, pady=4)
        ttk.Button(price_form, text="Update Rates Everywhere", command=self._update_everywhere).grid(
            row=2, column=4, padx=5, pady=4
        )

        self.price_tree = ttk.Treeview(right, columns=("code", "name", "price"), show="headings")
        for col, text, width in [
            ("code", "Code", 80),
            ("name", "Test Name", 220),
            ("price", "Price", 90),
        ]:
            self.price_tree.heading(col, text=text)
            self.price_tree.column(col, width=width, anchor=tk.W)
        self.price_tree.pack(fill=tk.BOTH, expand=True)

        ttk.Button(right, text="Delete Selected Price", command=self._delete_price).pack(anchor=tk.E, pady=6)

    def refresh(self):
        self._refresh_rates()
        self._refresh_tests()
        self._refresh_prices()

    def _refresh_rates(self):
        for item in self.rate_tree.get_children():
            self.rate_tree.delete(item)
        conn = get_connection()
        rows = conn.execute(
            "SELECT id, rate_id, client_type, year, kind FROM billing_rates ORDER BY rate_id"
        ).fetchall()
        conn.close()
        for row in rows:
            self.rate_tree.insert(
                "",
                tk.END,
                iid=row["id"],
                values=(row["rate_id"], row["client_type"], row["year"] or "", row["kind"] or ""),
            )

    def _refresh_tests(self):
        conn = get_connection()
        tests = conn.execute("SELECT id, code, name FROM tests ORDER BY code").fetchall()
        conn.close()
        self.test_map = {f"{t['code']} - {t['name']}": t["id"] for t in tests}
        self.test_combo["values"] = list(self.test_map.keys())

    def _refresh_prices(self):
        for item in self.price_tree.get_children():
            self.price_tree.delete(item)
        rate_id = self._get_selected_rate_id()
        if not rate_id:
            self.selected_rate_var.set("Selected Rate: None")
            return
        self.selected_rate_var.set(f"Selected Rate: {rate_id}")
        conn = get_connection()
        rows = conn.execute(
            """
            SELECT tr.id, t.code, t.name, tr.price
            FROM test_rates tr
            JOIN tests t ON t.id = tr.test_id
            WHERE tr.rate_id = ?
            ORDER BY t.code
            """,
            (rate_id,),
        ).fetchall()
        conn.close()
        for row in rows:
            self.price_tree.insert(
                "",
                tk.END,
                iid=row["id"],
                values=(row["code"], row["name"], f"{row['price']:.2f}"),
            )

    def _get_selected_rate_id(self):
        sel = self.rate_tree.selection()
        if not sel:
            return None
        rate_row = self.rate_tree.item(sel[0])["values"]
        return rate_row[0] if rate_row else None

    def _on_rate_select(self, _event):
        self._refresh_prices()

    def _save_rate(self):
        rate_id = self.rate_id.get().strip()
        if not rate_id:
            messagebox.showerror("Missing", "Rate ID is required.")
            return
        client_type = self.client_type.get().strip() or "CUSTOM"
        year = self.rate_year.get().strip() or None
        kind = self.rate_kind.get().strip() or None
        notes = self.rate_notes.get().strip() or None

        conn = get_connection()
        try:
            conn.execute(
                "INSERT INTO billing_rates (rate_id, client_type, year, kind, notes) VALUES (?, ?, ?, ?, ?)",
                (rate_id, client_type, year, kind, notes),
            )
            conn.commit()
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to save rate: {exc}")
        finally:
            conn.close()
        self._refresh_rates()
        self.on_rates_changed()

    def _delete_rate(self):
        rate_id = self._get_selected_rate_id()
        if not rate_id:
            return
        if not messagebox.askyesno("Confirm", f"Delete rate {rate_id} and all its prices?"):
            return
        conn = get_connection()
        conn.execute("DELETE FROM billing_rates WHERE rate_id = ?", (rate_id,))
        conn.commit()
        conn.close()
        self._refresh_rates()
        self._refresh_prices()
        self.on_rates_changed()

    def _save_price(self):
        rate_id = self._get_selected_rate_id()
        if not rate_id:
            messagebox.showerror("Missing", "Select a billing rate first.")
            return
        test_key = self.test_choice.get()
        if test_key not in self.test_map:
            messagebox.showerror("Missing", "Select a test.")
            return
        try:
            price = float(self.test_price.get().strip())
        except ValueError:
            messagebox.showerror("Invalid", "Price must be a number.")
            return

        test_id = self.test_map[test_key]
        conn = get_connection()
        conn.execute(
            """
            INSERT INTO test_rates (rate_id, test_id, price)
            VALUES (?, ?, ?)
            ON CONFLICT(rate_id, test_id) DO UPDATE SET price = excluded.price
            """,
            (rate_id, test_id, price),
        )
        conn.commit()
        conn.close()
        self._refresh_prices()
        self.on_rates_changed()

    def _delete_price(self):
        sel = self.price_tree.selection()
        if not sel:
            return
        price_id = int(sel[0])
        if not messagebox.askyesno("Confirm", "Delete this test price?"):
            return
        conn = get_connection()
        conn.execute("DELETE FROM test_rates WHERE id = ?", (price_id,))
        conn.commit()
        conn.close()
        self._refresh_prices()
        self.on_rates_changed()

    def _update_everywhere(self):
        self.on_rates_changed()
