import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.db import DB_PATH, backup_database, get_app_setting, set_app_setting


class SettingsTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.backup_dir_var = tk.StringVar(value="")
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        wrap = ttk.Frame(self)
        wrap.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        db_box = ttk.LabelFrame(wrap, text="Database")
        db_box.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(db_box, text=f"Active DB: {DB_PATH}").pack(anchor=tk.W, padx=8, pady=8)

        backup_box = ttk.LabelFrame(wrap, text="Backup")
        backup_box.pack(fill=tk.X)

        ttk.Label(backup_box, text="Backup Folder").grid(row=0, column=0, sticky=tk.W, padx=8, pady=8)
        ttk.Entry(backup_box, textvariable=self.backup_dir_var, width=70).grid(row=0, column=1, sticky=tk.W, padx=8, pady=8)
        ttk.Button(backup_box, text="Browse", command=self._browse_backup_dir).grid(row=0, column=2, sticky=tk.W, padx=8, pady=8)

        actions = ttk.Frame(backup_box)
        actions.grid(row=1, column=1, sticky=tk.W, padx=8, pady=(0, 10))
        ttk.Button(actions, text="Save Backup Folder", command=self._save_backup_dir).pack(side=tk.LEFT)
        ttk.Button(actions, text="Backup Now", command=self._backup_now).pack(side=tk.LEFT, padx=(8, 0))

        tip = (
            "Recommended: pick a cloud-synced folder (OneDrive/Dropbox) "
            "or a network share so backups are resilient."
        )
        ttk.Label(backup_box, text=tip).grid(row=2, column=0, columnspan=3, sticky=tk.W, padx=8, pady=(0, 8))

    def refresh(self):
        saved = get_app_setting("backup_dir", "")
        self.backup_dir_var.set(saved or "")

    def _browse_backup_dir(self):
        initial = self.backup_dir_var.get().strip() or str(Path.home())
        chosen = filedialog.askdirectory(initialdir=initial)
        if chosen:
            self.backup_dir_var.set(chosen)

    def _save_backup_dir(self):
        folder = self.backup_dir_var.get().strip()
        if not folder:
            messagebox.showerror("Missing Folder", "Select a backup folder first.")
            return
        set_app_setting("backup_dir", folder)
        messagebox.showinfo("Saved", f"Backup folder saved:\n{folder}")

    def _backup_now(self):
        folder = self.backup_dir_var.get().strip() or get_app_setting("backup_dir", "")
        if not folder:
            messagebox.showerror("Missing Folder", "Select and save a backup folder first.")
            return
        try:
            out_path = backup_database(folder)
        except Exception as exc:
            messagebox.showerror("Backup Failed", f"Could not create backup:\n{exc}")
            return
        if self.backup_dir_var.get().strip() != folder:
            self.backup_dir_var.set(folder)
        set_app_setting("backup_dir", folder)
        messagebox.showinfo("Backup Complete", f"Database backup created:\n{out_path}")
