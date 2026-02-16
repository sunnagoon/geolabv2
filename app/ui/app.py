import tkinter as tk
from tkinter import ttk
from pathlib import Path

from app.ui.projects import ProjectsTab
from app.ui.samples import SamplesTab
from app.ui.tests import TestsTab
from app.ui.billing import BillingTab
from app.ui.rates import RatesTab
from app.ui.results import ResultsTab
from app.ui.map_view import MapTab
from app.ui.worksheets import WorksheetsTab
from app.ui.calculations import CalculationsTab
from app.ui.settings import SettingsTab
from app.db import get_connection


class GeoLabApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GeoLab v2")
        self.geometry("1200x760")
        self.minsize(1100, 700)

        self.selected_project_id = None

        self._build_ui()

    def _build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # TerraPacific-inspired light blue palette with green accent.
        bg = "#ecf4fb"
        surface = "#f8fbff"
        surface_alt = "#e1eef9"
        text = "#15385b"
        muted = "#5a7b9a"
        accent = "#2f86de"
        border = "#bfd5ea"

        self.configure(bg=bg)

        style.configure(
            ".",
            background=bg,
            foreground=text,
            fieldbackground=surface,
            bordercolor=border,
            highlightcolor=border,
            insertcolor=text,
            selectbackground=accent,
            selectforeground="#0b1b14",
            font=("Segoe UI", 10),
        )

        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg, foreground=text)
        style.configure("TLabelframe", background=bg, foreground=muted)
        style.configure("TLabelframe.Label", background=bg, foreground=muted, font=("Segoe UI Semibold", 10))
        style.configure(
            "TButton",
            background=surface_alt,
            foreground=text,
            padding=(10, 6),
            borderwidth=1,
            relief="flat",
        )
        style.map(
            "TButton",
            background=[("active", "#d6e8f9")],
            foreground=[("disabled", muted)],
        )

        style.configure(
            "TEntry",
            fieldbackground=surface,
            background=surface,
            foreground=text,
            padding=4,
        )
        style.configure(
            "TCombobox",
            fieldbackground=surface,
            background=surface,
            foreground=text,
            arrowcolor=muted,
            padding=4,
        )

        style.configure(
            "Treeview",
            background=surface,
            fieldbackground=surface,
            foreground=text,
            bordercolor=border,
            rowheight=26,
        )
        style.configure(
            "Treeview.Heading",
            background=surface_alt,
            foreground=text,
            font=("Segoe UI Semibold", 10),
        )
        style.map(
            "Treeview",
            background=[("selected", accent)],
            foreground=[("selected", "#0b1b14")],
        )

        header = ttk.Frame(self)
        header.pack(fill=tk.X, padx=10, pady=8)
        self._logo_img = None
        logo_path = Path(__file__).resolve().parents[2] / "assets" / "terrapacific_logo.png"
        if logo_path.exists():
            try:
                self._logo_img = tk.PhotoImage(file=str(logo_path))
                ttk.Label(header, image=self._logo_img).pack(side=tk.LEFT, padx=(0, 10))
            except Exception:
                self._logo_img = None

        brand = ttk.Frame(header)
        brand.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(brand, text="TerraPacific Consultants Inc", font=("Segoe UI Semibold", 11), foreground="#2f86de").pack(
            anchor=tk.W
        )
        self.project_label_var = tk.StringVar(value="Selected Project: None")
        ttk.Label(brand, textvariable=self.project_label_var, font=("Segoe UI", 10)).pack(anchor=tk.W)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.projects_tab = ProjectsTab(self.notebook, on_project_selected=self._on_project_selected)
        self.samples_tab = SamplesTab(
            self.notebook,
            get_project_id=self._get_project_id,
            on_samples_changed=self._on_samples_changed,
        )
        self.tests_tab = TestsTab(
            self.notebook,
            get_project_id=self._get_project_id,
            on_tests_changed=self._on_tests_changed,
        )
        self.billing_tab = BillingTab(self.notebook, get_project_id=self._get_project_id)
        self.rates_tab = RatesTab(self.notebook, on_rates_changed=self._on_rates_changed)
        self.results_tab = ResultsTab(self.notebook, get_project_id=self._get_project_id)
        self.worksheets_tab = WorksheetsTab(
            self.notebook,
            get_project_id=self._get_project_id,
            on_saved=self._on_worksheets_saved,
        )
        self.calculations_tab = CalculationsTab(self.notebook, get_project_id=self._get_project_id)
        self.map_tab = MapTab(self.notebook)
        self.settings_tab = SettingsTab(self.notebook)

        self.notebook.add(self.projects_tab, text="Projects")
        self.notebook.add(self.samples_tab, text="Samples")
        self.notebook.add(self.tests_tab, text="Tests")
        self.notebook.add(self.results_tab, text="Results")
        self.notebook.add(self.worksheets_tab, text="Worksheets")
        self.notebook.add(self.calculations_tab, text="Calculations")
        self.notebook.add(self.billing_tab, text="Billing")
        self.notebook.add(self.rates_tab, text="Rates")
        self.notebook.add(self.map_tab, text="Map")
        self.notebook.add(self.settings_tab, text="Settings")

        self._set_project_enabled(False)

    def _on_project_selected(self, project_id):
        self.selected_project_id = project_id
        self._update_project_label()
        self._set_project_enabled(bool(project_id))
        self.samples_tab.refresh()
        self.tests_tab.refresh()
        self.results_tab.refresh()
        self.worksheets_tab.refresh()
        self.calculations_tab.refresh()
        self.billing_tab.refresh()
        self.rates_tab.refresh()
        self.map_tab.refresh()
        self.settings_tab.refresh()

    def _get_project_id(self):
        return self.selected_project_id

    def _set_project_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        self.notebook.tab(self.samples_tab, state=state)
        self.notebook.tab(self.tests_tab, state=state)
        self.notebook.tab(self.worksheets_tab, state=state)
        self.notebook.tab(self.calculations_tab, state=state)
        self.notebook.tab(self.billing_tab, state=state)

    def _update_project_label(self):
        if not self.selected_project_id:
            self.project_label_var.set("Selected Project: None")
            return
        conn = get_connection()
        row = conn.execute(
            "SELECT file_number, job_name FROM projects WHERE id = ?",
            (self.selected_project_id,),
        ).fetchone()
        conn.close()
        if not row:
            self.project_label_var.set("Selected Project: None")
            return
        self.project_label_var.set(f"Selected Project: {row['file_number']} - {row['job_name']}")

    def _on_rates_changed(self):
        self.projects_tab.refresh_rates()
        self.tests_tab.refresh()
        self.billing_tab.refresh()
        self.results_tab.refresh()
        self.worksheets_tab.refresh()
        self.calculations_tab.refresh()
        self.map_tab.refresh()

    def _on_tests_changed(self):
        self.tests_tab.refresh()
        self.results_tab.refresh()
        self.billing_tab.refresh()
        self.worksheets_tab.refresh()
        self.calculations_tab.refresh()

    def _on_samples_changed(self):
        self.tests_tab.refresh()
        self.results_tab.refresh()
        self.worksheets_tab.refresh()
        self.billing_tab.refresh()
        self.calculations_tab.refresh()
        self.map_tab.refresh()

    def _on_worksheets_saved(self):
        self.results_tab.refresh()
        self.billing_tab.refresh()
        self.tests_tab.refresh()
        self.calculations_tab.refresh()
