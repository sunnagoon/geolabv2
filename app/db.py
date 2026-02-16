import sqlite3
import sys
import os
from pathlib import Path
from datetime import datetime


def _resolve_db_path():
    # In a frozen executable, keep DB in a persistent user location.
    if getattr(sys, "frozen", False):
        root = Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / "GeoLab" / "data"
        return root / "geolab.db"
    return Path(__file__).resolve().parent.parent / "data" / "geolab.db"


DB_PATH = _resolve_db_path()


def get_connection():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_db():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_keys = ON;")

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_number TEXT NOT NULL UNIQUE,
            job_name TEXT NOT NULL,
            client_type TEXT NOT NULL,
            client_name TEXT NOT NULL,
            billing_rate_id TEXT NOT NULL,
            billing_year INTEGER,
            billing_kind TEXT,
            location_text TEXT,
            latitude REAL,
            longitude REAL,
            status TEXT NOT NULL DEFAULT 'Not Scheduled',
            created_at TEXT NOT NULL
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS samples (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            sample_name TEXT NOT NULL,
            sample_type TEXT,
            depth_raw TEXT,
            depth_from REAL,
            depth_to REAL,
            depth_unit TEXT,
            received_date TEXT,
            storage_location TEXT,
            disposal_date TEXT,
            status TEXT,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS tests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            default_cost REAL NOT NULL
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS sample_tests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sample_id INTEGER NOT NULL,
            test_id INTEGER NOT NULL,
            cost REAL NOT NULL,
            status TEXT NOT NULL,
            result_summary TEXT,
            completed_date TEXT,
            FOREIGN KEY(sample_id) REFERENCES samples(id) ON DELETE CASCADE,
            FOREIGN KEY(test_id) REFERENCES tests(id) ON DELETE RESTRICT
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS billing_rates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rate_id TEXT NOT NULL UNIQUE,
            client_type TEXT NOT NULL,
            year INTEGER,
            kind TEXT,
            notes TEXT
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS test_rates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rate_id TEXT NOT NULL,
            test_id INTEGER NOT NULL,
            price REAL NOT NULL,
            FOREIGN KEY(rate_id) REFERENCES billing_rates(rate_id) ON DELETE CASCADE,
            FOREIGN KEY(test_id) REFERENCES tests(id) ON DELETE CASCADE,
            UNIQUE(rate_id, test_id)
        );
        """
    )

    cur.execute("CREATE INDEX IF NOT EXISTS idx_samples_project ON samples(project_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_sample_tests_sample ON sample_tests(sample_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_sample_tests_test ON sample_tests(test_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_test_rates_rate ON test_rates(rate_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_test_rates_test ON test_rates(test_id);")

    _migrate_projects(cur)
    _migrate_samples(cur)
    _migrate_sample_tests(cur)
    _migrate_worksheets(cur)
    _migrate_settings(cur)
    _seed_tests(cur)
    _seed_rate_prices(cur)

    conn.commit()
    conn.close()


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


def get_app_setting(key, default=None):
    conn = get_connection()
    row = conn.execute("SELECT value FROM app_settings WHERE key = ?", (key,)).fetchone()
    conn.close()
    if not row:
        return default
    return row["value"]


def set_app_setting(key, value):
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO app_settings (key, value)
        VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
        """,
        (key, value),
    )
    conn.commit()
    conn.close()


def backup_database(target_dir):
    out_dir = Path(target_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"geolab_backup_{stamp}.db"

    src = get_connection()
    dst = sqlite3.connect(str(out_path))
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    return out_path


def _seed_tests(cur):
    test_names = [
        "Moisture Content",
        "Field Density/Moisture",
        "Sand Cone",
        "Sieve Part. Analysis",
        "-200 Washed Sieve",
        "LL/PL",
        "Atterberg Limits",
        "Max Density",
        "Expansion Index",
        "Consol",
        "Hydro Response",
        "Swell/Hydro",
        "Hydrometer",
        "Core Measurements",
        "Petrographic Analysis",
        "Direct Shear",
        "Chem",
        "R-Value",
    ]

    def make_code(name, used_codes):
        base = "".join(ch for ch in name.upper() if ch.isalnum() or ch == " ")
        parts = [p for p in base.split() if p]
        code = "".join(p[0] for p in parts)[:6]
        if not code:
            code = base[:6] or "TEST"
        candidate = code
        suffix = 2
        while candidate in used_codes:
            candidate = f"{code}{suffix}"
            suffix += 1
        used_codes.add(candidate)
        return candidate

    used = set()
    rows = cur.execute("SELECT code FROM tests").fetchall()
    used.update(r["code"] for r in rows)

    for name in test_names:
        exists = cur.execute("SELECT id FROM tests WHERE name = ?", (name,)).fetchone()
        if exists:
            continue
        code = make_code(name, used)
        cur.execute("INSERT INTO tests (code, name, default_cost) VALUES (?, ?, ?)", (code, name, 0.0))


def _seed_rate_prices(cur):
    seed_rates = [
        ("TCI2026T", "TCI", 2026, "T"),
        ("TCI2026TR", "TCI", 2026, "TR"),
        ("TCI2025T", "TCI", 2025, "T"),
        ("TCI2025TR", "TCI", 2025, "TR"),
        ("GCI", "GCI", None, None),
        ("SBCI", "SBCI", None, None),
    ]
    for rate_id, client_type, year, kind in seed_rates:
        cur.execute(
            """
            INSERT INTO billing_rates (rate_id, client_type, year, kind, notes)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(rate_id) DO NOTHING
            """,
            (rate_id, client_type, year, kind, "Seeded base rate"),
        )

    rate_id = "TCI2026T"
    existing = cur.execute(
        "SELECT COUNT(1) AS c FROM test_rates WHERE rate_id = ?",
        (rate_id,),
    ).fetchone()
    if existing and existing["c"] > 0:
        return

    prices = {
        "Moisture Content": 30.00,
        "Field Density/Moisture": 45.00,
        "Sieve Part. Analysis": 160.00,
        "-200 Washed Sieve": 80.00,
        "LL/PL": 175.00,
        "Max Density": 235.00,
        "Expansion Index": 160.00,
        "Consol": 265.00,
        "Swell/Hydro": 100.00,
        "Hydrometer": 160.00,
        "Core Measurements": 45.00,
        "Petrographic Analysis": 1320.00,
        "Direct Shear": 335.00,
        "Chem": 50.00,
        "R-Value": 245.00,
    }

    for name, price in prices.items():
        row = cur.execute("SELECT id FROM tests WHERE name = ?", (name,)).fetchone()
        if not row:
            continue
        cur.execute(
            """
            INSERT INTO test_rates (rate_id, test_id, price)
            VALUES (?, ?, ?)
            ON CONFLICT(rate_id, test_id) DO NOTHING
            """,
            (rate_id, row["id"], price),
        )


def _migrate_sample_tests(cur):
    cols = [r["name"] for r in cur.execute("PRAGMA table_info(sample_tests)").fetchall()]
    if "result_value" not in cols:
        cur.execute("ALTER TABLE sample_tests ADD COLUMN result_value REAL;")
    if "result_unit" not in cols:
        cur.execute("ALTER TABLE sample_tests ADD COLUMN result_unit TEXT;")
    if "result_value2" not in cols:
        cur.execute("ALTER TABLE sample_tests ADD COLUMN result_value2 REAL;")
    if "result_unit2" not in cols:
        cur.execute("ALTER TABLE sample_tests ADD COLUMN result_unit2 TEXT;")
    if "result_value3" not in cols:
        cur.execute("ALTER TABLE sample_tests ADD COLUMN result_value3 REAL;")
    if "result_unit3" not in cols:
        cur.execute("ALTER TABLE sample_tests ADD COLUMN result_unit3 TEXT;")
    if "result_value4" not in cols:
        cur.execute("ALTER TABLE sample_tests ADD COLUMN result_value4 REAL;")
    if "result_unit4" not in cols:
        cur.execute("ALTER TABLE sample_tests ADD COLUMN result_unit4 TEXT;")
    if "result_notes" not in cols:
        cur.execute("ALTER TABLE sample_tests ADD COLUMN result_notes TEXT;")


def _migrate_projects(cur):
    cols = [r["name"] for r in cur.execute("PRAGMA table_info(projects)").fetchall()]
    if "status" not in cols:
        cur.execute("ALTER TABLE projects ADD COLUMN status TEXT NOT NULL DEFAULT 'Not Scheduled';")
    if "location_text" not in cols:
        cur.execute("ALTER TABLE projects ADD COLUMN location_text TEXT;")
    if "latitude" not in cols:
        cur.execute("ALTER TABLE projects ADD COLUMN latitude REAL;")
    if "longitude" not in cols:
        cur.execute("ALTER TABLE projects ADD COLUMN longitude REAL;")


def _migrate_samples(cur):
    cols = [r["name"] for r in cur.execute("PRAGMA table_info(samples)").fetchall()]
    if "sample_type" not in cols:
        cur.execute("ALTER TABLE samples ADD COLUMN sample_type TEXT;")


def _migrate_worksheets(cur):
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS astm1557_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sample_test_id INTEGER NOT NULL UNIQUE,
            points_json TEXT NOT NULL,
            max_dry_density REAL,
            opt_moisture REAL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(sample_test_id) REFERENCES sample_tests(id) ON DELETE CASCADE
        );
        """
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_astm1557_sample_test ON astm1557_runs(sample_test_id);")
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS worksheet_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sample_test_id INTEGER NOT NULL UNIQUE,
            worksheet_key TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(sample_test_id) REFERENCES sample_tests(id) ON DELETE CASCADE
        );
        """
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_worksheet_runs_sample_test ON worksheet_runs(sample_test_id);")
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS grain_size_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sample_id INTEGER NOT NULL UNIQUE,
            payload_json TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(sample_id) REFERENCES samples(id) ON DELETE CASCADE
        );
        """
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_grain_size_runs_sample ON grain_size_runs(sample_id);")
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS calculations_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL UNIQUE,
            calc_key TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            computed_json TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        """
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_calculations_runs_project ON calculations_runs(project_id);")


def _migrate_settings(cur):
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );
        """
    )
