import json
from pathlib import Path

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
except Exception:
    canvas = None


WORKSHEET_SPECS = {
    "Sand Cone": {
        "key": "sand_cone",
        "title": "SAND CONE DENSITY TEST",
        "astm": "ASTM D 1556",
        "fields": [
            ("a_begin", "A Begin: Sand/Bottle/Cone", "lb"),
            ("b_end", "B End: Sand/Bottle/Cone", "lb"),
            ("d_cone", "D Sand in Cone (calibrated)", "lb"),
            ("f_sand_density", "F Density of Test Sand", "pcf"),
            ("h_moist_tare", "H Moist Soil + Tare", "lb"),
            ("i_tare", "I Tare", "lb"),
            ("l_rock", "L Rock >3/4 in (optional)", "lb"),
            ("moisture_pct", "Moisture Content", "%"),
        ],
        "computed": [
            ("c_sand_used", "C Sand Used (A-B)", "lb"),
            ("e_sand_hole", "E Sand in Hole (C-D)", "lb"),
            ("g_hole_vol", "G Volume of Hole (E/F)", "ft^3"),
            ("j_moist_soil", "J Moist Soil in Hole (H-I)", "lb"),
            ("k_total_density", "K Total Density (J/G)", "pcf"),
            ("m_rock_vol", "M Rock Volume (L/165.4)", "ft^3"),
            ("n_soil_wt", "N Soil Weight (J-L)", "lb"),
            ("o_soil_vol", "O Soil Volume (G-M)", "ft^3"),
            ("p_corr_total_density", "P Corrected Total Density (N/O)", "pcf"),
            ("dry_density", "Dry Density", "pcf"),
        ],
    },
    "Sieve Part. Analysis": {
        "key": "sieve_part",
        "title": "SIEVE PARTICLE ANALYSIS",
        "astm": "ASTM D 422",
        "fields": [
            ("dry_sample_weight", "Dry Sample Weight", "g"),
            ("minus200_weight", "Minus #200 Weight", "g"),
            ("uscs_class", "USCS Classification", ""),
        ],
        "computed": [
            ("passing_no200", "Passing No. 200", "%"),
        ],
    },
    "-200 Washed Sieve": {
        "key": "washed_200",
        "title": "-200 WASHED SIEVE",
        "astm": "ASTM D 1140",
        "fields": [
            ("a_wet_tare", "A Wet Sample + Tare", "g"),
            ("b_dry_tare", "B Dry Sample + Tare", "g"),
            ("c_tare", "C Tare", "g"),
            ("e_moist_soil", "E Weight of Moist Soil", "g"),
            ("h_dry40_tare", "H Dry + Tare (#40)", "g"),
            ("i_tare40", "I Tare (#40)", "g"),
            ("k_dry200_tare", "K Dry + Tare (#200)", "g"),
            ("l_tare200", "L Tare (#200)", "g"),
            ("uscs_class", "USCS Classification", ""),
        ],
        "computed": [
            ("d_moisture", "D Moisture Content", "%"),
            ("f_dry_sample", "F Dry Sample Weight", "g"),
            ("m_weight", "M Weight", "g"),
            ("n_minus200", "N -200 Sieve Weight", "g"),
            ("o_passing200", "Percent Passing #200", "%"),
        ],
    },
    "Expansion Index": {
        "key": "expansion_index",
        "title": "EXPANSION INDEX TEST",
        "astm": "ASTM D 4829",
        "fields": [
            ("expansion_index", "Expansion Index", "EI"),
        ],
        "computed": [
            ("expansion_potential", "Expansion Potential", ""),
        ],
    },
    "Hydro Response": {
        "key": "hydro_response",
        "title": "HYDRO RESPONSE",
        "astm": "ASTM D 4546",
        "fields": [
            ("hydro_response", "Hydro Response", "%"),
            ("normal_stress", "Normal Stress", "psf"),
        ],
        "computed": [],
    },
    "Field Density/Moisture": {
        "key": "field_moisture_density",
        "title": "FIELD DENSITY/MOISTURE",
        "astm": "ASTM D 2937",
        "fields": [
            ("ring_count", "A Number of Rings", "count", "", False, False),
            ("ring_moist_plus_rings", "B Moist Soil + Rings", "g", "", False, False),
            ("ring_weight_rings", "C Weight of Rings", "g", "", False, False),
            ("wet_sample_tare", "G Wet Sample + Tare", "g", "", False, False),
            ("dry_sample_tare", "H Dry Sample + Tare", "g", "", False, False),
            ("tare_weight", "I Tare", "g", "", False, False),
            ("ring_const", "Ring Constant", "", "5.8081", True, True),
            ("volume_divisor", "Volume Divisor", "", "2200", True, True),
            ("grams_per_lb", "Grams per Pound", "", "453.6", True, True),
            ("unit_wt_water", "Unit Weight of Water", "", "62.4", True, True),
            ("specific_gravity", "Specific Gravity", "", "2.7", True, True),
        ],
        "computed": [
            ("ring_moist_soil", "D Ring Moist Soil", "g"),
            ("ring_volume", "E Ring Volume", "ft^3"),
            ("ring_moist_density", "F Ring Moist Density", "pcf"),
            ("moist_density_used", "F Moist Density Used", "pcf"),
            ("water_weight", "J Weight of Water", "g"),
            ("dry_soil_weight", "K Weight of Dry Soil", "g"),
            ("moisture_content", "L Moisture Content", "%"),
            ("dry_density", "Dry Density", "pcf"),
            ("saturation", "Saturation", "%"),
        ],
    },
}


def get_spec(test_name):
    return WORKSHEET_SPECS.get(test_name)


def compute_values(test_name, payload):
    vals = dict(payload or {})
    out = {}
    if test_name == "Sand Cone":
        a = _num(vals.get("a_begin"))
        b = _num(vals.get("b_end"))
        d = _num(vals.get("d_cone"))
        f = _num(vals.get("f_sand_density"))
        h = _num(vals.get("h_moist_tare"))
        i = _num(vals.get("i_tare"))
        l = _num(vals.get("l_rock")) or 0.0
        moisture = _num(vals.get("moisture_pct"))
        c = _sub(a, b)
        e = _sub(c, d)
        g = _div(e, f)
        j = _sub(h, i)
        k = _div(j, g)
        m = _div(l, 165.4)
        n = _sub(j, l)
        o = _sub(g, m)
        p = _div(n, o)
        dry = None
        if p is not None and moisture is not None:
            dry = p / (1.0 + moisture / 100.0)
        out.update(
            {
                "c_sand_used": c,
                "e_sand_hole": e,
                "g_hole_vol": g,
                "j_moist_soil": j,
                "k_total_density": k,
                "m_rock_vol": m,
                "n_soil_wt": n,
                "o_soil_vol": o,
                "p_corr_total_density": p,
                "dry_density": dry,
            }
        )

    elif test_name == "Sieve Part. Analysis":
        dry = _num(vals.get("dry_sample_weight"))
        minus200 = _num(vals.get("minus200_weight"))
        out["passing_no200"] = _pct(minus200, dry)

    elif test_name == "-200 Washed Sieve":
        a = _num(vals.get("a_wet_tare"))
        b = _num(vals.get("b_dry_tare"))
        c = _num(vals.get("c_tare"))
        e = _num(vals.get("e_moist_soil"))
        h = _num(vals.get("h_dry40_tare"))
        i = _num(vals.get("i_tare40"))
        k = _num(vals.get("k_dry200_tare"))
        l = _num(vals.get("l_tare200"))
        d = None
        if a is not None and b is not None and c is not None and abs((b - c)) > 1e-12:
            d = ((a - b) / (b - c)) * 100.0
        f = None
        if e is not None and d is not None:
            f = e / (1.0 + d / 100.0)
        m = None
        h_minus_i = _sub(h, i)
        k_minus_l = _sub(k, l)
        if h_minus_i is not None and k_minus_l is not None:
            m = h_minus_i + k_minus_l
        n = _sub(f, m)
        o = _pct(n, f)
        out.update(
            {
                "d_moisture": d,
                "f_dry_sample": f,
                "m_weight": m,
                "n_minus200": n,
                "o_passing200": o,
            }
        )

    elif test_name == "Expansion Index":
        idx = _num(vals.get("expansion_index"))
        out["expansion_potential"] = _expansion_potential(idx)

    elif test_name in ("Hydro Response",):
        pass

    elif test_name == "Field Density/Moisture":
        ring_count = _num(vals.get("ring_count"))
        ring_b = _num(vals.get("ring_moist_plus_rings"))
        ring_c = _num(vals.get("ring_weight_rings"))
        ring_const = _num(vals.get("ring_const"))
        volume_divisor = _num(vals.get("volume_divisor"))
        grams_per_lb = _num(vals.get("grams_per_lb"))
        gamma_w = _num(vals.get("unit_wt_water"))
        if ring_const is None:
            ring_const = 5.8081
        if volume_divisor is None:
            volume_divisor = 2200.0
        if grams_per_lb is None:
            grams_per_lb = 453.6
        if gamma_w is None:
            gamma_w = 62.4
        ring_d = _sub(ring_b, ring_c)
        ring_e = None
        if ring_count is not None and volume_divisor:
            ring_e = ring_count * ring_const / volume_divisor
        ring_f = None
        if ring_d is not None and ring_e is not None and abs(ring_e) > 1e-12 and grams_per_lb:
            ring_f = (ring_d / ring_e) / grams_per_lb

        f_used = ring_f
        g = _num(vals.get("wet_sample_tare"))
        h = _num(vals.get("dry_sample_tare"))
        i = _num(vals.get("tare_weight"))
        j = _sub(g, h)
        k = _sub(h, i)
        l = _pct(j, k)
        dry = None
        if f_used is not None and l is not None:
            dry = f_used / (1.0 + l / 100.0)
        gs = _num(vals.get("specific_gravity"))
        if gs is None:
            gs = 2.7
        sat = None
        if l is not None and dry is not None:
            den = (gamma_w / dry) - (1.0 / gs)
            if abs(den) > 1e-12:
                sat = l / den
        if sat is not None:
            sat = max(0.0, min(100.0, sat))
        out.update(
            {
                "ring_moist_soil": ring_d,
                "ring_volume": ring_e,
                "ring_moist_density": ring_f,
                "moist_density_used": f_used,
                "water_weight": j,
                "dry_soil_weight": k,
                "moisture_content": l,
                "dry_density": dry,
                "saturation": sat,
            }
        )

    return out


def map_results(test_name, payload, computed):
    p = payload or {}
    c = computed or {}
    if test_name == "Sand Cone":
        dry = c.get("dry_density")
        if dry is None:
            dry = c.get("p_corr_total_density")
        return {
            "result_value": _round_or_none(dry),
            "result_unit": "pcf",
            "result_value2": _round_or_none(_num(p.get("moisture_pct"))),
            "result_unit2": "%",
            "result_value3": None,
            "result_unit3": "",
            "result_value4": None,
            "result_unit4": "",
            "result_notes": None,
        }
    if test_name == "Sieve Part. Analysis":
        uscs = (p.get("uscs_class") or "").strip()
        return {
            "result_value": None,
            "result_unit": uscs,
            "result_value2": _round_or_none(c.get("passing_no200")),
            "result_unit2": "%",
            "result_value3": None,
            "result_unit3": "",
            "result_value4": None,
            "result_unit4": "",
            "result_notes": None,
        }
    if test_name == "-200 Washed Sieve":
        return {
            "result_value": _round_or_none(c.get("o_passing200")),
            "result_unit": "%",
            "result_value2": None,
            "result_unit2": "",
            "result_value3": None,
            "result_unit3": "",
            "result_value4": None,
            "result_unit4": "",
            "result_notes": (p.get("uscs_class") or "").strip() or None,
        }
    if test_name == "Expansion Index":
        return {
            "result_value": _round_or_none(_num(p.get("expansion_index"))),
            "result_unit": "EI",
            "result_value2": None,
            "result_unit2": c.get("expansion_potential") or "",
            "result_value3": None,
            "result_unit3": "",
            "result_value4": None,
            "result_unit4": "",
            "result_notes": None,
        }
    if test_name == "Hydro Response":
        return {
            "result_value": _round_or_none(_num(p.get("hydro_response"))),
            "result_unit": "%",
            "result_value2": _round_or_none(_num(p.get("normal_stress"))),
            "result_unit2": "psf",
            "result_value3": None,
            "result_unit3": "",
            "result_value4": None,
            "result_unit4": "",
            "result_notes": None,
        }
    if test_name == "Field Density/Moisture":
        return {
            "result_value": _round_or_none(c.get("dry_density")),
            "result_unit": "pcf",
            "result_value2": _round_or_none(c.get("moisture_content")),
            "result_unit2": "%",
            "result_value3": _round_or_none(c.get("saturation")),
            "result_unit3": "%",
            "result_value4": None,
            "result_unit4": "",
            "result_notes": "Method: Ring",
        }
    return {
        "result_value": None,
        "result_unit": "",
        "result_value2": None,
        "result_unit2": "",
        "result_value3": None,
        "result_unit3": "",
        "result_value4": None,
        "result_unit4": "",
        "result_notes": None,
    }


def export_generic_pdf(path, project, sample_label, test_name, payload, computed):
    if canvas is None:
        raise RuntimeError("PDF export requires reportlab. Please install dependencies.")
    spec = get_spec(test_name)
    if not spec:
        raise RuntimeError("Worksheet type is not supported for export.")
    if test_name == "Field Density/Moisture":
        export_field_density_pdf(path, project, sample_label, payload, computed)
        return

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(path, pagesize=letter)
    w, h = letter

    y = h - 38
    c.setFont("Helvetica-Bold", 13)
    c.drawString(36, y, f"{spec['title']} - {spec['astm']}")
    y -= 16
    c.setFont("Helvetica", 9)
    c.drawString(36, y, f"Project: {project.get('job_name', '')}")
    y -= 12
    c.drawString(36, y, f"File Number: {project.get('file_number', '')}")
    y -= 12
    c.drawString(36, y, f"Sample: {sample_label}")
    y -= 18

    c.setFont("Helvetica-Bold", 9)
    c.drawString(36, y, "Input Data")
    y -= 12
    _draw_rows(c, y, spec["fields"], payload, left=36, width=w - 72)
    y -= 16 * max(1, len(spec["fields"])) + 8
    if spec["computed"]:
        c.setFont("Helvetica-Bold", 9)
        c.drawString(36, y, "Computed Values")
        y -= 12
        _draw_rows(c, y, spec["computed"], computed, left=36, width=w - 72)

    c.showPage()
    c.save()


def export_field_density_pdf(path, project, sample_label, payload, computed):
    if canvas is None:
        raise RuntimeError("PDF export requires reportlab. Please install dependencies.")
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(path, pagesize=letter)
    w, h = letter
    y = h - 38
    c.setFont("Helvetica-Bold", 12)
    c.drawString(36, y, "FIELD DENSITY / MOISTURE - ASTM D 2937")
    y -= 14
    c.setFont("Helvetica", 9)
    c.drawString(36, y, f"Project: {project.get('job_name', '')}")
    y -= 12
    c.drawString(36, y, f"File Number: {project.get('file_number', '')}")
    y -= 12
    c.drawString(36, y, f"Sample: {sample_label}")

    vals = dict(payload or {})
    inputs = [
        ("A Number of Rings", vals.get("ring_count"), "count"),
        ("B Moist Soil + Rings", vals.get("ring_moist_plus_rings"), "g"),
        ("C Weight of Rings", vals.get("ring_weight_rings"), "g"),
        ("G Wet Sample + Tare", vals.get("wet_sample_tare"), "g"),
        ("H Dry Sample + Tare", vals.get("dry_sample_tare"), "g"),
        ("I Tare", vals.get("tare_weight"), "g"),
        ("Ring Constant", vals.get("ring_const") or "5.8081", ""),
        ("Volume Divisor", vals.get("volume_divisor") or "2200", ""),
        ("Grams per Pound", vals.get("grams_per_lb") or "453.6", ""),
        ("Unit Weight Water", vals.get("unit_wt_water") or "62.4", ""),
        ("Specific Gravity", vals.get("specific_gravity") or "2.7", ""),
    ]
    outputs = [
        ("D Ring Moist Soil", computed.get("ring_moist_soil"), "g"),
        ("E Ring Volume", computed.get("ring_volume"), "ft^3"),
        ("F Ring Moist Density", computed.get("ring_moist_density"), "pcf"),
        ("J Weight of Water", computed.get("water_weight"), "g"),
        ("K Weight of Dry Soil", computed.get("dry_soil_weight"), "g"),
        ("L Moisture Content", computed.get("moisture_content"), "%"),
        ("M Dry Density", computed.get("dry_density"), "pcf"),
        ("N Saturation", computed.get("saturation"), "%"),
    ]

    top = h - 98
    _draw_rows(c, top, [(f"in{i}", label, unit) for i, (label, _v, unit) in enumerate(inputs)], {f"in{i}": v for i, (_l, v, _u) in enumerate(inputs)}, 36, 286)
    _draw_rows(c, top, [(f"out{i}", label, unit) for i, (label, _v, unit) in enumerate(outputs)], {f"out{i}": v for i, (_l, v, _u) in enumerate(outputs)}, 334, 242)

    eq_top = 86
    c.setFont("Helvetica-Bold", 8.8)
    c.drawString(334, eq_top + 76, "Computation Notes")
    c.setFont("Helvetica", 8)
    notes = [
        "D = B - C",
        "E = A * RingConstant / VolumeDivisor",
        "F = (D / E) / GramsPerPound",
        "J = G - H",
        "K = H - I",
        "L = (J / K) * 100",
        "M = F / (1 + L/100)",
        "N = L / ((UnitWeightWater/M) - (1/SpecificGravity))",
    ]
    y_note = eq_top + 62
    for line in notes:
        c.drawString(336, y_note, line)
        y_note -= 10

    c.showPage()
    c.save()


def _draw_rows(c, y_top, rows, values, left, width):
    label_w = int(width * 0.62)
    value_w = int(width * 0.2)
    unit_w = width - label_w - value_w
    row_h = 16
    c.setFont("Helvetica", 8.5)
    y = y_top
    for key, label, unit in rows:
        c.rect(left, y - row_h, label_w, row_h)
        c.rect(left + label_w, y - row_h, value_w, row_h)
        c.rect(left + label_w + value_w, y - row_h, unit_w, row_h)
        c.drawString(left + 4, y - 11, str(label))
        raw = (values or {}).get(key)
        txt = _fmt(raw)
        c.drawRightString(left + label_w + value_w - 4, y - 11, txt)
        c.drawString(left + label_w + value_w + 4, y - 11, unit or "")
        y -= row_h


def dumps_payload(payload):
    return json.dumps(payload or {})


def loads_payload(raw):
    if not raw:
        return {}
    try:
        obj = json.loads(raw)
    except Exception:
        return {}
    return obj if isinstance(obj, dict) else {}


def _num(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    text = str(val).strip()
    if not text:
        return None
    try:
        return float(text)
    except Exception:
        return None


def _fmt(val):
    if val is None:
        return ""
    if isinstance(val, str):
        return val
    try:
        return f"{float(val):.3f}".rstrip("0").rstrip(".")
    except Exception:
        return str(val)


def _sub(a, b):
    if a is None or b is None:
        return None
    return a - b


def _div(a, b):
    if a is None or b is None or abs(b) <= 1e-12:
        return None
    return a / b


def _pct(a, b):
    v = _div(a, b)
    if v is None:
        return None
    return v * 100.0


def _linear_fit(points):
    if len(points) < 2:
        return None, None
    n = float(len(points))
    sx = sum(p[0] for p in points)
    sy = sum(p[1] for p in points)
    sxx = sum(p[0] * p[0] for p in points)
    sxy = sum(p[0] * p[1] for p in points)
    den = n * sxx - sx * sx
    if abs(den) <= 1e-12:
        return None, None
    m = (n * sxy - sx * sy) / den
    b = (sy - m * sx) / n
    return m, b


def _viscosity_at_temp(temp_c):
    if temp_c is None:
        return None
    t = float(temp_c)
    keys = sorted(HYDRO_VISCOSITY.keys())
    if t <= keys[0]:
        return HYDRO_VISCOSITY[keys[0]]
    if t >= keys[-1]:
        return HYDRO_VISCOSITY[keys[-1]]
    lo = max(k for k in keys if k <= t)
    hi = min(k for k in keys if k >= t)
    if lo == hi:
        return HYDRO_VISCOSITY[lo]
    frac = (t - lo) / float(hi - lo)
    return HYDRO_VISCOSITY[lo] + frac * (HYDRO_VISCOSITY[hi] - HYDRO_VISCOSITY[lo])


def _round_or_none(v):
    if v is None:
        return None
    try:
        return round(float(v), 3)
    except Exception:
        return None


def _expansion_potential(value):
    if value is None:
        return ""
    if value <= 19:
        return "Very Low"
    if value <= 49:
        return "Low"
    if value <= 89:
        return "Medium"
    if value <= 130:
        return "High"
    return "Very High"


def _calc_saturation(dry_density_pcf, moisture_percent, gs):
    if dry_density_pcf is None or moisture_percent is None:
        return None
    gamma_w = 62.4
    if gs is None:
        gs = 2.65
    e = (gs * gamma_w / dry_density_pcf) - 1.0
    if e <= 0:
        return None
    sat = ((moisture_percent / 100.0) * gs / e) * 100.0
    if sat < 0:
        sat = 0.0
    if sat > 100:
        sat = 100.0
    return round(sat, 2)


def grain_size_section_flags(has_wash, has_sieve, has_hydrometer):
    include_wash = bool(has_wash or has_sieve or has_hydrometer)
    include_dry_sieve = bool(has_sieve)
    include_hydrometer = bool(has_sieve and has_hydrometer)
    return include_wash, include_dry_sieve, include_hydrometer


DRY_SIEVE_ORDER = [
    ("3_4in", '3/4"', 19.0),
    ("3_8in", '3/8"', 9.5),
    ("no4", "#4", 4.75),
    ("no10", "#10", 2.0),
    ("no16", "#16", 1.18),
    ("no40", "#40", 0.425),
    ("no50", "#50", 0.3),
    ("no100", "#100", 0.15),
    ("no200", "#200", 0.075),
]

HYDRO_TIMES = [1, 2, 5, 10, 15, 30, 60, 250, 1440]
HYDRO_VISCOSITY = {
    15: 1.16e-05,
    16: 1.133e-05,
    17: 1.104e-05,
    18: 1.076e-05,
    19: 1.05e-05,
    20: 1.025e-05,
    21: 1.0e-05,
    22: 9.76e-06,
    23: 9.53e-06,
    24: 9.31e-06,
    25: 9.1e-06,
    26: 8.9e-06,
    27: 8.7e-06,
    28: 8.51e-06,
    29: 8.32e-06,
    30: 8.14e-06,
}


def grain_size_default_payload():
    payload = {
        "wash_a_wet_tare": "",
        "wash_b_dry_tare": "",
        "wash_c_tare": "",
        "wash_e_moist_soil": "",
        "wash_h_dry40_tare": "",
        "wash_i_tare40": "",
        "wash_k_dry200_tare": "",
        "wash_l_tare200": "",
        "sieve_a_prewash_dry_weight": "",
        "sieve_uscs_class": "",
        "hydro_points": "",
        "hydro_gs": "2.67",
        "hydro_moist_sample_mass": "",
        "hydro_hydrostatic_moisture": "",
        "hydro_w": "",
        "hydro_pct_finer_no10": "1.0",
        "hydro_cal_t1": "20.5",
        "hydro_cal_c1": "0.0025",
        "hydro_cal_t2": "21.5",
        "hydro_cal_c2": "0.0020",
        "hydro_cal_t3": "19.6",
        "hydro_cal_c3": "0.0040",
        "hydro_cal_t4": "17.2",
        "hydro_cal_c4": "0.0045",
    }
    for key, _label, _size in DRY_SIEVE_ORDER:
        payload[f"sieve_pre_{key}"] = ""
        payload[f"sieve_post_{key}"] = ""
    for t in HYDRO_TIMES:
        payload[f"hydro_temp_{t}"] = ""
        payload[f"hydro_ra_{t}"] = ""
    return payload


def compute_grain_size(payload):
    vals = dict(grain_size_default_payload())
    vals.update(payload or {})
    a = _num(vals.get("wash_a_wet_tare"))
    b = _num(vals.get("wash_b_dry_tare"))
    c = _num(vals.get("wash_c_tare"))
    e = _num(vals.get("wash_e_moist_soil"))
    h = _num(vals.get("wash_h_dry40_tare"))
    i = _num(vals.get("wash_i_tare40"))
    k = _num(vals.get("wash_k_dry200_tare"))
    l = _num(vals.get("wash_l_tare200"))

    d = None
    if a is not None and b is not None and c is not None and abs((b - c)) > 1e-12:
        d = ((a - b) / (b - c)) * 100.0
    f = None
    if e is not None and d is not None:
        f = e / (1.0 + d / 100.0)
    m = None
    h_minus_i = _sub(h, i)
    k_minus_l = _sub(k, l)
    if h_minus_i is not None and k_minus_l is not None:
        m = h_minus_i + k_minus_l
    n = _sub(f, m)
    o = _pct(n, f)

    out = {
        "wash_d_moisture": d,
        "wash_f_dry_sample": f,
        "wash_m_weight": m,
        "wash_n_minus200": n,
        "wash_o_passing200": o,
    }
    # Dry sieve section from template style:
    # pre-test sieve weight + post-test (soil+sieve) -> retained per sieve.
    base = _num(vals.get("sieve_a_prewash_dry_weight"))
    if base is None:
        base = f
    out["sieve_a_prewash_dry_weight"] = base
    cum = 0.0
    for key, _label, _size in DRY_SIEVE_ORDER:
        pre = _num(vals.get(f"sieve_pre_{key}"))
        post = _num(vals.get(f"sieve_post_{key}"))
        retained = None
        if pre is not None and post is not None:
            retained = post - pre
            cum += retained
            out[f"sieve_ret_{key}"] = retained
            out[f"sieve_cum_{key}"] = cum
            if base is not None and abs(base) > 1e-12:
                out[f"sieve_pct_ret_{key}"] = (retained / base) * 100.0
                out[f"sieve_pct_pass_{key}"] = max(0.0, min(100.0, 100.0 - (cum / base) * 100.0))
            else:
                out[f"sieve_pct_ret_{key}"] = None
                out[f"sieve_pct_pass_{key}"] = None
        else:
            out[f"sieve_ret_{key}"] = None
            out[f"sieve_cum_{key}"] = None
            out[f"sieve_pct_ret_{key}"] = None
            out[f"sieve_pct_pass_{key}"] = None
    # If wash-derived passing is available, use it for #200.
    if out.get("wash_o_passing200") is not None:
        out["sieve_pct_pass_no200"] = out["wash_o_passing200"]

    # Hydrometer section (template equations)
    gs = _num(vals.get("hydro_gs"))
    if gs is None:
        gs = 2.67
    moist_mass = _num(vals.get("hydro_moist_sample_mass"))
    hydro_moist_pct = _num(vals.get("hydro_hydrostatic_moisture"))
    w_dry_input = _num(vals.get("hydro_w"))
    w_dry_from_moist = None
    if moist_mass is not None and hydro_moist_pct is not None:
        w_dry_from_moist = moist_mass / (1.0 + hydro_moist_pct / 100.0)
    w_dry = w_dry_input
    if w_dry is None:
        w_dry = w_dry_from_moist
    if w_dry is None:
        w_dry = f
    out["hydro_w_dry_used"] = w_dry
    out["hydro_w_dry_from_moist"] = w_dry_from_moist
    pct_finer_no10 = _num(vals.get("hydro_pct_finer_no10"))
    if pct_finer_no10 is None:
        pct_finer_no10 = 1.0
    cal_pts = []
    for idx in range(1, 5):
        ct = _num(vals.get(f"hydro_cal_t{idx}"))
        cc = _num(vals.get(f"hydro_cal_c{idx}"))
        if ct is not None and cc is not None:
            cal_pts.append((ct, cc))
    m_cal, b_cal = _linear_fit(cal_pts)
    out["hydro_cal_slope"] = m_cal
    out["hydro_cal_intercept"] = b_cal
    hydro_points = []
    for t in HYDRO_TIMES:
        ra = _num(vals.get(f"hydro_ra_{t}"))
        temp = _num(vals.get(f"hydro_temp_{t}"))
        rc = None
        if ra is not None:
            corr = 0.0
            if m_cal is not None and b_cal is not None and temp is not None:
                corr = m_cal * temp + b_cal
            rc = ra - corr
        n_vis = _viscosity_at_temp(temp)
        k_const = None
        if n_vis is not None and gs is not None and gs > 1.0:
            k_const = ((30.0 * n_vis) / (gs - 1.0)) ** 0.5
        l_eff = None
        if rc is not None:
            l_eff = (-264.516 * rc + 275.0) + 5.795
        d_mm = None
        if k_const is not None and l_eff is not None and t > 0 and l_eff > 0:
            d_mm = k_const * ((l_eff / float(t)) ** 0.5)
        partial = None
        if rc is not None and w_dry is not None and w_dry > 0 and gs > 1.0:
            partial = ((1000.0 / w_dry) * gs / (gs - 1.0)) * (rc - 1.0)
        total = None
        if partial is not None:
            total = partial * pct_finer_no10
        out[f"hydro_rc_{t}"] = rc
        out[f"hydro_d_{t}"] = d_mm
        out[f"hydro_partial_{t}"] = partial
        out[f"hydro_total_{t}"] = total
        if d_mm is not None and total is not None:
            hydro_points.append((d_mm, total))
    out["hydro_points_computed"] = hydro_points
    return out


def has_dry_sieve_entries(payload):
    vals = dict(payload or {})
    for key, _label, _size in DRY_SIEVE_ORDER:
        if _num(vals.get(f"sieve_pre_{key}")) is not None:
            return True
        if _num(vals.get(f"sieve_post_{key}")) is not None:
            return True
    return False


def parse_hydro_points(raw_text):
    text = (raw_text or "").strip()
    if not text:
        return []
    points = []
    for segment in text.split(";"):
        segment = segment.strip()
        if not segment:
            continue
        if ":" not in segment:
            continue
        x_text, y_text = segment.split(":", 1)
        x = _num(x_text)
        y = _num(y_text)
        if x is None or y is None:
            continue
        points.append((x, y))
    return points


def grain_curve_points(payload, include_dry_sieve, include_hydrometer, computed):
    pts = []
    vals = dict(payload or {})
    if include_dry_sieve:
        for key, _label, size_mm in DRY_SIEVE_ORDER:
            v = computed.get(f"sieve_pct_pass_{key}")
            if v is not None:
                pts.append((size_mm, v))
    if include_hydrometer:
        if computed.get("hydro_points_computed"):
            pts.extend(computed.get("hydro_points_computed"))
        else:
            pts.extend(parse_hydro_points(vals.get("hydro_points")))
    pts = [p for p in pts if p[0] > 0 and p[1] is not None]
    pts.sort(key=lambda p: p[0], reverse=True)
    return pts


def export_grain_size_pdf(path, project, sample_label, payload, computed, include_wash, include_dry_sieve, include_hydrometer):
    if canvas is None:
        raise RuntimeError("PDF export requires reportlab. Please install dependencies.")
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(path, pagesize=letter)
    w, h = letter
    y = h - 28
    c.setFont("Helvetica-Bold", 12)
    c.drawString(34, y, "GRAIN SIZE ANALYSIS WORKSHEET")
    y -= 16
    c.setFont("Helvetica", 9)
    c.drawString(34, y, f"Project: {project.get('job_name', '')}")
    y -= 14
    c.drawString(34, y, f"File Number: {project.get('file_number', '')}")
    y -= 14
    c.drawString(34, y, f"Sample: {sample_label}")
    y -= 18

    vals = dict(grain_size_default_payload())
    vals.update(payload or {})
    has_dry = has_dry_sieve_entries(vals)
    show_dry = bool(include_dry_sieve and has_dry)
    show_wash = bool(include_wash and not show_dry)

    if show_wash:
        c.setFont("Helvetica-Bold", 9)
        c.drawString(34, y, "-200 Wash Section")
        y -= 12
        rows = [
            ("D Moisture Content (%)", computed.get("wash_d_moisture")),
            ("F Dry Sample Weight (g)", computed.get("wash_f_dry_sample")),
            ("M Combined Retained (g)", computed.get("wash_m_weight")),
            ("N -200 Sieve Weight (g)", computed.get("wash_n_minus200")),
            ("O Percent Passing #200 (%)", computed.get("wash_o_passing200")),
        ]
        y = _draw_compact_rows(c, 34, y, rows)
        y -= 18

    if show_dry:
        c.setFont("Helvetica-Bold", 9)
        c.drawString(34, y, "Dry Sieve Section")
        y -= 12
        y = _draw_dry_sieve_table(c, 34, y, vals, computed)
        y -= 14
        c.setFont("Helvetica", 8)
        c.drawString(36, y, f"USCS Classification: {(vals.get('sieve_uscs_class') or '').strip()}")
        y -= 12

    if include_hydrometer:
        c.setFont("Helvetica-Bold", 9)
        c.drawString(34, y, "Hydrometer Section")
        y -= 12
        c.setFont("Helvetica", 8)
        c.drawString(
            36,
            y,
            "Rc = Ra - (m*T + b),  L = (-264.516*Rc + 275) + 5.795,  D = K*sqrt(L/t),  K=sqrt((30*n)/(Gs-1))",
        )
        y -= 11
        c.drawString(36, y, "Partial %Finer=((1000/W)*Gs/(Gs-1))*(Rc-1), Total %Finer=Partial*%Finer<#10")
        y -= 13
        slope = computed.get("hydro_cal_slope")
        intercept = computed.get("hydro_cal_intercept")
        c.drawString(36, y, f"Calibration (corrected hydrometer): m={_fmt(slope)}, b={_fmt(intercept)}")
        y -= 12
        w_used = computed.get("hydro_w_dry_used")
        w_from_m = computed.get("hydro_w_dry_from_moist")
        c.drawString(36, y, f"W used (oven-dry g): {_fmt(w_used)}   W from moist+hydrostatic moisture: {_fmt(w_from_m)}")
        y -= 12
        rows = []
        for t in HYDRO_TIMES:
            rc = computed.get(f"hydro_rc_{t}")
            dmm = computed.get(f"hydro_d_{t}")
            tf = computed.get(f"hydro_total_{t}")
            if rc is None and dmm is None and tf is None:
                continue
            rows.append((f"t={t} min", f"Rc={_fmt(rc)}  D={_fmt(dmm)} mm  %Finer={_fmt(tf)}"))
        if rows:
            y = _draw_compact_note_rows(c, 36, y, rows)
            y -= 8

    points = grain_curve_points(vals, show_dry, include_hydrometer, computed)
    if len(points) >= 2:
        _draw_grain_chart(c, 34, 42, w - 68, min(220, max(140, y - 52)), points)

    c.showPage()
    c.save()


def export_grouped_results_pdf(path, project, test_name, rows):
    if canvas is None:
        raise RuntimeError("PDF export requires reportlab. Please install dependencies.")
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(path, pagesize=letter)
    w, h = letter
    y = h - 36
    c.setFont("Helvetica-Bold", 12)
    c.drawString(34, y, f"{test_name} - Project Worksheet")
    y -= 14
    c.setFont("Helvetica", 9)
    astm = _astm_for_grouped_test(test_name)
    if astm:
        c.drawString(34, y, f"ASTM: {astm}")
        y -= 12
    c.drawString(34, y, f"Project: {project.get('job_name', '')}")
    y -= 12
    c.drawString(34, y, f"File Number: {project.get('file_number', '')}")
    y -= 16

    headers, extractors = _grouped_columns_for_test(test_name)
    col_w = [120, 90] + [90] * len(headers)
    left = 34
    row_h = 16

    c.setFont("Helvetica-Bold", 8.5)
    _draw_group_row(c, left, y, row_h, ["Sample", "Depth"] + headers, col_w)
    y -= row_h

    c.setFont("Helvetica", 8.5)
    for r in rows:
        vals = [
            r.get("sample_name", ""),
            r.get("depth_raw") or "",
        ]
        vals.extend([fn(r) for fn in extractors])
        _draw_group_row(c, left, y, row_h, vals, col_w)
        y -= row_h
        if y < 50:
            c.showPage()
            y = h - 36
            c.setFont("Helvetica-Bold", 12)
            c.drawString(34, y, f"{test_name} - Project Worksheet (cont.)")
            y -= 16
            c.setFont("Helvetica-Bold", 8.5)
            _draw_group_row(c, left, y, row_h, ["Sample", "Depth"] + headers, col_w)
            y -= row_h
            c.setFont("Helvetica", 8.5)

    c.showPage()
    c.save()


def _draw_group_row(c, left, y, row_h, values, widths):
    x = left
    for val, w in zip(values, widths):
        c.rect(x, y - row_h, w, row_h)
        c.drawCentredString(x + w / 2, y - 11, _fmt(val))
        x += w


def _grouped_columns_for_test(test_name):
    if test_name in ("Field Density/Moisture",):
        headers = ["Dry Density", "Moisture Content", "Saturation"]

        def v1(r):
            return r.get("result_value")

        def v2(r):
            return r.get("result_value2")

        def v3(r):
            return r.get("result_value3")

        return headers, [v1, v2, v3]
    if test_name == "Expansion Index":
        headers = ["Expansion Index", "Potential"]

        def v1(r):
            return r.get("result_value")

        def v2(r):
            return r.get("result_unit2") or r.get("result_value2")

        return headers, [v1, v2]
    if test_name == "Sand Cone":
        headers = ["Dry Density", "Moisture Content"]

        def v1(r):
            return r.get("result_value")

        def v2(r):
            return r.get("result_value2")

        return headers, [v1, v2]
    if test_name == "Moisture Content":
        headers = ["Moisture Content"]

        def v1(r):
            return r.get("result_value")

        return headers, [v1]
    return ["Value"], [lambda r: r.get("result_value")]


def _astm_for_grouped_test(test_name):
    mapping = {
        "Field Density/Moisture": "ASTM D 2937",
        "Expansion Index": "ASTM D 4829",
        "Sand Cone": "ASTM D 1556",
        "Moisture Content": "ASTM D 2216",
    }
    return mapping.get(test_name, "")


def _draw_compact_rows(c, left, y_top, rows):
    row_h = 14
    label_w = 250
    val_w = 120
    y = y_top
    c.setFont("Helvetica", 8)
    for label, value in rows:
        c.rect(left, y - row_h, label_w, row_h)
        c.rect(left + label_w, y - row_h, val_w, row_h)
        c.drawString(left + 3, y - 10, str(label))
        c.drawRightString(left + label_w + val_w - 4, y - 10, _fmt(value))
        y -= row_h
    return y


def _draw_dry_sieve_table(c, left, y_top, payload, computed):
    headers = ["Sieve", "mm", "Pre Wt", "Post Wt", "Retained", "Cum Ret", "% Ret", "% Pass"]
    col_w = [56, 42, 64, 64, 64, 64, 54, 54]
    row_h = 14
    x = left
    c.setFont("Helvetica-Bold", 7.6)
    for hdr, w in zip(headers, col_w):
        c.rect(x, y_top - row_h, w, row_h)
        c.drawCentredString(x + w / 2, y_top - 10, hdr)
        x += w
    y = y_top - row_h
    c.setFont("Helvetica", 7.5)
    for key, label, size in DRY_SIEVE_ORDER:
        y -= row_h
        row_vals = [
            label,
            _fmt(size),
            _fmt(payload.get(f"sieve_pre_{key}")),
            _fmt(payload.get(f"sieve_post_{key}")),
            _fmt(computed.get(f"sieve_ret_{key}")),
            _fmt(computed.get(f"sieve_cum_{key}")),
            _fmt(computed.get(f"sieve_pct_ret_{key}")),
            _fmt(computed.get(f"sieve_pct_pass_{key}")),
        ]
        x = left
        for val, w in zip(row_vals, col_w):
            c.rect(x, y, w, row_h)
            c.drawCentredString(x + w / 2, y + 4, str(val))
            x += w
    return y


def _draw_compact_note_rows(c, left, y_top, rows):
    row_h = 12
    w1 = 76
    w2 = 360
    y = y_top
    c.setFont("Helvetica", 7.6)
    for k, v in rows:
        c.rect(left, y - row_h, w1, row_h)
        c.rect(left + w1, y - row_h, w2, row_h)
        c.drawString(left + 3, y - 9, str(k))
        c.drawString(left + w1 + 3, y - 9, str(v))
        y -= row_h
    return y


def _draw_grain_chart(c, left, bottom, width, height, points):
    import math

    c.rect(left, bottom, width, height)
    minx = min(p[0] for p in points)
    maxx = max(p[0] for p in points)
    minx = max(minx, 0.001)
    maxx = max(maxx, minx * 10)
    miny = 0.0
    maxy = 100.0
    log_min = math.log10(minx)
    log_max = math.log10(maxx)
    if abs(log_max - log_min) < 1e-9:
        return

    def pxy(x, y):
        lx = math.log10(x)
        # Flip x-axis so larger sizes are on the left.
        px = left + ((log_max - lx) / (log_max - log_min)) * width
        py = bottom + ((y - miny) / (maxy - miny)) * height
        return px, py

    c.setFont("Helvetica", 7)
    c.setStrokeColorRGB(0.85, 0.88, 0.92)
    for yp in range(0, 101, 10):
        _x, py = pxy(minx, yp)
        c.line(left, py, left + width, py)
        c.setFillColorRGB(0.2, 0.2, 0.2)
        c.drawString(left - 18, py - 2, f"{yp}")
    tick_sizes = [100, 50, 25, 10, 4.75, 2.0, 0.85, 0.425, 0.25, 0.15, 0.075, 0.02, 0.005]
    for s in tick_sizes:
        if s < minx or s > maxx:
            continue
        px, _y = pxy(s, miny)
        c.line(px, bottom, px, bottom + height)
        c.drawCentredString(px, bottom - 9, f"{s:g}")

    c.setStrokeColorRGB(0.1, 0.35, 0.65)
    c.setFillColorRGB(0.1, 0.35, 0.65)
    ordered = sorted(points, key=lambda p: p[0], reverse=True)
    last = None
    for x, y in ordered:
        px, py = pxy(max(x, minx), max(miny, min(maxy, y)))
        c.circle(px, py, 1.7, stroke=1, fill=1)
        if last:
            c.line(last[0], last[1], px, py)
        last = (px, py)

    c.setFillColorRGB(0.7, 0.1, 0.1)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(left + width / 2, bottom - 21, "Particle Size (mm, log scale)")
    c.saveState()
    c.translate(left - 26, bottom + height / 2)
    c.rotate(90)
    c.drawCentredString(0, 0, "Percent Finer (%)")
    c.restoreState()
