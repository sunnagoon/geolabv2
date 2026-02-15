import math
from pathlib import Path

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
except Exception:
    canvas = None


def default_payload():
    return {
        "project_title": "",
        "project_engineer": "",
        "project_number": "",
        "project_date": "",
        "ll": "",
        "pl": "",
        "passing_200": "",
        "finer_2um": "",
        "dry_density": "",
        "fabric_factor": "1.0",
        "ko_drying": "0.33",
        "ko_wetting": "0.67",
        "layer_thickness": "",
        "layer_description": "",
        "suction_initial_surface": "2.9",
        "suction_final_surface": "4.5",
        "constant_suction": "4.1",
        "depth_constant_suction": "6.9",
        "vertical_barrier_depth": "0.0",
        "horizontal_barrier_length": "0.0",
        "moisture_index": "-40",
        "em_distance": "",
    }


def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    text = str(v).strip()
    if not text:
        return None
    try:
        return float(text)
    except Exception:
        return None


def _clamp(v, lo, hi):
    return max(lo, min(hi, v))


def _round(v, n=3):
    if v is None:
        return None
    return round(float(v), n)


def _fmt(v, n=3):
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    return f"{float(v):.{n}f}".rstrip("0").rstrip(".")


def compute_pti(payload):
    vals = dict(default_payload())
    vals.update(payload or {})

    ll = _num(vals.get("ll"))
    pl = _num(vals.get("pl"))
    p200 = _num(vals.get("passing_200"))
    finer2 = _num(vals.get("finer_2um"))
    dry_density = _num(vals.get("dry_density"))
    fabric = _num(vals.get("fabric_factor")) or 1.0
    ko_dry = _num(vals.get("ko_drying")) or 0.33
    ko_wet = _num(vals.get("ko_wetting")) or 0.67
    depth = _num(vals.get("layer_thickness")) or 0.0
    suct_wet = _num(vals.get("suction_initial_surface")) or 0.0
    suct_dry = _num(vals.get("suction_final_surface")) or 0.0
    suct_const = _num(vals.get("constant_suction")) or 0.0
    depth_const = _num(vals.get("depth_constant_suction")) or depth
    moisture_index = _num(vals.get("moisture_index")) or 0.0

    pi = None
    if ll is not None and pl is not None:
        pi = ll - pl

    fine_clay_corr = None
    if finer2 is not None:
        fine_clay_corr = _clamp(finer2 / 95.0, 0.30, 1.20)
    coarse_corr = 1.0
    gamma0 = None
    if pi is not None and finer2 is not None and ll is not None:
        gamma0 = _clamp(0.02 + 0.0010 * pi + 0.0008 * finer2 + 0.0002 * (ll - 30.0), 0.02, 0.25)

    gamma_h_mean = None
    if gamma0 is not None and fine_clay_corr is not None:
        gamma_h_mean = gamma0 * fine_clay_corr * coarse_corr
    gamma_h_shrink = gamma_h_mean * (0.90 + 0.10 * ko_dry) if gamma_h_mean is not None else None
    gamma_h_swell = gamma_h_mean * (1.00 + 0.10 * ko_wet) if gamma_h_mean is not None else None

    alpha_mean = None
    if pi is not None and finer2 is not None:
        alpha_mean = 0.001 + 0.00004 * pi + 0.000015 * finer2
    alpha_shrink = alpha_mean * (0.98 + 0.06 * ko_dry) if alpha_mean is not None else None
    alpha_swell = alpha_mean * (0.98 + 0.06 * ko_wet) if alpha_mean is not None else None
    s_term = -2500.0 * alpha_mean if alpha_mean is not None else None
    p_term = gamma_h_mean * alpha_mean * 3.2 if gamma_h_mean is not None and alpha_mean is not None else None
    koho = p_term * (0.75 + 0.25 * ko_wet) if p_term is not None else None

    delta_suction_surface = suct_dry - suct_wet
    active_depth = min(depth, depth_const) if depth_const > 0 else depth
    if active_depth < 0:
        active_depth = 0.0

    ym_center_in = None
    if gamma_h_shrink is not None and active_depth > 0:
        ym_center_in = -gamma_h_shrink * delta_suction_surface * active_depth * fabric * 2.45

    em_user = _num(vals.get("em_distance"))
    if em_user is not None and em_user > 0:
        em_center_ft = em_user
        em_method = "User Input"
    else:
        em_center_ft = _clamp(0.85 * active_depth + 0.03 * abs(moisture_index), 4.0, 30.0)
        em_method = "Modified PTI"

    distances = []
    ym_table = []
    steps = 10
    if em_center_ft > 0:
        for i in range(steps + 1):
            x = em_center_ft * i / steps
            distances.append(x)
            if ym_center_in is None:
                ym_table.append(None)
            else:
                ratio = max(0.0, 1.0 - (x / em_center_ft))
                ym_table.append(ym_center_in * (ratio ** 1.2))

    # Suction profiles (depth increasing downward) for plotting/reporting.
    profile_depth = max(depth, depth_const, 0.0)
    if profile_depth <= 0:
        profile_depth = 7.0
    prof_steps = 12
    suction_depth_ft = []
    suction_wet_pf = []
    suction_dry_pf = []
    suction_const_pf = []
    for i in range(prof_steps + 1):
        z = profile_depth * i / prof_steps
        if depth_const > 1e-9:
            r = _clamp(z / depth_const, 0.0, 1.0)
        else:
            r = 1.0
        wet = suct_wet + (suct_const - suct_wet) * r
        dry = suct_dry + (suct_const - suct_dry) * r
        suction_depth_ft.append(_round(z, 2))
        suction_wet_pf.append(_round(wet, 3))
        suction_dry_pf.append(_round(dry, 3))
        suction_const_pf.append(_round(suct_const, 3))

    out = {
        "pi": _round(pi),
        "gamma0_mean": _round(gamma0),
        "fine_clay_correction": _round(fine_clay_corr),
        "coarse_correction": _round(coarse_corr),
        "gamma_h_mean": _round(gamma_h_mean),
        "gamma_h_shrink": _round(gamma_h_shrink),
        "gamma_h_swell": _round(gamma_h_swell),
        "alpha_mean": _round(alpha_mean, 6),
        "alpha_shrink": _round(alpha_shrink, 6),
        "alpha_swell": _round(alpha_swell, 6),
        "s_term": _round(s_term),
        "p_term": _round(p_term, 6),
        "koho": _round(koho, 6),
        "delta_suction_surface": _round(delta_suction_surface),
        "active_depth_ft": _round(active_depth),
        "ym_center_shrink_in": _round(ym_center_in, 2),
        "ym_center_shrink_cm": _round((ym_center_in or 0.0) * 2.54, 2) if ym_center_in is not None else None,
        "em_center_ft": _round(em_center_ft, 2),
        "em_center_cm": _round(em_center_ft * 30.48, 2),
        "em_method": em_method,
        "distances_ft": [_round(d, 2) for d in distances],
        "distances_cm": [_round(d * 30.48, 0) for d in distances],
        "ym_profile_in": [_round(y, 2) if y is not None else None for y in ym_table],
        "ym_profile_cm": [_round((y or 0.0) * 2.54, 2) if y is not None else None for y in ym_table],
        "suction_wet_surface": _round(suct_wet, 2),
        "suction_dry_surface": _round(suct_dry, 2),
        "suction_constant": _round(suct_const, 2),
        "suction_depth_ft": suction_depth_ft,
        "suction_wet_pf": suction_wet_pf,
        "suction_dry_pf": suction_dry_pf,
        "suction_const_pf": suction_const_pf,
    }
    if dry_density is not None:
        out["dry_density_pcf"] = _round(dry_density, 2)
    if p200 is not None:
        out["passing_200_pct"] = _round(p200, 2)
    if finer2 is not None:
        out["finer_2um_pct"] = _round(finer2, 2)
    if ll is not None:
        out["ll"] = _round(ll, 2)
    if pl is not None:
        out["pl"] = _round(pl, 2)
    return out


def export_pti_pdf(path, project, payload, computed):
    if canvas is None:
        raise RuntimeError("PDF export requires reportlab. Please install dependencies.")
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(path, pagesize=letter)
    w, h = letter

    y = h - 34
    c.setFont("Helvetica-Bold", 13)
    c.drawString(34, y, "PTI-Style Slab-On-Grade Shrink/Swell Distortion Analysis")
    y -= 14
    c.setFont("Helvetica", 9)
    c.drawString(34, y, f"Project Title: {payload.get('project_title') or project.get('job_name', '')}")
    y -= 12
    c.drawString(34, y, f"Project Number: {payload.get('project_number') or project.get('file_number', '')}")
    y -= 12
    c.drawString(34, y, f"Project Engineer: {payload.get('project_engineer', '')}")
    y -= 12
    c.drawString(34, y, f"Project Date: {payload.get('project_date', '')}")
    y -= 18

    c.setFont("Helvetica-Bold", 10)
    c.drawString(34, y, "Shrink Calculation")
    y -= 12
    c.setFont("Helvetica", 9)
    c.drawString(
        36,
        y,
        f"Ym Center (Shrink) = {_fmt(computed.get('ym_center_shrink_in'), 2)} in "
        f"({_fmt(computed.get('ym_center_shrink_cm'), 2)} cm)",
    )
    y -= 12
    c.drawString(
        36,
        y,
        f"Em Center = {_fmt(computed.get('em_center_ft'), 2)} ft "
        f"({_fmt(computed.get('em_center_cm'), 2)} cm)  Method: {computed.get('em_method', '')}",
    )
    y -= 16

    c.setFont("Helvetica-Bold", 9)
    c.drawString(34, y, "Summary of Input Data - Soil Properties")
    y -= 12
    soil_rows = [
        ("Layer Thickness (ft)", payload.get("layer_thickness")),
        ("Layer Description", payload.get("layer_description")),
        ("Liquid Limit", payload.get("ll")),
        ("Plastic Limit", payload.get("pl")),
        ("Plasticity Index", computed.get("pi")),
        ("% Passing #200", payload.get("passing_200")),
        ("% Finer than 2 um", payload.get("finer_2um")),
        ("Dry Density (pcf)", payload.get("dry_density")),
        ("Fabric Factor", payload.get("fabric_factor")),
        ("Ko Drying", payload.get("ko_drying")),
        ("Ko Wetting", payload.get("ko_wetting")),
    ]
    y = _draw_rows(c, 34, y, soil_rows, width=270)

    y2 = h - 164
    c.setFont("Helvetica-Bold", 9)
    c.drawString(326, y2, "Summary of Input Data - Suction / Em")
    y2 -= 12
    suct_rows = [
        ("Wet Surface Suction (pF)", payload.get("suction_initial_surface")),
        ("Dry Surface Suction (pF)", payload.get("suction_final_surface")),
        ("Constant Suction (pF)", payload.get("constant_suction")),
        ("Depth to Constant Suction (ft)", payload.get("depth_constant_suction")),
        ("Vertical Barrier Depth (ft)", payload.get("vertical_barrier_depth")),
        ("Horizontal Barrier Length (ft)", payload.get("horizontal_barrier_length")),
        ("Thornthwaite Moisture Index", payload.get("moisture_index")),
        ("Em Distance (ft, input)", payload.get("em_distance")),
        ("Computed Em (ft)", computed.get("em_center_ft")),
        ("Active Depth (ft)", computed.get("active_depth_ft")),
    ]
    _draw_rows(c, 326, y2, suct_rows, width=250)

    y = min(y - 8, y2 - 8)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(34, y, "Derived Layer Geotechnical Properties")
    y -= 12
    derived_rows = [
        ("Gamma0 Mean", computed.get("gamma0_mean")),
        ("Fine Clay Correction", computed.get("fine_clay_correction")),
        ("Coarse Correction", computed.get("coarse_correction")),
        ("GammaH Mean", computed.get("gamma_h_mean")),
        ("GammaH Shrink", computed.get("gamma_h_shrink")),
        ("GammaH Swell", computed.get("gamma_h_swell")),
        ("Alpha Mean", computed.get("alpha_mean")),
        ("Alpha Shrink", computed.get("alpha_shrink")),
        ("Alpha Swell", computed.get("alpha_swell")),
        ("S", computed.get("s_term")),
        ("P", computed.get("p_term")),
        ("KoHo", computed.get("koho")),
    ]
    y = _draw_rows(c, 34, y, derived_rows, width=270)

    dist = computed.get("distances_ft") or []
    ym_in = computed.get("ym_profile_in") or []
    y -= 8
    c.setFont("Helvetica-Bold", 9)
    c.drawString(34, y, "Shrink at Distance X from Edge of Slab")
    y -= 12
    y = _draw_dist_table(c, 34, y, dist, ym_in, width=542)
    y -= 8
    _draw_profile_chart(c, 34, max(44, y - 140), 542, 130, dist, ym_in)

    c.showPage()
    c.setFont("Helvetica-Bold", 12)
    c.drawString(34, h - 34, "SUCTION PROFILES")
    c.setFont("Helvetica", 9)
    c.drawString(34, h - 48, f"Project: {payload.get('project_title') or project.get('job_name', '')}")
    c.drawString(34, h - 60, f"File Number: {payload.get('project_number') or project.get('file_number', '')}")
    depths = computed.get("suction_depth_ft") or []
    wet = computed.get("suction_wet_pf") or []
    dry = computed.get("suction_dry_pf") or []
    const = computed.get("suction_const_pf") or []
    _draw_suction_chart(c, 52, 120, w - 104, h - 220, depths, wet, dry, const)
    c.save()


def _draw_rows(c, left, y_top, rows, width=260):
    row_h = 14
    name_w = int(width * 0.66)
    val_w = width - name_w
    y = y_top
    c.setFont("Helvetica", 8)
    for name, value in rows:
        c.rect(left, y - row_h, name_w, row_h)
        c.rect(left + name_w, y - row_h, val_w, row_h)
        c.drawString(left + 3, y - 10, str(name))
        c.drawRightString(left + name_w + val_w - 3, y - 10, _fmt(value, 4))
        y -= row_h
    return y


def _draw_dist_table(c, left, y_top, dist_ft, ym_in, width):
    if not dist_ft or not ym_in:
        return y_top
    row_h = 14
    cols = min(len(dist_ft), len(ym_in))
    label_w = 96
    data_w = max(20, (width - label_w) / cols)
    y = y_top
    c.setFont("Helvetica", 7.4)

    c.rect(left, y - row_h, label_w, row_h)
    c.drawString(left + 3, y - 10, "Distance (ft)")
    for i in range(cols):
        x = left + label_w + i * data_w
        c.rect(x, y - row_h, data_w, row_h)
        c.drawCentredString(x + data_w / 2, y - 10, _fmt(dist_ft[i], 2))
    y -= row_h

    c.rect(left, y - row_h, label_w, row_h)
    c.drawString(left + 3, y - 10, "Shrink (in)")
    for i in range(cols):
        x = left + label_w + i * data_w
        c.rect(x, y - row_h, data_w, row_h)
        c.drawCentredString(x + data_w / 2, y - 10, _fmt(ym_in[i], 2))
    return y - row_h


def _draw_profile_chart(c, left, bottom, width, height, dist_ft, ym_in):
    if not dist_ft or not ym_in:
        return
    c.rect(left, bottom, width, height)
    x_min = min(dist_ft)
    x_max = max(dist_ft)
    if abs(x_max - x_min) <= 1e-9:
        return
    y_min = min(ym_in)
    y_max = max(ym_in)
    y_pad = max(0.2, 0.08 * max(abs(y_min), abs(y_max), 1.0))
    y_lo = y_min - y_pad
    y_hi = y_max + y_pad

    def pxy(xv, yv):
        px = left + ((xv - x_min) / (x_max - x_min)) * width
        py = bottom + ((yv - y_lo) / (y_hi - y_lo)) * height
        return px, py

    c.setStrokeColorRGB(0.85, 0.88, 0.92)
    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.setFont("Helvetica", 7)
    for i in range(6):
        yv = y_lo + (i / 5.0) * (y_hi - y_lo)
        _x, py = pxy(x_min, yv)
        c.line(left, py, left + width, py)
        c.drawString(left - 28, py - 2, _fmt(yv, 2))
    for xv in dist_ft:
        px, _y = pxy(xv, y_lo)
        c.line(px, bottom, px, bottom + height)
        c.drawCentredString(px, bottom - 10, _fmt(xv, 1))

    c.setStrokeColorRGB(0.10, 0.35, 0.65)
    c.setFillColorRGB(0.10, 0.35, 0.65)
    last = None
    for xv, yv in zip(dist_ft, ym_in):
        px, py = pxy(xv, yv)
        c.circle(px, py, 1.8, stroke=1, fill=1)
        if last is not None:
            c.line(last[0], last[1], px, py)
        last = (px, py)

    c.setFillColorRGB(0.7, 0.1, 0.1)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(left + width / 2, bottom - 20, "Distance from slab edge (ft)")
    c.saveState()
    c.translate(left - 34, bottom + height / 2)
    c.rotate(90)
    c.drawCentredString(0, 0, "Shrink (in)")
    c.restoreState()


def _draw_suction_chart(c, left, bottom, width, height, depth_ft, wet_pf, dry_pf, const_pf):
    if not depth_ft:
        return
    c.rect(left, bottom, width, height)
    z_min = min(depth_ft)
    z_max = max(depth_ft)
    if abs(z_max - z_min) <= 1e-9:
        return
    all_s = [v for v in (wet_pf + dry_pf + const_pf) if v is not None]
    if not all_s:
        return
    s_min = min(all_s)
    s_max = max(all_s)
    s_pad = max(0.25, 0.1 * (s_max - s_min if s_max > s_min else 1.0))
    x_lo = s_min - s_pad
    x_hi = s_max + s_pad

    def pxy(suction, depth):
        px = left + ((suction - x_lo) / (x_hi - x_lo)) * width
        # Depth positive downward.
        py = bottom + height - ((depth - z_min) / (z_max - z_min)) * height
        return px, py

    c.setStrokeColorRGB(0.85, 0.88, 0.92)
    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.setFont("Helvetica", 7)
    for i in range(8):
        sv = x_lo + (i / 7.0) * (x_hi - x_lo)
        px, _y = pxy(sv, z_min)
        c.line(px, bottom, px, bottom + height)
        c.drawCentredString(px, bottom - 10, _fmt(sv, 2))
    for i in range(7):
        zv = z_min + (i / 6.0) * (z_max - z_min)
        _x, py = pxy(x_lo, zv)
        c.line(left, py, left + width, py)
        c.drawString(left - 24, py - 2, _fmt(zv, 1))

    def draw_profile(vals, color):
        c.setStrokeColorRGB(*color)
        c.setFillColorRGB(*color)
        last = None
        for z, s in zip(depth_ft, vals):
            if s is None:
                continue
            px, py = pxy(s, z)
            c.circle(px, py, 1.5, stroke=1, fill=1)
            if last is not None:
                c.line(last[0], last[1], px, py)
            last = (px, py)

    draw_profile(wet_pf, (0.12, 0.45, 0.72))
    draw_profile(dry_pf, (0.74, 0.25, 0.22))
    draw_profile(const_pf, (0.20, 0.55, 0.28))

    c.setFont("Helvetica-Bold", 8)
    c.setFillColorRGB(0.7, 0.1, 0.1)
    c.drawCentredString(left + width / 2, bottom - 22, "Suction (pF)")
    c.saveState()
    c.translate(left - 30, bottom + height / 2)
    c.rotate(90)
    c.drawCentredString(0, 0, "Depth (ft)")
    c.restoreState()

    # Legend
    ly = bottom + height + 14
    c.setFont("Helvetica", 8)
    legend = [("Initial suction at edge", (0.12, 0.45, 0.72)), ("Final suction at edge", (0.74, 0.25, 0.22)), ("Constant suction", (0.20, 0.55, 0.28))]
    lx = left
    for text, color in legend:
        c.setStrokeColorRGB(*color)
        c.setFillColorRGB(*color)
        c.line(lx, ly, lx + 18, ly)
        c.circle(lx + 9, ly, 1.4, stroke=1, fill=1)
        c.setFillColorRGB(0.2, 0.2, 0.2)
        c.drawString(lx + 24, ly - 3, text)
        lx += 170
