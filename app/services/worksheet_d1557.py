from pathlib import Path

try:
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.pdfgen import canvas
except Exception:
    canvas = None


def compute_d1557_rows(raw_rows):
    rows = []
    for i, raw in enumerate(raw_rows, start=1):
        a = _to_float(raw.get("A"))
        b = _to_float(raw.get("B"))
        d = _to_float(raw.get("D"))
        e = _to_float(raw.get("E"))
        f = _to_float(raw.get("F"))

        c = None
        g = None
        h = None
        i_dry = None

        if a is not None and b is not None:
            c = a - b
        if d is not None and e is not None and f is not None:
            denom = e - f
            if abs(denom) > 1e-12:
                g = ((d - e) / denom) * 100.0
        if c is not None:
            h = c * 29.76 / 453.6
        if h is not None and g is not None:
            i_dry = h / (1.0 + g / 100.0)

        rows.append(
            {
                "test_no": i,
                "A": a,
                "B": b,
                "C": c,
                "D": d,
                "E": e,
                "F": f,
                "G": g,
                "H": h,
                "I": i_dry,
            }
        )
    return rows


def extract_points(rows):
    points = []
    for r in rows:
        if r["G"] is not None and r["I"] is not None:
            points.append((float(r["G"]), float(r["I"])))
    return points


def calculate_d1557(points):
    if len(points) < 3:
        max_pt = max(points, key=lambda p: p[1]) if points else (None, None)
        return {
            "max_dry_density": None if max_pt[1] is None else round(float(max_pt[1]), 2),
            "opt_moisture": None if max_pt[0] is None else round(float(max_pt[0]), 2),
            "method": "max-observed",
        }

    xs = [float(p[0]) for p in points]
    ys = [float(p[1]) for p in points]
    s1 = sum(xs)
    s2 = sum(x * x for x in xs)
    s3 = sum(x * x * x for x in xs)
    s4 = sum(x * x * x * x for x in xs)
    sy = sum(ys)
    sxy = sum(x * y for x, y in zip(xs, ys))
    sx2y = sum((x * x) * y for x, y in zip(xs, ys))
    n = len(xs)

    solved = _gauss3(
        [
            [s4, s3, s2, sx2y],
            [s3, s2, s1, sxy],
            [s2, s1, n, sy],
        ]
    )
    if not solved:
        max_pt = max(points, key=lambda p: p[1])
        return {
            "max_dry_density": round(float(max_pt[1]), 2),
            "opt_moisture": round(float(max_pt[0]), 2),
            "method": "max-observed",
        }

    a, b, c = solved
    if abs(a) < 1e-9:
        max_pt = max(points, key=lambda p: p[1])
        return {
            "max_dry_density": round(float(max_pt[1]), 2),
            "opt_moisture": round(float(max_pt[0]), 2),
            "method": "max-observed",
        }

    xv = -b / (2 * a)
    yv = a * xv * xv + b * xv + c
    lo, hi = min(xs), max(xs)
    if xv < lo or xv > hi or yv <= 0:
        max_pt = max(points, key=lambda p: p[1])
        return {
            "max_dry_density": round(float(max_pt[1]), 2),
            "opt_moisture": round(float(max_pt[0]), 2),
            "method": "max-observed",
        }
    return {
        "max_dry_density": round(float(yv), 2),
        "opt_moisture": round(float(xv), 2),
        "method": "quadratic-fit",
    }


def export_d1557_pdf(path, project, sample_label, rows, calc, g_values):
    if canvas is None:
        raise RuntimeError("PDF export requires reportlab. Please install dependencies.")
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(path, pagesize=landscape(letter))
    w, h = landscape(letter)

    top_margin = 24
    bottom_margin = 22
    section_gap = 20
    header_h = 84

    entered_rows = [r for r in rows if _has_raw_entry(r)]

    # Table metrics
    row_h = 14
    col_w = [28, 80, 64, 62, 80, 80, 70, 58, 58, 58]
    headers = [
        "Test",
        "A: Comp. Soil+Mold (gm)",
        "B: Mold (gm)",
        "C: Net Soil (gm)",
        "D: Wet Soil+Cont. (gm)",
        "E: Dry Soil+Cont. (gm)",
        "F: Container (gm)",
        "G Moisture (%)",
        "H Wet Dens. (pcf)",
        "I Dry Dens. (pcf)",
    ]
    table_w = sum(col_w)
    left = (w - table_w) / 2
    table_h = row_h * (1 + len(entered_rows))

    # Vertical layout: header block, table, graph with equal spacing.
    y_top = h - top_margin
    header_bottom = y_top - header_h
    table_top = header_bottom - section_gap
    table_bottom = table_top - table_h
    chart_top = table_bottom - section_gap
    chart_bottom = bottom_margin + 34  # reserve room for x-axis ticks/title
    chart_h = max(140, chart_top - chart_bottom - 10)
    if chart_bottom + chart_h > chart_top:
        chart_h = chart_top - chart_bottom
    top = table_top

    # Header block
    c.setFont("Helvetica-Bold", 13)
    c.drawString(30, y_top, "COMPACTION TEST - ASTM D1557")
    c.setFont("Helvetica", 9)
    c.drawString(30, y_top - 16, f"Project: {project.get('job_name', '')}")
    c.drawString(30, y_top - 29, f"File Number: {project.get('file_number', '')}")
    c.drawString(30, y_top - 42, f"Sample: {sample_label}")
    result_value_x = 220
    c.drawString(30, y_top - 55, "Maximum Dry Density (pcf):")
    c.setFillColorRGB(0.75, 0.1, 0.1)
    c.drawString(result_value_x, y_top - 55, f"{calc.get('max_dry_density', '')}")
    c.setFillColorRGB(0, 0, 0)
    c.drawString(30, y_top - 68, "Optimum Moisture Content (%):")
    c.setFillColorRGB(0.75, 0.1, 0.1)
    c.drawString(result_value_x, y_top - 68, f"{calc.get('opt_moisture', '')}")
    c.setFillColorRGB(0, 0, 0)

    # Header
    x = left
    c.setFont("Helvetica-Bold", 6.6)
    for j, head in enumerate(headers):
        c.rect(x, top - row_h, col_w[j], row_h)
        c.drawCentredString(x + col_w[j] / 2, top - 11, head)
        x += col_w[j]

    # Data rows: only include tests with entered raw data.
    c.setFont("Helvetica", 8)
    y = top - row_h
    for display_idx, r in enumerate(entered_rows, start=1):
        y -= row_h
        x = left
        vals = [
            str(display_idx),
            _fmt(r["A"]),
            _fmt(r["B"]),
            _fmt(r["C"]),
            _fmt(r["D"]),
            _fmt(r["E"]),
            _fmt(r["F"]),
            _fmt(r["G"]),
            _fmt(r["H"]),
            _fmt(r["I"]),
        ]
        for j, v in enumerate(vals):
            c.rect(x, y, col_w[j], row_h)
            c.drawCentredString(x + col_w[j] / 2, y + 5, str(v))
            x += col_w[j]

    # Chart with compaction points + zero-air-void lines, centered below table.
    points = extract_points(entered_rows)
    chart_w = 390
    chart_left = (w - chart_w) / 2
    c.rect(chart_left, chart_bottom, chart_w, chart_h)

    if points:
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        # User-requested axis ranges.
        max_data_x = max(xs)
        if max_data_x > 20:
            minx, maxx = 10.0, 30.0
        else:
            minx, maxx = 0.0, 20.0
        miny, maxy = 95.0, 155.0

        def pxy(xv, yv):
            px = chart_left + ((xv - minx) / (maxx - minx)) * chart_w
            py = chart_bottom + ((yv - miny) / (maxy - miny)) * chart_h
            return px, py

        # Axes ticks/grid for readability and proportional scaling.
        c.setStrokeColorRGB(0.82, 0.86, 0.9)
        for xv in _nice_ticks(minx, maxx, 6):
            px, _ = pxy(xv, miny)
            c.line(px, chart_bottom, px, chart_bottom + chart_h)
            c.setFillColorRGB(0.2, 0.2, 0.2)
            c.setFont("Helvetica", 7)
            c.drawCentredString(px, chart_bottom - 10, f"{xv:g}")
        for yv in _nice_ticks(miny, maxy, 6):
            _, py = pxy(minx, yv)
            c.line(chart_left, py, chart_left + chart_w, py)
            c.setFillColorRGB(0.2, 0.2, 0.2)
            c.setFont("Helvetica", 7)
            c.drawString(chart_left - 24, py - 2, f"{yv:g}")
        c.setFillColorRGB(0, 0, 0)

        # Zero-air-void reference lines
        for idx, g in enumerate(g_values):
            pts = []
            step = (maxx - minx) / 30.0
            xw = minx
            while xw <= maxx:
                den = 1.0 + (xw / 100.0) * g
                if den > 0:
                    yv = (62.43 * g) / den
                    if miny <= yv <= maxy:
                        pts.append(pxy(xw, yv))
                xw += step
            if len(pts) > 1:
                c.setDash(2, 2)
                c.setStrokeColorRGB(0.2, 0.45, 0.7 - idx * 0.1 if idx < 3 else 0.4)
                for k in range(1, len(pts)):
                    c.line(pts[k - 1][0], pts[k - 1][1], pts[k][0], pts[k][1])
                lx, ly = pts[-1]
                c.setFont("Helvetica", 7)
                c.drawString(lx - 25, ly + 2, f"G={g:g}")
                c.setDash()

        # Points
        c.setStrokeColorRGB(0.0, 0.2, 0.4)
        c.setFillColorRGB(0.0, 0.2, 0.4)
        sorted_points = sorted(points, key=lambda p: p[0])
        for xv, yv in sorted_points:
            px, py = pxy(xv, yv)
            c.circle(px, py, 2, stroke=1, fill=1)

        # Smooth interpreted proctor curve (quadratic fit) instead of point-to-point lines.
        fit = _fit_quadratic(points)
        if fit is not None:
            a, b, cc = fit
            c.setStrokeColorRGB(0.05, 0.35, 0.6)
            last = None
            step = (maxx - minx) / 80.0
            xw = minx
            while xw <= maxx:
                yfit = a * xw * xw + b * xw + cc
                if miny <= yfit <= maxy:
                    px, py = pxy(xw, yfit)
                    if last is not None:
                        c.line(last[0], last[1], px, py)
                    last = (px, py)
                xw += step

    # Axis titles in red and with more spacing from frame.
    c.setFillColorRGB(0.75, 0.1, 0.1)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(chart_left + chart_w / 2, chart_bottom - 22, "Moisture Content (%)")
    c.saveState()
    c.translate(chart_left - 34, chart_bottom + chart_h / 2)
    c.rotate(90)
    c.drawCentredString(0, 0, "Dry Density (pcf)")
    c.restoreState()
    c.setFillColorRGB(0, 0, 0)

    c.showPage()
    c.save()


def _to_float(value):
    if value is None:
        return None
    if isinstance(value, str):
        v = value.strip()
        if not v:
            return None
        return float(v)
    return float(value)


def _fmt(value):
    if value is None:
        return ""
    try:
        return f"{float(value):.2f}"
    except Exception:
        return str(value)


def _has_raw_entry(row):
    return any(row.get(k) is not None for k in ("A", "B", "D", "E", "F"))


def _gauss3(A):
    m = [row[:] for row in A]
    for col in range(3):
        pivot = max(range(col, 3), key=lambda r: abs(m[r][col]))
        if abs(m[pivot][col]) < 1e-12:
            return None
        if pivot != col:
            m[col], m[pivot] = m[pivot], m[col]
        div = m[col][col]
        for k in range(col, 4):
            m[col][k] /= div
        for r in range(3):
            if r == col:
                continue
            factor = m[r][col]
            for k in range(col, 4):
                m[r][k] -= factor * m[col][k]
    return [m[0][3], m[1][3], m[2][3]]


def _fit_quadratic(points):
    if len(points) < 3:
        return None
    xs = [float(p[0]) for p in points]
    ys = [float(p[1]) for p in points]
    s1 = sum(xs)
    s2 = sum(x * x for x in xs)
    s3 = sum(x * x * x for x in xs)
    s4 = sum(x * x * x * x for x in xs)
    sy = sum(ys)
    sxy = sum(x * y for x, y in zip(xs, ys))
    sx2y = sum((x * x) * y for x, y in zip(xs, ys))
    return _gauss3(
        [
            [s4, s3, s2, sx2y],
            [s3, s2, s1, sxy],
            [s2, s1, len(xs), sy],
        ]
    )


def _nice_ticks(minv, maxv, count):
    if maxv <= minv:
        return [minv]
    span = maxv - minv
    raw = span / max(1, (count - 1))
    mag = 10 ** int(__import__("math").floor(__import__("math").log10(abs(raw)))) if raw != 0 else 1
    norm = raw / mag
    if norm < 1.5:
        step = 1 * mag
    elif norm < 3:
        step = 2 * mag
    elif norm < 7:
        step = 5 * mag
    else:
        step = 10 * mag
    start = __import__("math").floor(minv / step) * step
    end = __import__("math").ceil(maxv / step) * step
    vals = []
    v = start
    n = 0
    while v <= end + 1e-9 and n < 100:
        vals.append(v)
        v += step
        n += 1
    return [x for x in vals if minv - 1e-9 <= x <= maxv + 1e-9]
