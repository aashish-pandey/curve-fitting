"""
dosy_curve_fit.py
─────────────────
Fits mono/bi-exponential decay curves to DOSY NMR decomposition Excel sheets.

Usage
-----
    from dosy_curve_fit import fit_dosy
    fit_dosy("output")                        # scans output/ for *_component_Decomposition*.xlsx
    fit_dosy("output", x_col="B_value", y_col="I_norm", clipped_col="clipped")

For every sheet in every matching file the function:
  • reads  B_value (x),  I_norm (y),  clipped (yes/no)
  • fits 5 curves:
      [1] All data    → mono-exp          → D_all_mono
      [2] All data    → bi-exp            → D1_all_bi,  D2_all_bi
      [3] Clipped=Yes → mono-exp only     → D_yes_mono
      [4] Clipped=No  → mono-exp          → D_no_mono
      [5] Clipped=No  → bi-exp            → D1_no_bi,   D2_no_bi
  • appends fitted-value columns (G–K) in the same rows as the data
  • writes a results summary table below the data
  • embeds one openpyxl ScatterChart per fit (5 charts total) below the summary
  • embeds 4 OVERLAY ScatterCharts, each combining one fit from All (Mono/Bi),
    the single Yes-Mono fit, and one fit from No (Mono/Bi) — i.e. all 2×1×2
    combinations of subset-fit choices — below the individual charts
  • saves back to the same file

Dependencies: math, os, glob  (stdlib)  +  numpy  +  openpyxl
"""

import math, os, glob
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.axis import ChartLines


def fit_dosy(
    output_folder: str = "output",
    x_col:         str = "B_value",
    y_col:         str = "I_norm",
    clipped_col:   str = "clipped",
) -> None:

    # ── 1. Models ────────────────────────────────────────────────────────────

    def mono(x, A, D):
        return A * np.exp(np.clip(-D * x, -700, 700))

    def bi(x, A1, D1, A2, D2):
        return (A1 * np.exp(np.clip(-D1 * x, -700, 700))
              + A2 * np.exp(np.clip(-D2 * x, -700, 700)))

    # ── 2. Nelder-Mead (pure Python, no scipy) ───────────────────────────────

    def nelder_mead(f, p0, tol=1e-12, max_iter=20000):
        n = len(p0)
        s = [list(p0)]
        for i in range(n):
            v = list(p0); v[i] += 0.1 * abs(v[i]) if abs(v[i]) > 1e-10 else 1e-4
            s.append(v)
        sc = [f(v) for v in s]
        for _ in range(max_iter):
            o = sorted(range(n+1), key=lambda i: sc[i])
            s, sc = [s[i] for i in o], [sc[i] for i in o]
            if sc[-1] - sc[0] < tol: break
            c = [sum(s[i][j] for i in range(n)) / n for j in range(n)]
            xr = [c[j] + (c[j] - s[-1][j]) for j in range(n)]; fr = f(xr)
            if fr < sc[0]:
                xe = [c[j] + 2*(xr[j]-c[j]) for j in range(n)]; fe = f(xe)
                s[-1], sc[-1] = (xe,fe) if fe < fr else (xr,fr)
            elif fr < sc[-2]:
                s[-1], sc[-1] = xr, fr
            else:
                if fr < sc[-1]: s[-1], sc[-1] = xr, fr
                xc = [c[j] + 0.5*(s[-1][j]-c[j]) for j in range(n)]; fc_ = f(xc)
                if fc_ < sc[-1]: s[-1], sc[-1] = xc, fc_
                else:
                    for i in range(1, n+1):
                        s[i] = [s[0][j]+0.5*(s[i][j]-s[0][j]) for j in range(n)]
                        sc[i] = f(s[i])
        return np.array(s[0])

    def fit_curve(model, p0, bx, iy):
        """Run Nelder-Mead; return (params, fitted_y, R²) or None on failure."""
        if len(bx) < len(p0) + 1:
            return None
        def ssr(p):
            try:    return float(np.sum((iy - model(bx, *p))**2))
            except: return 1e30
        try:
            p = nelder_mead(ssr, p0)
            yp = model(bx, *p)
            ss_res = float(np.sum((iy - yp)**2))
            ss_tot = float(np.sum((iy - iy.mean())**2))
            r2 = 1 - ss_res/ss_tot if ss_tot > 0 else 0.0
            return p, yp, r2
        except:
            return None

    def warm_p0_mono(bx, iy):
        amp  = float(iy[np.argmin(bx)]) if len(iy) else 1.0
        amp  = amp if amp > 0 else 1.0
        span = float(bx.max() - bx.min()) if len(bx) > 1 else 1.0
        y0   = float(iy[np.argmin(bx)]); y0 = y0 if abs(y0) > 1e-30 else 1e-30
        yn   = float(iy[np.argmax(bx)]); yn = yn if abs(yn) > 1e-30 else 1e-30
        rate = abs(math.log(abs(yn/y0))) / span if span > 0 else 1e-9
        rate = rate if rate > 1e-30 else 1e-9
        return [amp, rate]

    def warm_p0_bi(bx, iy):
        amp, rate = warm_p0_mono(bx, iy)
        return [amp*0.6, rate, amp*0.4, rate*10]

    # ── 3. Excel style helpers ────────────────────────────────────────────────

    HDR_FILL  = PatternFill("solid", fgColor="1F4E79")
    SUB_FILL  = PatternFill("solid", fgColor="BDD7EE")
    ROW_FILL  = PatternFill("solid", fgColor="EBF3FB")
    W_FONT    = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    B_FONT    = Font(name="Arial", bold=True, color="1F4E79", size=9)
    N_FONT    = Font(name="Arial", size=9)
    CTR       = Alignment(horizontal="center", vertical="center")

    def sc(cell, val, fill=None, font=None, align=CTR, nf=None):
        cell.value = val
        if fill: cell.fill = fill
        cell.font = font or N_FONT
        cell.alignment = align
        if nf: cell.number_format = nf

    def fp(v):
        if not isinstance(v, float): return str(v)
        if math.isnan(v) or math.isinf(v): return str(v)
        return f"{v:.4e}" if (abs(v) < 1e-3 or abs(v) > 9999) else f"{v:.6f}"

    # ── 4. Chart builders ─────────────────────────────────────────────────────

    def make_scatter(ws, title, xref, yref_data, yref_fit, anchor):
        ch = ScatterChart(); ch.scatterStyle = "smoothMarker"
        ch.title = title; ch.width = 14; ch.height = 10
        ch.x_axis.title = x_col; ch.y_axis.title = y_col
        ch.x_axis.numFmt = "0.00E+00"; ch.style = 2
        ch.x_axis.delete = False; ch.y_axis.delete = False
        ch.x_axis.majorTickMark = "out"; ch.y_axis.majorTickMark = "out"
        ch.x_axis.majorGridlines = ChartLines()

        raw = Series(yref_data, xref, title="Data")
        raw.marker.symbol = "circle"; raw.marker.size = 4
        raw.marker.graphicalProperties.solidFill = "4472C4"
        raw.graphicalProperties.line.noFill = True
        ch.series.append(raw)

        if yref_fit:
            fit = Series(yref_fit, xref, title="Fit")
            fit.marker.symbol = "none"
            fit.graphicalProperties.line.solidFill = "FF0000"
            fit.graphicalProperties.line.width = 18000
            ch.series.append(fit)

        ws.add_chart(ch, anchor)

    def make_overlay(ws, title, xref, curve_defs, raw_defs, anchor):
        """
        curve_defs: list of (col_idx, label, hex_color) — one LINE per fit, no markers.
        raw_defs:   list of (col_idx, label, hex_color) — one MARKER-only series per
                    raw data subset (e.g. Yes/No), no connecting line. Columns are
                    expected to contain blanks for rows outside that subset, so the
                    markers only appear at the correct B_value positions.
        """
        ch = ScatterChart(); ch.scatterStyle = "lineMarker"
        ch.title = title; ch.width = 16; ch.height = 10
        ch.x_axis.title = x_col; ch.y_axis.title = y_col
        ch.x_axis.numFmt = "0.00E+00"; ch.style = 2
        ch.x_axis.delete = False; ch.y_axis.delete = False
        ch.x_axis.majorTickMark = "out"; ch.y_axis.majorTickMark = "out"
        ch.x_axis.majorGridlines = ChartLines()

        # raw data points first (drawn behind the fit lines)
        for col_idx, label, color in raw_defs:
            yref = Reference(ws, min_col=col_idx, min_row=xref.min_row, max_row=xref.max_row)
            s = Series(yref, xref, title=label)
            s.marker.symbol = "circle"; s.marker.size = 5
            s.marker.graphicalProperties.solidFill = color
            s.graphicalProperties.line.noFill = True
            ch.series.append(s)

        # fit lines on top
        for col_idx, label, color in curve_defs:
            yref = Reference(ws, min_col=col_idx, min_row=xref.min_row, max_row=xref.max_row)
            s = Series(yref, xref, title=label)
            s.marker.symbol = "none"
            s.graphicalProperties.line.solidFill = color
            s.graphicalProperties.line.width = 20000
            ch.series.append(s)

        ws.add_chart(ch, anchor)

    # ── 5. Per-sheet processing ───────────────────────────────────────────────

    def process_sheet(ws, fname):
        # Clear charts from any previous run of this script on this file — otherwise
        # re-running on an already-processed workbook stacks a new chart set on top
        # of the old one at nearly the same anchors, which looks like overlap.
        ws._charts = []

        # 5a. locate header row ───────────────────────────────────────────────
        xl = x_col.strip().lower()
        yl = y_col.strip().lower()
        cl = clipped_col.strip().lower()
        hrow = xi = yi = ci = None
        for row in ws.iter_rows(min_row=1, max_row=min(20, ws.max_row)):
            found = {}
            for cell in row:
                if cell.value is None: continue
                raw = cell.value
                if isinstance(raw, (list, tuple)): raw = raw[0] if raw else None
                if raw is None: continue
                label = str(raw).strip().lower()
                if label == xl: found["x"] = cell.column
                elif label == yl: found["y"] = cell.column
                elif label == cl: found["c"] = cell.column
            if len(found) == 3:
                hrow = row[0].row
                xi, yi, ci = found["x"], found["y"], found["c"]
                break
        if hrow is None:
            print(f"  [SKIP] {ws.title}: headers not found"); return

        # 5b. read data ────────────────────────────────────────────────────────
        rows_b, rows_i, rows_clip, row_nums = [], [], [], []
        for row in ws.iter_rows(min_row=hrow+1, max_row=ws.max_row):
            bv = row[xi-1].value; iv = row[yi-1].value; cv = row[ci-1].value
            if bv is None or iv is None: continue
            try:
                rows_b.append(float(bv)); rows_i.append(float(iv))
                if isinstance(cv, (list, tuple)): cv = cv[0] if cv else None
                rows_clip.append(str(cv).strip().lower() if cv is not None else "")
                row_nums.append(row[0].row)
            except (ValueError, TypeError): continue

        if len(rows_b) < 4:
            print(f"  [SKIP] {ws.title}: only {len(rows_b)} rows"); return

        b_all = np.array(rows_b); i_all = np.array(rows_i)
        clip  = np.array(rows_clip)

        mask_yes = clip == "yes"; mask_no = clip == "no"
        b_yes, i_yes = b_all[mask_yes], i_all[mask_yes]
        b_no,  i_no  = b_all[mask_no],  i_all[mask_no]

        # 5c. run fits ─────────────────────────────────────────────────────────
        fits = {
            "all_mono": fit_curve(mono, warm_p0_mono(b_all, i_all), b_all, i_all),
            "all_bi":   fit_curve(bi,   warm_p0_bi(b_all, i_all),   b_all, i_all),
            "yes_mono": fit_curve(mono, warm_p0_mono(b_yes, i_yes), b_yes, i_yes),
            "no_mono":  fit_curve(mono, warm_p0_mono(b_no,  i_no),  b_no,  i_no),
            "no_bi":    fit_curve(bi,   warm_p0_bi(b_no,  i_no),    b_no,  i_no),
        }

        # 5d. write fitted columns G–K ─────────────────────────────────────────
        # For subset fits (yes/no), params are found from the subset but the
        # equation is evaluated over ALL B_value points.
        FIT_HDRS   = ["Fit_All_Mono","Fit_All_Bi","Fit_Yes_Mono","Fit_No_Mono","Fit_No_Bi"]
        FIT_KEYS   = ["all_mono","all_bi","yes_mono","no_mono","no_bi"]
        FIT_MODELS = [mono, bi, mono, mono, bi]
        COL_OFF    = 7   # G

        # key -> column index, used later for overlay charts too
        key_to_col = {key: COL_OFF + k for k, key in enumerate(FIT_KEYS)}

        for k, (key, hdr, model) in enumerate(zip(FIT_KEYS, FIT_HDRS, FIT_MODELS)):
            col = COL_OFF + k
            hcell = ws.cell(row=hrow, column=col)
            sc(hcell, hdr, fill=SUB_FILL, font=B_FONT)
            ws.column_dimensions[get_column_letter(col)].width = 16

            if fits[key] is None: continue
            params = fits[key][0]
            # evaluate fitted equation over ALL B values regardless of subset
            yp_all = model(b_all, *params)
            for di, rn in enumerate(row_nums):
                c = ws.cell(row=rn, column=col)
                c.value = round(float(yp_all[di]), 8)
                c.number_format = "0.00000000"
                c.font = N_FONT; c.alignment = CTR

        # 5d-2. raw Yes/No value columns (for overlay scatter markers) ─────────
        # Rows aren't grouped by clip status, so each row gets its I_norm value
        # written only into the matching column; the other column stays blank.
        # A contiguous Reference over these columns then plots markers only at
        # the rows that actually belong to that subset.
        col_yes_raw = COL_OFF + 5   # L
        col_no_raw  = COL_OFF + 6   # M
        sc(ws.cell(row=hrow, column=col_yes_raw), "Yes_Raw_I", fill=SUB_FILL, font=B_FONT)
        sc(ws.cell(row=hrow, column=col_no_raw),  "No_Raw_I",  fill=SUB_FILL, font=B_FONT)
        ws.column_dimensions[get_column_letter(col_yes_raw)].width = 14
        ws.column_dimensions[get_column_letter(col_no_raw)].width = 14

        for di, rn in enumerate(row_nums):
            target_col = col_yes_raw if rows_clip[di] == "yes" else (
                         col_no_raw if rows_clip[di] == "no" else None)
            if target_col is None:
                continue
            c = ws.cell(row=rn, column=target_col)
            c.value = round(float(i_all[di]), 8)
            c.number_format = "0.00000000"
            c.font = N_FONT; c.alignment = CTR

        # 5e. summary table ────────────────────────────────────────────────────
        last_row = max(row_nums)
        sr = last_row + 3

        # header
        ws.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=6)
        sc(ws.cell(row=sr, column=1),
           f"Curve Fit Results — {ws.title}", fill=HDR_FILL, font=W_FONT)
        sr += 1

        for ci_, h in enumerate(["Fit","D Coeff(s) [m²/s]","A","R²","Notes"], 1):
            sc(ws.cell(row=sr, column=ci_), h, fill=SUB_FILL, font=B_FONT)
        sr += 1

        def d_val(key, idx): # extract D from params safely
            if fits[key] is None: return "failed"
            return float(fits[key][0][idx])

        def r2_val(key):
            if fits[key] is None: return "—"
            return round(fits[key][2], 6)

        summary_rows = [
            ("All → Mono",     d_val("all_mono",1),
             d_val("all_mono",0), r2_val("all_mono"), ""),
            ("All → Bi (D1)",  d_val("all_bi",1),
             d_val("all_bi",0),  r2_val("all_bi"),  "fast component"),
            ("All → Bi (D2)",  d_val("all_bi",3),
             d_val("all_bi",2),  r2_val("all_bi"),  "slow component"),
            ("Yes → Mono",     d_val("yes_mono",1),
             d_val("yes_mono",0), r2_val("yes_mono"), ""),
            ("No  → Mono",     d_val("no_mono",1),
             d_val("no_mono",0),  r2_val("no_mono"),  ""),
            ("No  → Bi (D1)",  d_val("no_bi",1),
             d_val("no_bi",0),   r2_val("no_bi"),   "fast component"),
            ("No  → Bi (D2)",  d_val("no_bi",3),
             d_val("no_bi",2),   r2_val("no_bi"),   "slow component"),
        ]

        for row_data in summary_rows:
            for ci_, val in enumerate(row_data, 1):
                cell = ws.cell(row=sr, column=ci_)
                if isinstance(val, float):
                    sc(cell, val, fill=ROW_FILL, nf="0.00000E+00")
                else:
                    sc(cell, val, fill=ROW_FILL)
            sr += 1

        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 18

        # 5f. individual charts (one per fit) ─────────────────────────────────
        dmin, dmax = min(row_nums), max(row_nums)
        xref_all = Reference(ws, min_col=xi, min_row=dmin, max_row=dmax)
        yref_raw = Reference(ws, min_col=yi, min_row=dmin, max_row=dmax)

        def col_ref(col_idx):
            return Reference(ws, min_col=col_idx, min_row=dmin, max_row=dmax)

        # Chart height is fixed at 10cm; default Excel row height (~15pt ≈ 0.53cm)
        # means ~19 rows are needed just to clear one chart's height. 18 rows was
        # too tight and caused consecutive charts to overlap — use 26 for a clean gap.
        CHART_ROW_SPACING = 26
        chart_anchor_row = sr + 2
        chart_configs = [
            ("All Data — Mono-Exp",   xref_all, yref_raw, col_ref(COL_OFF+0)),
            ("All Data — Bi-Exp",     xref_all, yref_raw, col_ref(COL_OFF+1)),
            ("Clipped=Yes — Mono",    xref_all, yref_raw, col_ref(COL_OFF+2)),
            ("Clipped=No  — Mono",    xref_all, yref_raw, col_ref(COL_OFF+3)),
            ("Clipped=No  — Bi-Exp",  xref_all, yref_raw, col_ref(COL_OFF+4)),
        ]
        for k, (title, xr, yr, fr) in enumerate(chart_configs):
            anchor = f"A{chart_anchor_row + k*CHART_ROW_SPACING}"
            make_scatter(ws, title, xr, yr, fr, anchor)

        # 5g. overlay charts — all 2×1×2 combinations of subset-fit choices ───
        # All has 2 options (Mono/Bi), Yes has 1 option (Mono only), No has 2
        # options (Mono/Bi)  →  2 * 1 * 2 = 4 overlay plots, 3 curves each.
        KEY_LABEL = {
            "all_mono": "All-Mono", "all_bi": "All-Bi",
            "yes_mono": "Yes-Mono",
            "no_mono":  "No-Mono",  "no_bi":  "No-Bi",
        }
        KEY_COLOR = {
            "all_mono": "4472C4",  # blue
            "all_bi":   "1F4E79",  # dark blue
            "yes_mono": "70AD47",  # green
            "no_mono":  "ED7D31",  # orange
            "no_bi":    "C00000",  # dark red
        }
        overlay_combos = [
            ("Overlay: All-Mono + Yes-Mono + No-Mono", ["all_mono", "yes_mono", "no_mono"]),
            ("Overlay: All-Bi + Yes-Mono + No-Mono",   ["all_bi",   "yes_mono", "no_mono"]),
            ("Overlay: All-Mono + Yes-Mono + No-Bi",   ["all_mono", "yes_mono", "no_bi"]),
            ("Overlay: All-Bi + Yes-Mono + No-Bi",     ["all_bi",   "yes_mono", "no_bi"]),
        ]

        # raw Yes/No data-point markers, shown on every overlay plot
        raw_defs = [
            (col_yes_raw, "Yes-Raw", "70AD47"),  # green
            (col_no_raw,  "No-Raw",  "ED7D31"),  # orange
        ]

        overlay_anchor_row = chart_anchor_row + len(chart_configs) * CHART_ROW_SPACING
        placed = 0
        for title, keys in overlay_combos:
            if any(fits[k_] is None for k_ in keys):
                print(f"    [SKIP OVERLAY] {title}: missing fit(s)")
                continue
            curve_defs = [(key_to_col[k_], KEY_LABEL[k_], KEY_COLOR[k_]) for k_ in keys]
            anchor = f"A{overlay_anchor_row + placed*CHART_ROW_SPACING}"
            make_overlay(ws, title, xref_all, curve_defs, raw_defs, anchor)
            placed += 1

        print(f"  [OK] {ws.title} | N={len(b_all)} yes={mask_yes.sum()} no={mask_no.sum()} | overlays={placed}")

    # ── 6. File discovery & main loop ─────────────────────────────────────────

    pattern  = os.path.join(output_folder, "**", "*_component_Decomposition*.xlsx")
    files    = list(set(
        glob.glob(pattern, recursive=True)
        + glob.glob(os.path.join(output_folder, "*_component_Decomposition*.xlsx"))
    ))

    if not files:
        print(f"No matching files found in '{output_folder}'"); return

    print(f"Found {len(files)} file(s).")
    for fp_ in sorted(files):
        print(f"\n[FILE] {os.path.basename(fp_)}")
        wb = openpyxl.load_workbook(fp_)
        for sname in wb.sheetnames:
            process_sheet(wb[sname], os.path.basename(fp_))
        wb.save(fp_)
        print(f"  Saved → {fp_}")


if __name__ == "__main__":
    import sys
    fit_dosy(sys.argv[1] if len(sys.argv) > 1 else "output")
