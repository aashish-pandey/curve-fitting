import os
import numpy as np
from numpy.linalg import eig, svd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import ScatterChart, Reference, Series
import MnovaNMR

# --- CONFIG ------------------------------------------------------------------

OUTPUT_FOLDER = r"C:\Users\you\Documents\NMR_Results"

PEAKS = [
    ("peak_1", 3.70, 3.60),
    ("peak_2", 1.30, 1.10),
]

N_COMPONENTS = 2

# -----------------------------------------------------------------------------


def get_b_values(nmr_item):
    spectra  = nmr_item.spectra()
    b_values = np.zeros(len(spectra))
    # --- YOUR CODE HERE ---
    return b_values


def build_spectral_matrix(nmr_item, ppm_left, ppm_right):
    spectra  = nmr_item.spectra()
    ess      = spectra[0].coords[0]

    pt_left  = int(round(ess.ppmToPt(ppm_left)))
    pt_right = int(round(ess.ppmToPt(ppm_right)))
    pt_start = min(pt_left, pt_right)
    pt_end   = max(pt_left, pt_right)
    pts      = list(range(pt_start, pt_end + 1))

    ppm_axis = np.array([ess.ptToPpm(pt) for pt in pts])

    Y = np.zeros((len(spectra), len(pts)), dtype=float)
    for i, spc in enumerate(spectra):
        for j, pt in enumerate(pts):
            Y[i, j] = spc.reDataAt(pt)

    return Y, ppm_axis


def decra(Y, n_components=2):
    Y  = np.asarray(Y, dtype=float)
    K  = n_components

    Y1 = Y[:-1, :]
    Y2 = Y[1:,  :]

    U, s, Vt = svd(Y1, full_matrices=False)
    U  = U[:,  :K]
    s  = s[    :K]
    Vt = Vt[:K, :]

    Phi_red      = U.T @ Y2 @ Vt.T @ np.diag(1.0 / s)
    eigenvalues, eigenvectors = eig(Phi_red)
    eigenvalues  = eigenvalues.real
    eigenvectors = eigenvectors.real

    order        = np.argsort(eigenvalues)[::-1]
    eigenvalues  = eigenvalues[order]
    eigenvectors = eigenvectors[:, order]

    C1 = U @ np.diag(s) @ eigenvectors
    C  = np.vstack([C1, C1[-1, :] * eigenvalues])
    C  = C / C[0, :]

    S_comp = np.linalg.pinv(C) @ Y

    return C, S_comp, eigenvalues


def fit_diffusion(C, b_values):
    b         = np.asarray(b_values, dtype=float)
    D_values  = np.zeros(C.shape[1])
    I0_values = np.zeros(C.shape[1])

    for k in range(C.shape[1]):
        decay = C[:, k]
        mask  = decay > 0
        if mask.sum() < 2:
            continue
        coeffs       = np.polyfit(b[mask], np.log(decay[mask]), deg=1)
        D_values[k]  = -coeffs[0]
        I0_values[k] =  np.exp(coeffs[1])

    return D_values, I0_values


# --- EXCEL HELPERS -----------------------------------------------------------

def _hdr(cell, value, bg="FF1F4E79"):
    cell.value     = value
    cell.font      = Font(bold=True, color="FFFFFFFF", name="Arial", size=10)
    cell.fill      = PatternFill("solid", start_color=bg, end_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")

def _label(cell, value):
    cell.value = value
    cell.font  = Font(bold=True, name="Arial", size=10)

def _val(cell, value):
    cell.value = value
    cell.font  = Font(name="Arial", size=10)

def _thin_border():
    s = Side(style="thin", color="FF000000")
    return Border(left=s, right=s, top=s, bottom=s)

def _apply_border(ws, min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                             min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = _thin_border()


def write_peak_sheet(wb, peak_label, b_values, C, S_comp, eigenvalues,
                     ppm_axis, D_values, I0_values):
    ws     = wb.create_sheet(title=peak_label[:31])
    b      = np.asarray(b_values, dtype=float)
    labels = ["Polymer", "Small molecule"]

    # header
    ws.merge_cells("A1:F1")
    ws["A1"].value     = f"DECRA Results  |  {peak_label}  |  ppm {ppm_axis.min():.3f} - {ppm_axis.max():.3f}"
    ws["A1"].font      = Font(bold=True, size=12, name="Arial", color="FF1F4E79")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    # fit summary table (rows 3-5)
    for col, h in enumerate(["Component", "D (m2/s)", "I0", "Eigenvalue", "Integral at b=0", "Equation"], 1):
        _hdr(ws.cell(3, col), h)

    dppm = abs(ppm_axis[1] - ppm_axis[0]) if len(ppm_axis) > 1 else 1.0
    for k in range(N_COMPONENTS):
        r        = 4 + k
        integral = float(np.trapz(S_comp[k, :], dx=dppm))
        _val(ws.cell(r, 1), labels[k])
        _val(ws.cell(r, 2), float(D_values[k]))
        _val(ws.cell(r, 3), float(I0_values[k]))
        _val(ws.cell(r, 4), float(eigenvalues[k]))
        _val(ws.cell(r, 5), float(integral))
        _val(ws.cell(r, 6), f"I(b) = {I0_values[k]:.4f} * exp(-{D_values[k]:.4e} * b)")
        ws.cell(r, 2).number_format = "0.00E+00"
        ws.cell(r, 3).number_format = "0.0000"
        ws.cell(r, 4).number_format = "0.000000"
        ws.cell(r, 5).number_format = "0.00"

    _apply_border(ws, 3, 3 + N_COMPONENTS, 1, 6)

    # decay data table
    DR = 8
    _hdr(ws.cell(DR, 1), "b (s/m2)")
    for k in range(N_COMPONENTS):
        _hdr(ws.cell(DR, 2 + k), f"{labels[k]} C(b)")

    for i, bv in enumerate(b):
        r = DR + 1 + i
        ws.cell(r, 1).value         = float(bv)
        ws.cell(r, 1).number_format = "0.00E+00"
        for k in range(N_COMPONENTS):
            ws.cell(r, 2 + k).value          = float(C[i, k])
            ws.cell(r, 2 + k).number_format  = "0.0000"

    # fitted curve table (100 points)
    FC    = 2 + N_COMPONENTS + 2
    n_fit = 100
    b_fit = np.linspace(b.min(), b.max(), n_fit)

    _hdr(ws.cell(DR, FC), "b fit (s/m2)")
    for k in range(N_COMPONENTS):
        _hdr(ws.cell(DR, FC + 1 + k), f"{labels[k]} fit")

    for i, bv in enumerate(b_fit):
        r = DR + 1 + i
        ws.cell(r, FC).value         = float(bv)
        ws.cell(r, FC).number_format = "0.00E+00"
        for k in range(N_COMPONENTS):
            ws.cell(r, FC + 1 + k).value          = float(I0_values[k] * np.exp(-D_values[k] * bv))
            ws.cell(r, FC + 1 + k).number_format  = "0.0000"

    # chart
    chart              = ScatterChart()
    chart.scatterStyle = "smoothMarker"
    chart.title        = f"Decay curves - {peak_label}"
    chart.x_axis.title = "b (s/m2)"
    chart.y_axis.title = "Normalised intensity"
    chart.height       = 14
    chart.width        = 22
    chart.legend.position = "r"

    n_rows_data = DR + len(b)

    for k in range(N_COMPONENTS):
        # raw decay points - dots only
        xvals  = Reference(ws, min_col=1,     min_row=DR + 1, max_row=n_rows_data)
        yvals  = Reference(ws, min_col=2 + k, min_row=DR + 1, max_row=n_rows_data)
        s_data = Series(yvals, xvals, title=f"{labels[k]} data")
        s_data.marker.symbol = "dot"
        s_data.marker.size   = 5
        s_data.graphicalProperties.line.noFill = True
        chart.series.append(s_data)

        # fitted curve - smooth line, no markers
        xfit  = Reference(ws, min_col=FC,           min_row=DR + 1, max_row=DR + n_fit)
        yfit  = Reference(ws, min_col=FC + 1 + k,   min_row=DR + 1, max_row=DR + n_fit)
        s_fit = Series(yfit, xfit, title=f"{labels[k]} fit")
        s_fit.marker.symbol = "none"
        s_fit.graphicalProperties.line.solidFill = "FF0000" if k == 0 else "4472C4"
        s_fit.graphicalProperties.line.width     = 20000
        chart.series.append(s_fit)

    ws.add_chart(chart, "A13")

    # column widths
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 44


def write_summary_sheet(wb, sample_name, peak_results):
    ws       = wb.active
    ws.title = "Summary"

    ws.merge_cells("A1:I1")
    ws["A1"].value     = f"DECRA Summary - {sample_name}"
    ws["A1"].font      = Font(bold=True, size=14, name="Arial", color="FF1F4E79")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26

    headers = ["Peak", "ppm window",
               "D polymer (m2/s)", "D small mol (m2/s)",
               "Integral polymer", "Integral small mol",
               "Ratio (poly/sm)", "lambda polymer", "lambda small mol"]
    for col, h in enumerate(headers, 1):
        _hdr(ws.cell(3, col), h)

    for row, (label, D_vals, I0_vals, eigs, S_comp, ppm_axis) in enumerate(peak_results, 4):
        dppm = abs(ppm_axis[1] - ppm_axis[0]) if len(ppm_axis) > 1 else 1.0
        i0   = float(np.trapz(S_comp[0, :], dx=dppm))
        i1   = float(np.trapz(S_comp[1, :], dx=dppm))

        vals = [label,
                f"{ppm_axis.min():.3f} - {ppm_axis.max():.3f}",
                float(D_vals[0]), float(D_vals[1]),
                i0, i1,
                f"=E{row}/F{row}",
                float(eigs[0]), float(eigs[1])]

        for col, v in enumerate(vals, 1):
            _val(ws.cell(row, col), v)
            ws.cell(row, col).alignment = Alignment(horizontal="center")

        ws.cell(row, 3).number_format = "0.00E+00"
        ws.cell(row, 4).number_format = "0.00E+00"
        ws.cell(row, 5).number_format = "0.00"
        ws.cell(row, 6).number_format = "0.00"
        ws.cell(row, 7).number_format = "0.00"
        ws.cell(row, 8).number_format = "0.000000"
        ws.cell(row, 9).number_format = "0.000000"

    _apply_border(ws, 3, 3 + len(peak_results), 1, len(headers))

    for col, width in zip("ABCDEFGHI", [20, 18, 20, 20, 16, 16, 14, 14, 14]):
        ws.column_dimensions[col].width = width


# -----------------------------------------------------------------------------

def main():
    nmr         = MnovaNMR.NMRPlugin()
    item        = nmr.activeNMRItem()
    sample_name = item.title(False) or "sample"
    b_values    = get_b_values(item)

    wb           = Workbook()
    peak_results = []

    for (peak_label, ppm_left, ppm_right) in PEAKS:
        print(f"Processing {peak_label} ...")
        Y, ppm_axis     = build_spectral_matrix(item, ppm_left, ppm_right)
        C, S_comp, eigs = decra(Y, n_components=N_COMPONENTS)
        D_vals, I0_vals = fit_diffusion(C, b_values)
        peak_results.append((peak_label, D_vals, I0_vals, eigs, S_comp, ppm_axis))
        write_peak_sheet(wb, peak_label, b_values, C, S_comp, eigs, ppm_axis, D_vals, I0_vals)

    write_summary_sheet(wb, sample_name, peak_results)
    wb.move_sheet("Summary", offset=-len(wb.sheetnames) + 1)

    out_path = os.path.join(OUTPUT_FOLDER, f"{sample_name}_DECRA.xlsx")
    wb.save(out_path)
    print(f"Saved: {out_path}")


main()
