#!/usr/bin/env python3
"""
NMR Curve Fitting Report Generator
Pure Python - generates HTML report with SVG charts.
Open in browser and print to PDF.
"""

import os
import re
import zipfile
import xml.etree.ElementTree as ET


def read_xlsx(filepath):
    """Read xlsx without external libs. Returns {sheet_name: [[row], ...]}"""
    sheets = {}
    
    with zipfile.ZipFile(filepath, 'r') as z:
        # Shared strings
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            tree = ET.parse(z.open('xl/sharedStrings.xml'))
            ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
            for si in tree.findall(f'.//{ns}si'):
                shared_strings.append(''.join(t.text or '' for t in si.iter(f'{ns}t')))
        
        # Sheet names
        tree = ET.parse(z.open('xl/workbook.xml'))
        ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        sheet_names = [s.get('name') for s in tree.findall(f'.//{ns}sheet')]
        
        # Read each sheet
        for idx, name in enumerate(sheet_names):
            sheet_file = f'xl/worksheets/sheet{idx + 1}.xml'
            if sheet_file not in z.namelist():
                continue
            
            tree = ET.parse(z.open(sheet_file))
            rows_data = {}
            
            for row in tree.findall(f'.//{ns}row'):
                row_num = int(row.get('r'))
                rows_data[row_num] = {}
                
                for cell in row.findall(f'{ns}c'):
                    col = re.match(r'([A-Z]+)', cell.get('r')).group(1)
                    col_idx = sum((ord(c) - 64) * (26 ** i) for i, c in enumerate(reversed(col))) - 1
                    
                    val_elem = cell.find(f'{ns}v')
                    if val_elem is not None and val_elem.text:
                        if cell.get('t') == 's':
                            val = shared_strings[int(val_elem.text)]
                        else:
                            try:
                                val = float(val_elem.text)
                            except:
                                val = val_elem.text
                    else:
                        val = None
                    rows_data[row_num][col_idx] = val
            
            if rows_data:
                max_row = max(rows_data.keys())
                max_col = max(max(r.keys()) for r in rows_data.values() if r) + 1
                sheets[name] = [[rows_data.get(r, {}).get(c) for c in range(max_col)] 
                               for r in range(1, max_row + 1)]
    
    return sheets


def process_sheet(name, data):
    """Extract best model and plot data from sheet."""
    summary_idx = None
    for i, row in enumerate(data):
        row_str = ' '.join(str(c).lower() for c in row if c)
        if 'name' in row_str and 'formula' in row_str:
            summary_idx = i
            break
    
    if summary_idx is None:
        return None
    
    headers = [str(h).strip() if h else '' for h in data[summary_idx]]
    headers_lower = [h.lower() for h in headers]
    
    def find_col(keywords):
        for i, h in enumerate(headers_lower):
            if any(k in h for k in keywords):
                return i
        return None
    
    idx_col = find_col(['#'])
    name_col = find_col(['name'])
    formula_col = find_col(['formula'])
    r2_col = find_col(['r²', 'r2'])
    diff_col = find_col(['diffusion'])  # dedicated column
    param_col = find_col(['parameters'])
    
    best_r2, best_row = -1, None
    for row in data[summary_idx + 1:]:
        if r2_col is not None and row[r2_col] is not None:
            try:
                r2 = float(row[r2_col])
                if r2 > best_r2:
                    best_r2, best_row = r2, row
            except:
                pass
    
    if not best_row:
        return None
    
    # Get model name and find corresponding fit/res columns
    model_name = best_row[name_col] if name_col is not None else 'Unknown'
    
    # Map model name to column prefix (from curve_fit_function.py _MODEL_META)
    model_to_col = {
        'Mono-Exp': 'MonoExp',
        'Mono-Exp+Offset': 'MonoExpOffset', 
        'Bi-Exp': 'BiExp',
        'Linear': 'Linear',
        'Inv.Linear': 'InvLinear',
        'Inv.Lin+Offset': 'InvLinearOffset',
        'Intermediate': 'Intermediate',
        'Interm+Offset': 'IntermediateOffset',
    }
    col_prefix = model_to_col.get(model_name, '')
    
    data_headers = [str(h).lower() if h else '' for h in data[0]]
    x_col = next((i for i, h in enumerate(data_headers) if 'b_val' in h or 'bval' in h or 'bvalue' in h), 0)
    y_col = next((i for i, h in enumerate(data_headers) if 'integral' in h), 1)
    
    # Find fit/res columns using the model's column prefix
    fit_col = None
    res_col = None
    if col_prefix:
        fit_col = next((i for i, h in enumerate(data_headers) if f'{col_prefix.lower()}_fit' in h), None)
        res_col = next((i for i, h in enumerate(data_headers) if f'{col_prefix.lower()}_res' in h), None)
    
    x, y, fit, res = [], [], [], []
    for row in data[1:summary_idx]:
        try:
            xv = float(row[x_col]) if row[x_col] else None
            if xv is not None:
                x.append(xv)
                y.append(float(row[y_col]) if row[y_col] else None)
                fit.append(float(row[fit_col]) if fit_col is not None and row[fit_col] else None)
                res.append(float(row[res_col]) if res_col is not None and row[res_col] else None)
        except:
            pass
    
    # Get diffusion coefficient directly from dedicated column
    diff_coef = best_row[diff_col] if diff_col is not None else None
    
    # Get formula
    formula = best_row[formula_col] if formula_col is not None else ''
    
    return {
        'name': name,
        'model': model_name,
        'formula': formula,
        'r2': best_r2,
        'diff_coef': diff_coef,
        'x': x, 'y': y, 'fit': fit, 'res': res
    }


def make_svg_chart(x, y, fit, res, title, r2, width=800, height=300):
    """Generate SVG chart with fit (left) and residual (right) side by side."""
    
    if not x or not y:
        return ""
    
    valid = [(xi, yi, fi, ri) for xi, yi, fi, ri in zip(x, y, fit, res) 
             if xi is not None and yi is not None]
    if not valid:
        return ""
    
    x_vals = [v[0] for v in valid]
    y_vals = [v[1] for v in valid]
    
    pad = 60
    chart_w = (width - 3 * pad) // 2
    chart_h = height - 2 * pad
    
    def scale(vals, size):
        mn, mx = min(vals), max(vals)
        rng = mx - mn if mx != mn else 1
        return mn, mx, lambda v: pad + (v - mn) / rng * size
    
    x_min, x_max, sx = scale(x_vals, chart_w)
    y_min, y_max, sy = scale(y_vals, chart_h)
    
    svg = [f'<svg width="{width}" height="{height}" xmlns="http://www.w3.org/2000/svg">']
    svg.append('<rect width="100%" height="100%" fill="white"/>')
    
    # Left chart: Fitted curve
    svg.append(f'<text x="{pad + chart_w//2}" y="20" text-anchor="middle" font-size="14" font-weight="bold">{title} (R² = {r2:.5f})</text>')
    
    svg.append(f'<line x1="{pad}" y1="{height-pad}" x2="{pad+chart_w}" y2="{height-pad}" stroke="black"/>')
    svg.append(f'<line x1="{pad}" y1="{pad}" x2="{pad}" y2="{height-pad}" stroke="black"/>')
    
    svg.append(f'<text x="{pad + chart_w//2}" y="{height-10}" text-anchor="middle" font-size="11">b-value</text>')
    svg.append(f'<text x="15" y="{height//2}" text-anchor="middle" font-size="11" transform="rotate(-90 15 {height//2})">Signal</text>')
    
    svg.append(f'<text x="{pad}" y="{height-pad+15}" text-anchor="middle" font-size="9">{x_min:.0f}</text>')
    svg.append(f'<text x="{pad+chart_w}" y="{height-pad+15}" text-anchor="middle" font-size="9">{x_max:.0f}</text>')
    svg.append(f'<text x="{pad-5}" y="{height-pad}" text-anchor="end" font-size="9">{y_min:.2f}</text>')
    svg.append(f'<text x="{pad-5}" y="{pad+5}" text-anchor="end" font-size="9">{y_max:.2f}</text>')
    
    for xi, yi in zip(x_vals, y_vals):
        px, py = sx(xi), height - pad - (sy(yi) - pad)
        svg.append(f'<circle cx="{px}" cy="{py}" r="4" fill="#4a90d9" opacity="0.7"/>')
    
    fit_points = [(xi, fi) for xi, fi in zip(x_vals, [v[2] for v in valid]) if fi is not None]
    fit_points.sort(key=lambda p: p[0])
    if fit_points:
        path = f'M {sx(fit_points[0][0])} {height - pad - (sy(fit_points[0][1]) - pad)}'
        for xi, fi in fit_points[1:]:
            path += f' L {sx(xi)} {height - pad - (sy(fi) - pad)}'
        svg.append(f'<path d="{path}" stroke="red" stroke-width="2" fill="none"/>')
    
    svg.append(f'<circle cx="{pad+chart_w-60}" cy="{pad+10}" r="4" fill="#4a90d9"/>')
    svg.append(f'<text x="{pad+chart_w-50}" y="{pad+14}" font-size="10">Data</text>')
    svg.append(f'<line x1="{pad+chart_w-65}" y1="{pad+25}" x2="{pad+chart_w-55}" y2="{pad+25}" stroke="red" stroke-width="2"/>')
    svg.append(f'<text x="{pad+chart_w-50}" y="{pad+29}" font-size="10">Fit</text>')
    
    # Right chart: Residuals
    offset = pad + chart_w + pad
    res_points = [(xi, ri) for xi, ri in zip(x_vals, [v[3] for v in valid]) if ri is not None]
    
    if res_points:
        r_vals = [p[1] for p in res_points]
        r_min, r_max, sr = scale(r_vals, chart_h)
        
        svg.append(f'<text x="{offset + chart_w//2}" y="20" text-anchor="middle" font-size="14" font-weight="bold">Residuals</text>')
        
        svg.append(f'<line x1="{offset}" y1="{height-pad}" x2="{offset+chart_w}" y2="{height-pad}" stroke="black"/>')
        svg.append(f'<line x1="{offset}" y1="{pad}" x2="{offset}" y2="{height-pad}" stroke="black"/>')
        
        zero_y = height - pad - (sr(0) - pad) if r_min <= 0 <= r_max else height // 2
        svg.append(f'<line x1="{offset}" y1="{zero_y}" x2="{offset+chart_w}" y2="{zero_y}" stroke="red" stroke-dasharray="5,3"/>')
        
        svg.append(f'<text x="{offset + chart_w//2}" y="{height-10}" text-anchor="middle" font-size="11">b-value</text>')
        svg.append(f'<text x="{offset-45}" y="{height//2}" text-anchor="middle" font-size="11" transform="rotate(-90 {offset-45} {height//2})">Residual</text>')
        
        svg.append(f'<text x="{offset}" y="{height-pad+15}" text-anchor="middle" font-size="9">{x_min:.0f}</text>')
        svg.append(f'<text x="{offset+chart_w}" y="{height-pad+15}" text-anchor="middle" font-size="9">{x_max:.0f}</text>')
        svg.append(f'<text x="{offset-5}" y="{height-pad}" text-anchor="end" font-size="9">{r_min:.3f}</text>')
        svg.append(f'<text x="{offset-5}" y="{pad+5}" text-anchor="end" font-size="9">{r_max:.3f}</text>')
        
        for xi, ri in res_points:
            px = offset + (sx(xi) - pad)
            py = height - pad - (sr(ri) - pad)
            svg.append(f'<circle cx="{px}" cy="{py}" r="4" fill="#2ecc71" opacity="0.7"/>')
    
    svg.append('</svg>')
    return '\n'.join(svg)


def generate_pdf_report(input_folder, output_folder=None):
    """Generate HTML report from all Excel files in a folder."""
    
    if output_folder is None:
        output_folder = input_folder
    
    if output_folder:
        os.makedirs(output_folder, exist_ok=True)
    
    output_filename = "curve_fitting_report.html"
    output_path = os.path.join(output_folder, output_filename) if output_folder else output_filename
    
    # Find all xlsx files
    xlsx_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx') and not f.startswith('~')]
    xlsx_files.sort()
    
    print(f"Found {len(xlsx_files)} Excel files in: {input_folder}")
    
    all_results = []
    
    for xlsx_file in xlsx_files:
        excel_path = os.path.join(input_folder, xlsx_file)
        sample_name = os.path.splitext(xlsx_file)[0]
        
        print(f"\nReading: {xlsx_file}")
        try:
            sheets = read_xlsx(excel_path)
            print(f"  Found {len(sheets)} sheets")
            
            for sheet_name, data in sheets.items():
                r = process_sheet(sheet_name, data)
                if r:
                    r['sample_name'] = sample_name
                    r['peak_name'] = sheet_name
                    all_results.append(r)
                    print(f"    {sheet_name}: {r['model']} (R²={r['r2']:.4f})")
                else:
                    all_results.append({'sample_name': sample_name, 'peak_name': sheet_name, 'error': True})
        except Exception as e:
            print(f"  Error: {e}")
    
    print(f"\nTotal results: {len(all_results)}")
    
    html = ['''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Curve Fitting Report</title>
<style>
body { font-family: Arial, sans-serif; margin: 40px; }
h1 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
h2 { color: #34495e; margin-top: 30px; }
h3 { color: #7f8c8d; }
table { border-collapse: collapse; width: 100%; margin: 20px 0; }
th { background: #3498db; color: white; padding: 10px; text-align: left; }
td { padding: 8px; border-bottom: 1px solid #ddd; }
tr:nth-child(even) { background: #f9f9f9; }
.chart-container { margin: 20px 0; page-break-inside: avoid; }
@media print { .chart-container { page-break-inside: avoid; } }
</style>
</head>
<body>
<h1>Curve Fitting Report</h1>
''']
    
    html.append('<h2>Section 1: Summary</h2>')
    html.append('<table>')
    html.append('<tr><th>Sample Name</th><th>Peak Name</th><th>Model (Equation)</th><th>R²</th><th>Diffusion Coefficient</th></tr>')
    
    for r in all_results:
        if r.get('error'):
            html.append(f'<tr><td>{r["sample_name"]}</td><td>{r["peak_name"]}</td><td colspan="3">Error processing</td></tr>')
        else:
            model_eq = f"{r['model']} ({r['formula']})" if r['formula'] else r['model']
            if len(model_eq) > 50:
                model_eq = model_eq[:47] + '...'
            diff = f"{r['diff_coef']:.4e}" if isinstance(r['diff_coef'], float) else 'N/A'
            html.append(f'<tr><td>{r["sample_name"]}</td><td>{r["peak_name"]}</td><td>{model_eq}</td>'
                       f'<td>{r["r2"]:.5f}</td><td>{diff}</td></tr>')
    
    html.append('</table>')
    
    html.append('<h2>Section 2: Fitted Curves & Residuals</h2>')
    
    for r in all_results:
        if r.get('error') or not r.get('x'):
            continue
        
        html.append(f'<div class="chart-container">')
        html.append(f'<h3>{r["sample_name"]} / {r["peak_name"]} - {r["model"]}</h3>')
        svg = make_svg_chart(r['x'], r['y'], r['fit'], r['res'], 'Fitted Curve', r['r2'])
        html.append(svg)
        html.append('</div>')
    
    html.append('</body></html>')
    
    with open(output_path, 'w') as f:
        f.write('\n'.join(html))
    
    print(f"\nDone: {output_path}")
    print("Open in browser and print to PDF (Ctrl+P / Cmd+P)")
    return output_path


if __name__ == "__main__":
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    INPUT_FOLDER = SCRIPT_DIR  # folder containing xlsx files
    OUTPUT_FOLDER = SCRIPT_DIR  # where to save report
    
    generate_pdf_report(INPUT_FOLDER, OUTPUT_FOLDER)
