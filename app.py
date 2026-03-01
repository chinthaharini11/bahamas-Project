from flask import Flask, request, jsonify, render_template
import zipfile
import os
import pdfplumber
import openpyxl
import base64
from PIL import Image
import io
import re

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'})

    file = request.files['file']
    if not file.filename.endswith('.zip'):
        return jsonify({'error': 'Please upload a ZIP file'})

    zip_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(zip_path)

    results = {
        'company_info': {},
        'directors': [],
        'shareholders': [],
        'auditor': {},
        'business_activity': [],
        'subsidiaries': [],
        'employees': {},
        'financial': {},
        'images': [],
        'raw_texts': []
    }

    with zipfile.ZipFile(zip_path, 'r') as z:
        for name in z.namelist():
            if z.getinfo(name).is_dir():
                continue
            lower = name.lower()
            short = name.split('/')[-1]
            data = z.read(name)

            # ── PDF ──────────────────────────────────────
            if lower.endswith('.pdf'):
                text = ''
                tables = []
                try:
                    with pdfplumber.open(io.BytesIO(data)) as pdf:
                        for page in pdf.pages:
                            t = page.extract_text()
                            if t:
                                text += t + '\n'
                            for table in page.extract_tables():
                                if table:
                                    tables.append(table)
                except Exception as e:
                    text = f'Error reading PDF: {str(e)}'

                results['raw_texts'].append({
                    'name': short,
                    'text': text,
                    'tables': tables
                })

                # Detect file type and extract accordingly
                name_upper = short.upper()
                if 'AOC' in name_upper or 'AOC-4' in name_upper:
                    extract_aoc4(text, tables, results)
                elif 'MGT' in name_upper or 'MGT-7' in name_upper:
                    extract_mgt7(text, tables, results)
                else:
                    # Try extracting from any PDF
                    extract_general(text, results)

            # ── EXCEL ────────────────────────────────────
            elif lower.endswith(('.xlsx', '.xls', '.xlsm')):
                try:
                    wb = openpyxl.load_workbook(
                        io.BytesIO(data), data_only=True)
                    for sheet in wb.sheetnames:
                        ws = wb[sheet]
                        rows = []
                        headers = []
                        for i, row in enumerate(
                                ws.iter_rows(values_only=True)):
                            clean = [
                                str(c).strip() if c is not None else ''
                                for c in row
                            ]
                            if any(c for c in clean):
                                if i == 0:
                                    headers = clean
                                else:
                                    rows.append(clean)
                        if rows:
                            results['shareholders'].append({
                                'sheet': sheet,
                                'headers': headers,
                                'rows': rows,
                                'source': short
                            })
                except Exception as e:
                    pass

            # ── IMAGE ────────────────────────────────────
            elif lower.endswith(
                    ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')):
                try:
                    img = Image.open(io.BytesIO(data))
                    img.thumbnail((900, 900))
                    buf = io.BytesIO()
                    fmt = 'JPEG' if lower.endswith(
                        ('.jpg', '.jpeg')) else 'PNG'
                    img.save(buf, format=fmt)
                    b64 = base64.b64encode(
                        buf.getvalue()).decode('utf-8')
                    mime = 'jpeg' if fmt == 'JPEG' else 'png'
                    results['images'].append({
                        'name': short,
                        'data': b64,
                        'mime': mime
                    })
                except Exception as e:
                    pass

    return jsonify(results)


def extract_aoc4(text, tables, results):
    """Extract fields from AOC-4 PDF"""
    patterns = {
        'CIN': [
            r'Corporate [Ii]dentity [Nn]umber[^\n]*?([A-Z]{1}[0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6})',
            r'CIN[^\n:]*?:?\s*([A-Z]{1}[0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6})'
        ],
        'Company Name': [
            r'Name of the [Cc]ompany[^\n]*?\n([A-Z][^\n]{3,80})',
            r'\*?Name of the company[^\n]*?\n?\s*([A-Z][A-Z\s&(),.-]{3,80}(?:LIMITED|LTD|PRIVATE|PVT)[\s.]*(?:LIMITED|LTD)?)'
        ],
        'Registered Address': [
            r'Address of the registered office[^\n]*?\n([^\n]{10,150})',
            r'\*?Address of the registered[^\n]*?\n?\s*([A-Z0-9][^\n]{10,150})'
        ],
        'Email': [
            r'e-mail[^\n]*?([a-zA-Z0-9._%+\-*]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})',
            r'[Ee]mail[^\n]*?([a-zA-Z0-9._%+\-*]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})'
        ],
        'Financial Year From': [
            r'[Ff]inancial year[^\n]*?[Ff]rom[^\n]*?(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})',
            r'From \(DD/MM/YYYY\)[^\n]*?\n\s*(\d{2}\/\d{2}\/\d{4})'
        ],
        'Financial Year To': [
            r'[Ff]inancial year[^\n]*?[Tt]o[^\n]*?(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})',
            r'To \(DD/MM/YYYY\)[^\n]*?\n\s*(\d{2}\/\d{2}\/\d{4})'
        ],
        'Authorised Capital': [
            r'[Aa]uthorised capital[^\n]*?(\d[\d,\.]+)',
            r'[Aa]uthorised [Cc]apital of the company[^\n]*?(\d[\d,\.]+)'
        ],
        'Paid Up Capital': [
            r'[Pp]aid.?up capital[^\n]*?(\d[\d,\.]+)',
        ],
        'Turnover': [
            r'[Tt]urnover[^\n]*?(\d[\d,\.]+)',
        ],
    }

    for field, rxs in patterns.items():
        for rx in rxs:
            m = re.search(rx, text)
            if m:
                val = m.group(1).strip().replace('\n', ' ')
                if val and field not in results['company_info']:
                    results['company_info'][field] = val
                break

    # Auditor from AOC-4
    auditor_patterns = {
        'Auditor Firm Name': [r'[Ff]irm [Nn]ame[^\n]*?\n\s*([A-Z][^\n]{3,80})'],
        'Auditor PAN': [r'PAN of [Aa]uditor[^\n]*?([A-Z]{5}[0-9]{4}[A-Z])'],
        'Auditor Reg No': [r'[Rr]egistration [Nn]umber[^\n]*?(\d{6,})'],
        'Signing Member': [r'[Mm]embership [Nn]umber[^\n]*?(\d{4,8})'],
    }
    for field, rxs in auditor_patterns.items():
        for rx in rxs:
            m = re.search(rx, text)
            if m:
                results['auditor'][field] = m.group(1).strip()
                break


def extract_mgt7(text, tables, results):
    """Extract fields from MGT-7A PDF"""
    # Company info from MGT-7
    mgt_patterns = {
        'CIN': [r'([A-Z]{1}[0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6})'],
        'Company Name': [r'Name of [Cc]ompany[^\n]*?\n\s*([A-Z][^\n]{3,80})'],
        'AGM Date': [r'AGM[^\n]*?(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})'],
        'SRN': [r'SRN[^\n:]*?:?\s*([A-Z][0-9]{8,})'],
        'Financial Year': [r'[Ff]inancial [Yy]ear[^\n]*?(\d{4}[-–]\d{2,4})'],
    }
    for field, rxs in mgt_patterns.items():
        for rx in rxs:
            m = re.search(rx, text)
            if m:
                val = m.group(1).strip()
                if val and field not in results['company_info']:
                    results['company_info'][field] = val
                break

    # Employee info
    emp_patterns = {
        'Total Employees': [r'[Tt]otal[^\n]*?[Ee]mployees?[^\n]*?(\d+)'],
        'Male': [r'[Mm]ale[^\n]*?(\d+)'],
        'Female': [r'[Ff]emale[^\n]*?(\d+)'],
    }
    for field, rxs in emp_patterns.items():
        for rx in rxs:
            m = re.search(rx, text)
            if m:
                results['employees'][field] = m.group(1).strip()
                break

    # Business activity from tables
    for table in tables:
        if not table:
            continue
        for row in table:
            if not row:
                continue
            row_text = ' '.join(str(c) for c in row if c)
            if any(k in row_text.upper() for k in
                   ['NIC', 'ACTIVITY', 'TURNOVER', 'BUSINESS']):
                clean = [str(c).strip() if c else '' for c in row]
                if any(c for c in clean):
                    results['business_activity'].append(clean)

    # Directors from tables
    for table in tables:
        if not table:
            continue
        for row in table:
            if not row:
                continue
            row_text = ' '.join(str(c) for c in row if c)
            if any(k in row_text.upper() for k in
                   ['DIN', 'DIRECTOR', 'DESIGNATION']):
                clean = [str(c).strip() if c else '' for c in row]
                if any(c for c in clean):
                    results['directors'].append(clean)

    # Subsidiaries from tables
    for table in tables:
        if not table:
            continue
        for row in table:
            if not row:
                continue
            row_text = ' '.join(str(c) for c in row if c)
            if any(k in row_text.upper() for k in
                   ['SUBSIDIARY', 'ASSOCIATE', 'HOLDING']):
                clean = [str(c).strip() if c else '' for c in row]
                if any(c for c in clean):
                    results['subsidiaries'].append(clean)


def extract_general(text, results):
    """Extract common fields from any PDF"""
    patterns = {
        'CIN': r'([A-Z]{1}[0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6})',
        'Email': r'([a-zA-Z0-9._%+\-*]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})',
        'Phone': r'(\+?91[\s\-]?[6-9]\d{9}|[6-9]\d{9})',
        'GSTIN': r'([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z])',
        'PAN': r'([A-Z]{5}[0-9]{4}[A-Z])',
    }
    for field, rx in patterns.items():
        m = re.search(rx, text)
        if m and field not in results['company_info']:
            results['company_info'][field] = m.group(1).strip()


if __name__ == '__main__':
    app.run(debug=True)
