from flask import Flask, request, jsonify, render_template, send_file
import zipfile
import os
import tempfile
import pdfplumber
import openpyxl
import base64
import io
import re
import anthropic
import json
import uuid
import concurrent.futures
import threading
from pathlib import Path

# Read .env directly — handles BOM, quotes, CRLF, and spaces
_env_path = Path(__file__).parent / '.env'
if _env_path.exists():
    with open(_env_path, encoding='utf-8-sig') as _ef:  # utf-8-sig strips BOM automatically
        for _line in _ef:
            _line = _line.strip()
            if _line and not _line.startswith('#') and '=' in _line:
                _k, _, _v = _line.partition('=')
                _k = _k.strip()
                _v = _v.strip().strip('"').strip("'")  # remove surrounding quotes
                if _k and _v:
                    os.environ.setdefault(_k, _v)
                    print(f"[.env] Loaded: {_k} = {_v[:6]}...")

try:
    from openai import OpenAI as OpenAIClient
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

try:
    import pytesseract
    from pdf2image import convert_from_bytes
    from docx import Document
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

if OCR_AVAILABLE and os.name == 'nt':
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'bahamas_uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

_NO_TEXT_PLACEHOLDERS = frozenset([
    "[Scanned PDF - OCR returned no text]",
    "[No text could be extracted from this PDF]",
])

_anthropic_client = None
_openai_client = None
_client_lock = threading.Lock()

def get_anthropic_client():
    global _anthropic_client
    if _anthropic_client is None:
        with _client_lock:
            if _anthropic_client is None:
                api_key = os.environ.get('ANTHROPIC_API_KEY')
                if api_key:
                    _anthropic_client = anthropic.Anthropic(api_key=api_key)
    return _anthropic_client

def get_openai_client():
    global _openai_client
    if _openai_client is None and OPENAI_AVAILABLE:
        with _client_lock:
            if _openai_client is None and OPENAI_AVAILABLE:
                api_key = os.environ.get('OPENAI_API_KEY')
                if api_key:
                    _openai_client = OpenAIClient(api_key=api_key)
    return _openai_client

def parse_ai_json(raw):
    """Strip markdown fences and parse JSON from AI response."""
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r'^```[a-z]*\n?', '', raw)
        raw = re.sub(r'\n?```$', '', raw)
    return json.loads(raw)


def extract_text_with_ocr(pdf_data, max_pages=10):
    if not OCR_AVAILABLE:
        return ''
    try:
        images = convert_from_bytes(pdf_data, dpi=100, last_page=max_pages)
        # Parallelize OCR across pages — each page spawns its own tesseract process
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(len(images), 4)) as page_exec:
            page_texts = list(page_exec.map(pytesseract.image_to_string, images))
        return '\n'.join(page_texts)
    except Exception as e:
        print(f"OCR ERROR: {str(e)}")
        return ''


def extract_data_from_image(image_b64, mime, filename):
    """Use Claude vision to extract structured data from an image."""
    client = get_anthropic_client()
    if not client:
        return None

    prompt = (
        "This image is from a company compliance document (like MGT-7 or AOC-4 filings). "
        "Extract all structured data you can find. "
        "Return a JSON object with these keys (only include keys that have data):\n"
        "- company_info: dict of company details (CIN, name, address, email, PAN, etc.)\n"
        "- shareholders: list of {name, shares} dicts if a shareholder table exists\n"
        "- directors: list of {name, din, designation} dicts if a directors table exists\n"
        "- financial: dict of financial figures if present\n"
        "- other_tables: list of any other tables as list-of-lists\n"
        "Return ONLY valid JSON, no explanation."
    )

    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=2048,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": mime,
                            "data": image_b64,
                        },
                    },
                    {"type": "text", "text": prompt}
                ],
            }]
        )
        return parse_ai_json(response.content[0].text)
    except Exception:
        return None




def extract_structured_from_text(text):
    """Try Claude first, fall back to OpenAI GPT-4o for text structuring."""
    prompt = (
        "Look at this document text carefully. Extract ALL information and organize it.\n"
        "Return a JSON array of sections. Each section has:\n"
        "  - heading: a short descriptive title (e.g. 'Company Details', 'Directors')\n"
        "  - type: either 'fields', 'table', or 'text'\n"
        "  - For type='fields': include 'data' as an object of key-value pairs\n"
        "  - For type='table': include 'headers' (array of strings) and 'rows' (array of arrays)\n"
        "  - For type='text': include 'content' as a plain string\n"
        "Return ONLY valid JSON array, no explanation, no markdown fences.\n\n"
        "Document text:\n"
    )

    # ── Try Claude first ──────────────────────────────────
    claude = get_anthropic_client()
    if claude:
        try:
            response = claude.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=4000,
                messages=[{"role": "user", "content": prompt + text[:12000]}]
            )
            return parse_ai_json(response.content[0].text)
        except Exception as e:
            print(f"Claude text failed: {e} — trying OpenAI...")

    # ── Fallback: OpenAI GPT-4o ───────────────────────────
    oai = get_openai_client()
    if oai:
        try:
            response = oai.chat.completions.create(
                model="gpt-4o",
                max_tokens=4000,
                messages=[{"role": "user", "content": prompt + text[:12000]}]
            )
            return parse_ai_json(response.choices[0].message.content)
        except Exception as e:
            print(f"OpenAI text failed: {e}")

    return [{'heading': 'Extracted Text', 'type': 'text', 'content': text}]


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'})

    file = request.files['file']
    lower_name = file.filename.lower()

    # Single image upload — OCR only, no AI
    image_exts = ('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.webp')
    if lower_name.endswith(image_exts):
        image_bytes = file.read()
        text = ''
        if OCR_AVAILABLE:
            try:
                from PIL import Image
                img = Image.open(io.BytesIO(image_bytes))
                text = pytesseract.image_to_string(img).strip()
            except Exception as e:
                print(f"Image OCR error: {e}")
        if text:
            structured = [{'heading': 'Extracted Text', 'type': 'text', 'content': text}]
        else:
            structured = [{'heading': 'No Text Found', 'type': 'text',
                           'content': 'Could not extract text from this image. Make sure Tesseract is installed.'}]
        return jsonify({
            'type': 'image',
            'filename': file.filename,
            'structured': structured
        })

    # Standalone PDF
    if lower_name.endswith('.pdf'):
        pdf_data = file.read()
        text = ''
        try:
            with pdfplumber.open(io.BytesIO(pdf_data)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text += t + '\n'
        except Exception as e:
            print(f"pdfplumber error: {e}")
        if not text.strip():
            text = extract_text_with_ocr(pdf_data)
        if not text.strip():
            text = '[No text could be extracted from this PDF]'
        if text in _NO_TEXT_PLACEHOLDERS:
            structured = [{'heading': 'No Text Extracted', 'type': 'text',
                            'content': 'Could not extract text from this PDF.'}]
        else:
            structured = extract_structured_from_text(text)
        return jsonify({'type': 'image', 'filename': file.filename, 'structured': structured})

    # Standalone Word (.docx)
    if lower_name.endswith('.docx'):
        if not OCR_AVAILABLE:
            return jsonify({'error': 'Word file support requires python-docx, pytesseract, and pdf2image to be installed'})
        doc_data = file.read()
        try:
            doc = Document(io.BytesIO(doc_data))
            text = '\n'.join(para.text for para in doc.paragraphs if para.text.strip())
        except Exception as e:
            return jsonify({'error': f'Could not read Word file: {str(e)}'})
        structured = extract_structured_from_text(text)
        return jsonify({'type': 'image', 'filename': file.filename, 'structured': structured})

    # Standalone Excel
    if lower_name.endswith(('.xlsx', '.xls', '.xlsm')):
        excel_data = file.read()
        try:
            wb = openpyxl.load_workbook(io.BytesIO(excel_data), data_only=True)
            text = ''
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                text += f'Sheet: {sheet}\n'
                for row in ws.iter_rows(values_only=True):
                    row_text = ' | '.join(str(c) if c is not None else '' for c in row)
                    if row_text.strip('| '):
                        text += row_text + '\n'
        except Exception as e:
            return jsonify({'error': f'Could not read Excel file: {str(e)}'})
        structured = extract_structured_from_text(text)
        return jsonify({'type': 'image', 'filename': file.filename, 'structured': structured})

    if not lower_name.endswith('.zip'):
        return jsonify({'error': 'Unsupported file type. Please upload a PDF, image, Word (.docx), Excel, or ZIP file.'})

    safe_name = os.path.basename(file.filename)
    if not safe_name:
        return jsonify({'error': 'Invalid filename'})
    unique_name = f"{uuid.uuid4().hex}_{safe_name}"
    zip_path = os.path.join(UPLOAD_FOLDER, unique_name)
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

    if not zipfile.is_zipfile(zip_path):
        os.remove(zip_path)
        return jsonify({'error': 'Invalid or corrupted ZIP file'})

    # Read all entries up front so the ZipFile handle can be closed before threads run
    entries = []
    with zipfile.ZipFile(zip_path, 'r') as z:
        for name in z.namelist():
            entries.append((name, name.lower(), name.split('/')[-1], z.read(name)))

    try:
        os.remove(zip_path)
    except Exception:
        pass

    def process_zip_entry(entry):
        _, lower, short, data = entry
        p = {'company_info': {}, 'directors': [], 'shareholders': [],
             'auditor': {}, 'business_activity': [], 'subsidiaries': [],
             'employees': {}, 'financial': {}, 'raw_texts': []}
        try:
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

                if not text.strip() or text.startswith('Error reading PDF'):
                    ocr_text = extract_text_with_ocr(data)
                    print(f"OCR result length: {len(ocr_text)}, preview: {ocr_text[:100]}")
                    text = ocr_text.strip() or "[Scanned PDF - OCR returned no text]"

                name_upper = short.upper()
                text_upper = text[:2000].upper()
                # Detect type by filename OR text content (handles generically named files)
                is_aoc = 'AOC' in name_upper or 'FORM AOC' in text_upper or 'FORM NO. AOC' in text_upper
                is_mgt = 'MGT' in name_upper or 'FORM MGT' in text_upper or 'FORM NO. MGT' in text_upper or 'ANNUAL RETURN' in text_upper
                is_known_type = is_aoc or is_mgt

                if text in _NO_TEXT_PLACEHOLDERS:
                    structured = [{'heading': 'No Text Extracted', 'type': 'text',
                                    'content': 'Could not extract text from this PDF.'}]
                elif is_known_type:
                    # Regex extraction handles AOC-4/MGT-7; skip the extra AI call
                    structured = []
                else:
                    structured = extract_structured_from_text(text)
                p['raw_texts'].append({'name': short, 'text': text, 'tables': tables, 'structured': structured})

                if is_aoc:
                    extracted = extract_aoc4(text, tables)
                elif is_mgt:
                    extracted = extract_mgt7(text, tables)
                else:
                    extracted = extract_general(text)

                for k, v in extracted.get('company_info', {}).items():
                    p['company_info'].setdefault(k, v)
                for k, v in extracted.get('auditor', {}).items():
                    p['auditor'].setdefault(k, v)
                for k, v in extracted.get('employees', {}).items():
                    p['employees'].setdefault(k, v)
                p['business_activity'].extend(extracted.get('business_activity', []))
                p['subsidiaries'].extend(extracted.get('subsidiaries', []))
                p['directors'].extend(extracted.get('directors', []))

            # ── EXCEL ────────────────────────────────────
            elif lower.endswith(('.xlsx', '.xls', '.xlsm')):
                wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
                for sheet in wb.sheetnames:
                    ws = wb[sheet]
                    rows = []
                    headers = []
                    for i, row in enumerate(ws.iter_rows(values_only=True)):
                        clean = [str(c).strip() if c is not None else '' for c in row]
                        if any(c for c in clean):
                            if i == 0:
                                headers = clean
                            else:
                                rows.append(clean)
                    if rows:
                        p['shareholders'].append({'sheet': sheet, 'headers': headers, 'rows': rows, 'source': short})

            # ── IMAGE ────────────────────────────────────
            elif lower.endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp')):
                ext = lower.rsplit('.', 1)[-1]
                mime_map = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
                            'tiff': 'image/tiff', 'bmp': 'image/bmp'}
                mime_type = mime_map.get(ext, 'image/png')
                b64 = base64.b64encode(data).decode('utf-8')
                extracted = extract_data_from_image(b64, mime_type, short)
                if extracted:
                    for k, v in extracted.get('company_info', {}).items():
                        p['company_info'].setdefault(k, v)
                    for sh in extracted.get('shareholders', []):
                        p['shareholders'].append({'source': short, 'name': sh.get('name', ''), 'shares': sh.get('shares', '')})
                    for d in extracted.get('directors', []):
                        p['directors'].append([d.get('din', ''), d.get('name', ''), d.get('designation', '')])
                    for k, v in extracted.get('financial', {}).items():
                        p['financial'].setdefault(k, v)

        except Exception as e:
            print(f"Unhandled error processing {short}: {e}")
        return p

    # Process all entries in parallel (up to 8 concurrent workers)
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        partials = list(executor.map(process_zip_entry, entries))

    # Merge partial results into final results
    for p in partials:
        for k, v in p['company_info'].items():
            results['company_info'].setdefault(k, v)
        for k, v in p['auditor'].items():
            results['auditor'].setdefault(k, v)
        for k, v in p['employees'].items():
            results['employees'].setdefault(k, v)
        for k, v in p['financial'].items():
            results['financial'].setdefault(k, v)
        results['directors'].extend(p['directors'])
        results['shareholders'].extend(p['shareholders'])
        results['business_activity'].extend(p['business_activity'])
        results['subsidiaries'].extend(p['subsidiaries'])
        results['raw_texts'].extend(p['raw_texts'])

    return jsonify(results)


def extract_aoc4(text, tables):
    """Extract fields from AOC-4 PDF — returns partial results dict."""
    r = {'company_info': {}, 'auditor': {}}
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
                if val and field not in r['company_info']:
                    r['company_info'][field] = val
                break

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
                r['auditor'][field] = m.group(1).strip()
                break

    return r


def extract_mgt7(text, tables):
    """Extract fields from MGT-7A PDF — returns partial results dict."""
    r = {'company_info': {}, 'employees': {}, 'business_activity': [], 'subsidiaries': [], 'directors': []}
    mgt_patterns = {
        'CIN': [r'([A-Z]{1}[0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6})'],
        'Company Name': [
            r'Name of (?:the )?[Cc]ompany[^\n]*?\n\s*([A-Z][^\n]{3,80})',
            r'([A-Z][A-Z\s&(),.-]{3,80}(?:PRIVATE\s+LIMITED|PVT\.?\s*LTD\.?))',
        ],
        'Registered Address': [
            r'[Rr]egistered office address[^\n]*?\n\s*([^\n]{10,200})',
            r'[Rr]egistered [Oo]ffice[^\n]*?\n\s*([A-Z0-9][^\n]{10,200})',
            r'(?:[Rr]egistered office address)[^\n]*?[\n:]\s*([A-Z0-9][^\n,]{5,}(?:,[^\n]{5,}){2,})',
        ],
        'Latitude': [
            r'[Ll]atitude\s+details?[^\n]*?\n\s*([\d.]+)',
            r'[Ll]atitude[^\n]*?\n\s*([\d.]+)',
            r'[Ll]atitude[^\n]*?(2[0-9]\.[\d]{4,})',
        ],
        'Longitude': [
            r'[Ll]ongitude\s+details?[^\n]*?\n\s*([\d.]+)',
            r'[Ll]ongitude[^\n]*?\n\s*([\d.]+)',
            r'[Ll]ongitude[^\n]*?(7[0-9]\.[\d]{4,})',
        ],
        'PAN': [
            r'[Pp]ermanent [Aa]ccount [Nn]umber[^\n]*?\n\s*([A-Z*]{2,5}[\d*]{4}[A-Z*])',
            r'PAN of the company[^\n]*?\n\s*([A-Z*]{5}[\d*]{4}[A-Z*])',
            r'\bPAN\b[^\n]*?\n\s*([A-Z*]{5}[\d*]{4}[A-Z*])',
        ],
        'AGM Date': [r'AGM[^\n]*?(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})'],
        'SRN': [r'SRN[^\n:]*?:?\s*([A-Z][0-9]{8,})'],
        'Financial Year': [r'[Ff]inancial [Yy]ear[^\n]*?(\d{4}[-–]\d{2,4})'],
    }
    for field, rxs in mgt_patterns.items():
        for rx in rxs:
            m = re.search(rx, text)
            if m:
                val = m.group(1).strip()
                if val and field not in r['company_info']:
                    r['company_info'][field] = val
                break

    emp_patterns = {
        'Total Employees': [r'[Tt]otal[^\n]*?[Ee]mployees?[^\n]*?(\d+)'],
        'Female': [r'\b[Ff]emale\b[^\n]*?(\d+)'],   # Female before Male to avoid substring match
        'Male': [r'\b[Mm]ale\b[^\n]*?(\d+)'],
    }
    for field, rxs in emp_patterns.items():
        for rx in rxs:
            m = re.search(rx, text)
            if m:
                r['employees'][field] = m.group(1).strip()
                break

    for table in tables:
        if not table:
            continue
        for row in table:
            if not row:
                continue
            row_text = ' '.join(str(c) for c in row if c)
            if 'Registered office address' in row_text or 'registered office address' in row_text.lower():
                for col in row[1:]:
                    val = str(col).strip() if col else ''
                    if val and len(val) > 10 and 'Registered Address' not in r['company_info']:
                        r['company_info']['Registered Address'] = val
                        break
            if 'Latitude' in row_text or 'latitude' in row_text.lower():
                for col in row[1:]:
                    val = str(col).strip() if col else ''
                    if val and re.match(r'^\d{1,2}\.\d+$', val) and 'Latitude' not in r['company_info']:
                        r['company_info']['Latitude'] = val
                        break
            if 'Longitude' in row_text or 'longitude' in row_text.lower():
                for col in row[1:]:
                    val = str(col).strip() if col else ''
                    if val and re.match(r'^\d{2,3}\.\d+$', val) and 'Longitude' not in r['company_info']:
                        r['company_info']['Longitude'] = val
                        break

    for table in tables:
        if not table:
            continue
        for row in table:
            if not row:
                continue
            row_text = ' '.join(str(c) for c in row if c)
            if any(k in row_text.upper() for k in ['NIC', 'ACTIVITY', 'TURNOVER', 'BUSINESS']):
                clean = [str(c).strip() if c else '' for c in row]
                if any(c for c in clean):
                    r['business_activity'].append(clean)

    for table in tables:
        if not table:
            continue
        for row in table:
            if not row:
                continue
            row_text = ' '.join(str(c) for c in row if c)
            if any(k in row_text.upper() for k in ['DIN', 'DIRECTOR', 'DESIGNATION']):
                clean = [str(c).strip() if c else '' for c in row]
                if any(c for c in clean):
                    r['directors'].append(clean)

    for table in tables:
        if not table:
            continue
        for row in table:
            if not row:
                continue
            row_text = ' '.join(str(c) for c in row if c)
            if any(k in row_text.upper() for k in ['SUBSIDIARY', 'ASSOCIATE', 'HOLDING']):
                clean = [str(c).strip() if c else '' for c in row]
                if any(c for c in clean):
                    r['subsidiaries'].append(clean)

    return r


def extract_general(text):
    """Extract common fields from any PDF — returns partial results dict."""
    r = {'company_info': {}}
    patterns = {
        'CIN': r'([A-Z]{1}[0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6})',
        'Email': r'([a-zA-Z0-9._%+\-*]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})',
        'Phone': r'(\+?91[\s\-]?[6-9]\d{9}|[6-9]\d{9})',
        'GSTIN': r'([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z])',
        'PAN': r'([A-Z]{5}[0-9]{4}[A-Z])',
    }
    for field, rx in patterns.items():
        m = re.search(rx, text)
        if m and field not in r['company_info']:
            r['company_info'][field] = m.group(1).strip()
    return r


@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    """Convert a scanned PDF to a Word document using OCR (Tesseract)."""
    if not OCR_AVAILABLE:
        return jsonify({'error': 'OCR libraries not installed (pytesseract, pdf2image, python-docx required)'}), 500

    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Please upload a PDF file'}), 400

    pdf_data = file.read()

    try:
        images = convert_from_bytes(pdf_data, dpi=300)
    except Exception as e:
        return jsonify({'error': f'Failed to render PDF pages: {str(e)}'}), 500

    word_doc = Document()

    for i, image in enumerate(images):
        ocr_text = pytesseract.image_to_string(image)

        heading = word_doc.add_paragraph()
        heading_run = heading.add_run(f'Page {i + 1}')
        heading_run.bold = True

        word_doc.add_paragraph(ocr_text)

    output = io.BytesIO()
    word_doc.save(output)
    output.seek(0)

    download_name = file.filename.rsplit('.', 1)[0] + '_ocr.docx'
    return send_file(
        output,
        as_attachment=True,
        download_name=download_name,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


if __name__ == '__main__':
    app.run(debug=True)
