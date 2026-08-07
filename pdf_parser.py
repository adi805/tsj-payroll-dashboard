"""
MZTS/Phase 2: PDF Parser for TSJ Closingan
Extracts structured data from Berita Acara PDFs (denda, premi, kekurangan gaji, dll).
Works with pdfplumber (pure Python, no system deps).

Usage:
    python3 pdf_parser.py                                        # all PDFs in folder
    python3 pdf_parser.py "path/to/file.pdf"                     # single file
    python3 pdf_parser.py "path/to/folder" --output results.json

Output: JSON array of {filename, type, date, people[], amounts[], raw_text}
"""
import sys, os, re, json
import pdfplumber

# ---- PDF type detection ----
TYPE_PATTERNS = [
    (r'DENDA\s+ANCAN|ANCAN\s+DENDA',       'DENDA_ANCAN_PANEN'),
    (r'DENDA\s+SEMPROT|SEMPROT\s+DENDA',    'DENDA_SEMPROT'),
    (r'DENDA\s+TPH|TPH\s+DENDA',            'DENDA_TPH'),
    (r'DENDA\s+SENSUS|SENSUS\s+DENDA',      'DENDA_SENSUS'),
    (r'DENDA\s+APEL|APEL\s+DENDA',          'DENDA_APEL'),
    (r'KURANG\s+GAJI|KEKURANGAN\s+GAJI',    'KEKURANGAN_GAJI'),
    (r'LEBIH\s+GAJI|KELBIHAN\s+GAJI',       'LEBIH_GAJI'),
    (r'PEMANEN\s+SAKIT|SAKIT\s+PEMANEN',    'PEMANEN_SAKIT'),
    (r'PREMI\s+PERAWATAN|PERAWATAN\s+PREMI','PREMI_PERAWATAN'),
    (r'PREMI\s+MUAT|MUAT\s+PREMI',          'PREMI_MUAT_TBS'),
    (r'PREMI\s+MENGGANTIKAN|MENGGANTIKAN\s+PREMI','PREMI_MENGGANTIKAN_KCS'),
    (r'PREMI\s+KR\s+AFDELING|AFDELING\s+KR','PREMI_KR_AFDELING'),
    (r'PREMI\s+PENGGANTI|PENGGANTI\s+PREMI','PREMI_PENGGANTI_MANDOR'),
    (r'LANGSIR\s+ALONG|LANGSIR',             'LANGSIR_ALONG'),
    (r'REVISI\s+BA|BA\s+REVISI',             'REVISI_BA'),
    (r'BA\s+MENGGANTIKAN|MENGGANTIKAN\s+SHIFT|SHIFT\s+SECURITY','BA_SHIFF_SECURITY'),
    (r'SPL\s+LEMBUR|LEMBUR\s+SPL',           'SURAT_PERINTAH_LEMBUR'),
    (r'PENGGALIHAN\s+KERJA|KERJA\s+PENGGALIHAN','PENGGALIHAN_KERJA'),
]

def detect_type(filename, text):
    name = os.path.splitext(os.path.basename(filename))[0].upper()
    combined = name + ' ' + text[:500].upper()
    for pattern, label in TYPE_PATTERNS:
        if re.search(pattern, combined, re.IGNORECASE):
            return label
    return 'LAINNYA'

# ---- Date extraction ----
DATE_PATTERNS = [
    r'\d{1,2}\s+(?:Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+\d{4}',
    r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}',
    r'\d{4}-\d{2}-\d{2}',
]
BULAN = {
    'januari':'01','februari':'02','maret':'03','april':'04','mei':'05','juni':'06',
    'juli':'07','agustus':'08','september':'09','oktober':'10','november':'11','desember':'12'
}

def extract_date(text):
    for pat in DATE_PATTERNS:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            found = m.group()
            for bulan, num in BULAN.items():
                if bulan in found.lower():
                    dm = re.match(r'(\d{1,2})\s+' + bulan + r'\s+(\d{4})', found, re.IGNORECASE)
                    if dm:
                        return f"{dm.group(2)}-{num}-{dm.group(1).zfill(2)}"
            return found
    return None

# ---- Money extraction (3 patterns for OCR-compatibility) ----
def extract_amounts(text):
    amounts = []
    seen = set()

    # Pattern 1: Rp with comma decimals "Rp 15.000,00" or "Rp15.000,00"
    for m in re.finditer(r'Rp\.?\s*([\d\.]{4,}(?:,\d{2})?)', text, re.IGNORECASE):
        val = m.group(1).replace('.','').replace(',','.')
        try:
            v = float(val)
            if v not in seen and v >= 1000:
                seen.add(v); amounts.append(v)
        except ValueError:
            pass

    # Pattern 2: dot-as-thousands sep "15.000" (no decimals) — skip if preceded by Rp
    for m in re.finditer(r'(\d{1,3}\.\d{3}(?:\.\d{3})*)', text):
        prefix_start = max(0, m.start() - 4)
        prefix = text[prefix_start:m.start()]
        # Skip if preceded by Rp (avoid double-extracting Rp amounts)
        if re.search(r'Rp\.?\s+$', prefix, re.IGNORECASE):
            continue
        val = m.group(1).replace('.','')
        try:
            v = float(val)
            if v not in seen and v >= 1000:
                seen.add(v); amounts.append(v)
        except ValueError:
            pass

    # Pattern 3: any standalone 4+ digit number (fallback for OCR-garbled docs)
    for m in re.finditer(r'(?<![.\d])\b(\d{4,})\b', text):
        val = m.group(1)
        try:
            v = float(val)
            if v not in seen and 1000 <= v <= 999999999:
                seen.add(v); amounts.append(v)
        except ValueError:
            pass

    return sorted(set(amounts), reverse=True)[:10]

# ---- Person extraction ----
def extract_names(text, top_n=5):
    """Extract capitalized names (common Indonesian names pattern)."""
    names = []
    for m in re.finditer(r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})\b', text):
        n = m.group(1).strip()
        if len(n) > 4 and not re.match(r'^(PT|TSJ|Kebun|Afdeling|Juli|Agustus|Pdf|Tgl|No|Den)\w*$', n):
            names.append(n)
    seen = set()
    unique = []
    for n in names:
        if n not in seen:
            seen.add(n); unique.append(n)
    return unique[:top_n]

# ---- NIK extraction ----
NIK_PATTERN = re.compile(r'\b(\d{7,10})\b')
def extract_niks(text):
    nik_raw = NIK_PATTERN.findall(text)
    return sorted(set(nik_raw))

# ---- Main parser ----
def parse_pdf(path):
    filename = os.path.basename(path)
    try:
        with pdfplumber.open(path) as pdf:
            pages_text = []
            for page in pdf.pages:
                t = page.extract_text() or ''
                if t.strip():
                    pages_text.append(t)
    except Exception as e:
        return {'filename': filename, 'error': str(e)}

    full_text = '\n'.join(pages_text)
    return {
        'filename': filename,
        'path': path,
        'type': detect_type(filename, full_text),
        'date': extract_date(full_text),
        'people': extract_names(full_text),
        'amounts': extract_amounts(full_text),
        'niks': extract_niks(full_text),
        'pages': len(pages_text),
        'chars': len(full_text),
        'preview': full_text[:300].replace('\n', ' '),
    }

def parse_folder(folder):
    results = []
    folder = folder.rstrip('/\\')
    for fname in sorted(os.listdir(folder)):
        if fname.lower().endswith('.pdf'):
            fpath = os.path.join(folder, fname)
            result = parse_pdf(fpath)
            results.append(result)
    return results

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='TSJ PDF Parser Phase 2')
    parser.add_argument('path', nargs='?', default=None,
                        help='PDF file or folder path')
    parser.add_argument('--output', '-o', default=None,
                        help='Output JSON file (default: print to stdout)')
    args = parser.parse_args()

    if not args.path:
        SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
        BASE = os.path.join(SCRIPT_DIR, 'closingan-juli-2026')
        if not os.path.isdir(BASE):
            BASE = os.path.join(SCRIPT_DIR, 'pendukung')
        print(f"[PDF Parser] Scanning: {BASE}", flush=True)
        results = parse_folder(BASE)
    elif os.path.isdir(args.path):
        print(f"[PDF Parser] Scanning folder: {args.path}", flush=True)
        results = parse_folder(args.path)
    elif os.path.isfile(args.path):
        print(f"[PDF Parser] Single file: {args.path}", flush=True)
        results = [parse_pdf(args.path)]
    else:
        print(f"ERROR: path not found: {args.path}")
        sys.exit(1)

    # Print summary
    for r in results:
        if 'error' in r:
            print(f"  ERROR {r['filename']}: {r['error']}")
        else:
            print(f"  OK {r['type']:30s} | {r['date'] or '?':10s} | "
                  f"{', '.join(r['people'][:2]) or '?'} | "
                  f"Rp {r['amounts'][0] if r['amounts'] else '?':>12} | "
                  f"{r['filename']}")

    total = len(results)
    ok = sum(1 for r in results if 'error' not in r)
    print(f"\nPDF Parser: {ok}/{total} parsed OK", flush=True)

    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"JSON saved: {args.output}")
    elif len(results) == 1:
        print(json.dumps(results[0], ensure_ascii=False, indent=2))
