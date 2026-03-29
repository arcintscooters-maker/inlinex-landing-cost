import re
import io
import pdfplumber
import openpyxl
import traceback


# ─── Powerslide PDF Parser ────────────────────────────────────────────────────

def parse_powerslide_pdf(file_bytes):
    lines = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.split('\n'))

    invoice_no = ''
    supplier = 'Powerslide'
    invoice_total = 0.0
    items = []

    for line in lines:
        m = re.search(r'Invoice\s+(IN-[\w]+)', line)
        if m and not invoice_no:
            invoice_no = m.group(1)
        m = re.search(r'Total sum\s+([\d,]+\.\d+)', line)
        if m:
            invoice_total = float(m.group(1).replace(',', ''))

    # Fix lines where EAN is glued to description (no space before 13-digit number)
    # Handle case where digits before EAN create a longer number (e.g. "3x1254040333593240")
    # by specifically splitting on known EAN prefixes (4040333 for Powerslide, 693633 for Flying Eagle)
    fixed_lines = []
    for line in lines:
        # First, split known EAN prefixes that may be glued to preceding text/digits
        line = re.sub(r'(\S)(4040333\d{6})', r'\1 \2', line)
        line = re.sub(r'(\S)(693633\d{7})', r'\1 \2', line)
        # Fallback: generic 13-digit EAN split (non-digit followed by 13 digits)
        line = re.sub(r'([^\d\s])(\d{13})(?=\s)', r'\1 \2', line)
        fixed_lines.append(line)

    # Matches standard lines with optional discount column
    # Also handles: no-discount lines, -left/-right SKU suffixes
    item_re = re.compile(
        r'^(\d+)\s+'         # pos
        r'(\S+)\s+'          # sku (may end with -left, -right, or be partial)
        r'(.+?)\s+'           # description
        r'(\d{13})\s+'       # EAN
        r'\d+\s+'            # tariff
        r'([\d.]+)\s+'       # qty
        r'(?:Pair|pc\.|Pack|Set)\s+'  # unit type
        r'([\d.]+)\s+'       # unit price
        r'(?:-[\d.]+%)?\s*'  # optional discount
        r'([\d.]+)$'          # total
    )

    # SKU suffix continuation: e.g. '34', 'XL', '40-43', 'right'
    sku_cont_re = re.compile(r'^([0-9]{1,4}|[A-Z]{1,5}|[0-9]{1,2}-[0-9]{1,2})$')

    def extract_brand(lines, start, window=5):
        skip_patterns = [
            r'^Net item', r'^Gross item', r"Republic of China",
            r'^SITZ', r'^AMTSGERICHT', r'^UST', r'^Gesch',
            r'^Es gelten', r'HYPOVEREINSBANK', r'SPARKASSE',
            r'SWIFT', r'iBanFirst', r'^Page:', r'^Pos\.', r'^Invoice'
        ]
        for j in range(start, min(start + window, len(lines))):
            l = lines[j].strip()
            if not l:
                continue
            if any(re.search(p, l) for p in skip_patterns):
                continue
            words = l.split()
            if words and words[0].isupper() and len(words[0]) > 2:
                brand_words = []
                for w in words:
                    if w == "People's":
                        break
                    brand_words.append(w)
                return ' '.join(brand_words)
        return ''

    i = 0
    while i < len(fixed_lines):
        line = fixed_lines[i].strip()
        m = item_re.match(line)
        if m:
            pos = int(m.group(1))
            sku = m.group(2)
            desc = m.group(3).strip().rstrip(',')
            ean = m.group(4)
            qty = float(m.group(5))
            unit_usd = float(m.group(6))
            total_usd = float(m.group(7))

            # Check next line for SKU suffix (e.g. '34', 'XL')
            if i + 1 < len(fixed_lines):
                nxt = fixed_lines[i + 1].strip()
                if sku_cont_re.match(nxt) and not re.match(r'^\d{2}\s+', nxt):
                    sku = sku + nxt
                    i += 1
                # Handle split SKU where first part ends with dash e.g. '908050-' + 'right ...'
                elif sku.endswith('-'):
                    first_word = nxt.split()[0] if nxt else ''
                    if first_word and first_word.isalpha() and len(first_word) <= 8:
                        sku = sku + first_word

            brand = extract_brand(fixed_lines, i + 1)

            items.append({
                'pos': pos,
                'sku': sku,
                'ean': ean,
                'description': desc,
                'brand': brand,
                'qty': int(qty),
                'unit_usd': unit_usd,
                'total_usd': total_usd,
            })
        i += 1

    # If total not found via regex, sum items
    if invoice_total == 0 and items:
        invoice_total = round(sum(item['total_usd'] for item in items), 2)

    return {
        'invoice_no': invoice_no,
        'supplier': supplier,
        'invoice_total_usd': invoice_total,
        'items': items,
        'notes': ''
    }


# ─── Flying Eagle Excel Parser ────────────────────────────────────────────────

def parse_flying_eagle_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    invoice_no = ''
    supplier = 'Flying Eagle'
    invoice_total = 0.0
    items = []
    notes_list = []

    # Read all rows as flat values
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append([str(c).strip() if c is not None else '' for c in row])

    # Extract invoice number and date
    for row in rows:
        for cell in row:
            m = re.search(r'(SAC-\d+|INV[-\s]?\d+)', cell, re.IGNORECASE)
            if m and not invoice_no:
                invoice_no = m.group(1).upper()

    # Find TOTAL row
    for row in rows:
        joined = ' '.join(row)
        m = re.search(r'TOTAL[:\s]*([\d,]+\.?\d*)', joined, re.IGNORECASE)
        if m:
            invoice_total = float(m.group(1).replace(',', ''))
            break

    # Detect item rows: row where col 0 is a number (pos), and there's a USD amount
    # Flying Eagle format: pos | model_name_row above | size/color | ... | qty | unit | USD | amount | notes
    current_model = ''
    pos_counter = 0

    for idx, row in enumerate(rows):
        # Check if this row has a position number in first non-empty col
        first = row[0] if row[0] else ''
        if re.match(r'^\d+$', first):
            pos = int(first)
            # Get description from this row and model from row above if blank pos
            desc_parts = [c for c in row[1:6] if c and c not in ('None', '')]
            desc = ' '.join(desc_parts).strip()

            # Find qty, unit price, amount
            qty = 1
            unit_usd = 0.0
            total_usd = 0.0
            note = ''

            # Scan row for numbers
            nums = []
            for c in row:
                try:
                    v = float(c)
                    nums.append(v)
                except:
                    pass

            # Check for note in last columns
            for c in row[7:]:
                if c and not re.match(r'^[\d.]+$', c) and len(c) > 5:
                    note = c
                    break

            # Flying Eagle pattern: qty=1 PRS, unit price, USD label, amount
            # Try to find: qty | PRS | USD | price | USD | amount
            row_joined = ' '.join(row)

            # Extract qty
            m_qty = re.search(r'(\d+)\s*(?:PRS|prs|pairs?)', row_joined, re.IGNORECASE)
            if m_qty:
                qty = int(m_qty.group(1))

            # Extract unit and total from numeric columns
            # Usually format: qty=1, unit=46, total=46
            usd_vals = []
            for c in row:
                try:
                    v = float(c)
                    if v > 0:
                        usd_vals.append(v)
                except:
                    pass

            if len(usd_vals) >= 2:
                unit_usd = usd_vals[-2]
                total_usd = usd_vals[-1]
            elif len(usd_vals) == 1:
                unit_usd = usd_vals[0]
                total_usd = usd_vals[0]

            # Get model name from the row above if it looks like a model header
            if idx > 0:
                prev_row = rows[idx - 1]
                prev_joined = ' '.join(prev_row).strip()
                if prev_joined and not re.match(r'^\d+', prev_joined) and not any(
                    kw in prev_joined.lower() for kw in ['invoice', 'shipped', 'item', 'description', 'price', 'amount', 'qty', 'tel', 'fax', 'email', 'bank', 'swift', 'beneficiary']
                ):
                    model_candidate = ' '.join(c for c in prev_row if c).strip()
                    if model_candidate and len(model_candidate) < 30:
                        current_model = model_candidate

            # Build full description
            full_desc = current_model
            # Add size/color info from this row
            for c in row[1:6]:
                if c and c not in ('None', '', 'PRS', 'USD') and not re.match(r'^[\d.]+$', c):
                    if c.lower() not in full_desc.lower():
                        full_desc = full_desc + ' ' + c if full_desc else c

            full_desc = full_desc.strip()

            if total_usd > 0:
                items.append({
                    'pos': pos,
                    'sku': f"{current_model.replace(' ', '-')}-{desc}".replace(' ', '-').upper()[:30],
                    'ean': '',
                    'description': full_desc if full_desc else desc,
                    'brand': 'Flying Eagle',
                    'qty': qty if qty > 0 else 1,
                    'unit_usd': unit_usd,
                    'total_usd': total_usd,
                })
                if note:
                    notes_list.append(f"Item {pos}: {note}")

        # Track model header rows (non-numbered rows that look like model names)
        elif row[0] == '' and row[1]:
            candidate = ' '.join(c for c in row if c).strip()
            if candidate and len(candidate) < 25 and not any(
                kw in candidate.lower() for kw in ['invoice', 'shipped', 'item', 'description', 'price', 'tel', 'fax', 'bank', 'vvvv', 'buyer', 'seller']
            ):
                current_model = candidate

    if invoice_total == 0 and items:
        invoice_total = round(sum(i['total_usd'] for i in items), 2)

    return {
        'invoice_no': invoice_no,
        'supplier': supplier,
        'invoice_total_usd': invoice_total,
        'items': items,
        'notes': '; '.join(notes_list)
    }



# ─── Universkate / FR Skates PDF Parser ──────────────────────────────────────

def parse_universkate_pdf(file_bytes):
    lines = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.split('\n'))

    invoice_no = ''
    supplier = 'Universkate / FR Skates'
    invoice_total = 0.0
    items = []

    for line in lines:
        m = re.search(r'PROFORMA\s+(\d+)', line)
        if m and not invoice_no:
            invoice_no = 'PRO-' + m.group(1)
        m = re.search(r'Net\s+à\s+payer\s+€\s+([\d\s,]+)', line)
        if m:
            invoice_total = float(m.group(1).strip().replace(' ','').replace(',','.'))

    ean_re = re.compile(r'^(\S+)\s+(\d{13})\s+(.+)$')

    for line in lines:
        line = line.strip()
        m = ean_re.match(line)
        if not m:
            continue
        sku = m.group(1)
        if sku.startswith('SIRET') or sku.startswith('EORI'):
            continue
        ean = m.group(2)
        rest = m.group(3).strip()

        tokens = rest.split()
        if len(tokens) < 3:
            continue

        total_str = tokens[-1]
        unit_str = tokens[-2]
        qty_str = tokens[-3]

        if ',' not in total_str or ',' not in unit_str:
            continue
        if not qty_str.isdigit():
            continue

        desc = ' '.join(tokens[:-3]).strip()
        qty = int(qty_str)
        unit_eur = float(unit_str.replace(' ','').replace(',','.'))
        total_eur = float(total_str.replace(' ','').replace(',','.'))

        if desc.startswith('FR -'):
            brand = 'FR Skates'
        elif 'INTUITION' in desc:
            brand = 'Intuition'
        else:
            brand = 'Universkate'

        items.append({
            'pos': len(items) + 1,
            'sku': sku,
            'ean': ean,
            'description': desc,
            'brand': brand,
            'qty': qty,
            'unit_usd': unit_eur,   # EUR stored as usd field (user pays SGD anyway)
            'total_usd': total_eur,
        })

    if invoice_total == 0 and items:
        invoice_total = round(sum(i['total_usd'] for i in items), 2)

    return {
        'invoice_no': invoice_no,
        'supplier': supplier,
        'invoice_total_usd': invoice_total,
        'items': items,
        'notes': 'Prices are in EUR. Landing cost calculated from your SGD payment.'
    }

# ─── Flying Eagle PDF Parser (OCR-based) ─────────────────────────────────────

def _is_cid_pdf(file_bytes):
    """Check if a PDF has CID-encoded (unextractable) text."""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        text = pdf.pages[0].extract_text() or ''
    # CID PDFs produce (cid:N) patterns instead of real text
    return '(cid:' in text

def _decode_cid_eans(file_bytes):
    """Extract EANs from CID-encoded PDFs by decoding CID character mappings.
    Flying Eagle PDFs encode only EAN barcodes as text; everything else is vector graphics."""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        chars = pdf.pages[0].chars

    # Group characters by y-position (each row = one EAN)
    rows = {}
    for c in chars:
        y = round(c['top'], 0)
        if y not in rows:
            rows[y] = []
        rows[y].append(c['text'])

    # Build CID-to-digit mapping from known Flying Eagle EAN prefix 6936335
    # All Flying Eagle EANs start with 6936335, so first 7 chars give us the mapping
    cid_map = {}
    first_row = rows[min(rows.keys())] if rows else []
    if len(first_row) >= 7:
        expected = '6936335'
        for i, ch in enumerate(first_row[:7]):
            m = re.match(r'\(cid:(\d+)\)', ch)
            if m:
                cid_map[int(m.group(1))] = expected[i]

    # Collect all CID values used across all rows
    all_cids = set()
    for row_chars in rows.values():
        for ch in row_chars:
            m = re.match(r'\(cid:(\d+)\)', ch)
            if m:
                all_cids.add(int(m.group(1)))

    # Resolve unknown CIDs using EAN-13 check digit validation
    # Try all permutations of missing digits for unknown CIDs
    unknown_cids = sorted([c for c in all_cids if c not in cid_map])
    known_digits = set(cid_map.values())
    missing_digits = [str(d) for d in range(10) if str(d) not in known_digits]

    def ean13_check(digits_str):
        """Validate EAN-13 check digit."""
        if len(digits_str) != 13 or not digits_str.isdigit():
            return False
        total = sum(int(digits_str[i]) * (1 if i % 2 == 0 else 3) for i in range(12))
        check = (10 - (total % 10)) % 10
        return check == int(digits_str[12])

    def decode_rows_with_map(cmap):
        """Decode all rows and check if all produce valid EAN-13."""
        decoded = []
        for y in sorted(rows.keys()):
            ean = ''
            for ch in rows[y]:
                m_cid = re.match(r'\(cid:(\d+)\)', ch)
                if m_cid:
                    cid = int(m_cid.group(1))
                    ean += cmap.get(cid, '?')
                else:
                    ean += ch
            decoded.append(ean)
        return decoded

    if unknown_cids and missing_digits:
        from itertools import permutations
        best_map = None
        for perm in permutations(missing_digits, len(unknown_cids)):
            test_map = dict(cid_map)
            for cid_val, digit in zip(unknown_cids, perm):
                test_map[cid_val] = digit
            decoded = decode_rows_with_map(test_map)
            if all(ean13_check(e) for e in decoded):
                cid_map = test_map
                best_map = test_map
                break
        if best_map is None:
            # Fallback: if no valid permutation, try assigning sorted
            for cid_val, digit in zip(unknown_cids, missing_digits):
                cid_map[cid_val] = digit

    # Decode all EANs
    eans = []
    for y in sorted(rows.keys()):
        decoded = ''
        for ch in rows[y]:
            m = re.match(r'\(cid:(\d+)\)', ch)
            if m:
                cid = int(m.group(1))
                decoded += cid_map.get(cid, '?')
            else:
                decoded += ch
        if len(decoded) == 13 and decoded.isdigit():
            eans.append(decoded)

    return eans


def parse_flying_eagle_pdf(file_bytes):
    """Parse Flying Eagle PDF invoices using OCR (with CID EAN fallback)."""
    import fitz  # PyMuPDF
    try:
        import pytesseract
        from PIL import Image
        has_tesseract = True
    except ImportError:
        has_tesseract = False

    # Step 1: Try to decode EANs from CID data (always works for FE PDFs)
    cid_eans = _decode_cid_eans(file_bytes)

    # Step 2: OCR the PDF pages for descriptions, quantities, prices
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    ocr_lines = []
    for page in doc:
        pix = page.get_pixmap(dpi=300)
        if has_tesseract:
            try:
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                text = pytesseract.image_to_string(img)
                ocr_lines.extend(text.split('\n'))
            except Exception:
                has_tesseract = False  # tesseract binary not actually available
        if not has_tesseract:
            # Fallback: use PyMuPDF's built-in OCR if available
            try:
                text = page.get_textpage_ocr(dpi=300, full=True).extractText()
                ocr_lines.extend(text.split('\n'))
            except Exception:
                pass

    invoice_no = ''
    supplier = 'Flying Eagle'
    invoice_total = 0.0
    items = []

    # Extract invoice number
    for line in ocr_lines:
        m = re.search(r'(SAC[-—]?\d+)', line, re.IGNORECASE)
        if m and not invoice_no:
            invoice_no = m.group(1).upper().replace('—', '-')
        m = re.search(r'Amount\s*\(USD\)[:\s]*\$?([\d,]+\.?\d*)', line, re.IGNORECASE)
        if m:
            invoice_total = float(m.group(1).replace(',', ''))
        m = re.search(r'Total.*?\$\s*([\d,]+\.\d+)', line)
        if m:
            invoice_total = float(m.group(1).replace(',', ''))

    # Parse item rows from OCR text
    # Flying Eagle PDF format: # | Description | EAN | Qty | Unit | Price | Amount
    # The EAN column may be garbled in OCR, so we use CID-decoded EANs by position
    item_re = re.compile(
        r'^\s*(\d{1,2})\s+'           # position number
        r'(.+?)\s+'                     # description
        r'(\d{13})?\s*'                # EAN (may be garbled)
        r'(\d+)\s+'                     # qty
        r'(?:pair|prs?|pc|set)\w*\s+'  # unit
        r'\$?\s*([\d.]+)\s+'           # unit price
        r'\$?\s*([\d.]+)',             # total
        re.IGNORECASE
    )
    # Simpler pattern: pos, description, qty, pair, price, amount
    simple_re = re.compile(
        r'^\s*(\d{1,2})\s+'           # pos
        r'(.+?)\s+'                    # description
        r'(\d+)\s+'                    # qty
        r'(?:pair|prs?|Pair|pc)\s+'   # unit
        r'\$?\s*([\d.]+)\s+'           # unit price
        r'\$?\s*([\d.]+)',            # total
    )

    ean_idx = 0
    for line in ocr_lines:
        line = line.strip()
        if not line:
            continue

        m = item_re.match(line)
        if not m:
            m = simple_re.match(line)
            if m:
                pos = int(m.group(1))
                desc = m.group(2).strip()
                qty = int(m.group(3))
                unit_usd = float(m.group(4))
                total_usd = float(m.group(5))
                ean = cid_eans[ean_idx] if ean_idx < len(cid_eans) else ''
                ean_idx += 1
            else:
                continue
        else:
            pos = int(m.group(1))
            desc = m.group(2).strip()
            ean_ocr = m.group(3) or ''
            qty = int(m.group(4))
            unit_usd = float(m.group(5))
            total_usd = float(m.group(6))
            # Prefer CID-decoded EAN over OCR (more reliable)
            ean = cid_eans[ean_idx] if ean_idx < len(cid_eans) else ean_ocr
            ean_idx += 1

        # Clean description: remove EAN-like numbers and trailing size info already captured
        desc = re.sub(r'\d{13}', '', desc).strip()
        desc = re.sub(r'\s+', ' ', desc)

        items.append({
            'pos': pos,
            'sku': '',
            'ean': ean,
            'description': desc,
            'brand': 'Flying Eagle',
            'qty': qty,
            'unit_usd': unit_usd,
            'total_usd': total_usd,
        })

    # If OCR failed but we have CID EANs, create stub items
    if not items and cid_eans:
        for idx, ean in enumerate(cid_eans):
            items.append({
                'pos': idx + 1,
                'sku': '',
                'ean': ean,
                'description': f'Flying Eagle Item {idx + 1} (OCR unavailable)',
                'brand': 'Flying Eagle',
                'qty': 1,
                'unit_usd': 0.0,
                'total_usd': 0.0,
            })

    if invoice_total == 0 and items:
        invoice_total = round(sum(i['total_usd'] for i in items), 2)

    notes = ''
    if not has_tesseract and items and items[0]['total_usd'] == 0:
        notes = 'OCR not available - only EANs extracted. Prices and quantities need manual entry.'

    return {
        'invoice_no': invoice_no,
        'supplier': supplier,
        'invoice_total_usd': invoice_total,
        'items': items,
        'notes': notes
    }


# ─── Generic fallback (tries PDF line parsing) ───────────────────────────────

def parse_generic_pdf(file_bytes):
    """Generic PDF parser - extracts lines with numbers that look like invoice rows."""
    lines = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.split('\n'))

    # Fix glued EANs
    fixed = [re.sub(r'(\S)(\d{13})', r'\1 \2', l) for l in lines]

    invoice_no = ''
    invoice_total = 0.0
    items = []

    for line in fixed:
        m = re.search(r'(?:Invoice|INV|PI)[#\s:]*([A-Z0-9-]{4,20})', line, re.IGNORECASE)
        if m and not invoice_no:
            invoice_no = m.group(1)
        m = re.search(r'(?:Total|TOTAL)[:\s]+([\d,]+\.\d+)', line)
        if m:
            invoice_total = float(m.group(1).replace(',', ''))

    item_re = re.compile(
        r'^(\d+)\s+(\S+)\s+(.+?)\s+(\d{13})\s+\d+\s+([\d.]+)\s+(?:Pair|pc\.|Pack|Set|PRS)\s+([\d.]+)\s+(?:-[\d.]+%)?\s*([\d.]+)$'
    )

    i = 0
    while i < len(fixed):
        m = item_re.match(fixed[i].strip())
        if m:
            sku = m.group(2)
            if i + 1 < len(fixed):
                nxt = fixed[i+1].strip()
                if re.match(r'^[A-Z0-9]{1,6}$', nxt):
                    sku += nxt
                    i += 1
            items.append({
                'pos': int(m.group(1)),
                'sku': sku,
                'ean': m.group(4),
                'description': m.group(3).strip(),
                'brand': '',
                'qty': int(float(m.group(5))),
                'unit_usd': float(m.group(6)),
                'total_usd': float(m.group(7)),
            })
        i += 1

    if invoice_total == 0 and items:
        invoice_total = round(sum(i['total_usd'] for i in items), 2)

    return {
        'invoice_no': invoice_no,
        'supplier': 'Unknown',
        'invoice_total_usd': invoice_total,
        'items': items,
        'notes': ''
    }


# ─── Main dispatcher ──────────────────────────────────────────────────────────

def parse_invoice(filename, file_bytes):
    filename_lower = filename.lower()

    # Detect by magic bytes (reliable regardless of filename/mimetype)
    is_pdf = file_bytes[:4] == b'%PDF'
    is_xlsx = file_bytes[:4] == b'PK'

    if is_pdf or filename_lower.endswith('.pdf'):
        try:
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                first_page = pdf.pages[0].extract_text() or ''
        except Exception as e:
            raise ValueError(f"Could not read PDF: {e}")

        # Detect Flying Eagle PDFs: CID-encoded fonts with 6936335 EAN prefix
        is_flying_eagle_pdf = _is_cid_pdf(file_bytes) and '(cid:' in first_page
        # Also detect by filename hint
        if not is_flying_eagle_pdf and 'sac' in filename_lower:
            is_flying_eagle_pdf = _is_cid_pdf(file_bytes)

        if 'powerslide' in first_page.lower() or 'IN-' in first_page:
            result = parse_powerslide_pdf(file_bytes)
        elif 'universkate' in first_page.lower() or 'PROFORMA' in first_page:
            result = parse_universkate_pdf(file_bytes)
        elif is_flying_eagle_pdf:
            result = parse_flying_eagle_pdf(file_bytes)
        else:
            result = parse_generic_pdf(file_bytes)

    elif is_xlsx or filename_lower.endswith(('.xlsx', '.xls')):
        result = parse_flying_eagle_excel(file_bytes)
    else:
        raise ValueError("Unsupported file type. Please upload PDF or Excel (.xlsx).")

    if not result['items']:
        raise ValueError("No line items could be extracted from this invoice. The format may not be supported yet.")

    return result


# ─── Quick test ───────────────────────────────────────────────────────────────
if __name__ == '__main__':
    import json

    print("=== Powerslide PDF ===")
    with open('/mnt/user-data/uploads/I-IN-2602611-Custno_-43993.pdf', 'rb') as f:
        result = parse_invoice('I-IN-2602611.pdf', f.read())
    print(f"Invoice: {result['invoice_no']}, Items: {len(result['items'])}, Total: {result['invoice_total_usd']}")
    for item in result['items']:
        print(f"  {item['pos']:2}. {item['sku']:<25} qty={item['qty']} total={item['total_usd']}")

    print()
    print("=== Flying Eagle Excel ===")
    with open('/mnt/user-data/uploads/Copy_of_SAC-026_Adam__PI_Singapore.xlsx', 'rb') as f:
        result2 = parse_invoice('SAC-026.xlsx', f.read())
    print(f"Invoice: {result2['invoice_no']}, Items: {len(result2['items'])}, Total: {result2['invoice_total_usd']}")
    for item in result2['items']:
        print(f"  {item['pos']:2}. {item['sku']:<35} qty={item['qty']} total={item['total_usd']}")

