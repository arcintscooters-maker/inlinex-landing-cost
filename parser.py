import re
import io
import pdfplumber
import openpyxl


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
    fixed_lines = []
    for line in lines:
        line = re.sub(r'(\S)(\d{13})', r'\1 \2', line)
        fixed_lines.append(line)

    item_re = re.compile(
        r'^(\d+)\s+'
        r'(\S+)\s+'
        r'(.+?)\s+'
        r'(\d{13})\s+'
        r'\d+\s+'
        r'([\d.]+)\s+'
        r'(?:Pair|pc\.|Pack)\s+'
        r'([\d.]+)\s+'
        r'(?:-[\d.]+%)?\s*'
        r'([\d.]+)$'
    )

    # SKU suffix: short alphanumeric continuation on its own line
    sku_cont_re = re.compile(r'^([A-Z0-9]{1,4}(?:-[A-Z0-9]{1,4})*)$')

    # Brand line: ALLCAPS words, not weight/China lines
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
                # Take words before "People's"
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

            # Check next line for SKU suffix
            if i + 1 < len(fixed_lines):
                nxt = fixed_lines[i + 1].strip()
                if sku_cont_re.match(nxt):
                    sku = sku + nxt
                    i += 1

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
        r'^(\d+)\s+(\S+)\s+(.+?)\s+(\d{13})\s+\d+\s+([\d.]+)\s+(?:Pair|pc\.|Pack|PRS)\s+([\d.]+)\s+(?:-[\d.]+%)?\s*([\d.]+)$'
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

    if filename_lower.endswith('.pdf'):
        # Detect supplier from content
        try:
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                first_page = pdf.pages[0].extract_text() or ''
        except:
            first_page = ''

        if 'powerslide' in first_page.lower() or 'IN-' in first_page:
            result = parse_powerslide_pdf(file_bytes)
        else:
            result = parse_generic_pdf(file_bytes)

    elif filename_lower.endswith(('.xlsx', '.xls')):
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
