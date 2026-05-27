#!/usr/bin/env python3
"""
P&ID Data Entry Automation
Extract tags from a PDF (with real coordinates) -> validate -> place into a DXF
on a REVIEW layer -> export a tag list + BOM to Excel.

IMPORTANT: ezdxf reads/writes DXF, not DWG. In AutoCAD, SAVEAS -> DXF before
running this, then SAVEAS -> DWG afterwards (or use the ODA File Converter).

Usage:
    python pid_automation.py --pdf input.pdf --dxf drawing.dxf

Tag patterns can be overridden with --inst-pattern and --line-pattern.
"""

import re
import sys
import argparse
import logging
from pathlib import Path
from collections import defaultdict

import pdfplumber
import ezdxf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Default tag patterns -- tuned against SSP Engineering / Motiva drawings and
# the Jet Pack vendor package. Override via CLI for other drawing standards.
#
# Instrument/equipment: optional 4-digit area prefix, an explicit anti-doc-ref
# lookahead (rejects PID-/FID-/PFD-), 1-3 letter service code, 2-6 digit
# number, optional A/B/C suffix, and up to 3 "/B" alternates.
# Matches: TK-1016, P-1001A/B/C, V-1011A/B/C, HE-1113A, X-1111A, F-1001A,
#          B-1001, T-801, EH-900, SU-1019, YC-612066, 1000-B-1015, 1000-U-1015A
# Rejects: 1010-PID-0011, 91062-FID-00026, 0000-PID-00XX, 25MA02, 304L
DEFAULT_INST_PATTERN = (
    r'(?:\d{4}-)?(?!PID-|FID-|PFD-)[A-Z]{1,3}-\d{2,6}[A-Z]?(?:/[A-Z]){0,3}'
)
# Line numbers: handles two variants seen in the wild:
#   forward: 2"-WT-1010-050-A130-N, 1/2"-HC-1SA0S01-1006-NI, VT-1010-057-A520-N
#   reversed: XP-120A-133-0011-BW-"X (pdfplumber reverses words from rotated text)
DEFAULT_LINE_PATTERN = (
    r'(?:[\dX]+(?:/\d+)?"-)?[A-Z]{1,3}-[A-Z0-9X]{3,7}-[A-Z0-9X]{3,4}-[A-Z0-9]{2,4}(?:-[A-Z]{1,2})?'
    r'|[A-Z]{1,3}-[\dX]{3,4}[A-Z]?-[\dX]{3,4}-[A-Z]{1,2}-"[A-Z0-9X]'
)
# Drawing references appear on every sheet's title block ("see drawing X").
# We strip them out so they don't get counted as equipment.
DOC_REF_PATTERN = r'\d{4,5}-(?:PID|FID|PFD)-\d{4,5}[A-Z]?'

# PDF points -> drawing units. Tune to your title block / scale.
# 1.0 means "use PDF point coordinates directly"; adjust once you see the result.
COORD_SCALE = 1.0

# Prefixes that are equipment, not instrument loops. We skip numbering-gap
# detection for these because "P-1001" and "P-1016" are different pumps, not a
# sequence -- reporting 14 "missing" pumps in between is noise. Instrument
# loops like PT-, LI-, FT- ARE sequential and worth checking.
EQUIPMENT_PREFIXES = {'P', 'V', 'TK', 'T', 'E', 'C', 'K', 'D', 'R', 'F', 'S', 'X',
                      'HE', 'EH', 'SU', 'YC', 'B', 'M', 'DA', 'EQ', 'TR'}

# If a tag appears on more than this fraction of pages in a multi-page PDF,
# treat it as a title-block element (drawing number, project ID) and drop it.
TITLE_BLOCK_PAGE_FRACTION = 0.7

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)s  %(message)s',
    datefmt='%H:%M:%S',
)
log = logging.getLogger('pid')


def extract_tags(pdf_path, inst_pattern, line_pattern):
    """Extract tags with real (x, y) positions, one record per matched word.

    pdfplumber gives top-left-origin coordinates (y grows downward). We convert
    to bottom-left-origin (y grows upward) here so placement matches DXF space.
    Line numbers are tested first; a word that matches the line pattern is not
    re-tested as an instrument tag.
    Multi-page PDFs are also de-noised: any tag appearing on >70% of pages is
    treated as title-block boilerplate (drawing number) and dropped.
    """
    log.info('Extracting from %s', pdf_path)
    inst_re = re.compile(inst_pattern)
    line_re = re.compile(line_pattern)
    doc_re = re.compile(DOC_REF_PATTERN)
    tags = []
    skipped_doc = 0

    with pdfplumber.open(pdf_path) as pdf:
        n_pages = len(pdf.pages)
        for page_num, page in enumerate(pdf.pages, 1):
            page_height = page.height  # for the Y flip
            words = page.extract_words()
            if not words:
                log.warning('Page %d: no extractable text (scanned image?)', page_num)
                continue

            for w in words:
                # pdfplumber sometimes glues adjacent punctuation (":VT-1010-057-A520-N",
                # "TK-1016."). Strip punctuation/whitespace so fullmatch can succeed.
                text = w['text'].strip().strip(':;,.()[]{}"\'')
                if not text:
                    continue
                # Drop drawing references (1010-PID-0033, 91062-FID-00026) before
                # they reach the instrument/line classifiers.
                if doc_re.fullmatch(text):
                    skipped_doc += 1
                    continue
                x = w['x0'] * COORD_SCALE
                # flip Y: pdf "top" is distance from page top
                y = (page_height - w['top']) * COORD_SCALE

                if line_re.fullmatch(text):
                    ttype = 'line'
                elif inst_re.fullmatch(text):
                    ttype = 'instrument'
                else:
                    continue

                tags.append({
                    'tag': text, 'type': ttype, 'page': page_num,
                    'x': round(x, 2), 'y': round(y, 2), 'confidence': 0.95,
                })

    # Title-block de-noise: in multi-page docs, tags repeated on most pages are
    # almost certainly drawing numbers stamped in the title block, not real tags.
    if n_pages > 3:
        page_counts = defaultdict(set)
        for t in tags:
            page_counts[t['tag']].add(t['page'])
        threshold = max(2, int(n_pages * TITLE_BLOCK_PAGE_FRACTION))
        boilerplate = {tag for tag, pages in page_counts.items() if len(pages) >= threshold}
        if boilerplate:
            before = len(tags)
            tags = [t for t in tags if t['tag'] not in boilerplate]
            log.info('Dropped %d title-block occurrences across %d unique tags (e.g. %s)',
                     before - len(tags), len(boilerplate), ', '.join(sorted(boilerplate)[:3]))

    log.info('Extracted %d tags across %d page(s) (skipped %d drawing refs)',
             len(tags), n_pages, skipped_doc)
    return tags


def validate(tags):
    """Flag duplicates, format errors, and numbering gaps within each prefix."""
    log.info('Validating %d tags', len(tags))
    seen = set()
    issues = []

    # group instrument tags by alpha prefix to detect numbering gaps. Skip
    # equipment prefixes (P-, V-, TK-, ...) where numbers aren't sequential.
    by_prefix = defaultdict(list)
    for t in tags:
        if t['type'] != 'instrument':
            continue
        m = re.match(r'([A-Z]{1,3})-(\d+)', t['tag'])
        if not m:
            continue
        prefix = m.group(1)
        if prefix in EQUIPMENT_PREFIXES:
            continue
        by_prefix[prefix].append(int(m.group(2)))

    gaps = {}
    for prefix, nums in by_prefix.items():
        nums_sorted = sorted(set(nums))
        if len(nums_sorted) >= 2:
            full = set(range(nums_sorted[0], nums_sorted[-1] + 1))
            missing = sorted(full - set(nums_sorted))
            if missing:
                gaps[prefix] = missing

    for t in tags:
        tag = t['tag']
        status = 'ok'
        if tag in seen:
            status = 'duplicate'
            issues.append(f'duplicate: {tag}')
        else:
            seen.add(tag)
        if len(tag) < 2:
            status = 'format_error'
            issues.append(f'format error: {tag!r}')
        t['status'] = status

    ok = sum(1 for t in tags if t['status'] == 'ok')
    dup = sum(1 for t in tags if t['status'] == 'duplicate')
    err = sum(1 for t in tags if t['status'] == 'format_error')
    log.info('Valid: %d   Duplicates: %d   Format errors: %d', ok, dup, err)

    if gaps:
        for prefix, missing in gaps.items():
            preview = ', '.join(f'{prefix}-{n}' for n in missing[:5])
            more = '' if len(missing) <= 5 else f' (+{len(missing) - 5} more)'
            log.info('Numbering gap in %s-*: missing %s%s', prefix, preview, more)

    if issues:
        for i in issues[:5]:
            log.info('  %s', i)
        if len(issues) > 5:
            log.info('  ... and %d more', len(issues) - 5)

    return tags, gaps


def place_in_dxf(dxf_path, tags, out_path, text_height=2.5):
    """Place each valid tag as MTEXT at its real (x, y) on a REVIEW layer.

    The source drawing is never modified; output is a new DXF.
    """
    log.info('Placing tags into %s', dxf_path)
    try:
        doc = ezdxf.readfile(dxf_path)
    except IOError:
        log.error('Cannot open %s (is it a valid DXF?)', dxf_path)
        return 0
    except ezdxf.DXFStructureError:
        log.error('%s is not a valid DXF. Did you pass a DWG? SAVEAS to DXF first.', dxf_path)
        return 0

    msp = doc.modelspace()
    if 'REVIEW' not in doc.layers:
        doc.layers.new(name='REVIEW', dxfattribs={'color': 1})  # red

    placed = 0
    for t in tags:
        if t['status'] != 'ok':
            continue
        try:
            mtext = msp.add_mtext(t['tag'], dxfattribs={'layer': 'REVIEW'})
            mtext.set_location((t['x'], t['y']))
            mtext.dxf.char_height = text_height
            placed += 1
        except Exception as e:  # keep going; one bad tag shouldn't stop the run
            log.warning('Failed to place %s: %s', t['tag'], e)

    doc.saveas(out_path)
    log.info('Placed %d tags -> %s', placed, out_path)
    return placed


def export_excel(tags, gaps, out_path):
    """Tag list sheet + BOM summary sheet."""
    log.info('Exporting Excel %s', out_path)
    wb = Workbook()
    header_fill = PatternFill('solid', start_color='4472C4', end_color='4472C4')
    header_font = Font(bold=True, color='FFFFFF', name='Arial')
    body_font = Font(name='Arial')

    def style_header(ws):
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

    def autofit(ws, cap=60):
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(width + 2, cap)

    ws = wb.active
    ws.title = 'Tag List'
    ws.append(['Tag', 'Type', 'Page', 'X', 'Y', 'Status'])
    for t in tags:
        ws.append([t['tag'], t['type'], t['page'], t['x'], t['y'], t['status']])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font
    style_header(ws)
    autofit(ws)

    bom = wb.create_sheet('BOM')
    bom.append(['Type', 'Count', 'Tags'])
    grouped = defaultdict(list)
    for t in tags:
        if t['status'] == 'ok':
            grouped[t['type']].append(t['tag'])
    for ttype, names in sorted(grouped.items()):
        bom.append([ttype, len(names), ', '.join(sorted(names))])
    for row in bom.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font
    style_header(bom)
    autofit(bom)

    if gaps:
        gs = wb.create_sheet('Numbering Gaps')
        gs.append(['Prefix', 'Missing'])
        for prefix, missing in gaps.items():
            gs.append([prefix, ', '.join(f'{prefix}-{n}' for n in missing)])
        for row in gs.iter_rows(min_row=2):
            for cell in row:
                cell.font = body_font
        style_header(gs)
        autofit(gs)

    wb.save(out_path)
    log.info('Wrote %s', out_path)


def main():
    ap = argparse.ArgumentParser(description='P&ID data entry automation')
    ap.add_argument('--pdf', required=True, help='input PDF (text-based)')
    ap.add_argument('--dxf', required=True, help='input DXF (SAVEAS from DWG first)')
    ap.add_argument('--inst-pattern', default=DEFAULT_INST_PATTERN)
    ap.add_argument('--line-pattern', default=DEFAULT_LINE_PATTERN)
    ap.add_argument('--text-height', type=float, default=2.5)
    args = ap.parse_args()

    pdf_path, dxf_path = Path(args.pdf), Path(args.dxf)
    if not pdf_path.exists():
        log.error('PDF not found: %s', pdf_path)
        sys.exit(1)
    if not dxf_path.exists():
        log.error('DXF not found: %s', dxf_path)
        sys.exit(1)

    tags = extract_tags(pdf_path, args.inst_pattern, args.line_pattern)
    if not tags:
        log.error('No tags extracted. Check the PDF is text-based and patterns match.')
        sys.exit(1)

    tags, gaps = validate(tags)

    out_dxf = dxf_path.with_name(dxf_path.stem + '_tags_REVIEW.dxf')
    place_in_dxf(dxf_path, tags, out_dxf, text_height=args.text_height)

    out_xlsx = dxf_path.with_name(dxf_path.stem + '_tag_list.xlsx')
    export_excel(tags, gaps, out_xlsx)

    log.info('Done. Open %s in AutoCAD, verify the REVIEW layer, then commit.', out_dxf.name)
    log.info('Excel deliverable: %s', out_xlsx.name)


if __name__ == '__main__':
    main()
