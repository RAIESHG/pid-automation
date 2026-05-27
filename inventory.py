"""Inventory PDFs: pages, word counts, candidate tags."""
import re
import sys
from collections import Counter
from pathlib import Path

import pdfplumber

CAND_RE = re.compile(r'[A-Z0-9"/\-]{3,}')  # noisy on purpose


def inspect(path: Path):
    info = {'file': path.name, 'pages': 0, 'words': 0, 'samples': [], 'cand': Counter()}
    try:
        with pdfplumber.open(path) as pdf:
            info['pages'] = len(pdf.pages)
            for page in pdf.pages:
                words = page.extract_words()
                info['words'] += len(words)
                for w in words:
                    t = w['text'].strip().strip(':;,.()[]{}"\'')
                    if not t:
                        continue
                    if re.search(r'[A-Z].*\d|\d.*[A-Z]', t) and CAND_RE.fullmatch(t):
                        info['cand'][t] += 1
    except Exception as e:
        info['error'] = str(e)
    return info


def main():
    here = Path('.')
    pdfs = sorted(p for p in here.glob('*.pdf') if p.name != 'sample.pdf')
    for p in pdfs:
        info = inspect(p)
        print('=' * 80)
        print(f"FILE: {info['file']}")
        if 'error' in info:
            print(f"  ERROR: {info['error']}")
            continue
        print(f"  pages: {info['pages']}   words: {info['words']}")
        if info['words'] == 0:
            print('  (no extractable text -- scanned/image PDF, would need OCR)')
            continue
        for tag, count in info['cand'].most_common(40):
            print(f"  {count:4d}x  {tag!r}")


if __name__ == '__main__':
    main()
