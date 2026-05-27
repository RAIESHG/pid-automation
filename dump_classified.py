"""Show how each candidate is classified, per file."""
import sys
from collections import Counter
from pathlib import Path

import pdfplumber

from test_patterns import classify


def dump(path: Path, limit: int = 40):
    by_class = {'instrument': Counter(), 'line': Counter(), 'doc_ref': Counter(), None: Counter()}
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                for w in page.extract_words():
                    cls = classify(w['text'])
                    t = w['text'].strip().strip(':;,.()[]{}"\'')
                    if t:
                        by_class[cls][t] += 1
    except Exception as e:
        print(f'  ERROR: {e}')
        return

    for cls_name in ('instrument', 'line', 'doc_ref'):
        items = by_class[cls_name]
        if not items:
            continue
        print(f'\n  [{cls_name}] {sum(items.values())} matches across {len(items)} unique:')
        for t, c in items.most_common(limit):
            print(f'    {c:4d}x  {t}')


def main():
    targets = sys.argv[1:] if len(sys.argv) > 1 else [
        '1010-PID-0011-Master.pdf',
        'Equipment List.pdf',
        'Jet Pack P&ID Rev2.pdf',
        'P3ST P&ID.pdf',
    ]
    for name in targets:
        p = Path(name)
        if not p.exists():
            print(f'== MISSING: {name} ==')
            continue
        print(f'\n{"=" * 80}\n=== {name}\n{"=" * 80}')
        dump(p)


if __name__ == '__main__':
    main()
