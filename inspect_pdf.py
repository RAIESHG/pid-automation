"""Inspect a PDF: counts words, shows a sample, and finds candidate tag patterns."""
import argparse
import re
from collections import Counter

import pdfplumber


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pdf')
    ap.add_argument('--show', type=int, default=80, help='words to print')
    args = ap.parse_args()

    with pdfplumber.open(args.pdf) as pdf:
        print(f'pages: {len(pdf.pages)}')
        all_words = []
        for i, page in enumerate(pdf.pages, 1):
            words = page.extract_words()
            print(f'  page {i}: {len(words)} words, size={page.width:.0f}x{page.height:.0f}')
            all_words.extend(w['text'] for w in words)

        print(f'\ntotal words: {len(all_words)}')
        if not all_words:
            print('PDF appears to be image-only (no extractable text). OCR is needed.')
            return

        print(f'\nfirst {args.show} words:')
        for w in all_words[:args.show]:
            print(f'  {w!r}')

        candidates = [w for w in all_words if re.search(r'[A-Z].*\d|\d.*[A-Z]', w)]
        print(f'\nalphanumeric candidates ({len(candidates)}):')
        for w, c in Counter(candidates).most_common(60):
            print(f'  {c:3d}x  {w!r}')


if __name__ == '__main__':
    main()
