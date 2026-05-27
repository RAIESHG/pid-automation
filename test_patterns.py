"""Test candidate regex patterns against tags mined from all PDFs."""
import re
from pathlib import Path
from collections import Counter

import pdfplumber

INST_PATTERN = (
    r'(?:\d{4}-)?'                              # optional area prefix: 1000-
    r'(?!PID-|FID-|PFD-)'                       # exclude document references
    r'[A-Z]{1,3}-'                              # service: TK, P, V, HE, X, EH, SU, YC
    r'\d{2,6}'                                  # number: 1016, 612066
    r'[A-Z]?'                                   # optional suffix letter
    r'(?:/[A-Z]){0,3}'                          # optional /B/C
)

LINE_PATTERN = (
    r'(?:[\dX]+(?:/\d+)?"-)?'                   # optional size: 2"-, 1/2"-, X"-
    r'[A-Z]{1,3}-'                              # service
    r'[A-Z0-9X]{3,7}-'                          # area: 1010, 1SA0S01, XXXX
    r'[A-Z0-9X]{3,4}-'                          # seq
    r'[A-Z0-9]{2,4}'                            # spec
    r'(?:-[A-Z]{1,2})?'                         # optional suffix
    r'|'
    r'[A-Z]{1,3}-'                              # reversed variant (rotated text)
    r'[\dX]{3,4}[A-Z]?-'
    r'[\dX]{3,4}-'
    r'[A-Z]{1,2}-'
    r'"[A-Z0-9X]'
)

EXCLUDE_AS_DOC = re.compile(r'\d{4,5}-(?:PID|FID|PFD)-\d{4,5}[A-Z]?')


def classify(word: str):
    word = word.strip().strip(':;,.()[]{}"\'')
    if not word:
        return None
    if EXCLUDE_AS_DOC.fullmatch(word):
        return 'doc_ref'
    if re.fullmatch(LINE_PATTERN, word):
        return 'line'
    if re.fullmatch(INST_PATTERN, word):
        return 'instrument'
    return None


def main():
    pdfs = sorted(p for p in Path('.').glob('*.pdf') if p.name != 'sample.pdf')
    grand = Counter()
    per_file = {}
    for p in pdfs:
        cnt = Counter()
        try:
            with pdfplumber.open(p) as pdf:
                for page in pdf.pages:
                    for w in page.extract_words():
                        cls = classify(w['text'])
                        if cls:
                            cnt[cls] += 1
        except Exception as e:
            cnt['error'] = str(e)
        per_file[p.name] = cnt
        for k, v in cnt.items():
            if isinstance(v, int):
                grand[k] += v

    print(f'{"FILE":<70} {"inst":>6} {"line":>6} {"docref":>6}')
    print('-' * 92)
    for name, cnt in per_file.items():
        print(f'{name[:70]:<70} {cnt.get("instrument", 0):>6} {cnt.get("line", 0):>6} {cnt.get("doc_ref", 0):>6}')
    print('-' * 92)
    print(f'{"TOTAL":<70} {grand["instrument"]:>6} {grand["line"]:>6} {grand["doc_ref"]:>6}')


if __name__ == '__main__':
    main()
