"""Render selected PDF pages to PNG previews for vision review."""
from pathlib import Path

import pypdfium2 as pdfium

TARGETS = [
    ('1010-PID-0011-Master.pdf', [1]),
    ('1100-PID-0013-Model.pdf', [1]),
    ('1000-PID-0001 (gs 10-28-25)[dft 11-3](dh 11-06-25)[dft 11-7] (gs 11-09-25)[dft 11-11].pdf', [1, 2]),
    ('Jet Pack P&ID Rev2.pdf', [1, 2]),
    ('P3ST P&ID.pdf', [5]),
    ('_P&ID Cover Sheet.pdf', [1]),
    ('Equipment List.pdf', [1]),
]

out = Path('previews')
out.mkdir(exist_ok=True)

for fname, pages in TARGETS:
    p = Path(fname)
    if not p.exists():
        print(f'skip (missing): {fname}')
        continue
    pdf = pdfium.PdfDocument(p)
    for page_num in pages:
        if page_num > len(pdf):
            continue
        page = pdf[page_num - 1]
        image = page.render(scale=1.5).to_pil()
        safe = ''.join(c if c.isalnum() or c in '-_.' else '_' for c in p.stem)
        outpath = out / f'{safe}_p{page_num}.png'
        image.save(outpath, optimize=True)
        print(f'wrote {outpath} ({image.size[0]}x{image.size[1]})')
    pdf.close()
