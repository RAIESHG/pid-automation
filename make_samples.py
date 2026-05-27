"""Generate tiny sample PDF + DXF inputs to demo pid_automation.py."""
from pathlib import Path

import ezdxf
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4


def make_pdf(path: Path) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    width, height = A4
    c.setFont('Helvetica', 12)
    tags = [
        ('PT-101', 100, height - 100),
        ('PT-102', 100, height - 140),
        ('LI-1001', 250, height - 100),
        ('LI-1003', 250, height - 140),  # gap: LI-1002 missing
        ('FT-201', 400, height - 100),
        ('PT-101', 400, height - 140),  # duplicate
        ('6"-P-1001-A1A', 100, height - 220),
        ('4"-S-2003-B2B', 250, height - 220),
    ]
    for text, x, y in tags:
        c.drawString(x, y, text)
    c.save()


def make_dxf(path: Path) -> None:
    doc = ezdxf.new(dxfversion='R2010')
    msp = doc.modelspace()
    msp.add_line((0, 0), (500, 0))
    msp.add_line((0, 0), (0, 500))
    msp.add_text('SAMPLE TITLE BLOCK', dxfattribs={'height': 5}).set_placement((20, 20))
    doc.saveas(str(path))


if __name__ == '__main__':
    here = Path(__file__).parent
    make_pdf(here / 'sample.pdf')
    make_dxf(here / 'sample.dxf')
    print('Wrote sample.pdf and sample.dxf')
