# P&ID Data Entry Automation

Extract instrument/equipment tags and line numbers from text-based P&ID PDFs, validate them, place tags on a DXF `REVIEW` layer, and export Excel tag lists + BOM.

Tuned for SSP Engineering / Motiva-style drawings and Jet Pack vendor P&IDs.

## Setup (office laptop)

```bash
git clone https://github.com/RAIESHG/pid-automation.git
cd pid-automation
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

**Important:** Convert DWG to DXF in AutoCAD (`SAVEAS` → DXF) before running. Output can be saved back to DWG afterward.

```bash
python pid_automation.py --pdf "1010-PID-0011-Master.pdf" --dxf "your-drawing.dxf"
```

Outputs (next to the input DXF name):

- `*_tags_REVIEW.dxf` — tags on red `REVIEW` layer
- `*_tag_list.xlsx` — Tag List, BOM, Numbering Gaps sheets

Optional overrides:

```bash
python pid_automation.py --pdf drawing.pdf --dxf drawing.dxf \
  --inst-pattern '...' --line-pattern '...' --text-height 2.5
```

## Batch run all PDFs in folder

```bash
python run_all.py
```

Writes Excel/DXF under `out/`.

## Helper scripts

| Script | Purpose |
|--------|---------|
| `inventory.py` | Page/word counts and tag candidates per PDF |
| `inspect_pdf.py` | Quick text dump from one PDF |
| `render_pages.py` | PNG previews for visual review |
| `test_patterns.py` | Pattern match counts across all PDFs |
| `dump_classified.py` | Show how tags are classified |
| `run_all.py` | Batch process every PDF |

## Notes

- **P3ST** line numbers use a different format (size embedded mid-string); not captured by default patterns yet.
- Scanned/image-only PDFs need OCR (no extractable text).
- `COORD_SCALE` in `pid_automation.py` may need tuning for DXF placement vs PDF coordinates.
