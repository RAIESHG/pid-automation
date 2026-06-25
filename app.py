"""
P&ID Automation — Streamlit Web App
Upload a PDF + DXF, extract tags, validate, place into DXF review layer,
export Excel tag list, and download results.
"""

import tempfile
import logging
from pathlib import Path
from collections import defaultdict

import streamlit as st
import pandas as pd

import pid_automation as pia
import llm_markup

st.set_page_config(
    page_title="P&ID Automation",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ── Logging bridge ────────────────────────────────────────────────────────────
class _UILogHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        self.records: list[str] = []

    def emit(self, record):
        self.records.append(self.format(record))

    def clear(self):
        self.records.clear()


@st.cache_resource
def _get_log_handler() -> _UILogHandler:
    h = _UILogHandler()
    h.setFormatter(logging.Formatter("%(asctime)s  %(levelname)s  %(message)s", datefmt="%H:%M:%S"))
    pia.log.addHandler(h)
    return h


def _save_upload(uploaded, suffix: str) -> Path:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded.getbuffer())
    tmp.flush()
    tmp.close()
    return Path(tmp.name)


log_handler = _get_log_handler()

# ── Header ────────────────────────────────────────────────────────────────────
st.title("P&ID Automation")
st.caption("Upload a PDF drawing and a DXF template to extract tags, validate, place into a DXF review layer, and export an Excel tag list.")

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Inputs")
    pdf_file = st.file_uploader("PDF drawing (text-based)", type=["pdf"])
    dxf_file = st.file_uploader(
        "DXF template (optional)",
        type=["dxf"],
        help="Leave blank to use the bundled sample.dxf as the template.",
    )
    if not dxf_file:
        st.caption("No file uploaded — using **sample.dxf** (default)")

    st.divider()
    st.header("Tag Patterns")
    inst_pattern = st.text_input("Instrument / Equipment pattern", value=pia.DEFAULT_INST_PATTERN)
    line_pattern  = st.text_input("Line number pattern",           value=pia.DEFAULT_LINE_PATTERN)

    st.divider()
    st.header("DXF Options")
    text_height   = st.number_input("Text height (drawing units)", value=5.0, min_value=0.5, step=0.5)
    extract_lines = st.checkbox("Copy line geometry to REVIEW_LINES layer")
    symbol_layers = st.text_input("Symbol layers (comma-sep regex)", value="")
    symbol_blocks = st.text_input("Symbol blocks (comma-sep regex)", value="")

    st.divider()
    st.header("AI Markup Interpretation")
    ai_enabled = st.checkbox("Interpret markup comments with AI", value=True)
    _secret_key = ""
    try:
        _secret_key = st.secrets.get("OPENROUTER_API_KEY", "")
    except Exception:
        pass
    openrouter_key = st.text_input(
        "OpenRouter API key",
        value=_secret_key,
        type="password",
        help="Get a free key at openrouter.ai — leave blank to skip AI step.",
    )
    ai_model = st.selectbox("Model", llm_markup.AVAILABLE_MODELS)

    st.divider()
    run_btn = st.button("Run Automation", type="primary", use_container_width=True)

# ── Idle state ────────────────────────────────────────────────────────────────
if not run_btn:
    if not pdf_file:
        st.info("Upload a **PDF** file in the sidebar, then click **Run Automation**.")
    else:
        st.success("PDF ready — click **Run Automation** in the sidebar.")
    st.stop()

if not pdf_file:
    st.error("Please upload a PDF file.")
    st.stop()

log_handler.clear()

with st.spinner("Saving uploaded files…"):
    pdf_path = _save_upload(pdf_file, ".pdf")
    if dxf_file:
        dxf_path = Path(_save_upload(dxf_file, ".dxf"))
        dxf_label = Path(dxf_file.name).stem
        _dxf_is_temp = True
    else:
        dxf_path = Path(__file__).parent / "sample.dxf"
        dxf_label = "sample"
        _dxf_is_temp = False

with st.spinner("Scanning PDF annotations…"):
    shx_texts = pia.extract_shx_annotations(pdf_path)
    markup    = pia.extract_markup_annotations(pdf_path)

# ── Annotation Summary ────────────────────────────────────────────────────────
n_shx    = len(shx_texts)
n_clouds = len(markup["revision_clouds"])
n_cmt    = len(markup["comments"])
n_rl     = len(markup["redlines"])
n_stamp  = len(markup["stamps"])

if n_shx or n_clouds or n_cmt or n_rl or n_stamp:
    st.header("Annotation Summary")
    ac1, ac2, ac3, ac4, ac5 = st.columns(5)
    ac1.metric("SHX Text Entries",  n_shx,    help="AutoCAD SHX font text recovered from annotations")
    ac2.metric("Revision Clouds",   n_clouds)
    ac3.metric("Review Comments",   n_cmt)
    ac4.metric("Redlines",          n_rl)
    ac5.metric("Stamps",            n_stamp)

    if n_cmt:
        with st.expander(f"Markup action items ({n_cmt})"):
            action_rows = [
                {
                    "Action":          c.get("action_type", "note").upper(),
                    "Text":            c["text"],
                    "Tags Referenced": ", ".join(e["tag"] for e in c.get("entities", [])),
                    "Page":            c["page"],
                    "Author":          c.get("author", ""),
                }
                for c in markup["comments"]
            ]
            st.dataframe(pd.DataFrame(action_rows), use_container_width=True, hide_index=True)

    if n_shx:
        with st.expander(f"AutoCAD SHX text ({n_shx} entries — previously missing from extraction)"):
            shx_df = pd.DataFrame(shx_texts)[["text", "x", "y", "page"]]
            shx_df.columns = ["Text", "X", "Y", "Page"]
            st.dataframe(shx_df, use_container_width=True, hide_index=True)

    if n_clouds:
        with st.expander(f"Revision clouds ({n_clouds})"):
            cloud_rows = [
                {"Page": c["page"],
                 "Vertices": len(c["vertices"]),
                 "Author": c.get("author", ""),
                 "Bbox": f"({c['bbox'][0]:.0f},{c['bbox'][1]:.0f})→({c['bbox'][2]:.0f},{c['bbox'][3]:.0f})"}
                for c in markup["revision_clouds"]
            ]
            st.dataframe(pd.DataFrame(cloud_rows), use_container_width=True, hide_index=True)

# ── Step 1 — Extract ──────────────────────────────────────────────────────────
st.header("Step 1 — Tag Extraction")
with st.spinner("Extracting tags from PDF…"):
    try:
        tags = pia.extract_tags(pdf_path, inst_pattern, line_pattern)
    except Exception as e:
        st.error(f"Extraction failed: {e}")
        st.stop()

if not tags:
    st.error("No tags found. Check that the PDF is text-based and the patterns match your drawing standard.")
    with st.expander("Log output"):
        st.code("\n".join(log_handler.records))
    st.stop()

tag_df = pd.DataFrame(tags)
c1, c2, c3 = st.columns(3)
c1.metric("Total tags",              len(tag_df))
c2.metric("Instruments / Equipment", int((tag_df["type"] == "instrument").sum()))
c3.metric("Line numbers",            int((tag_df["type"] == "line").sum()))

with st.expander("Raw extracted tags", expanded=False):
    st.dataframe(tag_df, use_container_width=True, hide_index=True)

# ── Step 2 — Validate ─────────────────────────────────────────────────────────
st.header("Step 2 — Validation")
with st.spinner("Validating tags…"):
    try:
        tags, gaps = pia.validate(tags)
    except Exception as e:
        st.error(f"Validation failed: {e}")
        st.stop()

vdf   = pd.DataFrame(tags)
n_ok  = int((vdf["status"] == "ok").sum())
n_dup = int((vdf["status"] == "duplicate").sum())
n_err = int((vdf["status"] == "format_error").sum())

c1, c2, c3 = st.columns(3)
c1.metric("Valid tags",    n_ok)
c2.metric("Duplicates",    n_dup,  delta=f"-{n_dup}" if n_dup else None, delta_color="inverse")
c3.metric("Format errors", n_err,  delta=f"-{n_err}" if n_err else None, delta_color="inverse")

if n_dup or n_err:
    with st.expander("Issues detail"):
        st.dataframe(vdf[vdf["status"] != "ok"], use_container_width=True, hide_index=True)

if gaps:
    with st.expander(f"Numbering gaps ({len(gaps)} prefix(es))"):
        rows = [
            {
                "Prefix": p,
                "Missing": ", ".join(f"{p}-{n}" for n in ms[:10])
                           + (f" … (+{len(ms)-10} more)" if len(ms) > 10 else ""),
                "Count": len(ms),
            }
            for p, ms in gaps.items()
        ]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Step 2.5 — AI Markup Interpretation ──────────────────────────────────────
llm_interpretations: list = []
llm_applied: list         = []
pdf_texts: list           = []  # extracted in Step 3; passed as context here if available

if n_cmt and ai_enabled and openrouter_key:
    st.header("Step 2.5 — AI Markup Interpretation")
    with st.spinner(f"Sending {n_cmt} comment(s) to {ai_model}…"):
        try:
            llm_interpretations = llm_markup.interpret_markup_with_llm(
                markup["comments"],
                tags,
                pdf_texts,
                api_key=openrouter_key,
                model=ai_model,
            )
            st.success(f"Interpreted {len(llm_interpretations)} comment(s).")
        except Exception as _llm_err:
            st.error(f"AI interpretation failed: {_llm_err}")

    if llm_interpretations:
        auto_count   = sum(1 for i in llm_interpretations if i.get("can_automate"))
        manual_count = len(llm_interpretations) - auto_count
        la1, la2 = st.columns(2)
        la1.metric("Auto-applicable",  auto_count)
        la2.metric("Needs manual work", manual_count)

        for interp in llm_interpretations:
            conf     = interp.get("confidence", "?")
            can_auto = interp.get("can_automate", False)
            badge    = "Auto" if can_auto else "Manual"
            preview  = interp.get("interpretation", "")[:70]
            with st.expander(f"#{interp.get('index','?')} [{badge} · {conf}] {preview}"):
                st.write("**Original comment:**", interp.get("original_text", ""))
                st.write("**Interpretation:**",   interp.get("interpretation", ""))
                st.write("**Confidence:**",        conf)
                actions = interp.get("actions", [])
                if actions:
                    st.write("**Planned actions:**")
                    for act in actions:
                        atype  = act.get("type", "")
                        detail = act.get("text") or act.get("description") or act.get("pattern") or ""
                        icon   = {"add_text": "➕", "flag_remove": "🗑️",
                                  "flag_manual": "🔧", "note": "📝"}.get(atype, "•")
                        st.write(f"  {icon} `{atype}`: {detail}")
elif n_cmt and ai_enabled and not openrouter_key:
    st.info("Enter an OpenRouter API key in the sidebar to enable AI markup interpretation.")

# ── Step 3 — DXF Placement ────────────────────────────────────────────────────
st.header("Step 3 — DXF Placement")
with st.spinner("Extracting PDF geometry and placing tags in DXF…"):
    try:
        out_dxf_path = Path(tempfile.mktemp(suffix="_tags_REVIEW.dxf"))
        pdf_geometry = pia.extract_pdf_geometry(pdf_path)
        pdf_texts    = pia.extract_pdf_texts(pdf_path, inst_pattern, line_pattern)

        lines_data: list   = []
        symbols_data: list = []
        if extract_lines:
            lines_data = pia.extract_lines(dxf_path)
        if symbol_layers or symbol_blocks:
            l_pats = [p.strip() for p in symbol_layers.split(",") if p.strip()]
            b_pats = [p.strip() for p in symbol_blocks.split(",") if p.strip()]
            symbols_data = pia.extract_symbols(dxf_path, layer_patterns=l_pats, block_patterns=b_pats)

        placed = pia.place_in_dxf(
            dxf_path, tags, out_dxf_path,
            text_height=text_height,
            extract_lines=extract_lines,
            symbol_layers=symbol_layers,
            symbol_blocks=symbol_blocks,
            pdf_geometry=pdf_geometry,
            pdf_texts=pdf_texts,
            shx_texts=shx_texts,
            markup=markup,
        )
    except Exception as e:
        st.error(f"DXF placement failed: {e}")
        st.stop()

if llm_interpretations:
    with st.spinner("Applying AI actions to DXF…"):
        try:
            llm_applied = llm_markup.apply_llm_actions_to_file(
                out_dxf_path, llm_interpretations, text_height)
        except Exception as _ae:
            st.warning(f"AI action apply warning: {_ae}")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Tags placed in DXF", placed)
c2.metric("Geometry segments",  len(pdf_geometry) if pdf_geometry else 0)
c3.metric("Text elements",      len(pdf_texts)    if pdf_texts    else 0)
c4.metric("SHX annotations",    len(shx_texts))

if llm_applied:
    with st.expander(f"AI actions applied to DXF ({len(llm_applied)})"):
        st.dataframe(pd.DataFrame(llm_applied), use_container_width=True, hide_index=True)

# ── Step 4 — Excel Export ─────────────────────────────────────────────────────
st.header("Step 4 — Excel Export")
with st.spinner("Building Excel workbook…"):
    try:
        out_xlsx_path = Path(tempfile.mktemp(suffix="_tag_list.xlsx"))
        pia.export_excel(tags, gaps, out_xlsx_path, lines=lines_data, symbols=symbols_data,
                         shx_texts=shx_texts, markup=markup,
                         llm_data={"interpretations": llm_interpretations, "applied": llm_applied})
    except Exception as e:
        st.error(f"Excel export failed: {e}")
        st.stop()

st.success("All steps complete!")

# BOM summary
st.subheader("Bill of Materials")
grouped: dict = defaultdict(list)
for t in tags:
    if t["status"] == "ok":
        grouped[t["type"]].append(t["tag"])
bom_rows = [
    {"Type": k, "Count": len(v),
     "Tags (preview)": ", ".join(sorted(v)[:8]) + ("…" if len(v) > 8 else "")}
    for k, v in sorted(grouped.items())
]
if bom_rows:
    st.dataframe(pd.DataFrame(bom_rows), use_container_width=True, hide_index=True)

# Full tag list
st.subheader("Full Tag List")
display_df = vdf[["tag", "type", "page", "x", "y", "status"]].copy()
display_df.columns = ["Tag", "Type", "Page", "X", "Y", "Status"]


def _color(val):
    if val == "ok":
        return "background-color:#d4edda;color:#155724"
    if val == "duplicate":
        return "background-color:#fff3cd;color:#856404"
    return "background-color:#f8d7da;color:#721c24"


st.dataframe(
    display_df.style.map(_color, subset=["Status"]),
    use_container_width=True,
    hide_index=True,
)

# ── Downloads ─────────────────────────────────────────────────────────────────
st.header("Download Results")
dl1, dl2 = st.columns(2)

if out_dxf_path.exists():
    dl1.download_button(
        label="Download DXF (REVIEW layer)",
        data=out_dxf_path.read_bytes(),
        file_name=f"{dxf_label}_tags_REVIEW.dxf",
        mime="application/octet-stream",
        use_container_width=True,
    )

if out_xlsx_path.exists():
    dl2.download_button(
        label="Download Excel Tag List",
        data=out_xlsx_path.read_bytes(),
        file_name=f"{dxf_label}_tag_list.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── Processing log ────────────────────────────────────────────────────────────
with st.expander("Processing log"):
    st.code("\n".join(log_handler.records) or "(no log output)")

# Cleanup temp input files
try:
    pdf_path.unlink(missing_ok=True)
except Exception:
    pass
if _dxf_is_temp:
    try:
        dxf_path.unlink(missing_ok=True)
    except Exception:
        pass
