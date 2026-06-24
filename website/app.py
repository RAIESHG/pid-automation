"""
P&ID Automation — Streamlit Web App
GitHub OAuth login gate + full pipeline UI.
"""

import sys
import io
import tempfile
import logging
from pathlib import Path
from collections import defaultdict
from urllib.parse import urlencode

import requests
import streamlit as st
import pandas as pd

# Add parent directory so we can import pid_automation
sys.path.insert(0, str(Path(__file__).parent.parent))
import pid_automation as pia

# ── Page config (must be first Streamlit call) ────────────────────────────────
st.set_page_config(
    page_title="P&ID Automation",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── GitHub OAuth helpers ──────────────────────────────────────────────────────
GITHUB_AUTHORIZE_URL = "https://github.com/login/oauth/authorize"
GITHUB_TOKEN_URL     = "https://github.com/login/oauth/access_token"
GITHUB_USER_URL      = "https://api.github.com/user"


def _oauth_configured() -> bool:
    try:
        _ = st.secrets["github"]["client_id"]
        return True
    except Exception:
        return False


def _get_auth_url() -> str:
    params = {
        "client_id": st.secrets["github"]["client_id"],
        "redirect_uri": st.secrets["github"]["redirect_uri"],
        "scope": "read:user",
    }
    return f"{GITHUB_AUTHORIZE_URL}?{urlencode(params)}"


def _exchange_code(code: str) -> str | None:
    resp = requests.post(
        GITHUB_TOKEN_URL,
        data={
            "client_id": st.secrets["github"]["client_id"],
            "client_secret": st.secrets["github"]["client_secret"],
            "code": code,
            "redirect_uri": st.secrets["github"]["redirect_uri"],
        },
        headers={"Accept": "application/json"},
        timeout=10,
    )
    return resp.json().get("access_token")


def _get_github_user(token: str) -> dict:
    resp = requests.get(
        GITHUB_USER_URL,
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
        timeout=10,
    )
    return resp.json()


# ── Login / OAuth gate ────────────────────────────────────────────────────────
def show_login():
    """Render the login page and handle OAuth callback. Returns only if authenticated."""

    col_l, col_c, col_r = st.columns([1, 2, 1])
    with col_c:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.image(
            "https://github.githubassets.com/images/modules/logos_page/GitHub-Mark.png",
            width=64,
        )
        st.title("P&ID Automation")
        st.caption("Sign in with your GitHub account to continue.")
        st.markdown("---")

        if not _oauth_configured():
            st.warning(
                "GitHub OAuth is not configured yet. "
                "Copy `website/.streamlit/secrets.toml.example` to "
                "`website/.streamlit/secrets.toml` and fill in your GitHub "
                "OAuth App credentials, then restart the app.",
                icon="⚙️",
            )
            with st.expander("Quick setup guide"):
                st.markdown(
                    """
1. Go to **[GitHub → Settings → Developer settings → OAuth Apps](https://github.com/settings/developers)**
2. Click **New OAuth App** and fill in:
   - **Application name:** P&ID Automation
   - **Homepage URL:** `http://localhost:8501`
   - **Authorization callback URL:** `http://localhost:8501`
3. Click **Register application**, copy the **Client ID**, and generate a **Client Secret**.
4. Paste both into `website/.streamlit/secrets.toml` (use the `.example` file as a template).
5. Restart the Streamlit app.
"""
                )
            st.stop()

        # Handle OAuth callback (code in URL query params)
        params = st.query_params
        if "code" in params:
            with st.spinner("Authenticating with GitHub…"):
                token = _exchange_code(params["code"])
            if token:
                user = _get_github_user(token)
                st.session_state["gh_user"] = user
                st.session_state["gh_token"] = token
                st.query_params.clear()
                st.rerun()
            else:
                st.error("GitHub authentication failed — the code may have expired. Try again.")
                if st.button("Retry login"):
                    st.query_params.clear()
                    st.rerun()
                st.stop()

        # Not yet authenticated — show login button
        st.link_button(
            "Login with GitHub",
            _get_auth_url(),
            use_container_width=True,
            type="primary",
        )
        st.stop()


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


# ── Helper ────────────────────────────────────────────────────────────────────
def _save_upload(uploaded, suffix: str) -> Path:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded.getbuffer())
    tmp.flush()
    tmp.close()
    return Path(tmp.name)


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════
if "gh_user" not in st.session_state:
    show_login()

user = st.session_state["gh_user"]
log_handler = _get_log_handler()

# ── Top bar ───────────────────────────────────────────────────────────────────
avatar = user.get("avatar_url", "")
login  = user.get("login", "")
name   = user.get("name") or login

top_l, top_r = st.columns([6, 1])
with top_l:
    st.title("P&ID Automation")
    st.caption("Extract tags from a PDF, validate them, place into a DXF review layer, and export an Excel tag list.")
with top_r:
    if avatar:
        st.image(avatar, width=48)
    st.caption(f"**{name}**")
    if st.button("Logout", use_container_width=True):
        st.session_state.pop("gh_user", None)
        st.session_state.pop("gh_token", None)
        st.rerun()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Inputs")
    pdf_file = st.file_uploader("PDF drawing (text-based)", type=["pdf"])
    dxf_file = st.file_uploader("DXF template (SAVEAS from DWG first)", type=["dxf"])

    st.divider()
    st.header("Tag Patterns")
    inst_pattern = st.text_input(
        "Instrument / Equipment pattern",
        value=pia.DEFAULT_INST_PATTERN,
        help="Python regex that matches instrument and equipment tags.",
    )
    line_pattern = st.text_input(
        "Line number pattern",
        value=pia.DEFAULT_LINE_PATTERN,
        help="Python regex that matches piping line numbers.",
    )

    st.divider()
    st.header("DXF Options")
    text_height  = st.number_input("Text height (drawing units)", value=5.0, min_value=0.5, step=0.5)
    extract_lines = st.checkbox("Copy line geometry to REVIEW_LINES layer")
    symbol_layers = st.text_input("Symbol layers (comma-sep regex)", value="")
    symbol_blocks = st.text_input("Symbol blocks (comma-sep regex)", value="")

    st.divider()
    run_btn = st.button("Run Automation", type="primary", use_container_width=True)

# ── Main area ─────────────────────────────────────────────────────────────────
if not run_btn:
    if not pdf_file or not dxf_file:
        st.info("Upload a **PDF** and a **DXF** file in the sidebar, then click **Run Automation**.")
    else:
        st.success("Files ready — click **Run Automation** in the sidebar.")
    st.stop()

if not pdf_file:
    st.error("Please upload a PDF file.")
    st.stop()
if not dxf_file:
    st.error("Please upload a DXF file.")
    st.stop()

log_handler.clear()

with st.spinner("Saving uploaded files…"):
    pdf_path = _save_upload(pdf_file, ".pdf")
    dxf_path = _save_upload(dxf_file, ".dxf")

# ── Step 1 — Extract ──────────────────────────────────────────────────────────
st.header("Step 1 — Tag Extraction")
with st.spinner("Extracting tags from PDF…"):
    try:
        tags = pia.extract_tags(pdf_path, inst_pattern, line_pattern)
    except Exception as e:
        st.error(f"Extraction failed: {e}")
        st.stop()

if not tags:
    st.error("No tags found. Check that the PDF is text-based and that the patterns match your drawing standard.")
    with st.expander("Log output"):
        st.code("\n".join(log_handler.records))
    st.stop()

tag_df = pd.DataFrame(tags)
c1, c2, c3 = st.columns(3)
c1.metric("Total tags",               len(tag_df))
c2.metric("Instruments / Equipment",  int((tag_df["type"] == "instrument").sum()))
c3.metric("Line numbers",             int((tag_df["type"] == "line").sum()))

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

vdf = pd.DataFrame(tags)
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
        )
    except Exception as e:
        st.error(f"DXF placement failed: {e}")
        st.stop()

c1, c2, c3 = st.columns(3)
c1.metric("Tags placed in DXF",  placed)
c2.metric("Geometry segments",   len(pdf_geometry) if pdf_geometry else 0)
c3.metric("Text elements",       len(pdf_texts)    if pdf_texts    else 0)

# ── Step 4 — Excel Export ─────────────────────────────────────────────────────
st.header("Step 4 — Excel Export")
with st.spinner("Building Excel workbook…"):
    try:
        out_xlsx_path = Path(tempfile.mktemp(suffix="_tag_list.xlsx"))
        pia.export_excel(tags, gaps, out_xlsx_path, lines=lines_data, symbols=symbols_data)
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
    {"Type": k, "Count": len(v), "Tags (preview)": ", ".join(sorted(v)[:8]) + ("…" if len(v) > 8 else "")}
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
        file_name=f"{Path(dxf_file.name).stem}_tags_REVIEW.dxf",
        mime="application/octet-stream",
        use_container_width=True,
    )

if out_xlsx_path.exists():
    dl2.download_button(
        label="Download Excel Tag List",
        data=out_xlsx_path.read_bytes(),
        file_name=f"{Path(dxf_file.name).stem}_tag_list.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── Processing log ────────────────────────────────────────────────────────────
with st.expander("Processing log"):
    st.code("\n".join(log_handler.records) or "(no log output)")

# Cleanup temp input files
try:
    pdf_path.unlink(missing_ok=True)
    dxf_path.unlink(missing_ok=True)
except Exception:
    pass
