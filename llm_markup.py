"""
LLM-powered P&ID markup interpretation via OpenRouter API.

Flow:
  1. Build per-comment context: the comment text + nearby tags/text within a radius
  2. POST everything to an OpenRouter-hosted LLM in a single call
  3. Parse the structured JSON response into interpretation records
  4. apply_llm_actions()  – applies automatable edits to an open ezdxf doc
  5. apply_llm_actions_to_file() – convenience wrapper that opens/saves the DXF
"""

import json
import logging
import re
from typing import Optional

import requests

log = logging.getLogger("pid")

DEFAULT_MODEL = "openrouter/auto"

AVAILABLE_MODELS = [
    "openrouter/auto",                        # Auto — OpenRouter picks best available
    "google/gemini-2.0-flash-exp:free",       # Free tier
    "openai/gpt-4o-mini",
    "openai/gpt-4o",
    "anthropic/claude-3.5-sonnet:beta",
    "anthropic/claude-3-haiku:beta",
    "google/gemini-flash-1.5",
    "meta-llama/llama-3.1-8b-instruct:free",
    "meta-llama/llama-3.3-70b-instruct",
]

_SYSTEM_PROMPT = """You are a P&ID (Piping and Instrumentation Diagram) drawing automation assistant.
You receive markup review comments extracted from a PDF drawing, each with its (x, y) position and a list of nearby drawing elements.
Your task: interpret each comment and return structured JSON actions that a DXF automation tool can execute.

P&ID domain knowledge:
- Instrument tags: PT=pressure transmitter, FT=flow, LT=level, TT=temperature, FIC=flow controller, etc.
- Equipment prefixes: P=pump, V=vessel, E=exchanger, TK=tank, C=compressor, K=blower
- Line numbers format: <size>"-<service>-<seq>-<spec>  e.g. 4"-CWS-1001-A3
- Services: CWS/CWR=chilled water, IA=instrument air, N2=nitrogen, LP/HP=steam, FW=firewater
- Utility connector: off-sheet arrow symbol showing a utility tie-in (steam, CW, IA, etc.)
  shown with service label (e.g. "CWS") above or beside the arrow box
- Off-sheet connectors (OSCs) contain a drawing number + sheet ref inside the box
- Connection numbers are the reference numbers inside OSC boxes

Automatable action types (DXF can do these):
  add_text       – place a new text label. Needs: text, x, y, layer
  flag_remove    – mark existing text for drafter removal. Needs: pattern, layer_hint
  flag_manual    – needs drafter/symbol work. Needs: description, x, y
  note           – informational only. Needs: description

Return ONLY a JSON object { "interpretations": [...] }. No prose."""


def _parse_llm_json(text: str) -> dict:
    """
    Robustly extract a JSON object from an LLM response.
    Handles markdown fences, prose preambles, and truncated output.
    """
    # Strip code fences
    text = re.sub(r"^```(?:json)?\s*", "", text.strip()).rstrip("`").strip()

    # Try direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Find the first {...} block
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    # JSON is truncated — find the last complete interpretation object and close the structure
    # Locate the opening of "interpretations" array
    start = text.find('"interpretations"')
    if start != -1:
        arr_start = text.find("[", start)
        if arr_start != -1:
            fragment = text[arr_start:]
            # Count complete objects by finding balanced braces
            depth, last_close, i = 0, arr_start, arr_start
            objects = []
            obj_start = None
            for i, ch in enumerate(fragment):
                if ch == "{":
                    if depth == 0:
                        obj_start = i
                    depth += 1
                elif ch == "}":
                    depth -= 1
                    if depth == 0 and obj_start is not None:
                        try:
                            objects.append(json.loads(fragment[obj_start: i + 1]))
                        except json.JSONDecodeError:
                            pass
                        obj_start = None
            if objects:
                log.warning("LLM response was truncated — recovered %d of ~%d interpretation(s)",
                            len(objects), text.count('"index"'))
                return {"interpretations": objects}

    raise json.JSONDecodeError("Could not extract valid JSON from LLM response", text, 0)


def _nearby(tags, pdf_texts, cx, cy, radius=250):
    """Return drawing elements within radius of (cx, cy), capped at 35."""
    out = []
    for t in (tags or []):
        if abs(t.get("x", 0) - cx) < radius and abs(t.get("y", 0) - cy) < radius:
            out.append({"kind": "tag", "tag": t["tag"], "type": t["type"],
                        "x": round(t["x"], 1), "y": round(t["y"], 1)})
    for t in (pdf_texts or []):
        if abs(t.get("x", 0) - cx) < radius and abs(t.get("y", 0) - cy) < radius:
            out.append({"kind": "text", "text": t["text"], "layer": t.get("layer", ""),
                        "x": round(t["x"], 1), "y": round(t["y"], 1)})
    return out[:35]


def interpret_markup_with_llm(
    comments: list,
    tags: list,
    pdf_texts: list,
    api_key: str,
    model: str = DEFAULT_MODEL,
) -> list:
    """
    Send markup comments to the LLM and return a list of interpretation dicts.
    Each dict has: index, original_text, interpretation, confidence,
                   can_automate, actions[]
    Raises requests.HTTPError on API failure.
    """
    if not comments:
        return []

    blocks = []
    for i, c in enumerate(comments, 1):
        blocks.append({
            "index": i,
            "text": c["text"],
            "position": {"x": round(c["x"], 1), "y": round(c["y"], 1)},
            "author": c.get("author", ""),
            "nearby_drawing_elements": _nearby(tags, pdf_texts, c["x"], c["y"]),
        })

    user_msg = (
        "Interpret the following P&ID markup comments and return edit instructions.\n\n"
        "MARKUP COMMENTS:\n"
        + json.dumps(blocks, indent=2)
        + "\n\n"
        "For each comment return an object with:\n"
        "  index           (int)   – same as above\n"
        "  original_text   (str)   – verbatim comment\n"
        "  interpretation  (str)   – plain-English summary of the required change\n"
        "  confidence      (str)   – high / medium / low\n"
        "  can_automate    (bool)  – true only if fully doable without a symbol library\n"
        "  actions         (array) – list of DXF edit actions (see system prompt for types)\n\n"
        "For add_text: include text, x, y (use nearby element positions as reference), layer\n"
        "For flag_remove: include pattern (substring to match), layer_hint\n"
        "For flag_manual: include description, x, y\n"
        "Return ONLY { \"interpretations\": [...] }"
    )

    resp = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "X-Title": "P&ID Automation",
        },
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": _SYSTEM_PROMPT},
                {"role": "user",   "content": user_msg},
            ],
            "temperature": 0.1,
            "max_tokens": 4096,
        },
        timeout=120,
    )
    if not resp.ok:
        body = ""
        try:
            body = resp.json().get("error", {}).get("message", resp.text[:400])
        except Exception:
            body = resp.text[:400]
        raise requests.HTTPError(
            f"{resp.status_code} from OpenRouter — {body}",
            response=resp,
        )

    raw = resp.json()["choices"][0]["message"]["content"].strip()
    parsed = _parse_llm_json(raw)
    interps = parsed.get("interpretations", parsed if isinstance(parsed, list) else [])
    log.info("LLM returned %d interpretation(s) for %d comment(s)", len(interps), len(comments))
    return interps


def apply_llm_actions(doc, interpretations: list, text_height: float = 5.0) -> list:
    """
    Apply automatable LLM actions to an open ezdxf document.

    Layers created:
      MARKUP_LLM_ADDS      (magenta/6)  – text added by LLM instruction
      MARKUP_FLAG_REMOVE   (red/1)      – existing text flagged for drafter removal
      MARKUP_MANUAL_NEEDED (orange/40)  – items needing manual symbol/geometry work

    Returns a list of applied-action summary dicts.
    """
    msp = doc.modelspace()
    applied = []

    def _layer(name, color):
        if name not in doc.layers:
            doc.layers.new(name=name, dxfattribs={"color": color})

    for interp in interpretations:
        idx = interp.get("index", "?")
        for action in interp.get("actions", []):
            atype = action.get("type", "")
            try:
                if atype == "add_text":
                    layer = action.get("layer", "MARKUP_LLM_ADDS")
                    _layer(layer, 6)
                    mt = msp.add_mtext(action["text"], dxfattribs={"layer": layer})
                    mt.set_location((float(action["x"]), float(action["y"])))
                    mt.dxf.char_height = text_height
                    applied.append({
                        "comment": idx, "type": "add_text",
                        "detail": f'Added "{action["text"]}" at ({action["x"]:.0f},{action["y"]:.0f}) on {layer}',
                    })

                elif atype == "flag_remove":
                    _layer("MARKUP_FLAG_REMOVE", 1)
                    pattern = action.get("pattern", "").lower()
                    layer_hint = action.get("layer_hint", "")
                    flagged = 0
                    for ent in list(msp.query("MTEXT TEXT")):
                        try:
                            ent_text = ent.text if hasattr(ent, "text") else ent.dxf.text
                        except Exception:
                            continue
                        if pattern and pattern in ent_text.lower():
                            if not layer_hint or layer_hint in ent.dxf.layer:
                                ent.dxf.layer = "MARKUP_FLAG_REMOVE"
                                flagged += 1
                    applied.append({
                        "comment": idx, "type": "flag_remove",
                        "detail": f'Flagged {flagged} element(s) matching "{pattern}" → MARKUP_FLAG_REMOVE (red)',
                    })

                elif atype == "flag_manual":
                    _layer("MARKUP_MANUAL_NEEDED", 40)
                    desc = action.get("description", interp.get("interpretation", ""))
                    if "x" in action and "y" in action:
                        mt = msp.add_mtext(
                            f"[MANUAL: {desc}]",
                            dxfattribs={"layer": "MARKUP_MANUAL_NEEDED"})
                        mt.set_location((float(action["x"]), float(action["y"])))
                        mt.dxf.char_height = max(2.0, text_height * 0.8)
                    applied.append({
                        "comment": idx, "type": "flag_manual",
                        "detail": f"Manual action needed: {desc}",
                    })

                elif atype == "note":
                    applied.append({
                        "comment": idx, "type": "note",
                        "detail": action.get("description", ""),
                    })

            except Exception as e:
                log.warning("Failed to apply action %s (comment %s): %s", atype, idx, e)

    log.info("Applied %d LLM action(s) from %d interpretation(s)", len(applied), len(interpretations))
    return applied


def apply_llm_actions_to_file(dxf_path, interpretations: list, text_height: float = 5.0) -> list:
    """Open a saved DXF, apply LLM actions in-place, save. Returns applied-action list."""
    import ezdxf
    doc = ezdxf.readfile(str(dxf_path))
    applied = apply_llm_actions(doc, interpretations, text_height)
    doc.saveas(str(dxf_path))
    return applied
