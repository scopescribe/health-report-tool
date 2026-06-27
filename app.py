"""
app.py
======
Health Report Assistant — a deterministic macro-sheet → Word report tool.

Flow:   Upload  ▸  Review & Edit  ▸  Generate
No LLM anywhere in the pipeline; extraction is pure pdfplumber / regex and
report generation is pure python-docx.
"""

from __future__ import annotations

import base64
import json
import os
from datetime import date

import streamlit as st

from schema import SECTIONS, is_relevant, merge_record, blank_record, TEXT, NUMBER, CHOICE, NOTE
from extraction import bytes_to_text, extract_record, derive_flags, coverage
from report_generator import generate_report

HERE = os.path.dirname(__file__)
LOGO = os.path.join(HERE, "assets", "ap_logo.jpeg")
SAMPLE = os.path.join(HERE, "sample_data", "sample_filled_male.txt")

st.set_page_config(page_title="Health Report Assistant", page_icon="🩺",
                   layout="wide", initial_sidebar_state="collapsed")


# --------------------------------------------------------------------------- #
# Styling
# --------------------------------------------------------------------------- #
def inject_css():
    st.markdown("""
    <style>
      :root {
        --navy:#1B4D6B; --navy-deep:#143A52; --gold:#C2912A;
        --ink:#1F2A33; --muted:#5B6B78; --line:#E3E9EE;
        --bg-card:#FFFFFF; --good:#1F8A57; --warn:#C2912A;
      }
      .block-container {padding-top:1.4rem; padding-bottom:3rem; max-width:1180px;}
      #MainMenu, footer, header {visibility:hidden;}

      /* Hero */
      .hero {display:flex; align-items:center; gap:18px; padding:6px 2px 2px;}
      .hero h1 {font-size:1.55rem; margin:0; color:var(--ink); font-weight:700; letter-spacing:-.01em;}
      .hero p {margin:2px 0 0; color:var(--muted); font-size:.92rem;}
      .pill {display:inline-block; font-size:.7rem; font-weight:600; letter-spacing:.04em;
             text-transform:uppercase; color:var(--navy); background:#EAF1F6;
             border:1px solid #D5E2EC; border-radius:999px; padding:3px 10px; margin-left:8px;}

      /* Stepper */
      .stepper {display:flex; gap:10px; margin:18px 0 22px;}
      .step {flex:1; display:flex; align-items:center; gap:10px; padding:12px 14px;
             background:var(--bg-card); border:1px solid var(--line); border-radius:12px;}
      .step .num {width:26px; height:26px; border-radius:50%; display:flex; align-items:center;
             justify-content:center; font-size:.8rem; font-weight:700; color:#fff; background:#C7D2DA;}
      .step.active {border-color:var(--navy); box-shadow:0 1px 0 rgba(27,77,107,.12);}
      .step.active .num {background:var(--navy);}
      .step.done .num {background:var(--good);}
      .step .lbl {font-size:.86rem; font-weight:600; color:var(--ink);}
      .step .sub {font-size:.72rem; color:var(--muted);}

      /* Cards */
      .card {background:var(--bg-card); border:1px solid var(--line); border-radius:14px;
             padding:20px 22px; margin-bottom:16px;}
      .section-h {font-size:1.02rem; font-weight:700; color:var(--ink); margin:0 0 2px;}
      .section-sub {font-size:.82rem; color:var(--muted); margin:0 0 8px;}

      /* Coverage banner */
      .cov {display:flex; align-items:center; gap:16px; background:linear-gradient(90deg,#F4F8FB,#FFFFFF);
            border:1px solid var(--line); border-left:4px solid var(--navy);
            border-radius:12px; padding:14px 18px; margin-bottom:6px;}
      .cov .big {font-size:1.7rem; font-weight:800; color:var(--navy); line-height:1;}
      .cov .meta {font-size:.82rem; color:var(--muted);}
      .bar {height:8px; background:#EaEef2; border-radius:99px; overflow:hidden; flex:1; min-width:120px;}
      .bar > span {display:block; height:100%; background:linear-gradient(90deg,var(--navy),#3E7CA0);}

      /* Buttons */
      .stButton>button, .stDownloadButton>button {
        border-radius:10px; font-weight:600; border:1px solid var(--navy);
        background:var(--navy); color:#fff; padding:.5rem 1.1rem;}
      .stButton>button:hover, .stDownloadButton>button:hover {background:var(--navy-deep); border-color:var(--navy-deep); color:#fff;}
      div[data-testid="stFileUploader"] {border:1.5px dashed #C7D6E0; border-radius:14px; padding:8px 12px; background:#FBFDFE;}

      .chip {display:inline-block; font-size:.72rem; padding:2px 9px; border-radius:99px;
             background:#EAF6EF; color:var(--good); border:1px solid #CBE9D8; margin:2px 4px 2px 0;}
      .muted {color:var(--muted); font-size:.84rem;}
      label p {font-weight:600 !important; color:var(--ink) !important;}
    </style>
    """, unsafe_allow_html=True)


def logo_data_uri() -> str:
    if not os.path.exists(LOGO):
        return ""
    b = base64.b64encode(open(LOGO, "rb").read()).decode()
    return f"data:image/jpeg;base64,{b}"


# --------------------------------------------------------------------------- #
# State
# --------------------------------------------------------------------------- #
def init_state():
    st.session_state.setdefault("stage", "upload")   # upload | review
    st.session_state.setdefault("record", None)
    st.session_state.setdefault("source_name", "")
    st.session_state.setdefault("raw_text", "")


def reset():
    for k in list(st.session_state.keys()):
        if k.startswith("w_") or k in ("stage", "record", "source_name", "raw_text", "report_bytes"):
            st.session_state.pop(k, None)
    init_state()


# --------------------------------------------------------------------------- #
# UI pieces
# --------------------------------------------------------------------------- #
def hero():
    uri = logo_data_uri()
    img = f'<img src="{uri}" style="height:42px">' if uri else ""
    st.markdown(f"""
    <div class="hero">
      {img}
      <div>
        <h1>Health Report Assistant <span class="pill">deterministic · no LLM</span></h1>
        <p>Upload a completed macro sheet → review the extracted data → generate a patient-ready report.</p>
      </div>
    </div>""", unsafe_allow_html=True)


def stepper(active: str):
    steps = [("1", "Upload", "PDF · Word · text"),
             ("2", "Review &amp; Edit", "verify extracted values"),
             ("3", "Generate", "download Word report")]
    order = {"upload": 0, "review": 1}
    cur = order.get(active, 0)
    html = '<div class="stepper">'
    for i, (n, lbl, sub) in enumerate(steps):
        cls = "active" if i == cur else ("done" if i < cur else "")
        num = "✓" if i < cur else n
        html += f'<div class="step {cls}"><div class="num">{num}</div><div><div class="lbl">{lbl}</div><div class="sub">{sub}</div></div></div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def coverage_banner(rec):
    filled, total = coverage(rec)
    pct = int(round(100 * filled / total)) if total else 0
    st.markdown(f"""
    <div class="cov">
      <div><div class="big">{pct}%</div></div>
      <div style="min-width:170px"><div class="meta"><b>{filled}</b> of {total} fields captured</div>
        <div class="bar"><span style="width:{pct}%"></span></div></div>
      <div class="meta">Source: <b>{rec['_meta'].get('source_file','—')}</b></div>
    </div>""", unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# Stage 1 — Upload
# --------------------------------------------------------------------------- #
def stage_upload():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-h">Upload a completed macro sheet</div>'
                '<div class="section-sub">Accepts a digitally filled PDF, Word document, or plain-text export. '
                'Checkboxes marked like <code>[x]</code> and typed values are read deterministically.</div>',
                unsafe_allow_html=True)
    up = st.file_uploader("Macro sheet", type=["pdf", "docx", "txt"], label_visibility="collapsed")
    c1, c2 = st.columns([1, 3])
    with c1:
        use_sample = st.button("Try the sample sheet", use_container_width=True)
    with c2:
        st.markdown('<div class="muted" style="padding-top:.55rem">'
                    'No file handy? Load a filled example male macro sheet to see the full flow.</div>',
                    unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    data = name = None
    if up is not None:
        data, name = up.getvalue(), up.name
    elif use_sample:
        data, name = open(SAMPLE, "rb").read(), "sample_filled_male.txt"

    if data is not None:
        with st.spinner("Extracting fields deterministically…"):
            text = bytes_to_text(data, name)
            rec = extract_record(text, name)
        st.session_state.record = rec
        st.session_state.source_name = name
        st.session_state.raw_text = text
        st.session_state.stage = "review"
        st.rerun()


# --------------------------------------------------------------------------- #
# Stage 2 — Review & Edit (+ generate)
# --------------------------------------------------------------------------- #
def _widget(key, label, ftype, options, unit, value):
    lab = f"{label}" + (f" ({unit})" if unit else "")
    if ftype == CHOICE:
        opts = list(options)
        if value and value not in opts:
            opts = [value] + opts
        if "" not in opts:
            opts = [""] + opts
        idx = opts.index(value) if value in opts else 0
        return st.selectbox(lab, opts, index=idx, key=f"w_{key}")
    if ftype == NOTE:
        return st.text_area(lab, value=value, key=f"w_{key}", height=110)
    return st.text_input(lab, value=value, key=f"w_{key}")


def stage_review():
    rec = st.session_state.record
    coverage_banner(rec)

    extracted = set(rec["_meta"].get("extracted_fields", []))
    if extracted:
        chips = "".join(f'<span class="chip">✓ {k}</span>' for k in sorted(extracted)[:18])
        more = f' <span class="muted">+{len(extracted)-18} more</span>' if len(extracted) > 18 else ""
        st.markdown(f'<div style="margin:6px 0 14px">{chips}{more}</div>', unsafe_allow_html=True)

    # Action row
    a1, a2, a3 = st.columns([1.1, 1, 1])
    with a1:
        if st.button("Suggest reference-range flags", use_container_width=True):
            suggestions = derive_flags(rec)
            st.session_state.record = merge_record(rec, suggestions)
            rec = st.session_state.record
            # sync the suggestion into the live widget state (keys persist across reruns)
            for k, v in suggestions.items():
                st.session_state[f"w_{k}"] = v
            st.toast(f"Applied {len(suggestions)} reference-range suggestion(s) to empty flags.")
            st.rerun()
    with a2:
        if st.button("Start over", use_container_width=True):
            reset(); st.rerun()
    with a3:
        st.download_button("Download JSON", data=json.dumps(rec, indent=2),
                           file_name="patient_data.json", mime="application/json",
                           use_container_width=True)

    sex = str(rec.get("sex", "") or "")

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-h">Review &amp; edit extracted data</div>'
                '<div class="section-sub">Confirm or correct each value before generating the report. '
                'Sex-specific screening sections show automatically.</div>', unsafe_allow_html=True)

    tab_labels = [f"{s['icon']}  {s['title']}" for s in SECTIONS]
    tabs = st.tabs(tab_labels)
    edits = {}
    for tab, section in zip(tabs, SECTIONS):
        with tab:
            fields = [f for f in section["fields"] if is_relevant(f[0], sex)]
            # hide detail helper fields unless their parent is abnormal
            cols = st.columns(2)
            for i, (fkey, flabel, ftype, fopts, funit) in enumerate(fields):
                if fkey.endswith("_detail"):
                    parent = fkey.replace("_detail", "")
                    if "abnormal" not in str(rec.get(parent, "")).lower():
                        continue
                with cols[i % 2]:
                    edits[fkey] = _widget(fkey, flabel, ftype, fopts, funit, str(rec.get(fkey, "") or ""))
    st.markdown("</div>", unsafe_allow_html=True)

    # commit edits back into the record
    for k, v in edits.items():
        rec[k] = v
    st.session_state.record = rec

    # ---- Generate -------------------------------------------------------- #
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-h">Generate the report</div>'
                '<div class="section-sub">Produces a formatted Word document on Affiliated Physicians letterhead, '
                'matching the standard health-summary layout.</div>', unsafe_allow_html=True)

    if st.button("📄  Generate Word report", type="primary"):
        with st.spinner("Building the document…"):
            docx_bytes = generate_report(rec)
        st.session_state["report_bytes"] = docx_bytes
        st.success("Report generated.")

    if st.session_state.get("report_bytes"):
        name = (rec.get("name") or "patient").strip().replace(" ", "_")
        fn = f"Health_Report_{name}_{date.today().isoformat()}.docx"
        st.download_button("⬇️  Download report (.docx)", data=st.session_state["report_bytes"],
                           file_name=fn,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.markdown("</div>", unsafe_allow_html=True)

    with st.expander("View extracted JSON"):
        st.json(rec)
    with st.expander("View recovered source text (debug)"):
        st.code(st.session_state.get("raw_text", "")[:6000] or "—")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    inject_css()
    init_state()
    hero()
    stepper(st.session_state.stage)
    if st.session_state.stage == "upload":
        stage_upload()
    else:
        stage_review()


if __name__ == "__main__":
    main()
