# Health Report Assistant

A deterministic web tool that ingests a completed patient **macro sheet**
(PDF, Word, or text), extracts the data into structured JSON, and generates a
formatted **Word health-summary report** on Affiliated Physicians letterhead.

There is **no LLM anywhere in the pipeline.** Extraction is pure
`pdfplumber` + regex; report generation is pure `python-docx`. Every value is
auditable and reproducible.

```
   Upload                 Extract                Review              Generate
 ┌─────────┐   text   ┌──────────────┐  JSON  ┌──────────┐  edits  ┌───────────┐
 │ PDF /   │ ───────▶ │ deterministic│ ─────▶ │  human   │ ──────▶ │  .docx    │
 │ DOCX /  │          │  parser      │        │  review  │         │  report   │
 │ TXT     │          │ (no LLM)     │        │  & edit  │         │           │
 └─────────┘          └──────────────┘        └──────────┘         └───────────┘
```

---

## Why no LLM

The macro sheet is a fixed-layout form. Selections are encoded as marked
checkboxes (`[x]`) and typed values in labelled blanks. That structure is fully
recoverable with deterministic text parsing, so an LLM would only add cost,
latency, and non-determinism. A human review step (built in) handles the rare
ambiguous field instead.

---

## Architecture

| File | Responsibility |
|------|----------------|
| `schema.py` | Single source of truth — every field, type, option, and the sex-specific section rules. Drives extraction, the review UI, and the report. |
| `extraction.py` | `bytes → text → record`. Checkbox tokenisation + label-anchored regex. Also a transparent, rule-based reference-range flag *suggester*. |
| `report_generator.py` | `record → .docx`. Reproduces the letterhead, cover letter, two-column health summary, labs, and recommendations. |
| `app.py` | Streamlit UI: a 3-step Upload → Review → Generate workflow. |
| `assets/ap_logo.jpeg` | Letterhead logo. |
| `sample_data/` | A filled example macro sheet in `.txt`, `.docx`, and `.pdf`. |

### How extraction works

1. **Text recovery** — `pdfplumber` for PDF, `python-docx` for DOCX, direct
   decode for text.
2. **Checkbox tokenisation** — a line like `[ ] Normal [x] Elevated` is split
   into `(checked, label)` pairs; the checked label is matched to the field's
   known options.
3. **Label-anchored regex** — numeric blanks (`Total Cholesterol 187`) are
   captured by a regex anchored on the field label.
4. Everything the parser is confident about is recorded in
   `record["_meta"]["extracted_fields"]` so the UI can show coverage.

### Reference-range suggestions (optional, off by default)

`extraction.derive_flags()` proposes conventional adult flags (e.g. LDL < 100 →
"Good") **only for numeric fields that have a value but no flag**. These are
surfaced as suggestions the reviewer can accept or override — never applied
silently. They are deterministic thresholds, not medical advice.

---

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open `http://localhost:8501`, then click **Try the sample sheet** to see the
full flow, or upload your own completed macro sheet.

### Deploying to Streamlit Cloud

Point it at this repo. No secrets and no system packages are required
(`packages.txt` is intentionally empty).

---

## Input expectations

The deterministic parser reads **digitally completed** forms: typed values and
checkboxes marked with a glyph inside the brackets (`[x]`, `[X]`, `[✓]`, …). An
empty `[ ]` is treated as unchecked.

Scanned or handwritten forms would need an OCR front-end (`pdf2image` +
`tesseract`); the architecture leaves a clean seam for that — only
`extraction.bytes_to_text` would change. The review step already covers
correction of any field OCR gets wrong.

---

## Extending the schema

Add a lab value in **one** place — `schema.py`:

```python
("vitamin_b6", "Vitamin B6", NUMBER, None, None),
```

The review form renders it automatically. Add a matching `setv(...)` line in
`extraction.py` to parse it, and a line in `report_generator.build_*` to place
it in the report.
