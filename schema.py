"""
schema.py
=========
The single source of truth for the patient health-report data model.

Everything downstream is driven from here:
  * `extraction.py`  populates a dict shaped like `blank_record()`
  * the Streamlit review form renders editable widgets from `SECTIONS`
  * `report_generator.py` reads the same dict to build the Word document

Keeping the model declarative (rather than scattering field names across
the codebase) means a new lab value is added in exactly one place.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any, Dict, List


# --------------------------------------------------------------------------- #
# Field types
# --------------------------------------------------------------------------- #
TEXT = "text"          # free text / numeric value, e.g. "116/78"
CHOICE = "choice"      # single selection from `options`
NUMBER = "number"      # numeric value (kept as string to preserve "20/20")
NOTE = "note"          # long free-text block


# --------------------------------------------------------------------------- #
# Section + field definitions
#
# Each field: (key, label, type, options-or-None, unit-or-None)
# `key` is the storage key; `label` is what the reviewer sees.
# --------------------------------------------------------------------------- #
SECTIONS: List[Dict[str, Any]] = [
    {
        "id": "patient",
        "title": "Patient & Visit",
        "icon": "🪪",
        "fields": [
            ("name", "Patient name", TEXT, None, None),
            ("patient_id", "ID", TEXT, None, None),
            ("exam_date", "Exam date", TEXT, None, None),
            ("sex", "Sex", CHOICE, ["Male", "Female"], None),
            ("provider", "Reviewing provider", TEXT, None, None),
        ],
    },
    {
        "id": "vitals",
        "title": "Vital Signs",
        "icon": "❤️",
        "fields": [
            ("blood_pressure", "Blood pressure", TEXT, None, None),
            ("bp_status", "BP status", CHOICE,
             ["Normal", "Borderline Elevated", "Elevated", "Very high"], None),
            ("pulse", "Pulse", NUMBER, None, "bpm"),
            ("pulse_rhythm", "Rhythm", CHOICE, ["Regular", "Irregular"], None),
            ("pulse_ox", "Pulse Ox", NUMBER, None, "%"),
            ("temperature", "Temperature", NUMBER, None, "°F"),
            ("height_in", "Height", NUMBER, None, "in"),
            ("weight_lbs", "Weight", NUMBER, None, "lbs"),
            ("bmi", "BMI", NUMBER, None, None),
            ("on_bp_medication", "On BP medication", CHOICE, ["No", "Yes"], None),
        ],
    },
    {
        "id": "vision",
        "title": "Vision & Eyes",
        "icon": "👁️",
        "fields": [
            ("far_left", "Far vision — Left (20/_)", TEXT, None, None),
            ("far_right", "Far vision — Right (20/_)", TEXT, None, None),
            ("far_status", "Far vision status", CHOICE,
             ["Normal", "Decreased right", "Decreased left", "Decreased both"], None),
            ("near_left", "Near vision — Left (20/_)", TEXT, None, None),
            ("near_right", "Near vision — Right (20/_)", TEXT, None, None),
            ("near_status", "Near vision status", CHOICE,
             ["Normal", "Decreased right", "Decreased left", "Decreased both"], None),
            ("correction", "Correction", CHOICE, ["Without", "With"], None),
            ("tonometry", "Tonometry (eye pressure)", CHOICE,
             ["Normal", "High left", "High right", "High both"], None),
            ("color_vision", "Color vision", CHOICE,
             ["Normal", "Abnormal", "See below"], None),
        ],
    },
    {
        "id": "hearing",
        "title": "Hearing",
        "icon": "👂",
        "fields": [
            ("hearing_status", "Hearing result", CHOICE,
             ["Normal", "Decreased right", "Decreased left",
              "High frequency hearing loss right", "High frequency hearing loss left",
              "High frequency hearing loss", "Chronic hearing loss", "Bilateral loss"],
             None),
            ("hearing_referral", "Refer to ENT / Audiologist", CHOICE, ["No", "Yes"], None),
        ],
    },
    {
        "id": "cardiac",
        "title": "Cardiac",
        "icon": "🫀",
        "fields": [
            ("ekg", "EKG", CHOICE,
             ["Normal", "Abnormal, see below", "Normal, see below", "See below"], None),
            ("cst", "Cardiac stress test (CST)", CHOICE,
             ["Not performed", "Normal", "Abnormal, see below",
              "Normal, see below", "Negative for ischemia", "See below"], None),
        ],
    },
    {
        "id": "cholesterol",
        "title": "Cholesterol Profile",
        "icon": "🧪",
        "fields": [
            ("total_chol", "Total cholesterol", NUMBER, None, "mg/dl"),
            ("total_chol_flag", "Total flag", CHOICE,
             ["", "Good", "Normal", "Borderline", "Elevated", "High", "Low"], None),
            ("hdl", "HDL", NUMBER, None, "mg/dl"),
            ("hdl_flag", "HDL flag", CHOICE,
             ["", "Good", "Normal", "Borderline", "Elevated", "High", "Low"], None),
            ("ldl", "LDL", NUMBER, None, "mg/dl"),
            ("ldl_flag", "LDL flag", CHOICE,
             ["", "Good", "Normal", "Borderline", "Elevated", "High", "Low"], None),
            ("triglycerides", "Triglycerides", NUMBER, None, "mg/dl"),
            ("trig_flag", "Triglyceride flag", CHOICE,
             ["", "Good", "Normal", "Borderline", "Elevated", "High", "Low"], None),
            ("chol_hdl_ratio", "Cholesterol/HDL ratio", NUMBER, None, None),
            ("chol_hdl_flag", "Ratio flag", CHOICE,
             ["", "All Good", "Good", "Normal", "Borderline", "Elevated", "High"], None),
            ("on_chol_medication", "On cholesterol medication", CHOICE, ["No", "Yes"], None),
        ],
    },
    {
        "id": "glucose",
        "title": "Glucose / Diabetes",
        "icon": "🩸",
        "fields": [
            ("glucose", "Glucose", NUMBER, None, "mg/dl"),
            ("glucose_flag", "Glucose flag", CHOICE,
             ["", "Normal", "Borderline", "Elevated"], None),
            ("a1c", "A1c", NUMBER, None, "%"),
            ("a1c_flag", "A1c flag", CHOICE,
             ["", "Normal", "Borderline", "Elevated", "Pre-diabetes"], None),
        ],
    },
    {
        "id": "cancer",
        "title": "Cancer Screening",
        "icon": "🎗️",
        "fields": [
            ("smoking_status", "Smoking status", CHOICE,
             ["Non-Smoker", "Smoker", "Former Smoker"], None),
            ("colon_screening", "Colon cancer screening", CHOICE,
             ["Up to date", "Recommended", "Recommended at 45 years old", "See below"], None),
            # Male
            ("prostate_exam", "Prostate exam", CHOICE,
             ["", "Normal", "Up to date", "Declined", "N/A", "Other"], None),
            ("psa", "PSA", NUMBER, None, "ng/ml"),
            ("psa_flag", "PSA flag", CHOICE, ["", "Normal", "Elevated", "See below"], None),
            ("prostate_recommendation", "Prostate recommendation", CHOICE,
             ["", "Annual prostate screening", "Up to date", "Recommended at 40",
              "Recommended at 45", "Recommended at 50", "Follow up with Urologist"], None),
            # Female
            ("mammogram", "Mammogram", CHOICE,
             ["", "Normal", "Normal with fibrocystic tissue", "Abnormal, see attached report",
              "Incomplete - Dense breast(s)", "See below"], None),
            ("mammo_recommendation", "Breast screening recommendation", CHOICE,
             ["", "Up to date", "Recommended at 40 years old",
              "Recommended diagnostic mammogram / ultrasound",
              "Dense breast additional images needed", "See below"], None),
            ("pap", "Pap test", CHOICE,
             ["", "Normal", "Normal with HPV", "Abnormal, see below", "See below"], None),
            ("pap_recommendation", "Cervical screening recommendation", CHOICE,
             ["", "Up to date", "Pap smear recommended", "Recommended 2022"], None),
            ("gyn_exam", "Gynecological exam", CHOICE,
             ["", "Up to date", "Recommended", "See below"], None),
            # Both
            ("skin_screening", "Skin cancer screening", CHOICE,
             ["Up to date", "Recommended", "Family hx of skin cancer", "See below"], None),
        ],
    },
    {
        "id": "labs",
        "title": "Laboratory Screening",
        "icon": "🔬",
        "fields": [
            ("cbc", "Complete blood count", CHOICE,
             ["Normal", "Abnormal - see below"], None),
            ("cbc_detail", "CBC detail", TEXT, None, None),
            ("blood_chem", "Blood chemistries", CHOICE,
             ["Normal", "Abnormal - see below"], None),
            ("blood_chem_detail", "Blood chemistry detail", TEXT, None, None),
            ("urinalysis", "Urinalysis", CHOICE,
             ["Normal", "Abnormal - see below"], None),
            ("urinalysis_detail", "Urinalysis detail", TEXT, None, None),
            ("tsh", "TSH", NUMBER, None, None),
            ("t4", "T4", NUMBER, None, None),
            ("b12", "Vitamin B12", NUMBER, None, None),
            ("folate", "Folate", NUMBER, None, None),
            ("vitamin_d", "Vitamin D", NUMBER, None, None),
        ],
    },
    {
        "id": "other",
        "title": "Other Testing",
        "icon": "📋",
        "fields": [
            ("exercise_status", "Exercise status", CHOICE,
             ["Regular exercise", "Limited exercise", "No exercise"], None),
            ("exercise_recommendation", "Exercise recommendation", CHOICE,
             ["", "Continue exercise regimen", "Cardiovascular exercises recommended",
              "Resistance exercises recommended",
              "Cardiovascular and resistance exercises recommended", "See below"], None),
            ("crp", "CRP", NUMBER, None, None),
            ("crp_flag", "CRP flag", CHOICE, ["", "Normal", "Elevated"], None),
            ("pulmonary", "Pulmonary function test", CHOICE,
             ["", "Normal", "Obstructive pattern", "Restrictive pattern",
              "Mild", "Moderate"], None),
            ("chest_xray", "Chest X-ray", CHOICE, ["", "Normal", "Other"], None),
            ("bone_density", "Bone densitometry", CHOICE,
             ["", "Normal", "Up to date", "Recommended", "Other"], None),
            ("recommendations", "Additional notations & recommendations", NOTE, None, None),
        ],
    },
]


# Convenience: ordered (section_id -> list of field-spec) and key -> spec
def field_specs() -> Dict[str, tuple]:
    specs = {}
    for section in SECTIONS:
        for spec in section["fields"]:
            specs[spec[0]] = spec
    return specs


def blank_record() -> Dict[str, Any]:
    """An empty record with every key present (value = "")."""
    rec: Dict[str, Any] = {"_meta": {"source_file": "", "extracted_fields": []}}
    for section in SECTIONS:
        for key, *_ in section["fields"]:
            rec[key] = ""
    return rec


def merge_record(base: Dict[str, Any], updates: Dict[str, Any]) -> Dict[str, Any]:
    """Return base updated with non-empty values from `updates`."""
    out = deepcopy(base)
    for k, v in updates.items():
        if k == "_meta":
            out.setdefault("_meta", {}).update(v or {})
        elif v not in ("", None):
            out[k] = v
    return out


# Fields shown only for a given sex (keeps the review form & report clean).
MALE_ONLY = {"prostate_exam", "psa", "psa_flag", "prostate_recommendation"}
FEMALE_ONLY = {"mammogram", "mammo_recommendation", "pap", "pap_recommendation", "gyn_exam"}


def is_relevant(key: str, sex: str) -> bool:
    sex = (sex or "").strip().lower()
    if sex.startswith("m") and key in FEMALE_ONLY:
        return False
    if sex.startswith("f") and key in MALE_ONLY:
        return False
    return True
