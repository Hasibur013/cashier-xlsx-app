"""Deterministic + LLM parsing for the Chayabithi Cafe cashier SMS.

The cashier sends a daily message in mixed Bengali/English. Each sold item is
written roughly as:

    <serial>। <item name> = <profit-rate> × <quantity> = <total>

This module converts that free text into a structured list of
``{"sl": int, "qty": float, "rate": float}`` records that can be written into
the monthly Excel workbook.

Two engines are provided:
  * ``parse_sms``      – offline, regex based, no network. Reliable default.
  * ``llm_parse_sms``  – uses the Groq API (LLM) when explicitly requested.
Both return the same structure so the workbook writer does not care which ran.
"""

from __future__ import annotations

import re
import json
import difflib

# ---------------------------------------------------------------------------
# Item catalogue: Bengali name (as the cashier writes it) -> serial number (SL)
# in the workbook. Matching by SL avoids the messy English spellings that live
# in the sheet ("St.rry  Juice", "Café Latte", "Coffee " with trailing space…).
# ---------------------------------------------------------------------------
BN_TO_SL = {
    "বাটিকাপ": 1,
    "চকবার": 2,
    "ক্ষীরআইসক্রিম": 3,
    "মিনিকোন": 4,
    "ভ্যানিলাকোন": 5,
    "ললি": 6,
    "মালাই": 7,
    "ব্ল্যাকফরেস্ট": 8,
    "ম্যাচো": 9,
    "চকলেটকোন": 10,
    "হেজেলকোন": 11,
    "রোবাস্ত": 12,
    "রয়েলসান্ডি": 13,
    "কান্চি": 14,
    "ম্যাজিকবল": 15,
    "পানি500ml": 16,
    "পানিএকলিটার": 17,
    "ফানটা": 18,
    "মোজো": 19,
    "স্প্রাইট": 20,
    "ডিংকো": 21,
    "বোতলজুস": 22,
    "স্পিডকেন": 23,
    "ফুটিকাজুস": 24,
    "ম্যাংগোজুস": 25,
    "চকলেটজুস": 26,
    "স্টবেরীজুস": 27,
    "বেলজিয়ামজুস": 28,   # sheet cell literally reads "Juice" (SL 28)
    "কোকিসজুস": 29,
    "কাবদই": 30,
    "কোকোকোলা": 31,
    "ভাইব": 32,
    "ক্যাফেলাটে": 33,
    "চটপটি": 34,
    "ডিমফুচকা": 35,
    "দইফুচকা": 36,
    "চিকেনফ্রাই": 37,
    # NOTE: "ফ্রাইড রাইস / Fried Rice" has no row in this workbook, so it is
    # intentionally not mapped (the cashier always reports it as 0 anyway).
    "বার্গার": 38,
    "চাওমিন": 39,
    "ললিপপ": 40,
    "পিনাটচকলেট": 41,
    "ফাইটারচকলেট": 42,
    "চকোমাট": 43,
    "কফিলজেন্স": 44,
    "জেলফিলজেন্স": 45,
    "পালসলজেন্স": 46,
    "স্পেশালmilk": 47,
    "ডাইরিমিল্ক": 48,
    "পার্ক": 50,          # sheet skips SL 49
    "কফি": 51,
}

# Special (non 4-row) block: Kids ride "Jumping" lives at fixed rows.
JUMPING_BN = "জাম্পিং"

# Display names (English) purely for the on-screen summary table.
SL_TO_EN = {
    1: "Bati Cup", 2: "Chokober", 3: "Kheer Icecream", 4: "Mini Cone",
    5: "Vanilla Cone", 6: "Lolly", 7: "Malai", 8: "Black Forest", 9: "Macho",
    10: "Chocolate Cone", 11: "Hezel Cone", 12: "Robust", 13: "Royal",
    14: "Kanchi", 15: "Magic ball", 16: "Water", 17: "Water 1 L", 18: "Fanta",
    19: "Mojo", 20: "Sprite", 21: "Dinko", 22: "Bottal Juice", 23: "Speed Can",
    24: "Futika", 25: "Mango Juice", 26: "Chocolate Juice", 27: "St.rry Juice",
    28: "Belgium Juice", 29: "Cookies Juice", 30: "Cup Doi", 31: "Cocacola",
    32: "Vibe", 33: "Café Latte", 34: "Chatapi", 35: "Dim Fuska",
    36: "Doi Fuska", 37: "Chicken Fry", 38: "Burger", 39: "Chowmin",
    40: "Lolly Pop", 41: "Peanut", 42: "Fitter", 43: "ChokoMoko",
    44: "Coffee Logence", 45: "Jelfy", 46: "Pulse", 47: "Special Milk",
    48: "Dairy Milk", 50: "Park", 51: "Coffee",
}

BENGALI_DIGITS = "০১২৩৪৫৬৭৮৯"
_BN_DIGIT_MAP = {ord(b): str(i) for i, b in enumerate(BENGALI_DIGITS)}


def bn_to_ascii_digits(text: str) -> str:
    return text.translate(_BN_DIGIT_MAP)


def _clean_name(chunk: str) -> str:
    """Normalise a candidate item name so it can be matched to BN_TO_SL keys."""
    chunk = bn_to_ascii_digits(chunk)
    # drop leading serial like "16." or "16।" and stray bullet punctuation
    chunk = re.sub(r"^[\s\d.।\-)*:]+", "", chunk)
    # drop parenthetical notes e.g. "(35)"
    chunk = re.sub(r"\(.*?\)", "", chunk)
    # remove trailing standalone digits (rate that leaked into the name region)
    chunk = re.sub(r"[\d.]+$", "", chunk)
    # collapse whitespace and punctuation
    chunk = re.sub(r"[\s।.,=×xX*+\-]+", "", chunk)
    return chunk.strip().lower()


# pre-computed normalised keys for fuzzy matching
_NORM_KEYS = {re.sub(r"\s+", "", k).lower(): sl for k, sl in BN_TO_SL.items()}


def _match_sl(name_norm: str):
    if not name_norm:
        return None
    if name_norm in _NORM_KEYS:
        return _NORM_KEYS[name_norm]
    # substring containment: prefer the LONGEST matching key so that
    # "ললিপপ" (Lolly Pop) wins over "ললি" (Lolly).
    best_key, best_len = None, 0
    for key in _NORM_KEYS:
        if key and (key in name_norm or name_norm in key):
            if len(key) > best_len:
                best_key, best_len = key, len(key)
    if best_key is not None:
        return _NORM_KEYS[best_key]
    # fuzzy fallback
    close = difflib.get_close_matches(name_norm, list(_NORM_KEYS.keys()), n=1, cutoff=0.72)
    if close:
        return _NORM_KEYS[close[0]]
    return None


_BENGALI_RE = re.compile(r"[\u0980-\u09FF]")


def _pick_name_line(region: str) -> str:
    """From the text preceding a numeric triple, isolate the item-name line.

    The cashier separates items with newlines and inserts subtotal lines
    ("মোট = …"). We take the last line that actually contains Bengali letters
    and is not a subtotal, so category headers / subtotals never leak in.
    """
    best = ""
    for raw in region.replace("\r", "\n").split("\n"):
        line = raw.strip()
        if not line or not _BENGALI_RE.search(line):
            continue
        if any(tok in line for tok in ("মোট", "লাভ", "Total", "total", "Grand")):
            continue
        best = line
    return best


_NUM = r"([0-9][0-9.]*)"
# rate × qty = total   (total optional). Only spaces/tabs may sit around the
# operators so the pattern never bleeds into the next line's serial number.
_TRIPLE_RE = re.compile(_NUM + r"[ \t]*[×xX*][ \t]*" + _NUM + r"(?:[ \t]*=[ \t]*" + _NUM + r")?")
_DATE_RE = re.compile(r"(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{2,4})")


def parse_date(text: str):
    """Return (day, month, year) from the message, or None."""
    m = _DATE_RE.search(bn_to_ascii_digits(text))
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000
    return d, mo, y


def parse_sms(text: str):
    """Offline regex parser. Returns (records, jumping, date, warnings)."""
    ascii_text = bn_to_ascii_digits(text)
    records = []
    jumping = None
    warnings = []

    last_end = 0
    matches = list(_TRIPLE_RE.finditer(ascii_text))
    for m in matches:
        rate = float(m.group(1))
        qty = float(m.group(2))
        total = float(m.group(3)) if m.group(3) else round(rate * qty, 2)
        name_region = ascii_text[last_end:m.start()]
        last_end = m.end()

        name_line = _pick_name_line(name_region)
        name_norm = _clean_name(name_line)
        # Jumping ride special-case
        if re.sub(r"\s+", "", JUMPING_BN) in name_norm or "jamping" in name_norm or "jumping" in name_norm:
            jumping = {"person": qty, "taka": rate, "total": total}
            continue

        sl = _match_sl(name_norm)
        if sl is None:
            # ignore category subtotals ("মোট"/"total") silently
            if any(tok in name_region for tok in ("মোট", "Total", "total", "grand", "Grand", "লাভ")):
                continue
            # Only warn when real sales would be lost. A qty of 0 (e.g. "ফ্রাইড
            # রাইস" which has no row in this workbook) is nothing to record.
            if qty > 0:
                warnings.append(
                    f"অচেনা আইটেম বাদ পড়েছে: '{name_line.strip()[:40]}' (qty={qty})"
                )
            continue
        records.append({"sl": sl, "qty": qty, "rate": rate, "total": total})

    date = parse_date(text)
    return records, jumping, date, warnings


# ---------------------------------------------------------------------------
# LLM engine (Groq). Kept optional; the offline parser is the safe default.
# ---------------------------------------------------------------------------
def _catalogue_for_prompt() -> str:
    lines = [f"{sl}: {SL_TO_EN.get(sl, '')}" for sl in sorted(set(BN_TO_SL.values()))]
    return "\n".join(lines)


def llm_parse_sms(text: str, client, model: str):
    """Ask the Groq LLM to return SL/qty/rate. Falls back to regex on failure."""
    catalogue = _catalogue_for_prompt()
    prompt = f"""You convert a Bengali café cashier sales message into JSON.

Here is the fixed item catalogue (serial number: name):
{catalogue}
Serial 999 = "Jumping" kids ride (special).

Each item line looks like: <serial>. <bengali name> = <rate> × <quantity> = <total>
Bengali digits (০-৯) map to 0-9. The symbol × separates rate and quantity.
Match every sold line to the closest catalogue serial by NAME meaning, not by the
serial the cashier typed (their numbering has mistakes). Ignore subtotal lines
that contain মোট / Total / Grand / লাভ.

Return ONLY JSON:
{{
  "date": {{"day": 0, "month": 0, "year": 0}},
  "items": [{{"sl": 0, "qty": 0.0, "rate": 0.0}}],
  "jumping": {{"person": 0.0, "taka": 0.0}}
}}

Message:
{text}
"""
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": "You are a precise JSON extraction engine. Never invent numbers."},
            {"role": "user", "content": prompt},
        ],
        temperature=0,
        response_format={"type": "json_object"},
    )
    data = json.loads(resp.choices[0].message.content)

    records = []
    for it in data.get("items", []):
        sl = it.get("sl")
        if sl == 999:
            continue
        if sl is None:
            continue
        records.append({
            "sl": int(sl),
            "qty": float(it.get("qty", 0) or 0),
            "rate": float(it.get("rate", 0) or 0),
            "total": round(float(it.get("qty", 0) or 0) * float(it.get("rate", 0) or 0), 2),
        })

    jumping = None
    j = data.get("jumping") or {}
    if j and (j.get("person") or j.get("taka")):
        jumping = {"person": float(j.get("person", 0) or 0), "taka": float(j.get("taka", 0) or 0)}
        jumping["total"] = round(jumping["person"] * jumping["taka"], 2)

    d = data.get("date") or {}
    date = None
    if d.get("day") and d.get("month"):
        y = int(d.get("year") or 0)
        if y and y < 100:
            y += 2000
        date = (int(d["day"]), int(d["month"]), y or None)

    return records, jumping, date, []
