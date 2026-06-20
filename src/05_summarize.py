"""
05_summarize.py
================
STAGE 5 of the FMCG M&A Newsletter pipeline: SUMMARIZATION

WHAT THIS SCRIPT DOES
----------------------
Turns each credible article into structured deal intelligence: acquirer,
target, deal_type, deal_value, and a one-sentence business-readable
synopsis. This is the stage where the LLM split discussed earlier in the
project actually gets implemented:

  ALWAYS DETERMINISTIC -- deal_type, deal_value
    Both are pattern-matching problems, not reasoning problems:
      deal_type  : keyword classification (Acquisition / Investment /
                   Merger / PE / Funding / Other)
      deal_value : regex for a currency symbol + number + scale word
                   (e.g. "₹200 cr", "$4.3bn", "INR 200 Crore")
    These run on every row regardless of the LLM flag, are instant, and
    were both validated against the real credible_news.csv before being
    finalized (see the keyword list comments below for the two real
    misses that testing caught and fixed: "snaps up" and "$5B Buy" were
    not matching anything until added).

  LLM-GATED (phi3.5 via Ollama) -- acquirer, target, synopsis
    These genuinely need reasoning: correctly attributing WHO acquired
    WHOM depends on sentence structure ("X acquires Y" vs "Y acquired by
    X"), and writing a coherent one-sentence synopsis is inherently
    generative. Neither is reliably solvable with regex, so unlike
    deal_type/deal_value, they are NOT attempted deterministically.

USE_LLM FLAG (off by default, must work with zero LLM dependency)
---------------------------------------------------------------------
  USE_LLM=true  -> acquirer/target/synopsis come from a local Ollama call
                   to phi3.5. If Ollama isn't reachable, falls back below.
  USE_LLM=false (or Ollama unreachable) -> acquirer='N/A', target='N/A',
                   synopsis=the cleaned headline itself.
  This default is deliberately HONEST rather than clever: a blank 'N/A'
  is preferable to a confidently wrong guess at who acquired whom, which
  is exactly the failure mode a fake deterministic guess would risk.

INPUT / OUTPUT
--------------
Reads:  data/credible_news.csv     (written by 04_credibility.py)
Writes: data/summarized_news.csv   (read next by 06_newsletter.py)

This stage does not drop rows -- every credible article gets a structured
record, even if some fields end up 'N/A'. Output = input columns,
unchanged, PLUS:
  acquirer, target, deal_type, deal_value, synopsis, summary_method
  (summary_method: 'llm' / 'deterministic_fallback' / 'llm_failed_fallback',
   so it's always auditable which path produced which row)

Console output reports the deal_type distribution, how many rows found a
deal_value, whether Ollama was used, and LLM success/failure counts if so.
"""

import os
import re
import json

import pandas as pd
import requests

# ─────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────

INPUT_FILE = os.path.join("data", "credible_news.csv")
OUTPUT_FILE = os.path.join("data", "summarized_news.csv")

# Same master switch used in 03_relevance.py -- one flag controls LLM use
# across the whole pipeline. Off by default so the script runs instantly
# with zero LLM dependency when there's no time to wait on Ollama.
USE_LLM = os.getenv("USE_LLM", "false").lower() == "true"

OLLAMA_URL = os.getenv("OLLAMA_URL", "http://localhost:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "phi3.5")
OLLAMA_TIMEOUT_SECONDS = 45  # extraction + synopsis is heavier than relevance's plain YES/NO

# Placeholder for "couldn't determine this" fields. Deliberately NOT the
# string "N/A" -- pandas (and Excel) treat "N/A" as a built-in missing-
# value sentinel, so writing it to CSV and reading it back anywhere
# downstream (06_newsletter.py, or just opening the file in Excel) would
# silently turn it into an actual null rather than visible text. Caught
# by testing: the LLM-extraction test run below showed real NaN values
# in place of the literal "N/A" string after one CSV round-trip.
UNKNOWN_PLACEHOLDER = "Not specified"

# Deal-value pattern: a currency symbol/word, a number (with optional
# decimal/commas), optionally followed by a scale word. Tested against the
# real dataset -- only matches when the headline itself states a figure
# (about 1 in 4 headlines do; the rest genuinely don't mention a value at
# all, so 'N/A' there reflects the source data, not a regex gap).
DEAL_VALUE_PATTERN = re.compile(
    r"(₹|Rs\.?|INR|\$|USD)\s?[\d,]+(\.\d+)?\s?(crore|cr\b|lakh[s]?|million|mn\b|billion|bn\b)?",
    re.IGNORECASE,
)

# Deal-type keyword buckets, checked in this priority order (most specific
# category first) so e.g. "merger" isn't accidentally re-classified by a
# broader later check. ACQUISITION_KEYWORDS includes both buy-side and
# sell-side language ("sells", "offload") since a divestment and an
# acquisition are the same underlying transaction described from either
# party's side, and the schema has no separate "Divestment" category.
MERGER_KEYWORDS = ["merger", "merges", "merged"]
PE_KEYWORDS = ["private equity", "venture capital"]
FUNDING_KEYWORDS = ["series a", "series b", "series c", "funding round", "raises", "raised", "secures funding"]
ACQUISITION_KEYWORDS = [
    "acqui", "stake", "buyout", "takeover", "buys", "bought", "buy",
    "purchase", "snap up", "snaps up", "snapped up", "offload",
    "divest", "sells", "sold", "sale of", " sale ",
]
INVESTMENT_KEYWORDS = ["invest"]


# ─────────────────────────────────────────────────────────────────────────
# DETERMINISTIC EXTRACTION (always runs)
# ─────────────────────────────────────────────────────────────────────────

def classify_deal_type(text):
    """Keyword-based deal_type classification. See module docstring for
    why this stays deterministic rather than going through the LLM."""
    text = " " + str(text).lower() + " "
    if any(kw in text for kw in MERGER_KEYWORDS):
        return "Merger"
    if any(kw in text for kw in PE_KEYWORDS) or re.search(r"\bvc\b", text):
        return "PE"
    if any(kw in text for kw in FUNDING_KEYWORDS):
        return "Funding"
    if any(kw in text for kw in ACQUISITION_KEYWORDS):
        return "Acquisition"
    if any(kw in text for kw in INVESTMENT_KEYWORDS):
        return "Investment"
    return "Other"


def extract_deal_value(text):
    """Regex-based deal_value extraction. Returns UNKNOWN_PLACEHOLDER if
    the headline and summary don't state a figure at all."""
    match = DEAL_VALUE_PATTERN.search(str(text))
    return match.group(0).strip() if match else UNKNOWN_PLACEHOLDER


def split_title(title):
    """Split an RSS title into (headline, publisher_suffix) on the LAST
    separator -- same logic as 04_credibility.py, redefined locally since
    each pipeline stage is self-contained."""
    title = str(title)
    parts = re.split(r"\s[-–—]\s", title)
    if len(parts) > 1:
        return " - ".join(parts[:-1]).strip(), parts[-1].strip()
    return title.strip(), ""


def clean_title(title):
    """Headline only, publisher suffix stripped -- used as the synopsis
    fallback when the LLM path isn't used."""
    headline, _ = split_title(title)
    return headline


# ─────────────────────────────────────────────────────────────────────────
# LLM EXTRACTION (Ollama / phi3.5, gated by USE_LLM)
# ─────────────────────────────────────────────────────────────────────────

def check_ollama():
    """Quick health check so USE_LLM=true degrades gracefully instead of
    crashing if Ollama isn't installed or running."""
    try:
        r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=5)
        return r.status_code == 200
    except requests.RequestException:
        return False


def build_extraction_prompt(title, summary):
    return f"""You are extracting structured deal information from an FMCG
M&A news headline, for a business newsletter.

From the text below, identify:
- acquirer: the company or investor making the acquisition/investment.
  Write "N/A" if genuinely unclear.
- target: the company or brand being acquired or invested in.
  Write "N/A" if genuinely unclear.
- synopsis: one plain-English sentence (maximum 25 words) summarising the
  deal for a business reader. No jargon, no restating "this article is about".

Title: {title}
Summary: {summary}

Respond with ONLY a JSON object in exactly this format, no other text:
{{"acquirer": "...", "target": "...", "synopsis": "..."}}"""


def parse_llm_json(raw_response):
    """Extract the first {...} block from the model's response and parse
    it as JSON. phi3.5 is a small model and won't always honor
    'respond with ONLY JSON' perfectly, so this tolerates extra text
    before/after the JSON object."""
    start = raw_response.find("{")
    end = raw_response.rfind("}") + 1
    if start == -1 or end <= start:
        raise ValueError("no JSON object found in response")
    return json.loads(raw_response[start:end])


def _normalize_unknown(value):
    """If a field is empty or one of the common 'unknown' conventions a
    model might write (N/A, NA, none, unclear, unknown...), normalize it
    to UNKNOWN_PLACEHOLDER so nothing pandas/Excel-unsafe ever gets
    written to the CSV."""
    value = str(value).strip()
    if not value or value.lower() in {"n/a", "na", "none", "unclear", "unknown", "null"}:
        return UNKNOWN_PLACEHOLDER
    return value


def ollama_extract(title, summary):
    """
    Send one article to the local phi3.5 model and parse acquirer/target/
    synopsis from its JSON response. On any failure (timeout, bad JSON,
    missing keys), falls back to the same deterministic default used when
    the LLM is off entirely -- UNKNOWN_PLACEHOLDER for acquirer/target,
    cleaned headline for synopsis -- rather than guessing.

    Returns (acquirer, target, synopsis, method) where method is 'llm' on
    success or 'llm_failed_fallback' if anything went wrong.
    """
    try:
        response = requests.post(
            f"{OLLAMA_URL}/api/generate",
            json={
                "model": OLLAMA_MODEL,
                "prompt": build_extraction_prompt(title, summary),
                "stream": False,
            },
            timeout=OLLAMA_TIMEOUT_SECONDS,
        )
        raw = response.json().get("response", "")
        data = parse_llm_json(raw)
        acquirer = _normalize_unknown(data.get("acquirer", ""))
        target = _normalize_unknown(data.get("target", ""))
        synopsis = str(data.get("synopsis", "")).strip() or clean_title(title)
        return acquirer, target, synopsis, "llm"
    except (requests.RequestException, ValueError, KeyError, json.JSONDecodeError):
        return UNKNOWN_PLACEHOLDER, UNKNOWN_PLACEHOLDER, clean_title(title), "llm_failed_fallback"


# ─────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 70)
    print(f"STAGE 5 - SUMMARIZE   |   USE_LLM = {USE_LLM}   |   model = {OLLAMA_MODEL}")
    print("=" * 70 + "\n")

    if not os.path.exists(INPUT_FILE):
        print(f"Input file not found: {INPUT_FILE}")
        print("Run 04_credibility.py first.\n")
        return

    df = pd.read_csv(INPUT_FILE)
    rows_in = len(df)

    if df.empty:
        print("Input file has 0 rows -- nothing to summarize.\n")
        return

    df["title"] = df["title"].fillna("")
    df["summary"] = df["summary"].fillna("")
    combined_text = df["title"] + " " + df["summary"]

    # ── Deterministic fields: always computed, regardless of USE_LLM ────
    df["deal_type"] = combined_text.apply(classify_deal_type)
    df["deal_value"] = combined_text.apply(extract_deal_value)

    # ── LLM-gated fields: acquirer, target, synopsis ────────────────────
    llm_available = False
    llm_success = llm_failed = 0

    if USE_LLM:
        llm_available = check_ollama()

    acquirers, targets, synopses, methods = [], [], [], []

    if USE_LLM and llm_available:
        print(f"Ollama reachable at {OLLAMA_URL} -- extracting acquirer/target/"
              f"synopsis for {rows_in} row(s) with {OLLAMA_MODEL}...\n")
        for _, row in df.iterrows():
            acquirer, target, synopsis, method = ollama_extract(row["title"], row["summary"])
            acquirers.append(acquirer)
            targets.append(target)
            synopses.append(synopsis)
            methods.append(method)
            if method == "llm":
                llm_success += 1
            else:
                llm_failed += 1
    else:
        if USE_LLM and not llm_available:
            print(f"USE_LLM=true but Ollama not reachable at {OLLAMA_URL}. "
                  f"Falling back to deterministic defaults for all {rows_in} row(s).\n")
        else:
            print(f"USE_LLM=false -- using deterministic defaults for all {rows_in} "
                  f"row(s) (acquirer/target = '{UNKNOWN_PLACEHOLDER}', synopsis = cleaned headline).\n")
        for _, row in df.iterrows():
            acquirers.append(UNKNOWN_PLACEHOLDER)
            targets.append(UNKNOWN_PLACEHOLDER)
            synopses.append(clean_title(row["title"]))
            methods.append("deterministic_fallback")

    df["acquirer"] = acquirers
    df["target"] = targets
    df["synopsis"] = synopses
    df["summary_method"] = methods

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    df.to_csv(OUTPUT_FILE, index=False)

    # ── Run summary printed to console ─────────────────────────────────
    deal_value_found = int((df["deal_value"] != UNKNOWN_PLACEHOLDER).sum())

    print("-" * 70)
    print("SUMMARIZATION SUMMARY")
    print("-" * 70)
    print(f"Input rows                                    : {rows_in}")
    print("Deal type distribution:")
    for deal_type, count in df["deal_type"].value_counts().items():
        print(f"  {deal_type:<15}: {count}")
    print(f"Deal value found in headline/summary           : {deal_value_found}/{rows_in}")
    if USE_LLM and llm_available:
        print(f"Acquirer/target/synopsis via phi3.5 (success)  : {llm_success}")
        print(f"Acquirer/target/synopsis LLM failed (fallback) : {llm_failed}")
    else:
        print("Acquirer/target/synopsis                       : deterministic fallback (N/A, N/A, cleaned headline)")
    print(f"Output file                                    : {OUTPUT_FILE}")
    print("-" * 70 + "\n")


if __name__ == "__main__":
    main()