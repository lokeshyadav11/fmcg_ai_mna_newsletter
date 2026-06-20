"""
03_relevance.py
================
STAGE 3 of the FMCG M&A Newsletter pipeline: RELEVANCE FILTERING

WHAT THIS SCRIPT DOES
----------------------
Decides which deduplicated articles are genuinely about an FMCG-sector
M&A/investment deal, as opposed to off-sector M&A, general FMCG market
commentary with no deal attached, or unrelated news that slipped through
an RSS query.

CLASSIFICATION LOGIC -- a deliberately simple, explainable 2x2 matrix:

                         FMCG keyword present   FMCG keyword absent
    Deal keyword present       RELEVANT              UNCERTAIN
    Deal keyword absent       UNCERTAIN              NOT_RELEVANT

  - BOTH present    -> high-confidence RELEVANT  (e.g. "Dabur acquires...")
  - NEITHER present -> high-confidence NOT_RELEVANT (e.g. unrelated tech news)
  - Exactly ONE present -> UNCERTAIN. This is where keyword matching
    genuinely can't decide, and is where an LLM adds value, IF enabled:
      * deal keyword but no FMCG keyword -> may be an FMCG company not on
        our keyword list (e.g. "Lactalis", a dairy major, doing a real
        acquisition) or it may be a genuinely off-sector deal (e.g. an
        SAP consultancy acquisition). Rules can't tell; an LLM that knows
        what "Lactalis" is, can.
      * FMCG keyword but no deal keyword -> may be sector commentary with
        no specific transaction (e.g. "Beauty boom to fuel more
        dealmaking...") or a deal described in language the keyword list
        didn't anticipate. Rules can't tell; an LLM reading the actual
        sentence can.

LLM ESCALATION (optional, OFF by default)
-------------------------------------------
Controlled by the USE_LLM environment variable. When 'true', uncertain
rows are sent to a locally running Ollama model for a YES/NO relevance
call. When 'false', or when Ollama isn't reachable, uncertain rows
default to RELEVANT -- i.e. the pipeline is inclusive on uncertainty
rather than silently dropping a real deal. The assumption made explicit
here: a few extra non-deal articles surviving into the credibility/
newsletter stage are far less costly than missing a real FMCG deal.

INPUT / OUTPUT
--------------
Reads:  data/processed_news.csv   (written by 02_deduplicate.py)
Writes: data/relevant_news.csv    (read next by 04_credibility.py)

Output = input columns, unchanged, PLUS two new columns for auditability:
  relevance_decision : RELEVANT / NOT_RELEVANT
  relevance_reason   : how the decision was reached (rule match, LLM
                       escalation, or inclusive default), so every
                       decision can be traced back to why it was made.
Only RELEVANT rows are written to relevant_news.csv.

Console output reports how many rows landed in each quadrant of the 2x2
matrix, how many were escalated to the LLM (if enabled), and the final
relevant article count.
"""

import os

import pandas as pd
import requests

# ─────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────

INPUT_FILE = os.path.join("data", "processed_news.csv")
OUTPUT_FILE = os.path.join("data", "relevant_news.csv")

# Master switch for LLM escalation. OFF by default -- the pipeline must
# run correctly with zero LLM dependency, since Ollama isn't guaranteed to
# be installed/running on every machine this gets demoed on.
USE_LLM = os.getenv("USE_LLM", "false").lower() == "true"

OLLAMA_URL = os.getenv("OLLAMA_URL", "http://localhost:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "mistral:latest")
OLLAMA_TIMEOUT_SECONDS = 30

# FMCG sector vocabulary: categories + major brand/company names. This is
# necessarily incomplete -- that incompleteness is the whole reason the
# UNCERTAIN band and LLM escalation exist. It's meant to catch the common,
# obvious cases cheaply, not to be exhaustive.
FMCG_KEYWORDS = [
    "fmcg", "food", "beverage", "snack", "consumer goods", "personal care",
    "household", "packaged goods", "d2c", "direct to consumer", "dairy",
    "grocery", "kirana", "nutrition", "wellness", "beauty", "cosmetics",
    "unilever", "nestle", "pepsico", "coca-cola", "kraft", "heinz",
    "dabur", "marico", "emami", "godrej", "itc", "hindunilvr", "britannia",
    "reliance consumer", "tata consumer", "parle", "haldirams", "lactalis",
]

# Deal-type vocabulary, covering common ways a transaction gets reported,
# including plain-English phrasing ("buys", "bought") that press releases
# often use instead of "acquires".
DEAL_KEYWORDS = [
    "acquire", "acquisition", "acquires", "acquired", "merger", "buyout",
    "buys", "bought", "purchase", "investment", "invests", "funding",
    "raises", "stake", "private equity", "venture capital", "series a",
    "series b", "series c", "strategic investment", "majority stake",
    "minority stake", "takeover", "joint venture",
]


# ─────────────────────────────────────────────────────────────────────────
# RULE-BASED LAYER
# ─────────────────────────────────────────────────────────────────────────

def clean_text(title, summary):
    """Lowercase and combine title+summary into one text block for
    keyword matching."""
    return f"{title} {summary}".lower()


def has_keyword(text, keyword_list):
    return any(kw in text for kw in keyword_list)


def rule_classify(title, summary):
    """
    Apply the 2x2 matrix described in the module docstring.

    Returns (decision, reason, is_uncertain):
      decision     : 'RELEVANT' or 'NOT_RELEVANT' (provisional if uncertain)
      reason       : short human-readable explanation, kept in the output
                      so every decision is auditable
      is_uncertain : True if this row needs LLM escalation (or the
                      inclusive default) rather than a confident rule call
    """
    text = clean_text(title, summary)
    has_fmcg = has_keyword(text, FMCG_KEYWORDS)
    has_deal = has_keyword(text, DEAL_KEYWORDS)

    if has_deal and has_fmcg:
        return "RELEVANT", "rule: deal keyword + FMCG keyword both matched", False

    if not has_deal and not has_fmcg:
        return "NOT_RELEVANT", "rule: no deal keyword and no FMCG keyword matched", False

    if has_deal and not has_fmcg:
        return ("RELEVANT",
                "uncertain: deal keyword matched, no FMCG keyword "
                "(possible off-sector deal, or FMCG company not on keyword list)",
                True)

    return ("RELEVANT",
            "uncertain: FMCG keyword matched, no deal keyword "
            "(possible sector commentary, or unusually worded deal)",
            True)


# ─────────────────────────────────────────────────────────────────────────
# LLM ESCALATION LAYER (Ollama, optional)
# ─────────────────────────────────────────────────────────────────────────

def check_ollama():
    """
    Quick health check -- is a local Ollama server actually reachable?
    This lets USE_LLM=true degrade gracefully (fall back to the inclusive
    default) instead of crashing the run if Ollama isn't installed or
    isn't running.
    """
    try:
        r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=5)
        return r.status_code == 200
    except requests.RequestException:
        return False


def build_relevance_prompt(title, summary):
    return f"""You are screening news headlines for an FMCG (fast-moving
consumer goods: food, beverage, personal care, household products) M&A
and investment newsletter.

Decide if this article is reporting a SPECIFIC, real transaction
(acquisition, merger, investment, funding round, or stake purchase) where
at least one party is an FMCG-sector company or brand. Answer NO if the
article is general market commentary, a stock-price move with no deal,
or a deal in a clearly unrelated sector (e.g. IT services, logistics,
pharmacy retail with no FMCG product line).

Title: {title}
Summary: {summary}

Answer with exactly one word: YES or NO."""


def ollama_relevance_check(title, summary):
    """
    Send one UNCERTAIN article to the local LLM and parse a YES/NO answer.
    On any failure (timeout, bad response, unparseable text), default to
    RELEVANT -- the same inclusive-on-uncertainty principle used when the
    LLM is unavailable altogether.
    """
    try:
        response = requests.post(
            f"{OLLAMA_URL}/api/generate",
            json={
                "model": OLLAMA_MODEL,
                "prompt": build_relevance_prompt(title, summary),
                "stream": False,
            },
            timeout=OLLAMA_TIMEOUT_SECONDS,
        )
        answer = response.json().get("response", "").strip().upper()
        if answer.startswith("YES"):
            return "RELEVANT", "LLM escalation: model answered YES"
        if answer.startswith("NO"):
            return "NOT_RELEVANT", "LLM escalation: model answered NO"
        return ("RELEVANT",
                f"LLM escalation: unparseable response ('{answer[:30]}'), defaulted inclusive")
    except (requests.RequestException, ValueError, KeyError) as exc:
        return "RELEVANT", f"LLM escalation failed ({exc}), defaulted inclusive"


# ─────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 70)
    print(f"STAGE 3 - RELEVANCE FILTER   |   USE_LLM = {USE_LLM}")
    print("=" * 70 + "\n")

    if not os.path.exists(INPUT_FILE):
        print(f"Input file not found: {INPUT_FILE}")
        print("Run 02_deduplicate.py first.\n")
        return

    df = pd.read_csv(INPUT_FILE)
    rows_in = len(df)

    if df.empty:
        print("Input file has 0 rows -- nothing to filter.\n")
        return

    df["title"] = df["title"].fillna("")
    df["summary"] = df["summary"].fillna("")

    decisions, reasons, uncertain_flags = [], [], []
    for _, row in df.iterrows():
        decision, reason, is_uncertain = rule_classify(row["title"], row["summary"])
        decisions.append(decision)
        reasons.append(reason)
        uncertain_flags.append(is_uncertain)

    df["relevance_decision"] = decisions
    df["relevance_reason"] = reasons
    df["_uncertain"] = uncertain_flags

    confident_relevant = int(((df["relevance_decision"] == "RELEVANT") & (~df["_uncertain"])).sum())
    confident_not_relevant = int((df["relevance_decision"] == "NOT_RELEVANT").sum())
    uncertain_count = int(df["_uncertain"].sum())

    llm_available = False
    llm_yes = llm_no = llm_failed = 0

    if uncertain_count > 0 and USE_LLM:
        llm_available = check_ollama()
        if llm_available:
            print(f"Ollama reachable at {OLLAMA_URL} -- escalating "
                  f"{uncertain_count} uncertain row(s)...\n")
            for idx in df[df["_uncertain"]].index:
                decision, reason = ollama_relevance_check(df.at[idx, "title"], df.at[idx, "summary"])
                df.at[idx, "relevance_decision"] = decision
                df.at[idx, "relevance_reason"] = reason
                if "model answered YES" in reason:
                    llm_yes += 1
                elif "model answered NO" in reason:
                    llm_no += 1
                else:
                    llm_failed += 1
        else:
            print(f"USE_LLM=true but Ollama not reachable at {OLLAMA_URL}. "
                  f"Falling back to inclusive default for {uncertain_count} uncertain row(s).\n")
    elif uncertain_count > 0:
        print(f"USE_LLM=false -- {uncertain_count} uncertain row(s) kept via inclusive "
              f"default (see relevance_reason column for which ones).\n")

    df = df.drop(columns=["_uncertain"])
    relevant_df = df[df["relevance_decision"] == "RELEVANT"].reset_index(drop=True)

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    relevant_df.to_csv(OUTPUT_FILE, index=False)

    # ── Run summary printed to console ─────────────────────────────────
    print("-" * 70)
    print("RELEVANCE FILTER SUMMARY")
    print("-" * 70)
    print(f"Input rows                                     : {rows_in}")
    print(f"Confident RELEVANT (both keywords matched)      : {confident_relevant}")
    print(f"Confident NOT_RELEVANT (neither keyword matched): {confident_not_relevant}")
    print(f"Uncertain (exactly one keyword matched)         : {uncertain_count}")
    if USE_LLM and llm_available:
        print(f"  -> LLM said RELEVANT                          : {llm_yes}")
        print(f"  -> LLM said NOT_RELEVANT                       : {llm_no}")
        print(f"  -> LLM call failed / unparseable (defaulted)   : {llm_failed}")
    elif uncertain_count > 0:
        print("  -> LLM not used -- defaulted to RELEVANT (inclusive)")
    print(f"Final relevant article count                    : {len(relevant_df)}")
    print(f"Output file                                     : {OUTPUT_FILE}")
    print("-" * 70 + "\n")


if __name__ == "__main__":
    main()