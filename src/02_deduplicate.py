
"""
02_deduplicate.py
==================
STAGE 2 of the FMCG M&A Newsletter pipeline: DEDUPLICATION

WHAT THIS SCRIPT DOES
----------------------
01_ingest.py already removes EXACT duplicate URLs picked up by more than
one feed in the same run. It does NOT catch the much more common case: the
same real-world deal covered by multiple publishers under different URLs
and slightly reworded titles (e.g. one acquisition reported by Entrackr,
Business Standard, YourStory and Indiatimes on the same day). That is this
script's job, in two passes:

  PASS 1 -- Domain + date + title pre-pass (cheap, no ML)
      Drops rows that share the same domain, the same publish date, AND
      the same title. This catches near-zero-cost cases like a single
      feed re-publishing the identical entry, before the expensive
      embedding step runs on a smaller set.

  PASS 2 -- Semantic near-duplicate removal (sentence-transformers)
      Embeds each article's title+summary using the all-MiniLM-L6-v2
      model and computes pairwise cosine similarity across all surviving
      articles. Any pair scoring >= SIMILARITY_THRESHOLD (0.75) is treated
      as the same underlying story and collapsed into one record. This is
      what catches "Meesho acquires Kirana Club" being reported six
      different ways by six different outlets.

In BOTH passes, articles are sorted oldest-first beforehand, so within any
duplicate cluster the article that was published EARLIEST is always the
one kept. This preserves "who broke the story first" rather than an
arbitrary survivor based on row order.

INPUT / OUTPUT
--------------
Reads:  data/raw_news.csv        (written by 01_ingest.py)
Writes: data/processed_news.csv  (read next by 03_relevance.py)

Output schema is unchanged from ingestion: title, url, published, summary,
source_feed, ingested_at. (Helper columns used internally for sorting and
grouping are dropped before saving, so downstream scripts see exactly the
same columns they already expect.)

Console output reports row counts before/after each pass and the overall
reduction, so the dedup step's impact is auditable at a glance.
"""

import os
import re
from datetime import timezone
from email.utils import parsedate_to_datetime
from urllib.parse import urlparse

import numpy as np
import pandas as pd
from sentence_transformers import SentenceTransformer, util as st_util

# ─────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────

INPUT_FILE = os.path.join("data", "raw_news.csv")
OUTPUT_FILE = os.path.join("data", "processed_news.csv")

# Cosine similarity cutoff above which two articles are treated as the same
# underlying story. 0.75 (raised from an initial 0.70) was chosen to be
# strict enough that two DIFFERENT FMCG deals reported in similar language
# don't get wrongly merged, while still catching reworded headlines about
# the same deal.
SIMILARITY_THRESHOLD = 0.75

# all-MiniLM-L6-v2: a small, fast sentence-embedding model that is more
# than accurate enough for this use case and runs comfortably on CPU --
# no GPU or paid API required.
EMBEDDING_MODEL = "all-MiniLM-L6-v2"

# Columns 01_ingest.py is expected to have produced. Checked up front so a
# schema mismatch fails loudly here rather than silently downstream.
REQUIRED_COLUMNS = ["title", "url", "published", "summary", "source_feed", "ingested_at"]


# ─────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────

def extract_domain(url):
    """Return the bare domain of a URL (no 'www.') for grouping in Pass 1."""
    try:
        return urlparse(str(url)).netloc.replace("www.", "")
    except Exception:
        return ""


def parse_published_date(published_str):
    """
    Parse an RFC-822 RSS date string into a timezone-aware datetime.

    Deliberately uses email.utils.parsedate_to_datetime per-row rather than
    pandas.to_datetime() on the whole column: this dataset mixes Google
    News' "GMT" suffix with LiveMint's "+0530" offset format, and pandas'
    vectorized format-inference silently turns every non-matching row into
    NaT (verified against real pipeline output -- all 35 LiveMint rows
    were lost this way). Parsing per-row, the same way 01_ingest.py already
    does for the recency filter, handles both formats correctly.
    """
    try:
        dt = parsedate_to_datetime(published_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except (TypeError, ValueError):
        return pd.NaT


def normalize_title(title):
    """
    Lightly normalize a title for the Pass-1 exact-match check: lowercase,
    strip punctuation, collapse whitespace. This is intentionally NOT fuzzy
    matching -- it only catches titles that are identical apart from case
    or stray punctuation. Genuinely reworded headlines are left for the
    semantic pass (Pass 2), where they belong.
    """
    title = str(title).lower()
    title = re.sub(r"[^a-z0-9\s]", " ", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title


def domain_date_prepass(df):
    """
    PASS 1: drop rows that share domain + publish-date + normalized title.
    Assumes df is already sorted oldest-first, so keep='first' keeps the
    earliest-published copy.
    """
    before = len(df)
    df = df.drop_duplicates(subset=["domain", "pub_date", "title_norm"], keep="first")
    removed = before - len(df)
    return df.reset_index(drop=True), removed


def remove_semantic_duplicates(df, threshold=SIMILARITY_THRESHOLD):
    """
    PASS 2: embed title+summary for every surviving article and collapse
    any cluster of articles whose pairwise cosine similarity exceeds
    `threshold` into a single record (the earliest-published one).

    Assumes df is already sorted oldest-first. Greedy clustering: walk
    articles in order; for each one not already marked a duplicate, find
    every LATER article above the threshold and mark it as a duplicate of
    this one. Because of the date sort, the survivor of every cluster is
    always the earliest article in it.
    """
    if df.empty:
        return df, 0

    model = SentenceTransformer(EMBEDDING_MODEL)

    # Title + summary combined gives the embedding model the richest signal:
    # titles alone are sometimes too short to distinguish similar deals,
    # and summary can be empty for some feeds.
    texts = (df["title"].fillna("") + ". " + df["summary"].fillna("")).tolist()
    embeddings = model.encode(texts, show_progress_bar=False)

    sim_matrix = np.array(st_util.cos_sim(embeddings, embeddings))

    n = len(df)
    is_duplicate = np.zeros(n, dtype=bool)

    for i in range(n):
        if is_duplicate[i]:
            continue
        for j in range(i + 1, n):
            if is_duplicate[j]:
                continue
            if sim_matrix[i, j] >= threshold:
                is_duplicate[j] = True

    removed = int(is_duplicate.sum())
    df_deduped = df[~is_duplicate].reset_index(drop=True)
    return df_deduped, removed


# ─────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 70)
    print("STAGE 2 - DEDUPLICATE")
    print("=" * 70 + "\n")

    if not os.path.exists(INPUT_FILE):
        print(f"Input file not found: {INPUT_FILE}")
        print("Run 01_ingest.py first.\n")
        return

    df = pd.read_csv(INPUT_FILE)
    rows_in = len(df)

    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        print(f"Input is missing expected column(s): {missing_cols}")
        print(f"Expected schema: {REQUIRED_COLUMNS}\n")
        return

    if df.empty:
        print("Input file has 0 rows -- nothing to deduplicate.\n")
        return

    # Parse published date (RFC-822 strings from RSS) into real datetimes,
    # per-row (see parse_published_date docstring for why NOT to use
    # pandas.to_datetime() directly on the raw string column). Wrapping the
    # already-parsed objects in pd.to_datetime() here just casts the column
    # to a proper datetime64 dtype (needed for .dt/.sort_values) -- it does
    # NOT re-parse the original strings, so the mixed-format bug doesn't
    # reappear.
    df["published_dt"] = pd.to_datetime(df["published"].apply(parse_published_date), utc=True)
    unparseable_dates = df["published_dt"].isna().sum()

    # Sort oldest -> newest BEFORE any dedup step, so every pass below keeps
    # the earliest-published copy of a repeated story. Rows with an
    # unparseable date are pushed to the end rather than dropped.
    df = df.sort_values("published_dt", na_position="last").reset_index(drop=True)

    df["pub_date"] = df["published_dt"].dt.date
    df["domain"] = df["url"].apply(extract_domain)
    df["title_norm"] = df["title"].apply(normalize_title)

    # PASS 1 -- cheap domain+date+title pre-pass
    df, prepass_removed = domain_date_prepass(df)
    rows_after_prepass = len(df)

    # PASS 2 -- semantic near-duplicate removal
    df, semantic_removed = remove_semantic_duplicates(df, threshold=SIMILARITY_THRESHOLD)
    rows_after_semantic = len(df)

    # Drop helper columns -- output schema must match what 03_relevance.py
    # expects: the same 6 columns 01_ingest.py produced, nothing more.
    df = df.drop(columns=["published_dt", "pub_date", "domain", "title_norm"])

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    df.to_csv(OUTPUT_FILE, index=False)

    # ── Run summary printed to console ─────────────────────────────────
    total_removed = rows_in - rows_after_semantic
    pct_removed = (total_removed / rows_in * 100) if rows_in else 0

    print("-" * 70)
    print("DEDUPLICATION SUMMARY")
    print("-" * 70)
    print(f"Input rows                                   : {rows_in}")
    print(f"Rows with unparseable published date          : {unparseable_dates}")
    print(f"Removed in Pass 1 (domain+date+title)          : {prepass_removed}")
    print(f"Rows after Pass 1                              : {rows_after_prepass}")
    print(f"Removed in Pass 2 (semantic, threshold={SIMILARITY_THRESHOLD})  : {semantic_removed}")
    print(f"Rows after Pass 2 (final)                      : {rows_after_semantic}")
    print(f"Total removed                                  : {total_removed} ({pct_removed:.1f}%)")
    print(f"Output file                                    : {OUTPUT_FILE}")
    print("-" * 70 + "\n")


if __name__ == "__main__":
    main()
