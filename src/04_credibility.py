"""
04_credibility.py
==================
STAGE 4 of the FMCG M&A Newsletter pipeline: CREDIBILITY SCORING

WHAT THIS SCRIPT DOES
----------------------
Scores each relevance-filtered article on a simple, explainable source-
credibility scale (LOW / MEDIUM / HIGH), then keeps only MEDIUM+HIGH
articles for the newsletter -- a basic editorial check so the newsletter
doesn't surface a deal sourced only from an unverified blog or content farm.

THE BUG THIS REPLACES
----------------------
The original implementation checked the `source_feed` column directly
against SOURCE_WEIGHTS (publisher names like "reuters", "bloomberg").
source_feed actually holds the RSS QUERY label for 3 of the 8 feeds
("fmcg_acquisition", "food_beverage_deals", "consumer_investment") -- it
is NOT the publisher name for those rows, so every match against
SOURCE_WEIGHTS silently failed and every article fell into the "unknown
source" branch, capping every score at the same low ceiling and
collapsing all 84 articles into a single MEDIUM bucket with zero HIGH or
LOW differentiation.

THE FIX -- two-tier publisher detection
-----------------------------------------
  1. DIRECT FEEDS: 5 of the 8 RSS feeds in 01_ingest.py ARE a single
     named publisher (reuters_business, economic_times, livemint,
     moneycontrol, business_standard) -- for these, source_feed itself
     reliably tells you the publisher. No parsing needed, just a lookup.
  2. GOOGLE QUERY FEEDS: the remaining 3 feeds (fmcg_acquisition,
     food_beverage_deals, consumer_investment) are Google News SEARCH
     feeds covering many publishers, so source_feed can't identify the
     publisher there. For these, the publisher name is recovered from the
     TITLE instead -- Google News RSS titles end with " - Publisher Name"
     (e.g. "... - Reuters", "... - PR Newswire").

CREDIBILITY SCORE -- publisher weight + two small content signals
---------------------------------------------------------------------
  score = publisher_weight (1-3)
        + 1 if the cleaned headline has more than 6 words (a real,
              detailed headline vs. a thin stub)
        + 1 if the headline contains an explicit deal keyword
              (acquire / investment / funding / merger)

  Score 4-5 -> HIGH    | Score 2-3 -> MEDIUM   | Score 1 -> LOW
  (These cutoffs are a documented assumption, not a derived constant --
  see MIN_CREDIBILITY_LEVEL below to change what gets kept.)

INPUT / OUTPUT
--------------
Reads:  data/relevant_news.csv   (written by 03_relevance.py)
Writes: data/credible_news.csv   (read next by 05_summarize.py)

Output = input columns, unchanged, PLUS:
  publisher          : the publisher name the script identified (or
                        'unknown' if neither detection path matched)
  credibility_score  : the raw 1-5 integer score
  credibility         : LOW / MEDIUM / HIGH

Only rows at or above MIN_CREDIBILITY_LEVEL (default: MEDIUM) are written
to credible_news.csv. LOW-credibility rows are dropped here -- the same
row-count reduction the original pipeline showed going from 146 relevant
articles down to 84 credible ones.

Console output reports the publisher-detection method breakdown, which
publishers were actually identified, the score-tier distribution before
filtering, and the final kept/dropped count.
"""

import os
import re

import pandas as pd

# ─────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────

INPUT_FILE = os.path.join("data", "relevant_news.csv")
OUTPUT_FILE = os.path.join("data", "credible_news.csv")

# Only rows scoring at or above this tier are kept in the output. Set to
# "LOW" to keep everything -- useful for auditing what got dropped and why.
MIN_CREDIBILITY_LEVEL = "MEDIUM"
CREDIBILITY_RANK = {"LOW": 0, "MEDIUM": 1, "HIGH": 2}

# Publisher reputation weights (1-3). This is a judgment call, not an
# objective fact, and is documented here as exactly that: established
# international wires and India's top financial dailies score highest,
# specialist trade/startup press is mid-tier, and anything not on this
# list (regional sites, aggregators, PR-wire syndication) defaults to 1.
SOURCE_WEIGHTS = {
    # International wires / majors
    "reuters": 3,
    "bloomberg": 3,
    "financial times": 3,
    "wall street": 3,
    # Indian — Tier 1 financial dailies
    "economic times": 3,
    "business standard": 3,
    "mint": 2,
    "livemint": 2,
    "hindustan times": 2,
    "business line": 2,
    # Specialist / trade press
    "techcrunch": 2,
    "vccircle": 2,
    "inc42": 2,
    "entrackr": 2,
    # Generic aggregator
    "news.google": 1,
}

# Feeds in 01_ingest.py that ARE a single named publisher -- for these,
# source_feed reliably identifies the publisher with no title parsing
# needed. Keys must match the RSS_FEEDS keys in 01_ingest.py exactly.
DIRECT_PUBLISHER_FEEDS = {
    "reuters_business": "reuters",
    "economic_times": "economic times",
    "livemint": "livemint",
    "moneycontrol": "moneycontrol",
    "business_standard": "business standard",
}

# Deal keywords used ONLY for the credibility content-signal bonus. This
# is a tighter, deliberately separate list from 03_relevance.py's
# DEAL_KEYWORDS -- here it's just a thin signal of headline substance, not
# an FMCG-relevance decision, so it doesn't need to be exhaustive.
DEAL_SIGNAL_KEYWORDS = ["acquire", "acquisition", "investment", "funding", "merger"]


# ─────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────

def split_title(title):
    """
    Split an RSS title into (headline, publisher_suffix) on the LAST
    ' - ' / ' – ' / ' — ' separator, not the first. Using the last
    separator matters because a headline itself can legitimately contain
    a dash mid-sentence (e.g. "Reliance-backed venture..."); the
    publisher attribution is reliably the final segment, not the first
    one encountered.
    """
    title = str(title)
    parts = re.split(r"\s[-–—]\s", title)
    if len(parts) > 1:
        headline = " - ".join(parts[:-1]).strip()
        publisher_suffix = parts[-1].strip()
    else:
        headline, publisher_suffix = title.strip(), ""
    return headline, publisher_suffix


def clean_title(title):
    """Return just the headline portion of an RSS title (see
    split_title), used for the word-count and keyword-substance checks."""
    headline, _ = split_title(title)
    return headline


def detect_publisher(source_feed, title):
    """
    Two-tier publisher detection (see module docstring):
      1. Direct publisher feed -> source_feed itself names the publisher.
      2. Otherwise -> match the publisher SUFFIX of the title (everything
         after the final separator) against SOURCE_WEIGHTS, with spaces
         stripped on both sides so e.g. a feed rendering "Business
         Standard" as "BusinessStandard" or "Hindu BusinessLine" still
         matches "business standard" / "business line" in the dictionary.
         Matching only the suffix (not the whole title) also avoids a
         short key like "mint" accidentally matching mid-headline text.
      3. Neither matches -> 'unknown', weight 1.

    Returns (publisher_name, weight, detection_method).
    """
    source_feed_lower = str(source_feed).lower()

    if source_feed_lower in DIRECT_PUBLISHER_FEEDS:
        publisher = DIRECT_PUBLISHER_FEEDS[source_feed_lower]
        return publisher, SOURCE_WEIGHTS[publisher], "direct_feed"

    _, publisher_suffix = split_title(title)
    suffix_normalized = publisher_suffix.lower().replace(" ", "")

    for publisher, weight in SOURCE_WEIGHTS.items():
        if publisher.replace(" ", "") in suffix_normalized:
            return publisher, weight, "title_match"

    return "unknown", 1, "unknown"


def score_credibility(source_feed, title):
    """
    Combine publisher weight with two content signals into a single
    integer score, then map that score to a LOW/MEDIUM/HIGH tier.

    Returns (score, tier, publisher, detection_method).
    """
    publisher, weight, method = detect_publisher(source_feed, title)
    title_clean = clean_title(title)

    score = weight
    if len(title_clean.split()) > 6:
        score += 1
    if any(kw in title_clean.lower() for kw in DEAL_SIGNAL_KEYWORDS):
        score += 1

    if score >= 4:
        tier = "HIGH"
    elif score >= 2:
        tier = "MEDIUM"
    else:
        tier = "LOW"

    return score, tier, publisher, method


# ─────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 70)
    print("STAGE 4 - CREDIBILITY SCORING")
    print("=" * 70 + "\n")

    if not os.path.exists(INPUT_FILE):
        print(f"Input file not found: {INPUT_FILE}")
        print("Run 03_relevance.py first.\n")
        return

    df = pd.read_csv(INPUT_FILE)
    rows_in = len(df)

    if df.empty:
        print("Input file has 0 rows -- nothing to score.\n")
        return

    df["title"] = df["title"].fillna("")
    df["source_feed"] = df["source_feed"].fillna("")

    scores, tiers, publishers, methods = [], [], [], []
    for _, row in df.iterrows():
        score, tier, publisher, method = score_credibility(row["source_feed"], row["title"])
        scores.append(score)
        tiers.append(tier)
        publishers.append(publisher)
        methods.append(method)

    df["publisher"] = publishers
    df["credibility_score"] = scores
    df["credibility"] = tiers
    df["_detection_method"] = methods

    # ── Console diagnostics before filtering ────────────────────────────
    print("Publisher detection method:")
    print(df["_detection_method"].value_counts().to_string())
    print("\nPublisher identified:")
    print(df["publisher"].value_counts().to_string())
    print("\nCredibility tier (before filtering):")
    tier_counts = df["credibility"].value_counts().reindex(["HIGH", "MEDIUM", "LOW"]).fillna(0).astype(int)
    print(tier_counts.to_string())

    df = df.drop(columns=["_detection_method"])

    min_rank = CREDIBILITY_RANK[MIN_CREDIBILITY_LEVEL]
    keep_mask = df["credibility"].map(CREDIBILITY_RANK) >= min_rank
    credible_df = df[keep_mask].reset_index(drop=True)
    dropped = rows_in - len(credible_df)

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    credible_df.to_csv(OUTPUT_FILE, index=False)

    # ── Run summary printed to console ─────────────────────────────────
    print("\n" + "-" * 70)
    print("CREDIBILITY SUMMARY")
    print("-" * 70)
    print(f"Input rows                                    : {rows_in}")
    print(f"Minimum credibility kept                      : {MIN_CREDIBILITY_LEVEL}+")
    print(f"Rows dropped (below {MIN_CREDIBILITY_LEVEL})               : {dropped}")
    print(f"Final credible article count                  : {len(credible_df)}")
    print(f"Output file                                   : {OUTPUT_FILE}")
    print("-" * 70 + "\n")


if __name__ == "__main__":
    main()