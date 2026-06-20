"""
01_ingest.py
============
STAGE 1 of the FMCG M&A Newsletter pipeline: INGESTION

WHAT THIS SCRIPT DOES
---------------------
1. Pulls raw articles from 8 free RSS feeds — 3 Google News *search* feeds
   scoped to FMCG/consumer-deal keywords, plus 5 direct publisher business
   feeds (Reuters, Economic Times, LiveMint, Moneycontrol, Business Standard).
2. Keeps ONLY articles published within the last DAYS_LOOKBACK days (30 by
   default). This is the fix for the "450 old articles burying 8 recent
   ones" issue found during pipeline review — Google News RSS ranks by
   query relevance, not by date, so without this filter the pipeline
   ingests 9+ months of history every run.
3. Strips HTML out of the RSS <summary> field (Google News wraps summaries
   in raw <a>/<font> tags).
4. Drops exact duplicate URLs picked up by more than one feed in the same
   run (a cheap pre-pass — the real near-duplicate/semantic check happens
   later in 02_deduplicate.py).
5. Writes everything to data/raw_news.csv with ingestion metadata, and
   prints a per-feed + overall summary to the console.

OUTPUT
------
data/raw_news.csv with columns:
    title, url, published, summary, source_feed, ingested_at

Console output reports, per feed, how many entries it returned and how
many survived the recency filter, followed by an overall summary (total
articles kept, exact duplicates removed, final row count).

WHY FILTER AT INGESTION RATHER THAN LATER
------------------------------------------
If stale articles are allowed through, every downstream stage (dedup,
relevance scoring, credibility scoring) wastes time processing data that
can never appear in a newsletter claiming to cover "recent" deal activity.
Filtering here keeps the rest of the pipeline fast and keeps the data
genuinely real-time, as required by the problem statement.
"""

import os
import time
from datetime import datetime, timezone, timedelta
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser

import feedparser
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────

# How many days back an article's published date can be and still count as
# "recent" for this newsletter cycle. 30 days = one monthly issue.
DAYS_LOOKBACK = 30

# Where the raw ingested data lands — 02_deduplicate.py reads from this
# exact path next in the pipeline.
OUTPUT_DIR = "data"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "raw_news.csv")

# 8 free, no-API-key RSS feeds.
#   - The first 3 are Google News SEARCH feeds, scoped with FMCG/consumer
#     deal keywords. The "after:YYYY-MM-DD" param nudges Google toward
#     newer results but is not a hard guarantee, which is why the
#     DAYS_LOOKBACK filter below is still applied as the real safety net.
#   - The last 5 are direct publisher RSS feeds, which list articles
#     chronologically and are a more reliable "recent news" source than a
#     relevance-ranked Google search feed.
# NOTE: update the "after:" date below periodically (or compute it
# dynamically — see make_rss_feeds() if you want this automated).
RSS_FEEDS = {
    # Google News search feeds (query-relevance ranked, FMCG/consumer scoped)
    "fmcg_acquisition":    "https://news.google.com/rss/search?q=FMCG+acquisition+after:2026-05-01",
    "food_beverage_deals": "https://news.google.com/rss/search?q=food+beverage+acquisition+after:2026-05-01",
    "consumer_investment": "https://news.google.com/rss/search?q=consumer+brands+investment+after:2026-05-01",
    # Direct publisher RSS feeds (chronological, not relevance-ranked)
    "reuters_business":    "https://feeds.reuters.com/reuters/businessNews",
    "economic_times":      "https://economictimes.indiatimes.com/markets/rss.cms",
    "livemint":            "https://www.livemint.com/rss/companies",
    "moneycontrol":        "https://www.moneycontrol.com/rss/business.xml",
    "business_standard":   "https://www.business-standard.com/rss/latest.rss",
}


# ─────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────

class _HTMLStripper(HTMLParser):
    """Minimal HTML-tag stripper (stdlib only, no BeautifulSoup dependency).
    Google News RSS summaries arrive wrapped in raw HTML; this keeps just
    the visible text."""

    def __init__(self):
        super().__init__()
        self._chunks = []

    def handle_data(self, data):
        self._chunks.append(data)

    def get_text(self):
        return " ".join(self._chunks).strip()


def strip_html(raw_html):
    """Return plain text from an HTML-wrapped RSS summary field."""
    if not raw_html:
        return ""
    stripper = _HTMLStripper()
    stripper.feed(str(raw_html))
    return stripper.get_text()


def is_recent(published_str, days_lookback=DAYS_LOOKBACK):
    """
    Decide whether an article's published date falls inside the lookback
    window. RSS dates are normally RFC-822 strings (e.g. 'Wed, 17 Jun 2026
    09:00:00 GMT'), which email.utils.parsedate_to_datetime parses natively.

    If the date is missing or unparseable, the article is KEPT rather than
    silently dropped — a broken date should be flagged downstream, not lost
    here.
    """
    if not published_str:
        return True  # no date supplied -> don't penalise, keep it

    try:
        pub_dt = parsedate_to_datetime(published_str)
        if pub_dt.tzinfo is None:           # some feeds give naive datetimes
            pub_dt = pub_dt.replace(tzinfo=timezone.utc)
        cutoff = datetime.now(timezone.utc) - timedelta(days=days_lookback)
        return pub_dt >= cutoff
    except (TypeError, ValueError):
        return True  # unparseable date -> keep, don't lose the article


def fetch_feed(feed_name, feed_url):
    """
    Pull one RSS feed and return a list of article dicts, already filtered
    to the recency window with HTML stripped from summaries.

    A single broken/unreachable feed prints a warning and returns an empty
    list rather than crashing the whole ingestion run — ingestion should be
    resilient to one bad source among eight.
    """
    print(f"  -> Fetching '{feed_name}' ...", end=" ", flush=True)

    try:
        parsed = feedparser.parse(feed_url)
    except Exception as exc:
        print(f"FAILED ({exc})")
        return []

    if parsed.bozo and not parsed.entries:
        # 'bozo' is feedparser's flag for a malformed feed. Zero entries on
        # top of that means treat it as a hard failure for this run.
        print("FAILED (malformed feed, 0 entries returned)")
        return []

    total_entries = len(parsed.entries)
    kept_articles = []

    for entry in parsed.entries:
        published_raw = entry.get("published", "")
        if not is_recent(published_raw):
            continue  # outside the lookback window -> drop here, at source

        kept_articles.append({
            "title":       entry.get("title", "").strip(),
            "url":         entry.get("link", "").strip(),
            "published":   published_raw,
            "summary":     strip_html(entry.get("summary", "")),
            "source_feed": feed_name,
            "ingested_at": datetime.now(timezone.utc).isoformat(),
        })

    print(f"{total_entries} entries -> {len(kept_articles)} within last "
          f"{DAYS_LOOKBACK} days")
    return kept_articles


# ─────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 70)
    print(f"STAGE 1 - INGEST   |   lookback window = last {DAYS_LOOKBACK} days")
    print("=" * 70 + "\n")

    all_articles = []

    for feed_name, feed_url in RSS_FEEDS.items():
        all_articles.extend(fetch_feed(feed_name, feed_url))
        time.sleep(1)  # polite pacing -- avoid hammering feed servers

    if not all_articles:
        print("\nNo articles ingested. Check your network connection or "
              "feed URLs above.\n")
        return

    df = pd.DataFrame(all_articles)
    rows_before_dedup = len(df)

    # Cheap exact-duplicate pre-pass: the SAME article URL can appear under
    # more than one feed in a single run (e.g. a deal story matches both the
    # "fmcg_acquisition" and "consumer_investment" Google queries). This is
    # NOT semantic/near-duplicate dedup -- that happens in 02_deduplicate.py
    # -- it only removes identical links so one link isn't counted twice.
    df = df.drop_duplicates(subset=["url"], keep="first").reset_index(drop=True)
    exact_duplicates_removed = rows_before_dedup - len(df)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    df.to_csv(OUTPUT_FILE, index=False)

    # ── Run summary printed to console ─────────────────────────────────
    print("\n" + "-" * 70)
    print("INGESTION SUMMARY")
    print("-" * 70)
    print(f"Feeds queried                                : {len(RSS_FEEDS)}")
    print(f"Articles within last {DAYS_LOOKBACK} days (pre-dedup)    : {rows_before_dedup}")
    print(f"Exact-duplicate URLs removed                 : {exact_duplicates_removed}")
    print(f"Final article count written to disk          : {len(df)}")
    print(f"Output file                                  : {OUTPUT_FILE}")
    print("-" * 70 + "\n")


if __name__ == "__main__":
    main()