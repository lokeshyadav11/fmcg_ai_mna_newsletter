"""
06_newsletter.py
=================
STAGE 6 of the FMCG M&A Newsletter pipeline: NEWSLETTER GENERATION

WHAT THIS SCRIPT DOES
----------------------
Turns the structured deal records from 05_summarize.py into the final,
business-skimmable Word document -- the actual customer-facing
deliverable the rest of the pipeline exists to produce.

Capabilities implemented (agreed before writing this):
  1. Deals grouped by deal_type (Acquisition / Investment / PE / Funding /
     Merger / Other), so a reader can jump straight to the section they
     care about instead of scanning a flat list.
  2. Within each group, sorted by credibility (HIGH first) then by
     published date (most recent first) -- the most trustworthy, most
     recent items lead.
  3. Each entry shows: a bold "Acquirer -> Target" headline (falling back
     to the cleaned article title if both are unknown), the deal type and
     value, a one-line synopsis, the publisher with a colour-coded
     credibility badge, the published date, and a clickable link back to
     the original article.
  4. A front-page PIPELINE FUNNEL table (raw ingested -> deduped ->
     relevant -> credible -> final) -- this is specifically for the
     interviewer demo: it makes the screening the earlier stages did
     visible in the deliverable itself, rather than only in console logs
     nobody outside the pipeline ever sees.
  5. Exports to .docx via python-docx, the format already established for
     this deliverable (Excel/Word/PPT requirement -> Word chosen).

INPUT / OUTPUT
--------------
Reads:  data/summarized_news.csv               (written by 05_summarize.py)
        data/raw_news.csv, processed_news.csv,    (read ONLY for funnel
        relevant_news.csv, credible_news.csv       row counts; missing
                                                     files degrade
                                                     gracefully to 'N/A'
                                                     rather than failing)
Writes: data/newsletter_YYYY_MM.docx

This is the final pipeline stage -- nothing reads its output.
"""

import os
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime

import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE

# ─────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────

INPUT_FILE = os.path.join("data", "summarized_news.csv")
OUTPUT_FILE = os.path.join("data", f"newsletter_{datetime.now().strftime('%Y_%m')}.docx")

# Used ONLY to populate the funnel table -- read for a row count, nothing
# else. If a file is missing the funnel just shows 'N/A' for that stage
# rather than failing the whole newsletter.
FUNNEL_STAGES = [
    ("Raw ingested (RSS, last 30 days)", os.path.join("data", "raw_news.csv")),
    ("After deduplication", os.path.join("data", "processed_news.csv")),
    ("After relevance filter", os.path.join("data", "relevant_news.csv")),
    ("After credibility filter", os.path.join("data", "credible_news.csv")),
    ("Final newsletter items", INPUT_FILE),
]

DEAL_TYPE_ORDER = ["Acquisition", "Investment", "PE", "Funding", "Merger", "Other"]
CREDIBILITY_RANK = {"LOW": 0, "MEDIUM": 1, "HIGH": 2}
CREDIBILITY_COLORS = {
    "HIGH": RGBColor(0x1E, 0x7E, 0x34),     # green
    "MEDIUM": RGBColor(0xB7, 0x79, 0x0E),   # amber
    "LOW": RGBColor(0x6B, 0x6B, 0x6B),      # grey (shouldn't appear -- filtered upstream)
}

UNKNOWN_PLACEHOLDER = "Not specified"  # same convention as 05_summarize.py
FONT_NAME = "Arial"


# ─────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────

def split_title(title):
    """Split an RSS title into (headline, publisher_suffix) on the LAST
    separator -- same logic reused from 04/05, redefined locally since
    each pipeline stage is self-contained."""
    import re
    title = str(title)
    parts = re.split(r"\s[-–—]\s", title)
    if len(parts) > 1:
        return " - ".join(parts[:-1]).strip(), parts[-1].strip()
    return title.strip(), ""


def clean_title(title):
    headline, _ = split_title(title)
    return headline


def parse_published_date(published_str):
    """
    Per-row RFC-822 parsing, deliberately NOT pandas.to_datetime() on the
    whole column -- 02_deduplicate.py testing already found that pandas'
    vectorized format inference silently drops mixed GMT/+0530 timestamps
    (this dataset has both). Parsing per-row avoids that bug here too.
    """
    try:
        dt = parsedate_to_datetime(published_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except (TypeError, ValueError):
        return datetime.min.replace(tzinfo=timezone.utc)  # sorts last, never crashes


def get_funnel_counts():
    """Row count at each pipeline stage, for the transparency table.
    Returns a list of (label, count_or_None)."""
    counts = []
    for label, path in FUNNEL_STAGES:
        n = None
        if os.path.exists(path):
            try:
                n = len(pd.read_csv(path))
            except Exception:
                n = None
        counts.append((label, n))
    return counts


def add_hyperlink(paragraph, url, text, color="0563C1", underline=True):
    """
    Insert a clickable hyperlink run into a paragraph.

    python-docx (this version, 1.2.0) only exposes a READ-ONLY
    `paragraph.hyperlinks` property -- there is no built-in method to add
    one. This builds the required <w:hyperlink> XML directly, which is
    the standard, documented workaround for this library.
    """
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run_element = OxmlElement("w:r")
    run_props = OxmlElement("w:rPr")

    font_el = OxmlElement("w:rFonts")
    font_el.set(qn("w:ascii"), FONT_NAME)
    run_props.append(font_el)

    if color:
        color_el = OxmlElement("w:color")
        color_el.set(qn("w:val"), color)
        run_props.append(color_el)
    if underline:
        underline_el = OxmlElement("w:u")
        underline_el.set(qn("w:val"), "single")
        run_props.append(underline_el)

    run_element.append(run_props)
    text_el = OxmlElement("w:t")
    text_el.text = text
    run_element.append(text_el)
    hyperlink.append(run_element)
    paragraph._p.append(hyperlink)
    return hyperlink


def set_cell_background(cell, hex_color):
    """Shade a table cell. Uses w:shd directly since python-docx has no
    high-level API for cell shading."""
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def fix_zoom_settings(doc):
    """
    python-docx's default template ships a <w:zoom> element in
    settings.xml missing the schema-required w:percent attribute -- this
    is a pre-existing library quirk, confirmed by running the OOXML
    validator against a completely blank python-docx document and seeing
    the identical error. Word itself opens the file fine either way, but
    this one-line fix gets the document to a clean strict-schema
    validation pass, worth doing given this file is meant for an
    interviewer-facing demo.
    """
    settings = doc.settings.element
    zoom = settings.find(qn("w:zoom"))
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")


# ─────────────────────────────────────────────────────────────────────────
# DOCUMENT BUILDING
# ─────────────────────────────────────────────────────────────────────────

def set_default_style(doc):
    """Arial throughout -- universally supported, avoids font-substitution
    inconsistencies across machines/viewers."""
    style = doc.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = Pt(11)


def add_title_section(doc, deal_count):
    doc.add_heading("FMCG M&A Intelligence Newsletter", level=0)
    p = doc.add_paragraph()
    p.add_run(f"Issue: {datetime.now().strftime('%B %Y')}  |  {deal_count} deal(s) tracked").italic = True


def add_funnel_table(doc):
    doc.add_heading("How this issue was assembled", level=1)
    doc.add_paragraph(
        "Every article below passed through an automated screening pipeline: "
        "RSS ingestion, near-duplicate removal, FMCG-relevance filtering, and "
        "source-credibility scoring, before reaching this newsletter."
    )

    counts = get_funnel_counts()
    raw_count = counts[0][1]  # used as the denominator for "% retained"

    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    header_cells = table.rows[0].cells
    for i, heading in enumerate(["Pipeline stage", "Articles", "% of raw ingested"]):
        header_cells[i].text = heading
        header_cells[i].paragraphs[0].runs[0].font.bold = True
        set_cell_background(header_cells[i], "D9D9D9")

    for label, count in counts:
        row_cells = table.add_row().cells
        row_cells[0].text = label
        row_cells[1].text = str(count) if count is not None else "N/A"
        if count is not None and raw_count:
            row_cells[2].text = f"{count / raw_count * 100:.0f}%"
        else:
            row_cells[2].text = "N/A"

    doc.add_paragraph()  # spacing after table


def add_deal_entry(doc, row):
    headline_acquirer = str(row.get("acquirer", UNKNOWN_PLACEHOLDER))
    headline_target = str(row.get("target", UNKNOWN_PLACEHOLDER))

    p = doc.add_paragraph()
    if headline_acquirer != UNKNOWN_PLACEHOLDER and headline_target != UNKNOWN_PLACEHOLDER:
        p.add_run(f"{headline_acquirer} \u2192 {headline_target}").bold = True
    else:
        p.add_run(clean_title(row["title"])).bold = True

    # Deal type + value tag line
    tag_p = doc.add_paragraph()
    deal_value = str(row.get("deal_value", UNKNOWN_PLACEHOLDER))
    tag_text = str(row.get("deal_type", "Other"))
    if deal_value != UNKNOWN_PLACEHOLDER:
        tag_text += f"  \u2022  {deal_value}"
    tag_run = tag_p.add_run(tag_text)
    tag_run.italic = True
    tag_run.font.size = Pt(10)
    tag_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # Synopsis
    synopsis = str(row.get("synopsis", "")).strip() or clean_title(row["title"])
    doc.add_paragraph(synopsis)

    # Source line: publisher, credibility badge, date, link
    source_p = doc.add_paragraph()
    publisher = str(row.get("publisher", "unknown")).title()
    source_p.add_run(f"Source: {publisher}  \u2022  ").font.size = Pt(9)

    credibility = str(row.get("credibility", "MEDIUM")).upper()
    badge_run = source_p.add_run(credibility)
    badge_run.bold = True
    badge_run.font.size = Pt(9)
    badge_run.font.color.rgb = CREDIBILITY_COLORS.get(credibility, RGBColor(0x55, 0x55, 0x55))

    pub_dt = parse_published_date(row.get("published", ""))
    date_str = pub_dt.strftime("%d %b %Y") if pub_dt.year > 1 else "date unknown"
    source_p.add_run(f"  \u2022  {date_str}  \u2022  ").font.size = Pt(9)

    if str(row.get("url", "")).startswith("http"):
        add_hyperlink(source_p, str(row["url"]), "Read full article")

    doc.add_paragraph()  # spacing between entries


# ─────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 70)
    print("STAGE 6 - NEWSLETTER GENERATION")
    print("=" * 70 + "\n")

    if not os.path.exists(INPUT_FILE):
        print(f"Input file not found: {INPUT_FILE}")
        print("Run 05_summarize.py first.\n")
        return

    df = pd.read_csv(INPUT_FILE)
    if df.empty:
        print("Input file has 0 rows -- nothing to put in the newsletter.\n")
        return

    df["title"] = df["title"].fillna("")
    df["_published_dt"] = df["published"].apply(parse_published_date)
    df["_credibility_rank"] = df["credibility"].map(CREDIBILITY_RANK).fillna(0)
    df["deal_type"] = df["deal_type"].fillna("Other")

    doc = Document()
    fix_zoom_settings(doc)
    set_default_style(doc)
    add_title_section(doc, len(df))
    add_funnel_table(doc)

    section_counts = {}
    for deal_type in DEAL_TYPE_ORDER:
        subset = df[df["deal_type"] == deal_type].sort_values(
            ["_credibility_rank", "_published_dt"], ascending=[False, False]
        )
        section_counts[deal_type] = len(subset)
        if subset.empty:
            continue
        doc.add_heading(f"{deal_type} ({len(subset)})", level=1)
        for _, row in subset.iterrows():
            add_deal_entry(doc, row)

    # Methodology footer -- ties the deliverable back to the pipeline
    # narrative for the interviewer demo.
    doc.add_heading("Methodology", level=1)
    doc.add_paragraph(
        "This newsletter is generated by an automated pipeline: RSS ingestion with "
        "a 30-day recency filter, semantic near-duplicate removal, a rule-based "
        "FMCG-relevance filter with optional LLM escalation for ambiguous cases, "
        "publisher-credibility scoring, and AI-assisted deal summarization. "
        "Credibility tiers and source attribution are shown for every item so a "
        "reader can judge each deal's reliability independently."
    )

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    doc.save(OUTPUT_FILE)

    # ── Run summary printed to console ─────────────────────────────────
    print("-" * 70)
    print("NEWSLETTER SUMMARY")
    print("-" * 70)
    print(f"Total deals in this issue                     : {len(df)}")
    print("Breakdown by deal type:")
    for deal_type in DEAL_TYPE_ORDER:
        if section_counts.get(deal_type, 0) > 0:
            print(f"  {deal_type:<12}: {section_counts[deal_type]}")
    print(f"Output file                                   : {OUTPUT_FILE}")
    print("-" * 70 + "\n")


if __name__ == "__main__":
    main()