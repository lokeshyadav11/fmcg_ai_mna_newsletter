"""
app/streamlit_app.py
=====================
Demo app for the FMCG M&A Newsletter pipeline.

Run from the PROJECT ROOT (not from inside app/):
    streamlit run app/streamlit_app.py

This assumes the working directory is the project root, since every data
path below (data/raw_news.csv etc.) and every pipeline script path
(src/01_ingest.py etc.) is relative to it -- the same convention the
pipeline scripts themselves use.

WHAT THIS APP SHOWS
--------------------
1. Run Pipeline   -- trigger all 6 stages end-to-end (or one at a time),
                     with a live log and a toggle for the USE_LLM flag.
2. Pipeline Funnel -- row count at each stage, so the screening the
                     earlier stages did is visible, not just implied.
3. Deal Table      -- the final structured deals (summarized_news.csv),
                     filterable by deal type and credibility.
4. Downloads        -- CSV at every stage, plus the generated newsletter.
5. Methodology      -- brief, collapsed by default, explains dedup/
                     relevance/credibility logic for anyone reviewing.
"""

import os
import subprocess
import sys
from datetime import datetime

import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────

PIPELINE_STAGES = [
    ("01_ingest.py", "Ingesting articles from RSS feeds"),
    ("02_deduplicate.py", "Removing near-duplicate articles"),
    ("03_relevance.py", "Filtering for FMCG deal relevance"),
    ("04_credibility.py", "Scoring source credibility"),
    ("05_summarize.py", "Extracting deal details and synopsis"),
    ("06_newsletter.py", "Building the newsletter"),
]

DATA_FILES = {
    "Raw ingested": "data/raw_news.csv",
    "After deduplication": "data/processed_news.csv",
    "After relevance filter": "data/relevant_news.csv",
    "After credibility filter": "data/credible_news.csv",
    "Final (summarized)": "data/summarized_news.csv",
}

UNKNOWN_PLACEHOLDER = "Not specified"

st.set_page_config(page_title="FMCG M&A Newsletter Agent", layout="wide")


# ─────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────

def run_stage(script_name, use_llm, ollama_model, log_placeholder, log_lines):
    """
    Run one pipeline stage as a subprocess, streaming its stdout into the
    log area line by line as it runs (rather than only showing output
    after the process finishes, which would make a multi-minute LLM-
    enabled run look frozen).

    Returns True on success (exit code 0), False otherwise.
    """
    env = os.environ.copy()
    env["USE_LLM"] = "true" if use_llm else "false"
    env["OLLAMA_MODEL"] = ollama_model

    script_path = os.path.join("src", script_name)
    process = subprocess.Popen(
        [sys.executable, "-u", script_path],  # -u: unbuffered, so the log streams live
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        env=env,
        bufsize=1,
    )

    for line in process.stdout:
        log_lines.append(line.rstrip())
        log_placeholder.code("\n".join(log_lines[-25:]))  # keep the log area from growing forever

    process.wait()
    return process.returncode == 0


def get_funnel_counts():
    counts = {}
    for label, path in DATA_FILES.items():
        if os.path.exists(path):
            try:
                counts[label] = len(pd.read_csv(path))
            except Exception:
                counts[label] = None
        else:
            counts[label] = None
    return counts


def load_csv_safe(path):
    if os.path.exists(path):
        try:
            return pd.read_csv(path)
        except Exception:
            return None
    return None


# ─────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────

st.title("FMCG M&A Newsletter Agent")
st.caption("Ingest -> Deduplicate -> Relevance -> Credibility -> Summarize -> Newsletter")
st.divider()

# ─────────────────────────────────────────────────────────────────────────
# SECTION 1: RUN PIPELINE
# ─────────────────────────────────────────────────────────────────────────

st.header("1. Run Pipeline")

col1, col2 = st.columns([1, 2])
with col1:
    use_llm = st.toggle(
        "Use LLM (Ollama)",
        value=False,
        help="When on, 03_relevance.py escalates ambiguous cases and "
             "05_summarize.py extracts acquirer/target/synopsis via a "
             "local Ollama model instead of the deterministic fallback. "
             "Requires Ollama running locally.",
    )
with col2:
    ollama_model = st.text_input("Ollama model", value="phi3.5", disabled=not use_llm)

run_col1, run_col2 = st.columns(2)
with run_col1:
    run_full = st.button("Run Full Pipeline", type="primary", width="stretch")
with run_col2:
    with st.expander("Run a single stage instead"):
        stage_choice = st.selectbox(
            "Stage", options=[s[0] for s in PIPELINE_STAGES], label_visibility="collapsed"
        )
        run_single = st.button("Run Selected Stage", width="stretch")

if run_full:
    progress = st.progress(0.0)
    status = st.empty()
    log_area = st.empty()
    log_lines = []
    all_ok = True

    for i, (script, description) in enumerate(PIPELINE_STAGES):
        status.info(f"Step {i + 1}/{len(PIPELINE_STAGES)}: {description}...")
        success = run_stage(script, use_llm, ollama_model, log_area, log_lines)
        progress.progress((i + 1) / len(PIPELINE_STAGES))
        if not success:
            status.error(f"Step {i + 1} ({script}) failed -- see log above.")
            all_ok = False
            break

    if all_ok:
        status.success("Pipeline complete.")

if run_single:
    status = st.empty()
    log_area = st.empty()
    log_lines = []
    status.info(f"Running {stage_choice}...")
    success = run_stage(stage_choice, use_llm, ollama_model, log_area, log_lines)
    if success:
        status.success(f"{stage_choice} completed.")
    else:
        status.error(f"{stage_choice} failed -- see log above.")

st.divider()

# ─────────────────────────────────────────────────────────────────────────
# SECTION 2: PIPELINE FUNNEL
# ─────────────────────────────────────────────────────────────────────────

st.header("2. Pipeline Funnel")

counts = get_funnel_counts()
if any(v is not None for v in counts.values()):
    cols = st.columns(len(counts))
    for col, (label, count) in zip(cols, counts.items()):
        with col:
            st.metric(label, count if count is not None else "—")
    st.bar_chart(
        pd.DataFrame(
            [(label, count) for label, count in counts.items() if count is not None],
            columns=["Stage", "Articles"],
        ).set_index("Stage")
    )
else:
    st.info("Run the pipeline to see funnel counts here.")

st.divider()

# ─────────────────────────────────────────────────────────────────────────
# SECTION 3: DEAL TABLE
# ─────────────────────────────────────────────────────────────────────────

st.header("3. Deals in This Issue")

summarized_df = load_csv_safe(DATA_FILES["Final (summarized)"])
if summarized_df is not None and not summarized_df.empty:
    filter_col1, filter_col2, filter_col3 = st.columns(3)
    with filter_col1:
        deal_types = ["All"] + sorted(summarized_df["deal_type"].dropna().unique().tolist())
        selected_type = st.selectbox("Deal type", deal_types)
    with filter_col2:
        credibility_options = ["HIGH", "MEDIUM", "LOW"]
        selected_credibility = st.multiselect("Credibility", credibility_options, default=credibility_options)
    with filter_col3:
        search_term = st.text_input("Search (acquirer, target, or title)")

    filtered = summarized_df.copy()
    if selected_type != "All":
        filtered = filtered[filtered["deal_type"] == selected_type]
    if selected_credibility:
        filtered = filtered[filtered["credibility"].isin(selected_credibility)]
    if search_term:
        term = search_term.lower()
        mask = (
            filtered["title"].str.lower().str.contains(term, na=False)
            | filtered["acquirer"].str.lower().str.contains(term, na=False)
            | filtered["target"].str.lower().str.contains(term, na=False)
        )
        filtered = filtered[mask]

    display_cols = [c for c in [
        "acquirer", "target", "deal_type", "deal_value", "synopsis",
        "publisher", "credibility", "published",
    ] if c in filtered.columns]
    st.dataframe(filtered[display_cols], width="stretch", hide_index=True)
    st.caption(f"Showing {len(filtered)} of {len(summarized_df)} deals")
else:
    st.info("Run the pipeline to see deals here.")

st.divider()

# ─────────────────────────────────────────────────────────────────────────
# SECTION 4: DOWNLOADS
# ─────────────────────────────────────────────────────────────────────────

st.header("4. Downloads")

dl_cols = st.columns(len(DATA_FILES) + 1)
for col, (label, path) in zip(dl_cols, DATA_FILES.items()):
    with col:
        if os.path.exists(path):
            with open(path, "rb") as f:
                st.download_button(label, f, file_name=os.path.basename(path), mime="text/csv")
        else:
            st.button(label, disabled=True)

import glob
newsletters = sorted(glob.glob("data/newsletter_*.docx"), reverse=True)
with dl_cols[-1]:
    if newsletters:
        with open(newsletters[0], "rb") as f:
            st.download_button(
                "Newsletter (.docx)", f, file_name=os.path.basename(newsletters[0]),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    else:
        st.button("Newsletter (.docx)", disabled=True)

st.divider()

# ─────────────────────────────────────────────────────────────────────────
# SECTION 5: METHODOLOGY
# ─────────────────────────────────────────────────────────────────────────

with st.expander("Methodology"):
    st.markdown(
        "**Ingestion** pulls from 8 RSS feeds (3 Google News search queries + "
        "5 direct publisher feeds) and keeps only articles from the last 30 days.\n\n"
        "**Deduplication** runs two passes: an exact domain+date+title match, "
        "then semantic similarity (sentence-transformers, cosine threshold 0.75) "
        "to collapse the same deal reported by multiple outlets.\n\n"
        "**Relevance** uses a 2x2 keyword matrix (deal keyword x FMCG keyword); "
        "both present is RELEVANT, neither is NOT_RELEVANT, exactly one is "
        "UNCERTAIN and is escalated to an LLM only when `USE_LLM=true`, "
        "otherwise it's kept by an inclusive default.\n\n"
        "**Credibility** scores publisher reputation (1-3) plus two content "
        "signals (headline length, deal keyword present), mapped to LOW/"
        "MEDIUM/HIGH; only MEDIUM+ survives.\n\n"
        "**Summarization** extracts deal_type and deal_value deterministically "
        "(keyword match, regex); acquirer/target/synopsis come from a local "
        "phi3.5 call when `USE_LLM=true`, otherwise both fields show "
        f"'{UNKNOWN_PLACEHOLDER}' rather than a guessed value."
    )

st.caption(f"Last viewed: {datetime.now().strftime('%d %b %Y, %H:%M')}")