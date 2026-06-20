# FMCG M&A Intelligence Newsletter

An automated pipeline that aggregates FMCG (fast-moving consumer goods) deal news from public RSS sources and produces a structured, business-readable newsletter on recent M&A and investment activity.

## What it does

1. **Ingest** — pulls articles from RSS feeds (Google News search feeds + direct publisher feeds), keeping only articles published in the last 30 days.
2. **Deduplicate** — removes exact and near-duplicate articles, since the same deal is often reported by multiple outlets.
3. **Filter for relevance** — keeps only articles that are genuinely about an FMCG-sector deal, using a rule-based keyword check with an optional LLM escalation step for ambiguous cases.
4. **Score credibility** — rates each source (publisher reputation + headline signal) and keeps medium/high-credibility articles only.
5. **Summarize** — extracts deal type and deal value deterministically; extracts acquirer, target, and a one-line synopsis via a local LLM call (optional).
6. **Generate newsletter** — compiles the final deals into a structured Word document, grouped by deal type and sorted by credibility and recency.

A Streamlit app is included to run the pipeline and browse results interactively.

## Architecture

See [`architecture.md`](architecture.md) for the pipeline diagram.

## Project structure

```
.
├── .env.example
├── .gitignore
├── README.md
├── architecture.md
├── requirements.txt
├── app/
│   └── streamlit_app.py        # interactive dashboard to run the pipeline and browse results
├── data/                        # pipeline outputs (raw → processed → relevant → credible → summarized)
└── src/
    ├── 01_ingest.py
    ├── 02_deduplicate.py
    ├── 03_relevance.py
    ├── 04_credibility.py
    ├── 05_summarize.py
    └── 06_newsletter.py
```

Each script in `src/` reads the previous stage's CSV output from `data/` and writes its own output back to `data/`, so the pipeline can be run stage-by-stage or end-to-end.

## Setup

```bash
conda create -n fmcg_ai_mna_news python=3.11
conda activate fmcg_ai_mna_news
pip install -r requirements.txt
```

Copy `.env.example` to `.env` and fill in any required values.

## Running the pipeline

Run each stage in order from the project root:

```bash
python src/01_ingest.py
python src/02_deduplicate.py
python src/03_relevance.py
python src/04_credibility.py
python src/05_summarize.py
python src/06_newsletter.py
```

Or run the Streamlit app, which can trigger each stage and shows the pipeline funnel and final deal table:

```bash
streamlit run app/streamlit_app.py
```

This app runs locally; it is not deployed to a hosted URL.

## LLM usage (optional)

Two stages — relevance filtering and summarization — support an optional local LLM call for the parts of the task that need reasoning rather than pattern matching:

```bash
set USE_LLM=true        # Windows
export USE_LLM=true     # macOS/Linux
```

This requires [Ollama](https://ollama.com) running locally with a model pulled (e.g. `ollama pull phi3.5`). With `USE_LLM=false` (the default), the pipeline runs entirely deterministically — keyword rules for relevance, regex for deal value/type — with no LLM dependency.

## Output

- `data/raw_news.csv` — all articles ingested in the last 30 days
- `data/processed_news.csv` — after deduplication
- `data/relevant_news.csv` — after relevance filtering
- `data/credible_news.csv` — after credibility scoring
- `data/summarized_news.csv` — structured deal records (acquirer, target, deal type, deal value, synopsis)
- `data/newsletter_YYYY_MM.docx` — the final newsletter
