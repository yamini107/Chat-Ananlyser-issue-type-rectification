# Chat Analyzer Dashboard · Graas.ai

Streamlit dashboard for daily chat enquiry analysis across **Shopee**, **Lazada**, and **TikTok**.

---

## What's New (Latest Version)

### 1. JSON Message Extraction
All three platforms encode messages as JSON blobs (`{"txt":"..."}`, `{"text":"..."}`, `{"content":"..."}`).
The new `extract_plain_text()` helper unwraps these before any keyword matching, fixing the root cause of most misclassifications.

### 2. TSV-Driven Issue Classification
Place `table__2_.tsv` next to `chat_analyzer_dashboard.py` **or** upload it via the sidebar at runtime.
TSV keywords are weighted ×2 vs legacy keywords, giving the reference data priority for Delay, Promotions, Size, Exchange, Invoice, and Product Quality categories.

### 3. Scoring-Based Classification
Both TSV and legacy keywords contribute weighted scores. Ties are broken by priority order (High > Medium > Low), so urgent issues like Refund / Complaint / Damaged Item are never mis-classified as low-priority.

### 4. Improved Unresolved Detection
- Pre-parsed plain-text seller messages (not raw JSON) are now evaluated.
- Auto-replies, chatbot messages, and bot greetings are excluded from resolution detection.
- Conversations with zero human seller responses are correctly marked Unresolved.

---

## Quick Start

```bash
pip install -r requirements.txt
streamlit run chat_analyzer_dashboard.py
```

## Streamlit Cloud Deployment

1. Push this repository to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Select your repo, branch `main`, file `chat_analyzer_dashboard.py`.
4. Click **Deploy**.

## File Structure

```
├── chat_analyzer_dashboard.py   # Main app
├── table__2_.tsv                # Issue type → keyword reference (optional, can upload via sidebar)
├── requirements.txt
├── .streamlit/
│   └── config.toml              # Theme + upload size config
└── README.md
```

## TSV Reference Format

The TSV must have exactly two columns:

| Issue type          | Key words              |
|---------------------|------------------------|
| Shipment status     | Ship                   |
| Shipment status     | delivery               |
| Cancellation request| cancel                 |
| Return & refund     | return                 |
| Return & refund     | refund                 |
| promotions          | discount               |
| ...                 | ...                    |

Issue type uses forward-fill (only the first row per group needs a value).

## Issue Type → Internal Name Mapping

| TSV Label            | Internal Name      | Priority |
|----------------------|--------------------|----------|
| Shipment status      | Delay              | Medium   |
| Order status         | Delay              | Medium   |
| Cancellation request | Cancellation       | Medium   |
| change request       | Cancellation       | Medium   |
| Restock              | Product Inquiry    | Low      |
| size recommendation  | Product Inquiry    | Low      |
| Return & refund      | Return / Refund    | Medium/High |
| exchange             | Return             | Medium   |
| parcel damaged       | Damaged/Wrong Item | High     |
| product query        | Product Inquiry    | Low      |
| promotions           | Promotion Issue    | Low      |
| Invoice              | Payment Issue      | Low      |
| Product quality      | Damaged/Wrong Item | High     |

## Environment Variables

| Variable               | Default                   | Description                        |
|------------------------|---------------------------|------------------------------------|
| `CHAT_ANALYZER_TSV`    | `./table__2_.tsv`         | Path to the TSV reference file     |

## Team Assignment (effective 30 March 2026)

| Agent      | Stores                              |
|------------|-------------------------------------|
| Yeria      | GED MY · AACMH · FFH · IKU         |
| Syahira    | GED SG · EWG · HFC · AAISS         |
| Keerthana  | GED PH · AABIY · AABIW · AAFTP     |
| Alfian     | IGZ ID · AADMJ · AAEDD · AADWP     |
| Jaye       | GSK · DBC · IEI · FYW · ILL        |
| Ratchakorn | AABWU · AAFHU · AAFHB              |

