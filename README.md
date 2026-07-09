# Deal Dashboard

Internal Streamlit app for tracking active/owned CRE deals (as opposed to `acquisitions`,
which screens incoming OMs pre-acquisition). Shows per property: cash on hand, equity
balances, debt/loan balances, distributions, contributions, budget vs. actuals, and
(once the underlying models exist) leasing/investment outlook, projected waterfall, and
sources & uses.

## Setup

```
pip install -r requirements.txt
```

Create `.streamlit/secrets.toml` (gitignored):

```toml
APP_PASSWORD = "..."
ANTHROPIC_API_KEY = "..."   # only needed to run tools/extract_loan_abstract.py
```

Run:

```
streamlit run app.py
```

## Adding a new property

1. Copy `data/TEMPLATE/config.yaml` to `data/<property_code>/config.yaml` and fill it in —
   ownership tiers, investor splits, loan/JV document references, and `sheet_map` (the
   logical-name → actual-tab-name mapping, since tab names drift between quarters).
2. Drop the property's quarterly distribution workbook at
   `data/<property_code>/source_files/distribution_workbook.xlsx` (gitignored — never commit
   this, it holds LP/GP-sensitive financial data).
3. The property is auto-discovered on next app load — no code changes needed.

## Data refresh

Manual for now: replace the file under `source_files/` each quarter, or use the in-app
uploader override. Automated ingestion is not yet built.

## Loan & JV abstracts

Full loan/JV legal documents are narrative PDFs, not recurring statements — they're parsed
**once** into a small structured abstract rather than re-parsed on every dashboard load:

```
python tools/extract_loan_abstract.py <doc.pdf> --property <code> --doc-type loan --tranche "Note A" --out data/<code>/abstracts/note_a_loan_agreement
```

This requires `ANTHROPIC_API_KEY`. The running dashboard itself has no Claude API dependency —
it only reads whatever `.json` abstracts already exist under `data/<code>/abstracts/`.

## Known workbook quirks (see `pipeline/parsers/`)

- `Equity - LP` / `Equity - BHC`: only the left balance-sheet block is parsed; the second,
  structurally-inconsistent block (right side of each tab) is not.
- `Rev Labs Interest`-style debt tabs: the "Interest Rate - Note X" labeled rows near the top
  actually hold dollar balances; the real rates are a few rows below, after the `SOFR` row.
  `debt.py` anchors by position, not by re-searching that label text.
- Tab names drift between quarters (e.g. `2026Budget` vs `2026 Budget`) — always resolve
  sheet names through `PropertyConfig.sheet_map`, never hardcode a tab name in a parser call.

## Deployment

Streamlit Community Cloud, same as `ga-automation` and `acquisitions` — set secrets there
rather than committing them.
