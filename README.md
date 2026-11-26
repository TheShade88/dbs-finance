DRAFT

# dbs-finance

Download a DBS bank export, double-click one file, and an Excel expense tracker is filled in. Excel formulas and rules categorise known merchants, and Claude names the rest. Before anything is sent to Claude, a local script strips out card numbers, references, balances and the names of people.

On seven months of my own statements, the formulas and seed rules left 127 rows uncategorised. Those were 107 distinct merchants, which Claude handled in one call of 1 min 47 s. Afterwards 34 rows were left for me, most of them PayNow transfers to people.

```
DBS CSV in Downloads
  -> mask card numbers (CARD-XXXX-XXXX-XXXX), skip rows already imported
  -> append to the workbook through Excel, recalculate
  -> rows the Merchant_Map rules can't place: scrub to merchant labels only, leak-check
  -> Claude picks a subcategory per merchant
  -> high confidence: new Merchant_Map rule   medium: set on the row, Remarks "AI (medium)"
     low: suggestion in Remarks only          unknown: left for you
  -> Import_Log line, save
```

## Requirements

- Windows with Microsoft Excel. The workbook is driven through Excel itself, because saving it with openpyxl drops the charts.
- Python 3.11 or later and [uv](https://docs.astral.sh/uv/).
- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) signed in to a Claude subscription. The pipeline runs `claude -p` in headless mode, so no API key is needed.

## Setup

```powershell
git clone <this repo> dbs-finance
cd dbs-finance
uv sync
claude auth login
```

Then check `config.toml`. `downloads_glob` points at `~/Downloads/transaction_history_*.csv` by default. If `claude` is not on your PATH, put the full path to `claude.cmd` under `[claude] command`. On the first run, `workbook/expense_tracker.xlsx` is created from the blank template in `template/`.

## Monthly use

1. Export the period from DBS iBanking in a desktop browser (the phone app cannot export).
2. Close the workbook in Excel and double-click `import.cmd`, or run `uv run finance run`.
3. Open the workbook. In Check_Flag, filter for NEEDS SUBCATEGORY to see what is left. In Remarks, filter for "AI (" to check what Claude set.

Overlapping exports are safe, because rows already in the workbook are skipped. A transaction has no ID in the DBS export, so duplicates are matched on date, description and amounts.

## Commands

```
uv run finance run --dry-run        import and scrub, save nothing, send nothing; writes runs/<id>/payload.json
uv run finance run                  import new exports, then categorise
uv run finance run <file.csv>       import a specific export
uv run finance categorise           categorise the backlog only
uv run finance categorise --retry   re-ask rows Claude already answered (after editing [rules])
uv run finance status               health counters
uv run pytest                       tests on an invented export; the Excel test needs Excel
```

## What leaves your machine

For each distinct merchant, Claude gets:

- the cleaned label, e.g. `MCDONALD'S ORCHARD SI SGP`
- the statement code (POS or ADV) and the direction of the money
- the typical amount and how many times the merchant appeared
- the foreign currency, if any

Each run saves exactly what was sent in `runs/<id>/payload.json`.

These are never sent:

- card numbers, reference numbers, dates and balances
- salary rows and PayNow memos
- names of private people, which go as `PERSON_07` while the real names stay in `private/people.json`
- the other party in a bank-to-bank transfer
- your employer's name, which the pipeline finds on its own from salary credits

If anything that looks like a long number, card, phone number, email or known name survives scrubbing, the run stops before sending.

Claude runs with no tools, without your user settings or hooks, and with no saved transcript. The prompt itself does go to Anthropic under your Claude account.

## How answers land in the workbook

The workbook is pure formulas. You only ever type in two columns: U (subcategory override) and Z (remarks). Claude's answers use the same two columns plus the Merchant_Map sheet.

A high-confidence answer becomes a Merchant_Map rule only if the brand text appears in every description of that merchant, contains no Excel wildcard, and would not reclassify any row already in a different subcategory. Every later month then matches it with no call to Claude, so less is sent as the map grows. To undo a wrong rule, delete its row on Merchant_Map.

Your own preferences go in `[rules] notes` in `config.toml` and are sent with every prompt. For example, the defaults say that spending outside Singapore is travel.

## Limits

- PayNow transfers to private people mostly stay unknown, because the bank text names only the person. They make up most of what is left each month.
- The parser expects the current DBS iBanking CSV layout (12 columns from the "Transaction Date" header). If DBS changes the format, the run stops with an error rather than guessing.
- Only the one account in the export is tracked. Cash is invisible once withdrawn.
- I tested two local models through Ollama (qwen2.5 7B and 3B on a 4 GB laptop GPU). They agreed with Claude on 27 and 19 of the 107 merchants, so this version uses Claude only.

MIT licence. Not affiliated with DBS Bank.
