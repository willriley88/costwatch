# costwatch

Input-cost review tooling built entirely on public US federal data. No API keys,
no paid data tier, no proprietary feeds.

**Status: working prototype, 2026-09-27.** Generates a real multi-sheet XLSX from
live data. Not published, not sold, no buyer contacted.

## What it does

`src/build_workbook.py` produces an Excel workbook with four sheets:

| Sheet | Contents |
|---|---|
| **Cost Summary** | every tracked input, current level, 3-month and 12-month change, colour-coded |
| **History** | 60 months of aligned monthly history plus a line chart |
| **BOM Impact** | your bill of materials, with each input's index move applied to your own baseline unit cost |
| **Sources** | series ID, publisher, access terms and a FRED URL for every number |

The **BOM Impact** sheet is the actual product thesis. Anyone can look up a PPI
index; nobody wants an index. A cost accountant wants "your material cost is up
$1.59 a unit and copper did most of it."

Verified run against `data/example_bom.csv` (a 6-component product):

```
copper wiring harness        +32.99%   $2.56 -> $3.40   +$0.84
steel mounting bracket       +14.35%   $3.70 -> $4.23   +$0.53
injection-moulded housing     +6.04%   $2.72 -> $2.88   +$0.16
adhesive and solvents        +13.30%   $0.49 -> $0.56   +$0.07
plant electricity per unit    +3.16%   $0.69 -> $0.71   +$0.02
corrugated shipping carton    -3.48%   $0.94 -> $0.91   -$0.03
```

## Usage

```bash
python3 -m venv .venv && ./.venv/bin/pip install openpyxl

./.venv/bin/python src/build_workbook.py                                  # all series
./.venv/bin/python src/build_workbook.py --category metals energy          # subset
./.venv/bin/python src/build_workbook.py --bom data/example_bom.csv \
                                         --out review.xlsx                # with BOM
./.venv/bin/python src/build_workbook.py --refresh                        # bypass 12h cache
```

BOM CSV format:

```csv
component,input_series,qty_per_unit,unit_cost_usd,baseline_date
housing,WPU072,0.85,3.20,2025-09-01
```

Cold run: ~15s for 14 series. Warm run off cache: 0.14s.

## Data access — verified 2026-09-27

Keyless and working:

- **FRED CSV** (`fredgraph.csv?id=SERIES`) — the workhorse. Serves BLS PPI/CPI,
  EIA energy and market prices through one keyless interface. 16/17 candidate
  series resolved on the first try.
- **BLS Public Data API v2** — POST JSON, keyless (25 queries/day unregistered).
- **CFTC Commitments of Traders** — plain annual ZIPs, keyless, 191 columns.

Confirmed to **require registration**, so deliberately unused:

- EIA API v2 → `403`
- USDA NASS QuickStats → `401`
- USDA AMS marsapi → `403`

### Trap: FRED blackholes unrecognised User-Agents

Do not set a custom UA on FRED requests. An unrecognised UA is silently
blackholed — TCP connects, TLS completes, then the read hangs to timeout. No 403,
no 429, so it looks exactly like a network fault or a dead series.

Measured on the same host within the same second, two rounds:

| User-Agent | Result |
|---|---|
| *(urllib default, no header)* | 200 in 0.0–0.3s |
| `curl/8.5.0` | 200 in 0.0–0.1s |
| `python-requests/2.32` | 200 in 0.0–0.1s |
| `costwatch/0.1` | TimeoutError, every time |
| `Mozilla/5.0 (X11; Linux x86_64) ...` | TimeoutError, every time |

A browser-looking UA is *worse* than none — the TLS fingerprint contradicts the
claim. `series.py` sends no UA header and throttles to ~1 req/sec.

## Positioning

The line this lives or dies on: **selling grain basis to grain people walks
straight into Barchart.** Putting public industrial input-cost series into a
*non-agricultural* manufacturer's cost accountant's spreadsheet does not. Keep
the catalog pointed at resin, steel, nonferrous, packaging, chemicals and energy.

Incumbents (Barchart, Expana, S&P) sell the enterprise tier quote-only. Nobody
sells self-serve at the low end.

## Not done

- No Excel custom functions / add-in yet. This emits a workbook; the add-in
  (AppSource / Google Workspace Marketplace) is the distribution step and is a
  separate build.
- No buyer contacted. No listing. No pricing published.
- Series catalog is 14 entries and hand-curated. That curation *is* the value —
  a cost accountant does not know `WPU072` means plastic resin — but it needs to
  grow with real buyer input rather than guesses.
