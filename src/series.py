"""
series.py — keyless fetch and cache for public US input-cost series.

Verified 2026-09-27, no API key and no paid tier:

  * BLS Public Data API v2 — POST, JSON, keyless (rate-limited: 25 queries/day
    and 10 years/query without registration). Returns PPI/CPI levels.
  * FRED CSV endpoint (`fredgraph.csv?id=SERIES`) — keyless GET, full history.
    This is the workhorse: it serves BLS PPI, EIA energy and market prices
    through one interface with no registration at all.
  * CFTC Commitments of Traders — plain annual ZIPs, keyless.

Confirmed to REQUIRE a key (so they are deliberately not used here):
  EIA API v2 (403), USDA NASS QuickStats (401), USDA AMS marsapi (403).

The thesis this supports: an input-cost tool for a NON-agricultural
manufacturer's cost accountant. Grain basis walks into Barchart; industrial
resin, steel, energy and packaging costs in a cost accountant's spreadsheet
do not.
"""

from __future__ import annotations

import csv
import json
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path

CACHE = Path(__file__).resolve().parent.parent / "data" / "cache"
CACHE.mkdir(parents=True, exist_ok=True)
FRED_CSV = "https://fred.stlouisfed.org/graph/fredgraph.csv?id={sid}"
CACHE_TTL_S = 12 * 3600

# DO NOT set a custom User-Agent on FRED requests.
#
# fredgraph.csv is behind Akamai, and an unrecognised UA string is silently
# BLACKHOLED — the TCP connection is accepted, TLS completes, and then the read
# hangs until timeout. There is no 403 and no 429, so it reads like a network
# fault or a dead series when it is actually a bot rule.
#
# Measured 2026-09-27, same host, same second, WPU101 and PPIACO, two rounds:
#   UA ''  (python-urllib default) -> 200 in 0.0-0.3s
#   UA 'curl/8.5.0'                -> 200 in 0.0-0.1s
#   UA 'python-requests/2.32'      -> 200 in 0.0-0.1s
#   UA 'costwatch/0.1'             -> TimeoutError every time
#   UA 'Mozilla/5.0 (X11...)'      -> TimeoutError every time
#
# So: send no UA header at all and let urllib use its default. A browser-looking
# UA is WORSE than none, because the TLS fingerprint does not match the claim.
UA: str | None = None


@dataclass(frozen=True)
class Series:
    """One input-cost series, described for a cost accountant not an economist."""

    sid: str
    label: str
    unit: str
    category: str
    note: str = ""


# Curated because discoverability is the actual product. A cost accountant does
# not know that "WPU072" is plastic resin; that mapping is the value being sold.
CATALOG: dict[str, Series] = {
    s.sid: s
    for s in [
        Series("WPU101", "Iron and steel", "PPI index (1982=100)", "metals",
               "broad steel input cost; use for fabrication and structural"),
        Series("WPU1017", "Steel mill products", "PPI index", "metals",
               "narrower than WPU101 — mill output specifically"),
        Series("WPU102", "Nonferrous metals", "PPI index", "metals",
               "copper/aluminium exposure"),
        Series("WPU3012", "Fabricated structural metal", "PPI index", "metals",
               "closer to a purchased fabricated part than raw steel"),
        Series("WPU072", "Plastic resins and materials", "PPI index", "plastics",
               "the series most injection moulders actually need"),
        Series("WPU061", "Industrial chemicals", "PPI index", "chemicals", ""),
        Series("WPU0911", "Paperboard containers", "PPI index", "packaging",
               "corrugated box cost — hits every shipper"),
        Series("WPU0571", "Industrial natural gas", "PPI index", "energy", ""),
        Series("WPUIP2311001", "Industrial electric power", "PPI index", "energy", ""),
        Series("APU000072610", "Electricity, per kWh (US city avg)", "USD/kWh", "energy",
               "dollar-denominated, unlike the PPI indices"),
        Series("DCOILWTICO", "WTI crude oil", "USD/barrel", "energy",
               "daily; freight and resin lead indicator"),
        Series("PPIACO", "PPI, all commodities", "PPI index", "benchmark",
               "use as the deflator/benchmark line on every chart"),
        Series("WPU0561", "Crude petroleum", "PPI index", "energy", ""),
        Series("CUSR0000SETB01", "Gasoline (CPI, seasonally adj)", "CPI index", "energy",
               "proxy for own-fleet delivery cost"),
    ]
}


class FetchError(RuntimeError):
    pass


# FRED sits behind Akamai and will start timing out (not 429-ing) after a rapid
# burst — observed 2026-09-27 after ~17 back-to-back requests, with curl still
# fine from the same host seconds later. So throttle deliberately and retry with
# backoff rather than treating a timeout as a dead series.
_MIN_INTERVAL_S = 1.1
_last_call = 0.0


def _throttle() -> None:
    global _last_call
    wait = _MIN_INTERVAL_S - (time.monotonic() - _last_call)
    if wait > 0:
        time.sleep(wait)
    _last_call = time.monotonic()


def _cache_path(sid: str) -> Path:
    return CACHE / f"{sid}.json"


def _fetch_fred_csv(sid: str, timeout: int = 30, tries: int = 4) -> list[tuple[str, float]]:
    url = FRED_CSV.format(sid=sid)
    last: Exception | None = None
    for attempt in range(1, tries + 1):
        try:
            _throttle()
            headers = {"User-Agent": UA} if UA else {}
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as fh:
                text = fh.read().decode("utf-8", "replace")
            break
        except Exception as exc:  # noqa: BLE001
            last = exc
            if attempt == tries:
                raise FetchError(f"{sid}: {type(exc).__name__} {exc}") from exc
            time.sleep(2.0 * (2 ** (attempt - 1)))  # 2s, 4s, 8s
    else:  # pragma: no cover
        raise FetchError(f"{sid}: {last}")

    rows = list(csv.reader(text.strip().splitlines()))
    if len(rows) < 2:
        raise FetchError(f"{sid}: empty response")
    out: list[tuple[str, float]] = []
    for r in rows[1:]:
        if len(r) < 2 or r[1].strip() in (".", ""):
            continue  # FRED marks missing observations with a bare dot
        try:
            out.append((r[0].strip(), float(r[1])))
        except ValueError:
            continue
    if not out:
        raise FetchError(f"{sid}: no numeric observations")
    return out


def load(sid: str, refresh: bool = False) -> list[tuple[str, float]]:
    """Observations as [(iso_date, value)], newest last. Cached to disk."""
    path = _cache_path(sid)
    if path.exists() and not refresh:
        age = time.time() - path.stat().st_mtime
        if age < CACHE_TTL_S:
            payload = json.loads(path.read_text())
            return [(d, v) for d, v in payload["obs"]]
    obs = _fetch_fred_csv(sid)
    path.write_text(json.dumps({"sid": sid, "fetched_utc": datetime.now(timezone.utc).isoformat(), "obs": obs}))
    return obs


def latest(sid: str, refresh: bool = False) -> tuple[str, float]:
    obs = load(sid, refresh)
    return obs[-1]


def pct_change(sid: str, months: int, refresh: bool = False) -> float | None:
    """
    Percent change over the last `months` of published data.

    Uses the observation nearest the target date rather than an index offset,
    because these series mix daily, monthly and weekly frequencies and an
    offset silently means different things per series.
    """
    obs = load(sid, refresh)
    if len(obs) < 2:
        return None
    end_date = date.fromisoformat(obs[-1][0])
    y, m = end_date.year, end_date.month - months
    while m <= 0:
        m += 12
        y -= 1
    target = date(y, m, 1)
    base = min(obs, key=lambda o: abs((date.fromisoformat(o[0]) - target).days))
    if base[1] == 0:
        return None
    return (obs[-1][1] / base[1] - 1.0) * 100.0


def change_since(sid: str, baseline: str, refresh: bool = False) -> dict | None:
    """Percent change from a CALLER-SUPPLIED baseline date to the latest print.

    `pct_change` answers "how much has this moved in N months", which is the
    right question for a market summary and the WRONG one for a BOM: the user's
    baseline unit cost was struck on a specific date, and applying a fixed
    12-month move to an 18-month-old baseline understates their exposure while
    applying it to a 2-month-old baseline invents one.

    Returns the matched baseline observation too, so the workbook can show which
    date it actually priced against and how far that sits from what was asked —
    the series may simply not publish on the requested date.
    """
    obs = load(sid, refresh)
    if len(obs) < 2:
        return None
    try:
        want = date.fromisoformat(baseline)
    except (TypeError, ValueError):
        return None
    base = min(obs, key=lambda o: abs((date.fromisoformat(o[0]) - want).days))
    if base[1] == 0:
        return None
    latest_date, latest_val = obs[-1]
    drift = abs((date.fromisoformat(base[0]) - want).days)
    return {
        "pct": (latest_val / base[1] - 1.0) * 100.0,
        "baseline_requested": baseline,
        "baseline_matched": base[0],
        "baseline_value": base[1],
        "latest_date": latest_date,
        "latest_value": latest_val,
        "match_drift_days": drift,
        "span_days": (date.fromisoformat(latest_date) - date.fromisoformat(base[0])).days,
    }


def series_table(sids: list[str], refresh: bool = False) -> list[dict]:
    """One row per series with the changes a cost review actually reports."""
    rows = []
    for sid in sids:
        meta = CATALOG.get(sid) or Series(sid, sid, "?", "uncategorised")
        try:
            obs = load(sid, refresh)
            row = {
                "sid": sid,
                "label": meta.label,
                "unit": meta.unit,
                "category": meta.category,
                "as_of": obs[-1][0],
                "value": obs[-1][1],
                "obs_count": len(obs),
                "chg_3m_pct": pct_change(sid, 3, refresh=False),
                "chg_12m_pct": pct_change(sid, 12, refresh=False),
                "error": None,
            }
        except FetchError as exc:
            row = {
                "sid": sid, "label": meta.label, "unit": meta.unit, "category": meta.category,
                "as_of": None, "value": None, "obs_count": 0,
                "chg_3m_pct": None, "chg_12m_pct": None, "error": str(exc),
            }
        rows.append(row)
    return rows
