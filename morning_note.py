#!/usr/bin/env python3
"""
Titan Wealth — Morning Note
Automated daily email: macro overview, market prices, and watchlist RNS analysis.
Runs at 7:20am UK time each weekday via GitHub Actions.
"""

import csv
import json
import logging
import os
import re
import smtplib
import time
from datetime import date, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Optional

import anthropic
import requests
import yfinance as yf
from bs4 import BeautifulSoup

# ── Logging ────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Configuration ──────────────────────────────────────────────────────────────

GMAIL_USER = "morningstocknote@gmail.com"
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
RECIPIENT = "will.peatfield@titanwci.com"
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
WORLD_NEWS_API_KEY = os.environ["WORLD_NEWS_API_KEY"]

WATCHLIST_PATH = Path(__file__).parent / "watchlist.csv"

# ── Brand colours ──────────────────────────────────────────────────────────────

DEEP_PURPLE = "#31135E"
EMPOWERED_PURPLE = "#8A3FFC"
LIGHT_PURPLE_BG = "#f8f5ff"
WHITE = "#FFFFFF"

# ── Market instruments (Yahoo Finance tickers) ─────────────────────────────────

MARKET_GROUPS = [
    {
        "name": "UK",
        "instruments": [
            {"name": "FTSE 100", "ticker": "^FTSE"},
            {"name": "FTSE 250", "ticker": "^FTMC"},
        ],
    },
    {
        "name": "Global",
        "instruments": [
            {"name": "MSCI World", "ticker": "URTH", "note": "ETF proxy"},
        ],
    },
    {
        "name": "Europe",
        "instruments": [
            {"name": "Germany DAX", "ticker": "^GDAXI"},
            {"name": "France CAC 40", "ticker": "^FCHI"},
            {"name": "EuroStoxx 50", "ticker": "^STOXX50E"},
            {"name": "EuroStoxx 600", "ticker": "^STOXX"},
        ],
    },
    {
        "name": "Americas",
        "instruments": [
            {"name": "Dow Jones", "ticker": "^DJI"},
            {"name": "S&P 500", "ticker": "^GSPC"},
            {"name": "Nasdaq 100", "ticker": "^NDX"},
        ],
    },
    {
        "name": "Asia-Pacific",
        "instruments": [
            {"name": "Nikkei 225", "ticker": "^N225"},
            {"name": "Hang Seng", "ticker": "^HSI"},
            {"name": "Kospi", "ticker": "^KS11"},
            {"name": "NSE Nifty 50", "ticker": "^NSEI"},
            {"name": "MSCI China", "ticker": "MCHI", "note": "ETF proxy"},
            {"name": "MSCI Emerging Markets", "ticker": "EEM", "note": "ETF proxy"},
        ],
    },
    {
        "name": "Commodities",
        "instruments": [
            {"name": "Brent Crude Oil", "ticker": "BZ=F"},
            {"name": "Gold", "ticker": "GC=F"},
            {"name": "Silver", "ticker": "SI=F"},
            {"name": "Copper", "ticker": "HG=F"},
        ],
    },
    {
        "name": "FX",
        "instruments": [
            {"name": "GBP/USD", "ticker": "GBPUSD=X"},
            {"name": "GBP/EUR", "ticker": "GBPEUR=X"},
        ],
    },
]

# Announcement categories that almost never move a share price — quick pre-filter
EXCLUDED_CATEGORIES = {
    "holding(s) in company",
    "total voting rights",
    "director/pdmr shareholding",
    "director shareholding",
    "pdmr shareholding",
    "director dealing",
    "rule 8",
    "form 8",
    "form 8.3",
    "disclosure of interests",
    "change of director details",
    "pif disclosure",
    "notification of major holdings",
}

# ── Watchlist ──────────────────────────────────────────────────────────────────


def load_watchlist() -> list[dict]:
    """Load watchlist CSV — only UK-listed stocks with real tickers."""
    stocks, seen = [], set()
    with open(WATCHLIST_PATH, newline="") as f:
        for row in csv.DictReader(f):
            ticker = row["ticker"].strip()
            exchange = row["exchange"].strip()
            if ticker == "N/A" or exchange not in ("LSE", "AIM", "TISE"):
                continue
            if ticker not in seen:
                seen.add(ticker)
                stocks.append(
                    {
                        "ticker": ticker,
                        "company": row["company"].strip(),
                        "exchange": exchange,
                    }
                )
    log.info(f"Watchlist: {len(stocks)} UK-listed stocks loaded")
    return stocks


# ── Macro news ─────────────────────────────────────────────────────────────────


def fetch_macro_news() -> str:
    """Pull today's top macro/market stories from worldnewsapi.com."""
    today = date.today().isoformat()
    base_params = {
        "api-key": WORLD_NEWS_API_KEY,
        "text": (
            "economy OR markets OR interest rates OR inflation OR GDP "
            "OR central bank OR Federal Reserve OR Bank of England OR ECB"
        ),
        "language": "en",
        "sort": "relevance",
        "sort-direction": "DESC",
        "number": 8,
        "earliest-publish-date": today,
    }
    for country in ("gb", None):
        params = dict(base_params)
        if country:
            params["country"] = country
        try:
            resp = requests.get(
                "https://api.worldnewsapi.com/search-news",
                params=params,
                timeout=15,
            )
            resp.raise_for_status()
            articles = resp.json().get("news", [])
            if articles:
                chunks = []
                for a in articles[:8]:
                    title = a.get("title", "")
                    body = a.get("text", "")[:600]
                    if title:
                        chunks.append(f"Title: {title}\nSummary: {body}")
                return "\n\n---\n\n".join(chunks)
        except Exception as e:
            log.warning(f"worldnewsapi fetch failed ({country=}): {e}")
    return ""


# ── Market prices ──────────────────────────────────────────────────────────────


def fetch_market_prices() -> dict[str, dict]:
    """Download last-close prices and day-over-day % change via yfinance."""
    all_tickers = [
        inst["ticker"] for g in MARKET_GROUPS for inst in g["instruments"]
    ]
    prices: dict[str, dict] = {t: {"price": None, "change_pct": None} for t in all_tickers}

    try:
        raw = yf.download(
            all_tickers,
            period="5d",       # 5 days to account for weekends / holidays
            interval="1d",
            progress=False,
            auto_adjust=True,
        )
        close = raw["Close"]
        for ticker in all_tickers:
            try:
                series = close[ticker].dropna()
                if len(series) >= 2:
                    prev, curr = float(series.iloc[-2]), float(series.iloc[-1])
                    prices[ticker] = {
                        "price": curr,
                        "change_pct": ((curr - prev) / prev) * 100,
                    }
                elif len(series) == 1:
                    prices[ticker] = {"price": float(series.iloc[-1]), "change_pct": None}
            except Exception:
                pass
    except Exception as e:
        log.error(f"yfinance batch download failed: {e} — falling back to individual")
        for ticker in all_tickers:
            try:
                hist = yf.Ticker(ticker).history(period="5d")
                series = hist["Close"].dropna()
                if len(series) >= 2:
                    prev, curr = float(series.iloc[-2]), float(series.iloc[-1])
                    prices[ticker] = {
                        "price": curr,
                        "change_pct": ((curr - prev) / prev) * 100,
                    }
                elif len(series) == 1:
                    prices[ticker] = {"price": float(series.iloc[-1]), "change_pct": None}
            except Exception as e2:
                log.warning(f"Price fetch failed for {ticker}: {e2}")

    fetched = sum(1 for v in prices.values() if v["price"] is not None)
    log.info(f"Prices fetched: {fetched}/{len(all_tickers)}")
    return prices


# ── RNS scraping (Investegate) ─────────────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-GB,en;q=0.9",
}


def fetch_rns_for_ticker(ticker: str, company: str, today: date) -> list[dict]:
    """Fetch today's RNS announcements for a single ticker from Investegate."""
    url = f"https://www.investegate.co.uk/{ticker}.L/"
    announcements = []
    try:
        resp = requests.get(url, headers=HEADERS, timeout=12)
        if resp.status_code == 404:
            # Try without .L suffix
            resp = requests.get(
                f"https://www.investegate.co.uk/{ticker}/", headers=HEADERS, timeout=12
            )
        if not resp.ok:
            return []

        soup = BeautifulSoup(resp.text, "lxml")
        today_str = today.strftime("%d/%m/%Y")

        # Investegate news rows are usually in a table; date appears in a cell
        for row in soup.find_all("tr"):
            cells = row.find_all("td")
            if len(cells) < 3:
                continue
            row_text = row.get_text(" ", strip=True)
            if today_str not in row_text:
                continue

            # Extract headline link
            link_tag = None
            for cell in cells:
                a = cell.find("a", href=True)
                if a and "/RegulatoryNews/" in a.get("href", ""):
                    link_tag = a
                    break

            if not link_tag:
                continue

            headline = link_tag.get_text(strip=True)
            href = link_tag["href"]
            full_url = (
                f"https://www.investegate.co.uk{href}"
                if href.startswith("/")
                else href
            )
            category = cells[-1].get_text(strip=True) if cells else ""

            # Quick exclusion based on category text
            if any(ex in category.lower() for ex in EXCLUDED_CATEGORIES):
                continue

            announcements.append(
                {
                    "ticker": ticker,
                    "company": company,
                    "headline": headline,
                    "category": category,
                    "url": full_url,
                }
            )
    except Exception as e:
        log.debug(f"Investegate fetch for {ticker}: {e}")

    return announcements


def fetch_all_rns(watchlist: list[dict]) -> list[dict]:
    """
    Fetch RNS for all watchlist stocks. Uses the LSE full-day feed first,
    then falls back to per-ticker Investegate requests.
    """
    today = date.today()
    announcements = []

    # ── Approach 1: LSE news API (JSON) ───────────────────────────────────────
    # The LSE website uses this internal endpoint — may change without notice.
    lse_succeeded = False
    try:
        since = datetime.combine(today, datetime.min.time()).strftime(
            "%Y-%m-%dT07:00:00"
        )
        resp = requests.get(
            "https://api.londonstockexchange.com/api/gw/lse/newsapi/v1/news",
            params={
                "categories": "regulatory-news-service",
                "pageSize": 200,
                "from": since,
            },
            headers=HEADERS,
            timeout=15,
        )
        if resp.ok:
            data = resp.json()
            items = data.get("news", data.get("items", []))
            if items:
                watch_map = {s["ticker"].upper(): s for s in watchlist}
                for item in items:
                    ticker_raw = (
                        item.get("ticker", "")
                        or item.get("instrument", {}).get("tidm", "")
                    ).upper()
                    if ticker_raw not in watch_map:
                        continue
                    stock = watch_map[ticker_raw]
                    category = item.get("category", item.get("subCategory", ""))
                    if any(ex in category.lower() for ex in EXCLUDED_CATEGORIES):
                        continue
                    announcements.append(
                        {
                            "ticker": ticker_raw,
                            "company": stock["company"],
                            "headline": item.get("headline", item.get("title", "")),
                            "category": category,
                            "url": item.get("url", ""),
                        }
                    )
                lse_succeeded = True
                log.info(
                    f"LSE API: {len(items)} total RNS, "
                    f"{len(announcements)} matched watchlist"
                )
    except Exception as e:
        log.info(f"LSE API unavailable: {e}")

    # ── Approach 2: per-ticker Investegate (fallback) ──────────────────────────
    if not lse_succeeded:
        log.info("Falling back to per-ticker Investegate scraping …")
        for stock in watchlist:
            items = fetch_rns_for_ticker(stock["ticker"], stock["company"], today)
            announcements.extend(items)
            if items:
                log.info(f"  {stock['ticker']}: {len(items)} announcement(s)")
            time.sleep(0.3)  # polite rate limit

        log.info(
            f"Investegate: {len(announcements)} watchlist announcements found"
        )

    return announcements


# ── AI analysis ────────────────────────────────────────────────────────────────

ai_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def fetch_announcement_text(url: str) -> str:
    """Fetch the body text of an RNS announcement page."""
    if not url:
        return ""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=12)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
        # Try known content containers across Investegate / LSE
        for selector in [
            ("div", {"class": "announcementBody"}),
            ("div", {"id": "announcement-body"}),
            ("div", {"class": "news-detail"}),
            ("article", {}),
        ]:
            tag, attrs = selector
            el = soup.find(tag, attrs)
            if el:
                return el.get_text("\n", strip=True)[:4000]
        # Last resort — strip everything
        return soup.get_text("\n", strip=True)[:4000]
    except Exception as e:
        log.debug(f"Could not fetch announcement text from {url}: {e}")
        return ""


def analyse_announcement(ann: dict) -> Optional[dict]:
    """
    Ask Claude whether this RNS is price-moving and, if so, to summarise it.
    Returns enriched dict or None if not price-moving.
    """
    body = fetch_announcement_text(ann["url"])

    prompt = f"""You are a UK equity analyst reading an RNS announcement.

Company: {ann['company']} ({ann['ticker']})
Headline: {ann['headline']}
Category: {ann['category']}
Full announcement text:
{body or '(not available — use headline and category only)'}

TASK:
1. Decide whether this announcement is likely to materially move the share price.
   - INCLUDE: trading updates, results (full year / interim / quarter), profit warnings, \
guidance changes, acquisitions/disposals, contract wins/losses, CEO changes, dividend changes, \
strategic updates, capital raises, material regulatory or legal developments.
   - EXCLUDE: routine director shareholding notifications, total voting rights updates, \
holdings in company notifications, Rule 8 / Form 8 disclosures, and other admin filings \
that convey no new business information.

2. If NOT price-moving, reply with exactly: NOT_PRICE_MOVING

3. If IS price-moving, reply with ONLY valid JSON (no markdown, no explanation):
{{
  "summary": "2–3 clear sentences on what the news says and what it means for the stock",
  "impact_direction": "positive" | "negative" | "neutral",
  "impact_magnitude": "high" | "medium" | "low",
  "key_metrics": "key numbers / figures mentioned, or empty string if none"
}}"""

    try:
        msg = ai_client.messages.create(
            model="claude-opus-4-6",
            max_tokens=512,
            messages=[{"role": "user", "content": prompt}],
        )
        text = msg.content[0].text.strip()

        if text == "NOT_PRICE_MOVING":
            return None

        result = json.loads(text)
        return {**ann, **result}

    except json.JSONDecodeError:
        log.warning(
            f"Claude returned non-JSON for {ann['ticker']}: {text[:120]!r}"
        )
        return None
    except Exception as e:
        log.error(f"Claude analysis failed for {ann['ticker']}: {e}")
        return None


def generate_macro_paragraph(raw_news: str) -> str:
    """Ask Claude to turn raw news snippets into a single macro summary paragraph."""
    if not raw_news:
        return (
            "No major macro news available this morning. "
            "Markets are expected to open quietly."
        )

    prompt = f"""You are writing the opening paragraph of a morning market note \
for UK investment professionals at Titan Wealth.

Based on the news snippets below, write ONE concise paragraph (4–6 sentences) covering \
the most important macro-economic and market developments with a UK-centric perspective. \
Focus on what matters for UK investors. Be clear and direct — no sensationalism, no fluff.

News snippets:
{raw_news}

Write ONLY the paragraph. No heading, no bullet points, no preamble."""

    try:
        msg = ai_client.messages.create(
            model="claude-opus-4-6",
            max_tokens=350,
            messages=[{"role": "user", "content": prompt}],
        )
        return msg.content[0].text.strip()
    except Exception as e:
        log.error(f"Macro paragraph generation failed: {e}")
        return "Unable to generate macro summary this morning."


# ── Email HTML ─────────────────────────────────────────────────────────────────


def _fmt_price(price: Optional[float], ticker: str) -> str:
    if price is None:
        return "—"
    if "=X" in ticker:        # FX
        return f"{price:.4f}"
    if "=F" in ticker:        # Futures / commodities
        return f"{price:,.2f}"
    return f"{price:,.2f}"    # Indices / ETFs


def _fmt_change(chg: Optional[float]) -> tuple[str, str]:
    """Returns (label, hex-colour)."""
    if chg is None:
        return "—", "#9ca3af"
    sign = "+" if chg >= 0 else ""
    colour = "#16a34a" if chg > 0 else ("#dc2626" if chg < 0 else "#6b7280")
    return f"{sign}{chg:.2f}%", colour


def _market_table_html(prices: dict) -> str:
    rows = ""
    for group in MARKET_GROUPS:
        # Group header row
        rows += (
            f'<tr style="background:{LIGHT_PURPLE_BG};">'
            f'<td colspan="3" style="padding:9px 14px 5px; color:{DEEP_PURPLE}; '
            f'font-weight:700; font-size:11px; text-transform:uppercase; '
            f'letter-spacing:0.09em; border-top:2px solid {EMPOWERED_PURPLE};">'
            f'{group["name"]}</td></tr>'
        )
        for inst in group["instruments"]:
            t = inst["ticker"]
            p = prices.get(t, {})
            price_str = _fmt_price(p.get("price"), t)
            chg_str, chg_col = _fmt_change(p.get("change_pct"))
            note = f' <span style="color:#9ca3af; font-size:10px;">({inst.get("note","")})</span>' if inst.get("note") else ""
            rows += (
                f"<tr>"
                f'<td style="padding:7px 14px; border-bottom:1px solid #f0ecf9; '
                f'color:#1a1a2e; font-size:14px;">{inst["name"]}{note}</td>'
                f'<td style="padding:7px 14px; border-bottom:1px solid #f0ecf9; '
                f'text-align:right; color:#1a1a2e; font-size:13px; font-family:\'Courier New\',monospace;">'
                f'{price_str}</td>'
                f'<td style="padding:7px 14px; border-bottom:1px solid #f0ecf9; '
                f'text-align:right; color:{chg_col}; font-size:13px; font-weight:700; font-family:\'Courier New\',monospace;">'
                f'{chg_str}</td>'
                f"</tr>"
            )
    return rows


def _rns_cards_html(announcements: list[dict]) -> str:
    if not announcements:
        return (
            '<p style="color:#6b7280; font-style:italic; margin:0; font-size:14px;">'
            "No price-moving RNS announcements this morning for the watchlist."
            "</p>"
        )

    direction_styles = {
        "positive": ("#dcfce7", "#15803d", "▲ Positive"),
        "negative": ("#fee2e2", "#b91c1c", "▼ Negative"),
        "neutral":  ("#f3f4f6", "#374151", "◆ Neutral"),
    }
    magnitude_label = {"high": "High Impact", "medium": "Medium Impact", "low": "Low Impact"}

    cards = ""
    for ann in announcements:
        direction = ann.get("impact_direction", "neutral")
        bg, fg, label = direction_styles.get(direction, direction_styles["neutral"])
        mag = magnitude_label.get(ann.get("impact_magnitude", "low"), "")

        metrics_html = ""
        if ann.get("key_metrics"):
            metrics_html = (
                f'<p style="margin:10px 0 0; font-size:13px; color:#4b5563; '
                f'background:{LIGHT_PURPLE_BG}; border-left:3px solid {EMPOWERED_PURPLE}; '
                f'padding:7px 10px; border-radius:0 4px 4px 0;">'
                f'<strong>Key figures:</strong> {ann["key_metrics"]}</p>'
            )

        link_html = ""
        if ann.get("url"):
            link_html = (
                f'<a href="{ann["url"]}" style="font-size:12px; color:{EMPOWERED_PURPLE}; '
                f'text-decoration:none; font-weight:600;">→ Full announcement</a>'
            )

        cards += f"""
<div style="border:1px solid #e5e7eb; border-radius:8px; padding:16px 18px;
            margin-bottom:14px; background:#ffffff;
            box-shadow:0 1px 3px rgba(49,19,94,0.06);">
  <div style="display:flex; justify-content:space-between; align-items:flex-start;
              gap:10px; margin-bottom:10px; flex-wrap:wrap;">
    <div>
      <span style="font-weight:700; color:{DEEP_PURPLE}; font-size:15px;">{ann['company']}</span>
      <span style="color:{EMPOWERED_PURPLE}; font-weight:600; font-size:13px;
                  margin-left:7px;">({ann['ticker']})</span>
    </div>
    <div>
      <span style="background:{bg}; color:{fg}; font-size:10px; font-weight:700;
                  padding:3px 9px; border-radius:100px; text-transform:uppercase;
                  letter-spacing:0.06em; white-space:nowrap; display:inline-block;
                  margin-bottom:3px;">{label}</span>
      <span style="background:#ede9fe; color:#5b21b6; font-size:10px; font-weight:700;
                  padding:3px 9px; border-radius:100px; text-transform:uppercase;
                  letter-spacing:0.06em; white-space:nowrap; display:inline-block;
                  margin-left:4px;">{mag}</span>
    </div>
  </div>
  <p style="margin:0 0 8px; font-size:13px; font-weight:600; color:#374151;">{ann['headline']}</p>
  <p style="margin:0; font-size:14px; color:#1a1a2e; line-height:1.65;">{ann['summary']}</p>
  {metrics_html}
  <div style="margin-top:12px;">{link_html}</div>
</div>"""

    return cards


def build_email_html(
    macro_para: str,
    prices: dict,
    announcements: list[dict],
    today: date,
) -> str:
    date_str = today.strftime("%A %d %B %Y")
    market_rows = _market_table_html(prices)
    rns_cards = _rns_cards_html(announcements)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Morning Note — {date_str}</title>
</head>
<body style="margin:0; padding:24px 0; background:#f0edf7;
             font-family:Inter,-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,sans-serif;">

  <div style="max-width:680px; margin:0 auto; background:#ffffff;
              border-radius:10px; overflow:hidden;
              box-shadow:0 4px 24px rgba(49,19,94,0.12);">

    <!-- ── Header ─────────────────────────────────────────────────────── -->
    <div style="background:{DEEP_PURPLE}; padding:28px 32px 24px;">
      <div style="color:{EMPOWERED_PURPLE}; font-size:10px; font-weight:700;
                  letter-spacing:0.18em; text-transform:uppercase; margin-bottom:6px;">
        Titan Wealth
      </div>
      <div style="color:{WHITE}; font-size:24px; font-weight:700; letter-spacing:-0.02em;">
        Morning Note
      </div>
      <div style="color:#c4b5e0; font-size:13px; margin-top:5px;">{date_str}</div>
    </div>

    <!-- Purple accent line -->
    <div style="height:3px; background:linear-gradient(90deg, {EMPOWERED_PURPLE}, {DEEP_PURPLE});"></div>

    <!-- ── Macro Overview ─────────────────────────────────────────────── -->
    <div style="padding:28px 32px 24px; border-bottom:1px solid #ede9f8;">
      <div style="color:{EMPOWERED_PURPLE}; font-size:10px; font-weight:700;
                  letter-spacing:0.14em; text-transform:uppercase; margin-bottom:12px;">
        Macro Overview
      </div>
      <p style="margin:0; font-size:15px; color:#1a1a2e; line-height:1.75;">
        Good morning,
      </p>
      <p style="margin:12px 0 0; font-size:15px; color:#1a1a2e; line-height:1.75;">
        {macro_para}
      </p>
    </div>

    <!-- ── Market Prices ──────────────────────────────────────────────── -->
    <div style="padding:24px 32px 20px; border-bottom:1px solid #ede9f8;">
      <div style="color:{EMPOWERED_PURPLE}; font-size:10px; font-weight:700;
                  letter-spacing:0.14em; text-transform:uppercase; margin-bottom:14px;">
        Market Prices — Last Close
      </div>
      <table style="width:100%; border-collapse:collapse;">
        <thead>
          <tr style="background:{DEEP_PURPLE};">
            <th style="padding:9px 14px; text-align:left; color:{WHITE};
                       font-size:11px; font-weight:600; text-transform:uppercase;
                       letter-spacing:0.08em;">Instrument</th>
            <th style="padding:9px 14px; text-align:right; color:{WHITE};
                       font-size:11px; font-weight:600; text-transform:uppercase;
                       letter-spacing:0.08em;">Price</th>
            <th style="padding:9px 14px; text-align:right; color:{WHITE};
                       font-size:11px; font-weight:600; text-transform:uppercase;
                       letter-spacing:0.08em;">Change</th>
          </tr>
        </thead>
        <tbody>
          {market_rows}
        </tbody>
      </table>
      <p style="margin:8px 0 0; font-size:10px; color:#9ca3af;">
        MSCI World, MSCI China and MSCI EM shown as iShares ETF proxies (URTH, MCHI, EEM).
        Prices as at previous close.
      </p>
    </div>

    <!-- ── Watchlist Company News ─────────────────────────────────────── -->
    <div style="padding:24px 32px 32px;">
      <div style="color:{EMPOWERED_PURPLE}; font-size:10px; font-weight:700;
                  letter-spacing:0.14em; text-transform:uppercase; margin-bottom:16px;">
        Watchlist — Company News (RNS)
      </div>
      {rns_cards}
    </div>

    <!-- ── Footer ─────────────────────────────────────────────────────── -->
    <div style="background:{DEEP_PURPLE}; padding:18px 32px; text-align:center;">
      <p style="margin:0; color:#c4b5e0; font-size:11px; line-height:1.7;">
        This note is generated automatically for internal use only.<br>
        <span style="color:{EMPOWERED_PURPLE}; font-weight:600;">Titan Wealth Investment Management</span>
        &nbsp;·&nbsp; Powering Ambitions
      </p>
    </div>

  </div>
</body>
</html>"""


# ── Email dispatch ─────────────────────────────────────────────────────────────


def send_email(html: str, subject: str) -> None:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"Titan Wealth Morning Note <{GMAIL_USER}>"
    msg["To"] = RECIPIENT
    msg.attach(MIMEText(html, "html", "utf-8"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, RECIPIENT, msg.as_string())
    log.info(f"Email sent → {RECIPIENT}")


# ── Main ───────────────────────────────────────────────────────────────────────


def main() -> None:
    today = date.today()
    log.info(f"=== Morning Note — {today} ===")

    if today.weekday() >= 5:
        log.info("Weekend — nothing to send.")
        return

    watchlist = load_watchlist()

    log.info("Fetching macro news …")
    raw_news = fetch_macro_news()

    log.info("Fetching market prices …")
    prices = fetch_market_prices()

    log.info("Fetching RNS announcements …")
    rns_raw = fetch_all_rns(watchlist)
    log.info(f"{len(rns_raw)} candidate announcements after category filter")

    log.info("Running AI analysis on announcements …")
    analysed = []
    for ann in rns_raw:
        result = analyse_announcement(ann)
        if result:
            analysed.append(result)
        time.sleep(0.4)   # gentle rate-limit on Claude API
    log.info(f"{len(analysed)} price-moving announcements identified by Claude")

    log.info("Generating macro paragraph …")
    macro_para = generate_macro_paragraph(raw_news)

    log.info("Building email …")
    html = build_email_html(macro_para, prices, analysed, today)

    subject = f"Morning Note — {today.strftime('%d %B %Y')}"
    send_email(html, subject)
    log.info("Done ✓")


if __name__ == "__main__":
    main()
