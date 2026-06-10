"""
NYRR Volunteer Opportunity Tracker
Monitors https://www.nyrr.org/getinvolved/volunteeropportunities

For each race listed, visits its detail page and checks every individual
volunteer role. Alerts on roles that pass the filters below.

Requirements:
    pip install playwright
    playwright install chromium

Schedule via Windows Task Scheduler to run every 30-60 minutes.
"""

import asyncio
import json
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# ── Configuration ─────────────────────────────────────────────────────────────

VOLUNTEER_URL = "https://www.nyrr.org/getinvolved/volunteeropportunities"

STATE_FILE = Path(__file__).parent / "nyrr_seen_opportunities.json"
LOG_FILE   = Path(__file__).parent / "nyrr_tracker.log"
ALERT_FILE = Path(__file__).parent / "nyrr_new_alerts.txt"

# ── Toggle ────────────────────────────────────────────────────────────────────
# True  → only include roles that count for 9+1 (NYC Marathon guaranteed entry)
# False → include all volunteer roles regardless of credit
NINE_PLUS_ONE_ONLY = False

import os as _os
NTFY_TOPIC = _os.environ.get("NTFY_TOPIC", "nyrr-volunteer-kz3069")

# ── Filters (always applied, regardless of mode) ──────────────────────────────
# Role names containing any of these are excluded (case-insensitive)
EXCLUDE_ROLE_KEYWORDS = [
    "volunteer leader",
    "leader in training",
    "leaders in training",
    "medical",
    "med tent",
    "first aid",
]

# ── Logging ───────────────────────────────────────────────────────────────────

def log(msg: str):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

# ── Windows notification ──────────────────────────────────────────────────────

def notify(title: str, body: str):
    """Send push notification via ntfy.sh (works everywhere, including GitHub Actions)."""
    import urllib.request as _ur
    try:
        req = _ur.Request(
            f"https://ntfy.sh/{NTFY_TOPIC}",
            data=body.encode("utf-8"),
            headers={"Title": title, "Priority": "high", "Tags": "runner,tada"},
            method="POST",
        )
        _ur.urlopen(req, timeout=10)
        log(f"Push notification sent → ntfy.sh/{NTFY_TOPIC}")
    except Exception as e:
        log(f"ntfy notification failed (non-fatal): {e}")

    # Also send Windows toast when running locally
    try:
        title_safe = title.replace("'", "''")
        body_safe  = body.replace("'", "''")
        ps = (
            f"Add-Type -AssemblyName System.Windows.Forms; "
            f"$n = New-Object System.Windows.Forms.NotifyIcon; "
            f"$n.Icon = [System.Drawing.SystemIcons]::Information; "
            f"$n.Visible = $true; "
            f"$n.ShowBalloonTip(10000, '{title_safe}', '{body_safe}', "
            f"[System.Windows.Forms.ToolTipIcon]::Info); "
            f"Start-Sleep -Seconds 10; $n.Dispose()"
        )
        subprocess.Popen(
            ["powershell", "-NonInteractive", "-Command", ps],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass

# ── State persistence ─────────────────────────────────────────────────────────

def load_seen() -> set:
    if STATE_FILE.exists():
        try:
            return set(json.loads(STATE_FILE.read_text(encoding="utf-8")).get("seen", []))
        except Exception:
            pass
    return set()

def save_seen(seen: set):
    STATE_FILE.write_text(
        json.dumps({"seen": sorted(seen), "updated": datetime.now().isoformat()}, indent=2),
        encoding="utf-8",
    )

# ── Filtering ─────────────────────────────────────────────────────────────────

def is_nine_plus_one(role_text: str) -> bool:
    lower = role_text.lower()
    return "9+1" in lower or ("credit" in lower and "9" in lower)

def is_excluded(role_text: str) -> bool:
    lower = role_text.lower()
    return any(kw in lower for kw in EXCLUDE_ROLE_KEYWORDS)

def role_id(race_name: str, race_date: str, role_name: str) -> str:
    return f"{race_name}|{race_date}|{role_name}"

# ── Scraping ──────────────────────────────────────────────────────────────────

def wait_for_page(page, timeout=90_000):
    """Navigate guard: bail if we land in the waiting room."""
    from playwright.sync_api import TimeoutError as PWTimeout
    time.sleep(4)
    url = page.url
    if "virtualcorral" in url or "waitingroom" in url.lower():
        log("Landed in virtual waiting room — try again when traffic is lower.")
        return False
    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
    except PWTimeout:
        pass  # proceed anyway
    return True


def scrape_event_list(page) -> list[dict]:
    """
    Extract the list of races from the volunteer opportunities listing page.
    Uses Playwright's native locator API (pierces shadow DOM) instead of raw JS.
    Returns: [{race_name, race_date, detail_url}, ...]
    """
    all_links = page.locator("a").all()

    events = []
    seen_urls = set()

    for link in all_links:
        try:
            text = link.inner_text().strip()
            href = link.get_attribute("href") or ""
        except Exception:
            continue

        # Match "LEARN MORE" buttons (the definitive detail page links)
        if text.upper() == "LEARN MORE" and href:
            full_url = href if href.startswith("http") else f"https://www.nyrr.org{href}"

            # Get the card text by reading the closest ancestor with date info
            # Walk up via JS from this specific element
            card_text = page.evaluate("""(el) => {
                let node = el;
                for (let i = 0; i < 10; i++) {
                    if (!node.parentElement) break;
                    node = node.parentElement;
                    const t = node.innerText || '';
                    if (t.length > 50 &&
                        /\\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\\b/i.test(t)) {
                        return t;
                    }
                }
                return node.innerText || '';
            }""", link.element_handle())

            lines = [l.strip() for l in card_text.split("\n") if l.strip()]
            race_name = next((l for l in lines if "volunteer" in l.lower()), lines[0] if lines else href)

            date_match = re.search(
                r'\b(\d{1,2})\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b',
                card_text, re.IGNORECASE
            )
            race_date = " ".join(date_match.group(0).split()) if date_match else ""

            if full_url not in seen_urls:
                seen_urls.add(full_url)
                events.append({"race_name": race_name, "race_date": race_date, "detail_url": full_url})

    # Log all found links if still 0 (for diagnosis)
    if not events:
        log("  No LEARN MORE links found. Dumping first 30 links for diagnosis:")
        for link in all_links[:30]:
            try:
                log(f"    '{link.inner_text().strip()[:60]}' → {link.get_attribute('href')}")
            except Exception:
                pass

    return events


def scrape_roles_from_detail(page, race: dict) -> list[dict]:
    """
    On an NYRR volunteer detail page, extract each role card from the
    "Assignment Options" section.

    Detail page structure (confirmed from screenshot):
      - Section heading: "Assignment Options"
      - Role cards each contain:
          * Status badge: "ALL SPOTS FILLED" (red) or absent if open
          * Role name: bold heading text
          * Credit badge: "9+1" pill (orange) if applicable

    Only returns roles that still have open spots.
    """
    roles = page.evaluate("""() => {
        const results = [];

        // Role cards are div.category-box elements inside ul.race-list
        // data-filterable-status="SOL" means all spots filled — skip these
        const cards = document.querySelectorAll('div.category-box');

        cards.forEach(card => {
            // Skip filled roles
            if (card.getAttribute('data-filterable-status') === 'SOL') return;

            // Role name is in div.category-name
            const nameEl = card.querySelector('.category-name');
            const role_name = nameEl ? nameEl.innerText.trim() : '';
            if (!role_name) return;

            // 9+1 credit badge is in span.tag-box
            const tags = Array.from(card.querySelectorAll('.tag-box'))
                              .map(t => t.innerText.trim());
            const has_credit = tags.some(t => /9\+1/i.test(t));

            results.push({ role_name, has_credit });
        });

        return results;
    }""")

    enriched = []
    for r in roles:
        role_name = r.get("role_name", "").strip()
        if not role_name:
            continue
        enriched.append({
            "race_name":  race["race_name"],
            "race_date":  race["race_date"],
            "role_name":  role_name,
            "credit":     "9+1" if r.get("has_credit") else "",
            "raw":        role_name + (" 9+1" if r.get("has_credit") else ""),
            "detail_url": race["detail_url"],
        })
    return enriched


def is_nine_plus_one_js(text: str) -> bool:
    lower = text.lower()
    return "9+1" in lower or ("credit" in lower and "9" in lower)


async def _scrape_detail_async(event: dict, context) -> list[dict]:
    """Scrape one event detail page asynchronously."""
    page = await context.new_page()
    try:
        await page.goto(event["detail_url"], timeout=60_000, wait_until="domcontentloaded")
        url = page.url
        if "virtualcorral" in url or "waitingroom" in url.lower():
            log(f"  {event['race_name']} — in queue, skipping")
            return []
        # Wait for role cards specifically — faster than networkidle
        try:
            await page.wait_for_selector("div.category-box, ul.race-list", timeout=8_000)
        except Exception:
            pass

        roles = await page.evaluate("""() => {
            const results = [];
            const cards = document.querySelectorAll('div.category-box');
            cards.forEach(card => {
                if (card.getAttribute('data-filterable-status') === 'SOL') return;
                const nameEl = card.querySelector('.category-name');
                const role_name = nameEl ? nameEl.innerText.trim() : '';
                if (!role_name) return;
                const tags = Array.from(card.querySelectorAll('.tag-box'))
                                  .map(t => t.innerText.trim());
                const has_credit = tags.some(t => /9\\+1/i.test(t));
                results.push({ role_name, has_credit });
            });
            return results;
        }""")

        enriched = []
        for r in roles:
            role_name = r.get("role_name", "").strip()
            if not role_name:
                continue
            enriched.append({
                "race_name":  event["race_name"],
                "race_date":  event["race_date"],
                "role_name":  role_name,
                "credit":     "9+1" if r.get("has_credit") else "",
                "raw":        role_name + (" 9+1" if r.get("has_credit") else ""),
                "detail_url": event["detail_url"],
            })

        log(f"  {event['race_name']} ({event['race_date']}) — {len(enriched)} available role(s)")
        return enriched

    except Exception as e:
        log(f"  {event['race_name']} — error: {e}")
        return []
    finally:
        await page.close()


async def _scrape_all_async() -> list[dict]:
    from playwright.async_api import async_playwright, TimeoutError as PWTimeout

    all_roles = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        )
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 900},
        )

        # ── Step 1: load listing page ─────────────────────────────────────────
        log("Loading listing page...")
        page = await context.new_page()
        try:
            await page.goto(VOLUNTEER_URL, timeout=90_000, wait_until="domcontentloaded")
        except PWTimeout:
            log("Listing page timed out.")
            await browser.close()
            return []

        url = page.url
        if "virtualcorral" in url or "waitingroom" in url.lower():
            log("In virtual waiting room — try again later.")
            await browser.close()
            return []
        # Wait specifically for event links (not just any <a>) before extracting
        try:
            await page.wait_for_selector('a[href*="events.nyrr.org"]', timeout=15_000)
        except Exception:
            pass

        await page.screenshot(path=str(Path(__file__).parent / "nyrr_debug_screenshot.png"))

        # Single JS call to extract all events — no per-link round trips
        raw_events = await page.evaluate("""() => {
            const results = [];
            const seen = new Set();
            document.querySelectorAll('a').forEach(a => {
                if ((a.innerText || '').trim().toUpperCase() !== 'LEARN MORE') return;
                const href = a.href || '';
                if (!href || seen.has(href)) return;
                seen.add(href);

                let node = a;
                for (let i = 0; i < 10; i++) {
                    if (!node.parentElement) break;
                    node = node.parentElement;
                    const t = node.innerText || '';
                    if (t.length > 50 &&
                        /\\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\\b/i.test(t)) break;
                }
                const cardText = node.innerText || '';
                const lines = cardText.split('\\n').map(l => l.trim()).filter(l => l.length > 2);
                const name = lines.find(l => l.toLowerCase().includes('volunteer')) || lines[0] || href;
                const dm = cardText.match(/\\b(\\d{1,2})\\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\\b/i);
                const date = dm ? dm[0].replace(/\\s+/g, ' ') : '';
                results.push({ race_name: name, race_date: date, detail_url: href });
            });
            return results;
        }""")
        events = raw_events

        await page.close()
        log(f"Found {len(events)} events — scraping detail pages in parallel...")

        # ── Step 2: scrape all detail pages concurrently (5 at a time) ────────
        sem = asyncio.Semaphore(10)

        async def guarded(event):
            async with sem:
                return await _scrape_detail_async(event, context)

        results = await asyncio.gather(*[guarded(e) for e in events])
        for r in results:
            all_roles.extend(r)

        await browser.close()

    return all_roles


def scrape_all() -> list[dict]:
    import asyncio
    return asyncio.run(_scrape_all_async())

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    log("=== NYRR Volunteer Tracker run started ===")
    mode = "9+1 credit only" if NINE_PLUS_ONE_ONLY else "all opportunities"
    log(f"Mode: {mode} | Excluding: leadership/medical roles")

    # TEST MODE — remove this block once notifications are confirmed working
    notify("NYRR Tracker: test ping", f"Script ran at {datetime.now().strftime('%I:%M %p')}. Notifications are working!")

    try:
        all_roles = scrape_all()
    except ImportError:
        log("ERROR: Playwright not installed. Run: pip install playwright && playwright install chromium")
        sys.exit(1)

    if not all_roles:
        log("No roles found — page may be down or layout changed.")
        return

    log(f"Total roles scraped across all events: {len(all_roles)}")

    # Apply filters — check both role name AND race name
    qualifying = []
    for r in all_roles:
        if is_excluded(r["role_name"]) or is_excluded(r["race_name"]):
            continue
        if NINE_PLUS_ONE_ONLY and not is_nine_plus_one_js(r["raw"]):
            continue
        qualifying.append(r)

    log(f"Qualifying roles after filters: {len(qualifying)}")

    seen = load_seen()
    new_ones = [r for r in qualifying if role_id(r["race_name"], r["race_date"], r["role_name"]) not in seen]

    # Group qualifying roles by race
    from collections import defaultdict
    by_race = defaultdict(list)
    for r in qualifying:
        by_race[(r["race_name"], r["race_date"], r["detail_url"])].append(r)

    # ── Final output ──────────────────────────────────────────────────────────
    log("─" * 60)
    if by_race:
        log("AVAILABLE QUALIFYING OPPORTUNITIES:")
        for (race_name, race_date, url), roles in by_race.items():
            log(f"  {race_name} ({race_date}) — {len(roles)} role(s) available")
    else:
        log("No qualifying opportunities available right now.")
    log("─" * 60)

    # Alert on new ones
    if new_ones:
        new_by_race = defaultdict(list)
        for r in new_ones:
            new_by_race[(r["race_name"], r["race_date"])].append(r)

        summary_lines = [
            f"• {name} ({date}): {len(roles)} role(s)"
            for (name, date), roles in new_by_race.items()
        ]
        notify(
            f"NYRR: New volunteer opportunities available!",
            "\n".join(summary_lines),
        )
        with open(ALERT_FILE, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.now().isoformat()}] NEW\n")
            f.write("\n".join(summary_lines) + "\n")
        for r in new_ones:
            seen.add(role_id(r["race_name"], r["race_date"], r["role_name"]))
    else:
        log("No new qualifying roles since last check.")

    for r in qualifying:
        seen.add(role_id(r["race_name"], r["race_date"], r["role_name"]))
    save_seen(seen)

    log("=== Run complete ===\n")


if __name__ == "__main__":
    main()
