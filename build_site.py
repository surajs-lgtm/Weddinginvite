#!/usr/bin/env python3
"""
Convert wedding-details.xlsx into site/data.js.

The workbook is the source of truth. Edit the spreadsheet, run this script,
and refresh the page.

Run:  python3 build_site.py
"""
import hashlib
import json
import os
import re
import sys
from datetime import datetime

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("openpyxl is required:  pip3 install openpyxl")

HERE = os.path.dirname(os.path.abspath(__file__))
BOOK = os.path.join(HERE, "wedding-details.xlsx")
OUT = os.path.join(HERE, "docs", "data.js")

FILLIN = "-- FILL IN --"

# Which drawing each printed ritual gets. Matched loosely so OCR variants and
# spelling differences on the card still find the right icon.
RITUAL_ICONS = [
    ("tilak", ["tilak", "tilak ceremony"]),
    ("matkor", ["matkor"]),
    ("madwa", ["madwa", "madhwa", "madwa homa"]),
    ("haldi", ["haldi"]),
    ("kalra", ["kalra", "devpuji", "dev pooja", "devpuj"]),
    ("dhidhari", ["dhid hari", "dhidhari", "ghee dhari", "gheehari", "dheem"]),
    ("janau", ["janau", "jaanu", "janoi"]),
    ("bhungalawa", ["bhunga lawa", "bhungla", "bhunga"]),
    ("parat", ["parat", "praat", "parath"]),
    ("lawabhujai", ["lawa bhujai", "lawabhujai", "bhujai", "bhujia", "lawa bhujia"]),
    ("jaimala", ["jaimala", "varmala", "varimala"]),
    ("sindoordaan", ["sindoor daan", "sindoor day", "sindur", "sindoor"]),
    ("janeu", ["janeu", "janeu ceremony", "janu", "yajnopavita"]),
    ("darwagar", ["darwagar", "darwached", "darwargar", "darwargarh",
                  "baraat", "baraat prasthan", "baraat prasan", "prasthan"]),
    ("vidai", ["vidai", "milap", "widaai", "vidaai"]),
]


def icon_for(event_name):
    """Best icon key for a function name, or None."""
    low = (event_name or "").lower()
    for key, needles in RITUAL_ICONS:
        for needle in needles:
            if needle in low:
                return key
    return None


ICON_DIR = os.path.join(HERE, "docs", "assets", "icons")
_icon_svg = {}


def icon_svg(key):
    """Inline a ritual icon as markup.

    The icons are stroked with currentColor so the page can theme them. That
    only works when the markup is in the document; an <img> would pin them to
    black. Inlining at build time also keeps the page working from file://,
    where fetch() of a sibling SVG would be blocked.
    """
    if not key:
        return ""
    if key in _icon_svg:
        return _icon_svg[key]
    path = os.path.join(ICON_DIR, key + ".svg")
    try:
        with open(path, "r", encoding="utf-8") as fh:
            raw = fh.read()
    except OSError:
        _icon_svg[key] = ""
        return ""
    head, _, rest = raw.partition(">")
    vb = "0 0 64 64"
    # the value contains spaces, so match the whole attribute, not a token
    found = re.search(r'viewBox\s*=\s*"([^"]+)"', head) or re.search(
        r"viewBox\s*=\s*'([^']+)'", head
    )
    if found:
        vb = found.group(1).strip()
    # inner markup: everything between the opening <svg ...> and </svg>
    inner = rest.rsplit("</svg>", 1)[0]
    inner = inner.rsplit("<", 1)[0] if not inner.rstrip().endswith(">") else inner
    _icon_svg[key] = (
        '<svg class="icon" viewBox="%s" fill="none" stroke="currentColor" '
        'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" '
        'aria-hidden="true" focusable="false">%s</svg>' % (vb, inner)
    )
    return _icon_svg[key]


def is_blank(value):
    """Empty, or still an untouched placeholder."""
    if value is None:
        return True
    text = str(value).strip()
    return text == "" or text == FILLIN or text.startswith("-- example")


def is_confirmed(value):
    """The Verified? column. Only a Yes means the row has been checked."""
    return bool(value) and str(value).strip().lower().startswith("y")


# ------------------------------------------------------------------ read
def read_details(wb):
    """Sheet 1: Field / Value pairs."""
    ws = wb["Wedding Details"]
    details = {}
    for field, value, *_ in ws.iter_rows(min_row=2, values_only=True):
        if not field:
            continue
        key = str(field).strip()
        if is_blank(value):
            details[key] = ""
        else:
            details[key] = str(value).strip()
    return details


def read_venues(wb):
    """Sheet 3: venues, keyed by name."""
    if "Venues" not in wb.sheetnames:
        return []
    ws = wb["Venues"]
    venues = []
    for name, address, maps, which in ws.iter_rows(min_row=2, values_only=True):
        if is_blank(name):
            continue
        venues.append(
            {
                "name": str(name).strip(),
                "address": "" if is_blank(address) else str(address).strip(),
                "maps": "" if is_blank(maps) else str(maps).strip(),
                "functions": "" if is_blank(which) else str(which).strip(),
            }
        )
    return venues


def read_rsvps(wb):
    """Sheet 4: optional guest list, shown as a counter only."""
    if "RSVP List" not in wb.sheetnames:
        return []
    ws = wb["RSVP List"]
    rows = []
    for name, relation, side, phone, status, guests, meal in ws.iter_rows(
        min_row=2, values_only=True
    ):
        if is_blank(name):
            continue
        rows.append(
            {
                "name": str(name).strip(),
                "relation": "" if is_blank(relation) else str(relation).strip(),
                "side": "" if is_blank(side) else str(side).strip(),
                "phone": "" if is_blank(phone) else str(phone).strip(),
                "status": "" if is_blank(status) else str(status).strip(),
                "guests": int(guests) if isinstance(guests, (int, float)) else 0,
                "meal": "" if is_blank(meal) else str(meal).strip(),
            }
        )
    return rows


# ------------------------------------------------------------------ dates
MONTHS = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}


def parse_date(text, default_year):
    """'09 Dec 2026', '09 Dec', '2026-12-09' -> (y, m, d) or None."""
    if not text:
        return None
    text = str(text).strip().replace(",", " ")
    text = re.sub(r"(\d)(st|nd|rd|th)", r"\1", text)

    iso = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", text)
    if iso:
        return int(iso.group(1)), int(iso.group(2)), int(iso.group(3))

    num = re.search(r"\b(\d{1,2})\b", text)
    mon = re.search(r"[A-Za-z]+", text)
    year = re.search(r"\b(20\d{2})\b", text)
    if num and mon and mon.group(0).lower() in MONTHS:
        return (
            int(year.group(1)) if year else default_year,
            MONTHS[mon.group(0).lower()],
            int(num.group(1)),
        )
    return None


def parse_time(text):
    """'04:00 PM', '4 PM', '16:00' -> minutes from midnight, or None."""
    if not text:
        return None
    text = str(text).strip().upper().replace(".", "")
    m = re.search(r"(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?", text)
    if not m:
        return None
    hour = int(m.group(1))
    minute = int(m.group(2) or 0)
    meridiem = m.group(3)
    if meridiem == "PM" and hour < 12:
        hour += 12
    if meridiem == "AM" and hour == 12:
        hour = 0
    if not meridiem and m.group(2) is None and 0 < hour <= 6:
        # A bare "4" next to a wedding event is almost always PM.
        hour += 12
    return hour * 60 + minute


def split_time_range(text):
    """Returns (start_minutes, end_minutes). Times without a range get no end."""
    if not text:
        return None, None
    parts = re.split(r"\s*(?:-|–|—|to)\s*", str(text).strip(), maxsplit=1)
    start = parse_time(parts[0])
    end = parse_time(parts[1]) if len(parts) > 1 else None
    if end is not None and start is not None and end < start:
        end += 24 * 60  # crosses midnight, e.g. 10:00 PM - 12:00 AM
    return start, end


def format_clock(minutes):
    if minutes is None:
        return ""
    minutes %= 24 * 60
    hour, minute = divmod(minutes, 60)
    suffix = "AM" if hour < 12 else "PM"
    hour12 = hour % 12 or 12
    return f"{hour12}:{minute:02d} {suffix}"


# ------------------------------------------------------------------ functions
def read_functions(wb, default_year):
    """Sheet 2: the event schedule, grouped by day."""
    if "Wedding Functions" not in wb.sheetnames:
        return [], []
    ws = wb["Wedding Functions"]
    events = []
    for order, date, name, time, venue, note, verified in ws.iter_rows(
        min_row=2, values_only=True
    ):
        if is_blank(name) or is_blank(date):
            continue
        parsed = parse_date(date, default_year)
        if not parsed:
            continue
        year, month, day = parsed
        start, end = split_time_range(time)
        events.append(
            {
                "order": int(order) if isinstance(order, (int, float)) else len(events) + 1,
                "date": f"{year:04d}-{month:02d}-{day:02d}",
                "event": str(name).strip(),
                "time": "" if is_blank(time) else str(time).strip(),
                "timeStart": format_clock(start),
                "timeEnd": format_clock(end),
                "startMinutes": start,
                "endMinutes": end,
                "venue": "" if is_blank(venue) else str(venue).strip(),
                # the note is guest-facing; keep it only once it is confirmed
                "note": "" if is_blank(note) or not is_confirmed(verified) else str(note).strip(),
                "verified": bool(verified) and str(verified).strip().lower().startswith("y"),
                "icon": icon_for(name),
                "iconSvg": icon_svg(icon_for(name)),
            }
        )
    events.sort(key=lambda e: (e["date"], e["startMinutes"] if e["startMinutes"] is not None else 0))

    days = []
    for event in events:
        if not days or days[-1]["date"] != event["date"]:
            days.append({"date": event["date"], "events": []})
        days[-1]["events"].append(event)
    return events, days


# ------------------------------------------------------------------ build
def build():
    if not os.path.exists(BOOK):
        sys.exit(f"Missing {BOOK}\nRun:  python3 make_workbook.py  (once), then edit it.")
    wb = load_workbook(BOOK, data_only=True)

    details = read_details(wb)
    venues = read_venues(wb)
    rsvps = read_rsvps(wb)

    year_text = details.get("Wedding Year", "")
    m = re.search(r"(20\d{2})", year_text)
    default_year = int(m.group(1)) if m else datetime.now().year
    details["Wedding Year"] = str(default_year)

    events, days = read_functions(wb, default_year)

    # "Countdown Event" names the function the hero countdown targets. Match it
    # against the Functions sheet (loose compare, so "darwagar" finds
    # "Baraat Prasthan"). Fall back to the main event, then to the earliest function.
    countdown = None
    wanted = str(details.get("Countdown Event", "")).strip().lower()
    if wanted:
        for ev in events:
            if wanted in str(ev.get("event", "")).strip().lower():
                countdown = ev
                break
        if countdown is None:
            print(f"  warning: Countdown Event {details['Countdown Event']!r} "
                  f"matches no function; using the earliest instead")
    if countdown is None:
        countdown = events[0] if events else None
    if countdown is None:
        parsed = parse_date(details.get("Wedding Date"), default_year)
        if parsed:
            countdown = {
                "date": f"{parsed[0]:04d}-{parsed[1]:02d}-{parsed[2]:02d}",
                "event": details.get("Wedding Main Event", "Wedding"),
            }

    known = [label for label, value in details.items() if str(value).strip()]
    missing = [label for label, value in details.items() if not str(value).strip()]

    data = {
        "generated": datetime.now().strftime("%d %b %Y, %H:%M"),
        "details": details,
        "events": events,
        "days": days,
        "venues": venues,
        "rsvps": rsvps,
        "countdown": countdown,
        "missing": missing,
    }

    payload = json.dumps(data, indent=2, ensure_ascii=False)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("// Generated by build_site.py from wedding-details.xlsx\n")
        fh.write("// Edit the spreadsheet and re-run the script; do not hand-edit this file.\n")
        fh.write("window.WEDDING_DATA = ")
        fh.write(payload)
        fh.write(";\n")

    print(f"wrote {OUT}")
    print(f"  {len(events)} functions across {len(days)} day(s)")
    for day in days:
        print(f"    {day['date']}: {len(day['events'])} function(s)")

    # Cache busting. GitHub Pages serves these with a max-age, so a guest who
    # opened the page before an update can keep seeing the old version. Stamp a
    # hash of each asset's own *content* into its tag, so any real edit always
    # produces a new URL and a no-op rebuild leaves it alone.
    #
    # This has to cover styles.css and app.js, not just data.js. Those two were
    # unstamped, which let a browser pair fresh HTML with a cached old
    # stylesheet — the combination that made the layout look scattered.
    index = os.path.join(os.path.dirname(OUT), "index.html")
    if os.path.exists(index):
        with open(index, encoding="utf-8") as fh:
            html = fh.read()
        original = html

        # data.js: hash the payload, excluding the "generated" clock, otherwise
        # the stamp churns every run and stays frozen for two edits in a minute.
        fingerprint = {k: v for k, v in data.items() if k != "generated"}
        stamps = {
            "data.js": hashlib.sha1(
                json.dumps(fingerprint, indent=2, ensure_ascii=False).encode("utf-8")
            ).hexdigest()[:8]
        }
        for asset in ("styles.css", "app.js"):
            path = os.path.join(os.path.dirname(OUT), asset)
            if os.path.exists(path):
                with open(path, "rb") as fh:
                    stamps[asset] = hashlib.sha1(fh.read()).hexdigest()[:8]

        for asset, digest in stamps.items():
            attr = "href" if asset.endswith(".css") else "src"
            html = re.sub(
                rf'({attr}="){re.escape(asset)}(\?v=[0-9a-f]+)?(")',
                lambda m, d=digest: f"{m.group(1)}{asset}?v={d}{m.group(3)}",
                html,
            )

        if html != original:
            with open(index, "w", encoding="utf-8") as fh:
                fh.write(html)
            print("stamped " + ", ".join(f"{k}={v}" for k, v in sorted(stamps.items())))

    if missing:
        print(f"\nStill to fill in ({len(missing)}):")
        for label in missing:
            print(f"  - {label}")
    else:
        print("\nAll detail fields are filled in.")


if __name__ == "__main__":
    build()
