#!/usr/bin/env python3
"""
Convert groom-details.xlsx into site/data.js.

The workbook is the source of truth. Edit the spreadsheet, run this script,
and refresh the page.

Run:  python3 build_site.py
      python3 build_site.py --book bride-details.xlsx --out docs/data-bride.js

The two-argument form is how the bride's family gets their own card. They fill
in a copy of the workbook and the same script builds a second data file from
it; the "Card Flank" cell inside that copy is what flips the order, so nothing
about the layout is decided here.
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
BOOK = os.path.join(HERE, "groom-details.xlsx")
OUT = os.path.join(HERE, "docs", "data.js")

FILLIN = "-- FILL IN --"

# Which drawing each function on the schedule gets. Matched loosely so OCR
# variants and spelling differences on the card still find the right icon.
# These are keyed on the function, not on any ritual: the Rituals sheet is
# gone, but the schedule still marks itself with a drawing per row.
FUNCTION_ICONS = [
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
    for key, needles in FUNCTION_ICONS:
        for needle in needles:
            if needle in low:
                return key
    return None


ICON_DIR = os.path.join(HERE, "docs", "assets", "icons")
_icon_svg = {}


def icon_svg(key):
    """Inline a function icon as markup.

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


# Every record-bearing sheet ends with these two columns. Publish is the hard
# gate: a row that is anything but a Yes is dropped here, at build time, so an
# unpublished record never reaches the generated data file at all. Audience
# then decides which of the two pages a surviving record is allowed to render
# on. Verified? used to do this job, but it only ever gated the guest-facing
# note on Wedding Functions while the event itself was always published, and it
# gated whole rows elsewhere. Merging the two means a published event now shows
# its note too.
AUDIENCES = ("friends", "family", "both")


def is_published(value):
    """The Publish column. Only a Yes means the row goes on the site."""
    return bool(value) and str(value).strip().lower().startswith("y")


def norm_audience(value):
    """The Audience column, normalised to friends/family/both.

    An unset or misspelled cell becomes "both" rather than hiding the record:
    a typo should not silently delete something the family was expecting, and
    the column has a dropdown so the typo case should not arise in practice.
    """
    if is_blank(value):
        return "both"
    text = str(value).strip().lower()
    if text in AUDIENCES:
        return text
    if text.startswith("friend"):
        return "friends"
    if text.startswith("fam"):
        return "family"
    return "both"


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
    return parents_view(details)


# Which parent leads inside each set. Deliberately not the same on both sides,
# because this is the printed-card convention: the groom's card leads with the
# mother, the bride's with the father. Earlier this led with the father on both
# sides, which was a change of mind, not a correction of a fault.
#
# This has to agree with the order of the family__line blocks in index.html and
# family/index.html, because those are what the family page renders and this is
# what the card renders. The bride's pages are generated from the same two
# sources, so both blocks and the joined string change together.
PARENT_LEAD = {"Groom": "mother", "Bride": "father"}


def parents_view(details):
    """Add the four parent fields plus a joined 'Groom's Parents' display value.

    Three jobs, all in one place because getting them out of step is how a
    father's name ends up printed as a mother's:

    1. Normalise a workbook that still has the old single "Groom's Parents"
       field. That value is a father's -- it was typed into a field with no
       gender -- so it fills Father and Mother is left empty.
    2. Join both parents into the old display key so app.js and family.js keep
       reading one field. The lead differs per side, see PARENT_LEAD: the
       groom's card reads "Smt. Pushpa Singh & Shri Randhir Prasad Singh" and
       the bride's reads "Shri Samsher Bahadur Singh & Smt. Shubhawati Devi".
    3. A blank parent drops out of the join instead of leaving a dangling
       "&", and a side with neither stays empty so the card still removes the
       whole line.
    """
    out = dict(details)
    for side in ("Groom", "Bride"):
        old = out.pop("%s's Parents" % side, "")
        father = out.get("%s's Father" % side, "")
        mother = out.get("%s's Mother" % side, "")
        if not father and not mother and old:
            father, mother = old, ""
        out["%s's Father" % side] = father
        out["%s's Mother" % side] = mother
        pair = (father, mother) if PARENT_LEAD[side] == "father" else (mother, father)
        names = [n for n in pair if n]
        # "&" between them, "and" for a married couple reads oddly
        out["%s's Parents" % side] = " & ".join(names) if names else ""
    return out


def read_field_sheet(wb, name):
    """Any sheet shaped Field / Value / Notes. Returns {} when the sheet is absent.

    Same reader as read_details, factored out because the family page has more
    than one such sheet. Missing sheet must stay non-fatal: the friends page has
    to keep working if the family sheets are deleted.
    """
    if name not in wb.sheetnames:
        return {}
    ws = wb[name]
    out = {}
    for field, value, *_ in ws.iter_rows(min_row=2, values_only=True):
        if not field:
            continue
        key = str(field).strip()
        if not key or key.lower().startswith("field"):
            continue
        out[key] = "" if is_blank(value) else str(value).strip()
    return out


def read_members(wb):
    """The Family Members sheet, one row per person, name required.

    These are the names printed under "Eagerly Awaiting Your Presence".
    A row is only printed when its Publish cell says Yes, so a relative's name
    typed in but not yet marked for publication does not reach the page by
    accident.
    """
    if "Family Members" not in wb.sheetnames:
        return []
    ws = wb["Family Members"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        def cell(i):
            return row[i] if i < len(row) else None

        name, relation, side, frm = cell(0), cell(1), cell(2), cell(3)
        publish, audience = cell(4), cell(5)
        if is_blank(name) or not is_published(publish):
            continue
        rows.append(
            {
                "name": str(name).strip(),
                "relation": "" if is_blank(relation) else str(relation).strip(),
                "side": "" if is_blank(side) else str(side).strip(),
                "from": "" if is_blank(frm) else str(frm).strip(),
                "audience": norm_audience(audience),
            }
        )
    return rows


def read_contacts(wb):
    """The Family Contacts sheet, one row per person, name required.

    Family-only by default: a phone number on the family page is not something
    a guest should be able to find, so an unstated Audience still normalises
    through norm_audience and the column has a dropdown for the rest.
    """
    if "Family Contacts" not in wb.sheetnames:
        return []
    ws = wb["Family Contacts"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        def cell(i):
            return row[i] if i < len(row) else None

        name, relation, phone = cell(0), cell(1), cell(2)
        about, note = cell(3), cell(4)
        publish, audience = cell(5), cell(6)
        if is_blank(name) or not is_published(publish):
            continue
        rows.append(
            {
                "name": str(name).strip(),
                "relation": "" if is_blank(relation) else str(relation).strip(),
                "phone": "" if is_blank(phone) else str(phone).strip(),
                "about": "" if is_blank(about) else str(about).strip(),
                "note": "" if is_blank(note) else str(note).strip(),
                "audience": norm_audience(audience),
            }
        )
    return rows


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
    """Sheet 4: optional guest list, shown as a counter only.

    Rows are read positionally rather than unpacked, because the sheet carries
    trailing Publish and Audience columns and a fixed 7-tuple unpack would
    raise the moment anyone filled one of them in.
    """
    if "RSVP List" not in wb.sheetnames:
        return []
    ws = wb["RSVP List"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        def cell(i):
            return row[i] if i < len(row) else None

        name, relation, side = cell(0), cell(1), cell(2)
        phone, status, guests = cell(3), cell(4), cell(5)
        publish, audience = cell(7), cell(8)
        if is_blank(name) or not is_published(publish):
            continue
        rows.append(
            {
                "name": str(name).strip(),
                "relation": "" if is_blank(relation) else str(relation).strip(),
                "side": "" if is_blank(side) else str(side).strip(),
                "phone": "" if is_blank(phone) else str(phone).strip(),
                "status": "" if is_blank(status) else str(status).strip(),
                "guests": int(guests) if isinstance(guests, (int, float)) else 0,
                "audience": norm_audience(audience),
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
    """Sheet 2: the event schedule, grouped by day.

    Read by position rather than unpacked into named variables: the sheet has
    grown trailing Publish and Audience columns over time, and a fixed-length
    unpack would raise ValueError the moment anyone typed into one of them.
    """
    if "Wedding Functions" not in wb.sheetnames:
        return [], []
    ws = wb["Wedding Functions"]
    events = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        def cell(i):
            return row[i] if i < len(row) else None

        order, date, name, time = cell(0), cell(1), cell(2), cell(3)
        venue, note = cell(4), cell(5)
        family_note, publish, audience = cell(6), cell(7), cell(8)
        # Publish is the record's gate: an unpublished function is not written
        # to the data file at all, so it cannot reach either page by accident.
        if is_blank(name) or is_blank(date) or not is_published(publish):
            continue
        parsed = parse_date(date, default_year)
        if not parsed:
            continue
        year, month, day = parsed
        start, end = split_time_range(time)
        event_name = str(name).strip()
        icon = icon_for(name)
        events.append(
            {
                "order": int(order) if isinstance(order, (int, float)) else len(events) + 1,
                "date": f"{year:04d}-{month:02d}-{day:02d}",
                "event": event_name,
                "time": "" if is_blank(time) else str(time).strip(),
                "timeStart": format_clock(start),
                "timeEnd": format_clock(end),
                "startMinutes": start,
                # "" rather than None when a function has no end time. The .ics
                # export still needs to tell "no end" from "ends at midnight",
                # so app.js reads this with endMins() instead of == null.
                "endMinutes": end if end is not None else "",
                "venue": "" if is_blank(venue) else str(venue).strip(),
                # the note is guest-facing and rides along with the record;
                # before Publish existed this was gated separately, because
                # Verified? only ever hid the note and never the event
                "note": "" if is_blank(note) else str(note).strip(),
                "audience": norm_audience(audience),
                # "" not None: nothing reads this key, but a null here would be
                # the one place a bare null reaches the generated data file
                "icon": "" if is_blank(icon) else icon,
                "iconSvg": icon_svg(icon),
                # family-page only, and not gated on Publish because these are
                # written for relatives rather than derived from the card
                "familyNote": "" if is_blank(family_note) else str(family_note).strip(),
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
def card_order(details):
    """Work out which side of the couple leads the card, and in what order.

    One family prints the card, but the same wedding gets sent out from both
    sides, and each family puts its own person first: the groom's family sends
    "Suraj & Priyanka" with the groom's parents above the names, the bride's
    family sends "Priyanka & Suraj" with the bride's parents above. Nothing
    about the data changes between the two, only the order, so it is worked out
    once here from the workbook's "Card Flank" field rather than being decided
    in three separate places in JavaScript.

    Everything the renderers need is emitted together under data["order"]:
    the two names, the two parent blocks, the two family-member side groups,
    and a couple line. The scripts read that block and never re-derive it, so
    the names, the parents, the relations list and the hashtag cannot end up
    disagreeing about which family leads.
    """
    raw = str(details.get("Card Flank", "") or "").strip().lower()
    if raw in ("bride",):
        first = "bride"
    elif raw in ("groom", ""):
        first = "groom"
    else:
        first = "groom"
        print(f"  warning: Card Flank {details.get('Card Flank')!r} is neither"
              " Groom nor Bride; using Groom")
    second = "groom" if first == "bride" else "bride"

    def name(side):
        return str(details.get("%s Name" % side.title(), "") or "").strip()

    a, b = name(first.title()), name(second.title())
    pair = [n for n in (a, b) if n]

    return {
        "flank": first,
        # Index 0 is whoever the card is for. The scripts assign names to
        # fixed ids in this order and reorder the DOM to match.
        "names": pair,
        "sides": ["%s's side" % first.title(), "%s's side" % second.title(), "Both"],
        "coupleLine": " weds ".join(pair) if pair else "",
        "hashtag": "#%s" % "".join(n for n in pair) if pair else "",
    }


def write_variant_pages(folder, data_file):
    """Copy index.html and family.html with their data tag repointed.

    Everything else -- markup, styles, scripts -- is byte-identical to the
    groom's pages on purpose. The two cards differ in one thing, which is whose
    name is on it, and that is entirely down to the workbook. Generating the
    copies here rather than keeping two sets by hand is what stops them drifting
    apart after the next edit.

    The bride's pages live under bride/ and bride/family/ because the site is
    served from a custom domain at its own root, where /, /family/, /bride/ and
    /bride/family/ are the four addresses worth handing out. A page a directory
    deep cannot reach ../styles.css, so the sources reference every asset from
    the root and the copy keeps those absolute paths untouched.

    The stamp loop below then rewrites the ?v= on the new data file in these
    copies exactly as it does for the originals, so the cache busting is not
    something the second card has to remember to do.
    """
    written = []
    for src_name, out_name in (("index.html", "bride/index.html"),
                               ("family/index.html", "bride/family/index.html")):
        src = os.path.join(folder, src_name)
        if not os.path.exists(src):
            continue
        with open(src, encoding="utf-8") as fh:
            html = fh.read()
        # Only the data tag. app.js and family.js are shared, and their stamps
        # are already identical in the source page. The leading slash is
        # captured and put back, or the copy would point at data-bride.js
        # relative to itself and find nothing.
        html, n = re.subn(
            r'(src=")(/?)data\.js(\?v=[^"]*)?"',
            lambda m: f'{m.group(1)}{m.group(2)}{data_file}"',
            html,
        )
        if not n:
            print(f"  warning: no data.js tag found in {src_name}; "
                  "left it alone rather than guessing")
            continue
        out = os.path.join(folder, out_name)
        os.makedirs(os.path.dirname(out), exist_ok=True)
        with open(out, "w", encoding="utf-8") as fh:
            fh.write(html)
        written.append(out_name)
    if written:
        print("  variant pages: " + ", ".join(written))
    return written


def build(book=None, out=None):
    book = book or BOOK
    out = out or OUT
    if not os.path.exists(book):
        sys.exit(f"Missing {book}\nRun:  python3 make_workbook.py  (once), then edit it.")
    wb = load_workbook(book, data_only=True)

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
    # "Groom's Parents" is derived from the four parent fields, so reporting it
    # would double-count a side that is already filled. Report the real fields.
    derived = {("%s's Parents" % s) for s in ("Groom", "Bride")}
    # "Card Flank" is a setting, not something to be written. Blank means Groom.
    derived.add("Card Flank")
    missing = [
        label
        for label, value in details.items()
        if not str(value).strip() and label not in derived
    ]

    # "Wedding Date" is a human string like "11 Dec 2026", which is what belongs
    # in the workbook. Every date the scripts format goes through
    # new Date(iso), and appending an ISO "T00:00:00" to a non-ISO string is
    # left to each browser's fallback parser, so the same card can print its
    # date on one screen and not another. Emit the ISO form once, here, next to
    # the string, and let the scripts use this one.
    wedding_iso = ""
    parsed = parse_date(details.get("Wedding Date"), default_year)
    if parsed:
        wedding_iso = f"{parsed[0]:04d}-{parsed[1]:02d}-{parsed[2]:02d}"

    data = {
        "generated": datetime.now().strftime("%d %b %Y, %H:%M"),
        "details": details,
        "weddingDateISO": wedding_iso,
        "order": card_order(details),
        "events": events,
        "days": days,
        "venues": venues,
        "rsvps": rsvps,
        "countdown": countdown,
        "missing": missing,
        # Extra material for docs/family.html. Kept in one block so the friends
        # page, which never reads it, is unaffected if these sheets go away.
        "family": {
            "details": read_field_sheet(wb, "Family Details"),
            "contacts": read_contacts(wb),
            "members": read_members(wb),
        },
    }

    payload = json.dumps(data, indent=2, ensure_ascii=False)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(f"// Generated by build_site.py from {os.path.basename(book)}\n")
        fh.write("// Edit the spreadsheet and re-run the script; do not hand-edit this file.\n")
        fh.write("window.WEDDING_DATA = ")
        fh.write(payload)
        fh.write(";\n")

    print(f"wrote {out}")
    print(f"  {len(events)} functions across {len(days)} day(s)")
    for day in days:
        print(f"    {day['date']}: {len(day['events'])} function(s)")

    # A second card is the same two pages pointing at a different data file.
    # Generated rather than hand-copied so the two can never drift apart, and
    # a separate file rather than a ?variant= query so the bride's family never
    # downloads the groom's relatives list to get to their own.
    if out != OUT:
        write_variant_pages(os.path.dirname(out), os.path.basename(out))

    # Cache busting. GitHub Pages serves these with a max-age, so a guest who
    # opened the page before an update can keep seeing the old version. Stamp a
    # hash of each asset's own *content* into its tag, so any real edit always
    # produces a new URL and a no-op rebuild leaves it alone.
    #
    # This has to cover styles.css and app.js, not just data.js. Those two were
    # unstamped, which let a browser pair fresh HTML with a cached old
    # stylesheet — the combination that made the layout look scattered.
    family = data["family"]
    family_missing = sorted(
        label for label, value in family["details"].items() if not str(value).strip()
    )

    # The data file is stamped under its own name, which is not always
    # "data.js": a variant card is written to data-bride.js and its pages point
    # there. Keying this on the hard-coded name left those pages pointing at an
    # un-stamped URL, so a browser was free to keep serving the previous version
    # of the card after a rebuild.
    data_name = os.path.basename(out)
    data_stamp = hashlib.sha1(
        json.dumps(
            {k: v for k, v in data.items() if k != "generated"},
            indent=2,
            ensure_ascii=False,
        ).encode("utf-8")
    ).hexdigest()[:8]

    # The stylesheet and the two scripts are shared by every card, so they carry
    # the same stamp on all of them. Only the data file differs.
    stamps = {}
    for asset in ("styles.css", "app.js", "family.js"):
        path = os.path.join(os.path.dirname(out), asset)
        if os.path.exists(path):
            with open(path, "rb") as fh:
                stamps[asset] = hashlib.sha1(fh.read()).hexdigest()[:8]
    stamps[data_name] = data_stamp

    # Each build stamps its own pages and no others, so a variant build never
    # rewrites the originals and the original build never reaches for a page
    # that may not exist. family/index.html is optional: before it exists there
    # is nothing to rewrite.
    if data_name == "data.js":
        pages = ("index.html", "family/index.html")
    else:
        pages = ("bride/index.html", "bride/family/index.html")

    for page in pages:
        path = os.path.join(os.path.dirname(out), page)
        if not os.path.exists(path):
            continue
        with open(path, encoding="utf-8") as fh:
            html = fh.read()
        original = html

        for asset, digest in stamps.items():
            attr = "href" if asset.endswith(".css") else "src"
            # \?v=[^"]* rather than \?v=[0-9a-f]+ because the first build of a
            # new page leaves a bare `family.js?v=` in the template. Requiring
            # hex digits there would not match, and the empty query would sit
            # in the served tag forever, defeating the cache busting.
            # The optional slash keeps the asset rooted: pages a directory deep
            # are written with /app.js, and matching the bare name would strip
            # the slash and leave a page that cannot load its own scripts.
            html = re.sub(
                rf'({attr}=")(/?){re.escape(asset)}(\?v=[^"]*)?"',
                lambda m, d=digest: f'{m.group(1)}{m.group(2)}{asset}?v={d}"',
                html,
            )

        if html != original:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(html)
            print(f"stamped {page}: " + ", ".join(f"{k}={v}" for k, v in sorted(stamps.items())))

    if missing:
        print(f"\nStill to fill in ({len(missing)}):")
        for label in missing:
            print(f"  - {label}")
    else:
        print("\nAll detail fields are filled in.")

    members = data["family"]["members"]
    if members:
        print(f"\nFamily Members listed ({len(members)}):")
        for m in members:
            bits = [m["name"]]
            if m["relation"]:
                bits.append(m["relation"])
            if m["side"]:
                bits.append(m["side"])
            print("  - " + " — ".join(bits))
    else:
        print(
            "\nNo family members listed. Fill the 'Family Members' sheet and set"
            "\nVerified? to Yes; the section hides itself until you do."
        )

    if family_missing:
        print(f"\nFamily page still to fill in ({len(family_missing)}):")
        for label in family_missing:
            print(f"  - {label}")


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Build site data from the workbook.")
    ap.add_argument("--book", default=BOOK, help="workbook to read")
    ap.add_argument("--out", default=OUT, help="data.js to write")
    args = ap.parse_args()
    build(args.book, args.out)
