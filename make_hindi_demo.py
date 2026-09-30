#!/usr/bin/env python3
"""
Build a Hindi demo of the invitation, for looking at before deciding on it.

This is a demonstration, not a translation of the real card. It exists so the
couple can see how the card looks set in Devanagari -- whether the type holds
together, whether the spacing still works, whether the names and the parents
still read as a wedding card -- before anyone commits to doing the work
properly.

What is translated here
  The chrome and the card: the blessing, the parents' labels, the request, the
  countdown, the section headings, the buttons, the form, the weekdays and
  months, and the names of the twelve functions. Those are all fixed phrases
  that carry no meaning of their own, and mistranslating one is obvious
  immediately.

Nothing here is deployed. The output is written into docs/ but every file it
produces is listed in .gitignore, so it lives on this machine and never reaches
GitHub Pages. The English card is untouched by all of this.

Run:  python3 make_hindi_demo.py
      then open  http://localhost:8000/hindi-demo-family.html
"""
import argparse
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(HERE, "docs")
DATA = os.path.join(DOCS, "data-hindi.js")
PAGE = os.path.join(DOCS, "demo-hindi.html")

# Fixed phrases only. See the module docstring for what is left out and why.
# Keys are the real field names in groom-details.xlsx, so a missing key here
# shows up as English still on the page rather than as a silent no-op.
STRINGS = {
    # --- the card, and the wording both pages share -----------------------
    "Website Title": "हम विवाह कर रहे हैं",
    "Opening Blessing Line": "हमारे पारिवारिक आशीर्वाद के साथ",
    "Invitation Line": "आपकी उपस्थिति चाहते हैं",
    "Couple Line (short)": "सुरज की शादी प्रियंका से",
    # --- family page only -------------------------------------------------
    "Family Greeting": "हमारे पारिवारिक आशीर्वाद के साथ",
    "Family Invitation Line": "अपने संबंधित परिवारों के साथ, आपकी उपस्थिति की प्रार्थना करते हैं",
    "Awaiting Section Title": "आपकी उपस्थिति की उत्सुकता",
    "Ashirwad Line": "आपका आशीर्वाद और उपस्थिति हमें नई शुरुआत के लिए अमूल्य उपहार होगी",
    # --- dates and places -------------------------------------------------
    # The weekday and month names are not listed here. Those come from the
    # browser, in Hindi, once data.lang is set -- see DATE_LOCALE in app.js.
    # Only the wording this site composes itself needs a translation.
    "Wedding Date": "11 दिसंबर 2026",
    # Wedding Start Time is not spelled out here: it is passed through
    # hindi_time() so the AM/PM half is handled in one place.
    "Countdown Event": "बारात प्रस्थान",
    "Function Dates": "09 दिसंबर 2026 से 12 दिसंबर 2026",
    "City": "पुणे",
    # --- blank in the English workbook, blank in the demo too -------------
    "Quote / Verse": "",
    "Contact Name": "",
    "RSVP Deadline": "",
    "Email": "",
    "Awaiting Section Note": "",
    "Family Note": "",
    "Baraat Route": "",
    "Parking": "",
    "Accommodation": "",
    "Gift / Shagun Note": "",
}

# Chrome that lives in the markup rather than the data, so it has to be
# rewritten in the page as well. Keyed on the exact English text, with the same
# capitalisation it has in docs/family.html -- these are matched literally, and
# a case that does not match is simply left in English.
MARKUP = {
    # --- section eyebrows, exactly as written in the markup ---------------
    "We invite you to celebrate with us": "आप हमारे साथ उत्सव करने को आमंत्रित हैं",
    "The invitation": "निमंत्रण",
    "With folded hands": "हाथ जोड़कर",
    "If you need us": "यदि आपको हमारी आवश्यकता हो",
    "On the day": "उस दिन",
    "Celebrations": "उत्सव",
    "The main event": "मुख्य कार्यक्रम",
    "Where": "कहाँ",
    "Ashirwad": "आशीर्वाद",
    "Blessings for the Journey Ahead": "आगामी यात्रा के लिए आशीर्वाद",
    "Kindly respond": "कृपया उत्तर दें",
    "Counting down to": "तक गिनती हो रही है",
    "Who To Call": "किसे बुलाएँ",
    "Practical Details": "व्यावहारिक जानकारी",
    # --- headings ---------------------------------------------------------
    "You Are Invited": "आप आमंत्रित हैं",
    "Wedding Schedule": "कार्यक्रम",
    "Venue & Directions": "स्थान और रास्ता",
    "Venue & Map": "स्थान और नक्शा",
    # --- buttons and links -------------------------------------------------
    # The capitalisation here is the markup's own. Much of the site's small
    # caps look is CSS text-transform, so the source text is sentence case.
    "View Schedule": "कार्यक्रम देखें",
    "Send RSVP": "उत्तर भेजें",
    "Calendar": "कैलेंडर",
    "Share": "साझा करें",
    "Map": "नक्शा",
    "WhatsApp": "व्हाट्सएप",
    "Send on WhatsApp": "व्हाट्सएप पर भेजें",
    "Open Invitation": "निमंत्रण खोलें",
    "Open": "खोलें",
    "Sound": "ध्वनि",
    "Clear all": "सब हटाएँ",
    # --- the opening card's own labels -------------------------------------
    "Groom’s Parents": "वर के माता-पिता",
    "Bride’s Parents": "वधू के माता-पिता",
    "request the honour of your presence": "आपकी उपस्थिति चाहते हैं",
    "Tap the card to open": "खोलने के लिए टैप करें",
    "With the blessings of our families": "हमारे पारिवारिक आशीर्वाद के साथ",
    "Together with their families": "अपने संबंधित परिवारों के साथ",
    # --- countdown units ---------------------------------------------------
    "Days": "दिन",
    "Hours": "घंटे",
    "Minutes": "मिनट",
    "Seconds": "सेकंड",
    # --- RSVP form ---------------------------------------------------------
    "Your name": "आपका नाम",
    "Number of guests": "आगंतुकों की संख्या",
    "Relation": "रिश्ता",
    "Phone / WhatsApp": "फ़ोन / व्हाट्सएप",
    "Will you attend?": "क्या आप पधारेंगे?",
    "Meal preference": "भोजन की पसंद",
    "Message for the couple": "जोड़े के लिए संदेश",
    "No preference": "कोई विशेष पसंद नहीं",
    "Vegetarian": "शाकाहारी",
    "Jain": "जैन",
    "Non-vegetarian": "मांसाहारी",
    "Joyfully accepts": "सहर्ष स्वीकार करता हूँ",
    "Not sure yet": "अभी निश्चित नहीं",
    "Regretfully declines": "क्षमा करें, नहीं आ पाऊँगा",
    "Your saved RSVPs on this device (": "इस डिवाइस पर सहेजे गए उत्तर (",
    "With love & blessings": "स्नेह और आशीर्वाद के साथ",
}

# Phrases app.js and family.js compose for themselves, rather than reading
# whole from the workbook. They live in data.ui, so the scripts stay the same
# in both languages. {shown} and {total} are filled in by the script.
UI = {
    "countingDownTo": "तक गिनती हो रही है",
    "countingDownToWedding": "शादी तक गिनती हो रही है",
    "allCelebrations": "सभी कार्यक्रम,",
    "venueCaption": "विवाह / मुख्य कार्यक्रम",
    "openInMaps": "नक्शे में खोलें",
    "copyAddress": "पता कॉपी करें",
}


# The twelve functions by name. Transliterations only -- the description under
# each one stays English, as explained above.
FUNCTIONS = {
    "Matkor": "माटकोर",
    "Madwa": "मांडवा",
    "Ghee Dhari": "घी धारी",
    "Devpuji": "देव पूजा",
    "Haldi Kalsa": "हल्दी कल्सा",
    "Dholki": "ढोलकी",
    "Mehndi": "मेहंदी",
    "Janau": "जनेऊ",
    "Baraat Prasthan": "बारात प्रस्थान",
    "Saptapadi": "सप्तपदी",
    "Tilak": "तिलक",
    "Jaimala": "जयमाला",
    "Sindoor Daan": "सिंदूर दान",
    "Vidai": "विदाई",
    "Lawa Bhujai": "लवा भुजाई",
    "Janeu ceremony": "जनेऊ",
    "Kalra": "कलरा",
    "Bhunga Lawa": "भुंगा लवा",
    "Parat": "परात",
    "Ganesh Puja": "गणेश पूजा",
    "Hawan": "हवन",
}

# The family blocks and the parent role labels. The parents' names and the
# relatives' names are not touched: those are people's names, and they stay as
# the family writes them.
LABELS = {
    "Groom's Family": "वर का पारिवार",
    "Bride's Family": "वधू का पारिवार",
    "Father": "पिता",
    "Mother": "माता",
}


def load_data(path):
    raw = open(path, encoding="utf-8").read()
    payload = raw.split("window.WEDDING_DATA = ", 1)[1].rstrip().rstrip(";")
    return json.loads(payload)


def hindi_time(text):
    """'2:00 PM' -> 'दोपहर 2:00 बजे'. A time left as '6:00 PM' in the middle of a
    Devanagari sentence is the sort of thing that makes a translated card look
    unfinished, and the AM/PM half is the only part that changes."""
    if not text:
        return text
    raw = str(text).strip()
    low = raw.lower().replace(".", "")
    if low.endswith("pm"):
        label, rest = "दोपहर", raw[:-2].strip()
    elif low.endswith("am"):
        label, rest = "सुबह", raw[:-2].strip()
    else:
        return raw
    hour, _, minute = rest.partition(":")
    try:
        h = int(hour)
    except ValueError:
        return raw
    if not minute:
        return "%s %d बजे" % (label, h)
    return "%s %d:%s बजे" % (label, h, minute)


def translate(data):
    for key, value in STRINGS.items():
        if key in data.get("details", {}):
            data["details"][key] = value
        if key in data.get("family", {}).get("details", {}):
            data["family"]["details"][key] = value

    # The same function names appear in more than one place -- the
    # schedule's day panels, and the main-event card -- and each carries its own
    # copy of the name, so all three are translated here.
    for group in (data.get("events", []), data.get("days", [])):
        for ev in group:
            for ev2 in ([ev] + list(ev.get("events", []))):
                name = str(ev2.get("event", "")).strip()
                if name in FUNCTIONS:
                    ev2["eventEn"] = name
                    ev2["event"] = FUNCTIONS[name]
    if data.get("countdown", {}).get("event") in FUNCTIONS:
        name = data["countdown"]["event"]
        data["countdown"]["eventEn"] = name
        data["countdown"]["event"] = FUNCTIONS[name]

    # The AM/PM half of a time is the only part that changes in Hindi, and a
    # lone "6:00 PM" in the middle of a Devanagari sentence reads as unfinished.
    # The 24-hour fields are left alone -- nothing on the page shows them.
    for group in (data.get("events", []), data.get("days", [])):
        for ev in group:
            # A day's own time fields and its nested events' fields, because
            # the schedule reads the nested copies and the section cards read
            # the top-level ones. Missing either leaves AM/PM on the page.
            for ev2 in ([ev] + list(ev.get("events", []))):
                # `time` is usually a 24-hour value the page never shows, and
                # hindi_time leaves those alone. On a few events it is a
                # display string instead, and app.js falls back to it when
                # timeStart is missing, so it gets translated too.
                for key in ("time", "timeStart", "timeEnd"):
                    if ev2.get(key):
                        ev2[key] = hindi_time(ev2[key])
    for holder in [data.get("countdown", {})]:
        for key in ("time", "timeStart", "timeEnd"):
            if holder.get(key):
                holder[key] = hindi_time(holder[key])
    data.setdefault("details", {})["Wedding Start Time"] = hindi_time(
        data.get("details", {}).get("Wedding Start Time", "")
    )

    # Dates are not translated here on purpose. app.js and family.js ask the
    # browser for the weekday and month names in DATE_LOCALE, which is hi-IN
    # once data.lang is set, so the names -- and the order the parts come in --
    # come from the platform's own data rather than from a table here.
    # The relative lists are grouped by the side label family.js looks up, so
    # that label is translated in the data too, not only in the markup.
    sides = {"Groom's side": "वर पक्ष", "Bride's side": "वधू पक्ष", "Both": "दोनों"}
    for m in data.get("family", {}).get("members", []):
        side = m.get("side")
        if side in sides:
            m["sideEn"] = side
            m["side"] = sides[side]

    data["ui"] = UI

    data["lang"] = "hi"
    return data


def entity_forms(text):
    """The markup's own spelling of a phrase, plus its entity-encoded variant.

    A literal "&" or "'" in the markup is written as &amp; or &rsquo;, so a
    phrase written the readable way here does not match the file byte for byte.
    Both forms are returned so the caller can try each.
    """
    forms = [text]
    amp = text.replace("&", "&amp;")
    if amp != text:
        forms.append(amp)
    apos = text.replace("’", "&rsquo;")
    if apos != text:
        forms.append(apos)
    return forms


# Attributes a screen reader reads out loud, and the browser's own tooltips.
# These are invisible on the page, so a page can look finished and still be
# read out in English.
ATTRS = {
    "aria-label": {
        "Open the wedding invitation": "शादी का निमंत्रण खोलें",
        "Countdown to the wedding": "शादी तक गिनती",
        "Families": "परिवार",
        "Wedding days": "शादी के दिन",
        "Open in Maps": "नक्शे में खोलें",
    },
    "title": {
        "Add to calendar": "कैलेंडर में जोड़ें",
        "Mute music": "संगीत बंद करें",
        "Open in Maps": "नक्शे में खोलें",
        "Share invitation": "निमंत्रण साझा करें",
        "WhatsApp": "व्हाट्सएप",
    },
    "placeholder": {
        "e.g. Uncle": "जैसे चाचा जी",
    },
    "alt": {},
}


def translate_attrs(html):
    for attr, table in ATTRS.items():
        for en, hi in table.items():
            for en_form in entity_forms(en):
                html = html.replace('%s="%s"' % (attr, en_form),
                                    '%s="%s"' % (attr, hi))
    return html


def add_hindi_font(html):
    """A Devanagari face, and a fallback stack that can actually render it.

    Cormorant Garamond and Jost carry no Devanagari, so without this the text
    falls back to whatever the system has, which differs per machine -- and on
    a card the type is the whole point of the demo. Tiro Devanagari Hindi is a
    serif with the same kind of proportions as the display face, which keeps
    the card from looking like two different cards stacked together.
    """
    html = html.replace(
        'href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond',
        'href="https://fonts.googleapis.com/css2?family=Tiro+Devanagari+Hindi'
        '&family=Cormorant+Garamond',
        1,
    )
    extra = """<style>
/* Demo only: put the Devanagari serif ahead of the Latin display face, so
   Devanagari picks it up and the names and English words alongside it keep
   Cormorant. The two sit together well because both are low-contrast
   transitional serifs. */
html[lang="hi"] {
  --display: "Tiro Devanagari Hindi", "Cormorant Garamond", Georgia, serif;
  --sans: "Tiro Devanagari Hindi", "Jost", -apple-system, sans-serif;
}
html[lang="hi"] body { letter-spacing: 0; }
/* The card's small caps labels are set with letter-spacing, which on
   Devanagari pulls the matras apart and looks broken. Matras are already
   positioned by the font, so the tracking has to go. */
html[lang="hi"] .card__role,
html[lang="hi"] .family__side,
html[lang="hi"] .section__eyebrow,
html[lang="hi"] .topbar__tag,
html[lang="hi"] .family__role { letter-spacing: 0; text-transform: none; }
/* Devanagari runs taller at the same font-size and the conjuncts need the
   descender room, or the top bar clips. */
html[lang="hi"] .topbar { line-height: 1.9; }
html[lang="hi"] .card__parent { line-height: 1.5; }
</style>"""
    return html.replace("</head>", extra + "\n</head>", 1)


def build_page(src_page, out_page, data_file, report=False):
    with open(src_page, encoding="utf-8") as fh:
        html = fh.read()

    html = html.replace('<html lang="en">', '<html lang="hi">', 1)
    html = html.replace("<title>We're Getting Married</title>",
                        "<title>हमारा विवाह</title>", 1)
    html, n = re.subn(r'(src=")data\.js(\?v=[^"]*)?"',
                      lambda m: f'{m.group(1)}{data_file}"', html)
    if not n:
        print("  warning: no data.js tag in %s" % os.path.basename(src_page))

    # The fixed phrases that live in the markup rather than the data. Both
    # tables are matched literally, so a phrase whose capitalisation differs
    # from the markup stays in English rather than being mangled.
    missed = []
    for table in (MARKUP, LABELS):
        for en, hi in table.items():
            if not hi:
                continue
            before = html
            # The markup stores some of these as HTML entities -- "Venue &
            # Directions" is written Venue &amp; Directions, and the apostrophe
            # in "Groom's Parents" as &rsquo;. Both spellings have to be tried,
            # or those phrases silently stay in English.
            for en_form in entity_forms(en):
                # Whitespace-tolerant on purpose. A label may be hugged by its
                # tags or sit on its own indented line inside them, and
                # matching only one of those shapes is how half the phrases go
                # unmatched and stay in English.
                pattern = r"(?<=>)\s*%s\s*(?=<)" % re.escape(en_form)
                html = re.sub(pattern, lambda _m, t=hi: t, html)
            if html == before:
                missed.append(en)
    # Only the family page is reported on. The guest page has no formal
    # invitation or family sections, so most of the table legitimately
    # has nothing to match there -- reporting that as a miss every build would
    # bury the real failures under expected noise.
    if missed and report:
        # Silent misses are how half a page stays in English, so they are
        # reported rather than passed over.
        print("  %d phrase(s) not found in the markup, left in English:"
              % len(missed))
        for phrase in missed:
            print("    - %s" % phrase)
    html = add_hindi_font(html)
    html = translate_attrs(html)

    with open(out_page, "w", encoding="utf-8") as fh:
        fh.write(html)
    return out_page


def main():
    ap = argparse.ArgumentParser(description="Build the Hindi demo card.")
    ap.add_argument("--out-dir", default=DOCS)
    args = ap.parse_args()

    src_data = os.path.join(DOCS, "data.js")
    if not os.path.exists(src_data):
        sys.exit("Missing %s\nRun:  python3 build_site.py  first" % src_data)

    data = translate(load_data(src_data))
    data["generated"] = "hindi demo"

    with open(DATA, "w", encoding="utf-8") as fh:
        fh.write("// Hindi DEMO. Generated by make_hindi_demo.py.\n")
        fh.write("window.WEDDING_DATA = ")
        fh.write(json.dumps(data, indent=2, ensure_ascii=False))
        fh.write(";\n")
    print("wrote %s" % DATA)

    for name in ("family.html", "index.html"):
        src = os.path.join(DOCS, name)
        if not os.path.exists(src):
            continue
        stem = os.path.splitext(name)[0]
        out = os.path.join(args.out_dir, "hindi-demo-%s.html" % stem)
        build_page(src, out, os.path.basename(DATA), report=(stem == "family"))
        print("wrote %s" % out)


if __name__ == "__main__":
    main()
