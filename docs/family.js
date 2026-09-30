/* ==========================================================================
   family.js — the longer invitation, for relatives.

   Loaded only by docs/family.html, after app.js. app.js renders the shared
   parts (hero, countdown, schedule, venue, RSVP) and is used unchanged by both
   pages; this file adds what only the family version carries:

     - the formal invitation wording, in printed-card order
     - what each of the twelve functions actually is
     - who to call
     - baraat route, parking, accommodation, gift note

   Nothing here writes to the workbook or the page in a way a visitor can
   change, matching the rule app.js follows: this page is read-only, and every
   edit happens in groom-details.xlsx followed by build_site.py.

   Every block hides itself when its data is blank, so a half-finished workbook
   produces a shorter page rather than a page full of empty headings.
   ========================================================================== */
(function () {
  "use strict";

  var data = window.WEDDING_DATA || {};
  var D = data.details || {};

  /* Matches app.js: the language's own weekday and month names, en-GB unless
     the data says the card is in Hindi. */
  var DATE_LOCALE = data.lang === "hi" ? "hi-IN" : "en-GB";

  /* The workbook's Audience column decides which card a record is allowed to
     appear on: Friends, Family or Both. build_site.py has already dropped every
     Publish = No row, so this only splits what survives between the two pages.
     The filter lives here, on the data, rather than hiding elements in CSS: a
     friends-only record hidden with display:none is still in the page and
     still readable in view-source, which is the opposite of what the column is
     for. A record with no audience counts as Both, matching norm_audience()
     on the Python side. */
  function forAudience(records, side) {
    return (records || []).filter(function (r) {
      if (!r) return false;
      var a = String(r.audience || "both").trim().toLowerCase();
      if (a === "both") return true;
      if (a === "friends") return side === "friends";
      if (a === "family") return side === "family";
      return true;
    });
  }
  /* No event filtering here. The schedule is rendered by app.js, which picks
     the side from the data-card marker, so the family page's functions are
     already filtered for this card. This file used to declare its own
     `events` for that job and never used it. */
  var family = data.family || {};
  var F = family.details || {};
  /* Which family this card was printed for, worked out once by build_site.py
     from the workbook's "Card Flank". The groom's family sends the card with
     Suraj first, the bride's family with Priyanka first. */
  var order = data.order || { flank: "groom", names: [], sides: [] };

  function $(id) { return document.getElementById(id); }

  function text(node, value) {
    if (!node) return;
    var s = (value == null ? "" : String(value)).trim();
    node.textContent = s;
  }

  /* el() used to live here, for the practical-details blocks. Those moved into
     app.js along with the markup they built, so nothing in this file creates an
     element any more: it reads text into the nodes already in the page. */

  function niceDate(iso) {
    var d = new Date(iso + "T00:00:00");
    if (isNaN(d.getTime())) return iso;
    return d.toLocaleDateString(DATE_LOCALE, {
      weekday: "long", day: "numeric", month: "long", year: "numeric"
    });
  }

  function telHref(phone) {
    return "tel:" + String(phone).replace(/[^\d+]/g, "");
  }

  /* ------------------------------------------------- formal invitation */
  /* The friends page puts the ask at the foot of the hero card. A printed
     invitation puts it after both sets of parents, and that is the order a
     relative is reading for. */
  (function formalInvitation() {
    var sec = $("formalInvite");
    if (!sec) return;

    var blessing = F["Family Greeting"] || D["Opening Blessing Line"] || "";
    text($("formalBlessing"), blessing);

    /* The parents are not repeated in the formal invitation: the families
       block just above already names both sets, under a FATHER / MOTHER
       label. Printing them again here read as a mistake. */

    /* Names lead with whoever printed the card, from the same data.order the
       hero and the card cover read. */
    var names = $("formalNames");
    if (names) {
      var pair = (order && order.names) || [];
      if (pair.length) {
        names.textContent = pair.join("  &  ");
        names.hidden = false;
      } else {
        names.hidden = true;
      }
    }

    /* The long wording, e.g. "together with their respective families, request
       the honour of your presence". Falls back to the short couple line so the
       invitation still reads as an invitation if the field is blank. */
    var ask = F["Family Invitation Line"] || D["Couple Line (short)"] || "";
    text($("formalAsk"), ask);

    var dates = D["Function Dates"] || D["Wedding Date"] || "";
    text($("formalDates"), dates);

    /* A message to relatives, if one was written. */
    var note = F["Family Note"] || "";
    var noteBlock = $("familyNoteBlock");
    if (noteBlock) {
      text(noteBlock, note);
      noteBlock.hidden = !note;
    }

    /* Hide the whole section only when it would be entirely empty, which
       cannot happen once the names are filled in. Kept for a blank workbook.

       `parents` used to be in this chain. It was a leftover from before the
       parents were dropped from this block, and was never declared in this
       file. It did no damage only because `blessing` is truthy today, so the
       `||` short-circuited before the name was reached. Blank out
       "Family Greeting" and "Opening Blessing Line" and the ReferenceError
       escapes this IIFE and the outer one, which silently skipped the
       family list, the contacts and the logistics after it. */
    var any = blessing || (order.names || []).length || ask || dates;
    sec.hidden = !any;
  })();


  /* ---------------------------------------------------- eagerly awaiting */
  /* The named relatives, grouped by side. Grouping is what makes this read
     like a printed card instead of a guest list: the two families sit side by
     side, each name with its relation underneath. */
  (function awaitingSection() {
    var sec = $("awaiting");
    var list = $("awaitingList");
    if (!sec || !list) return;

    var members = forAudience(family.members, "family").filter(function (m) {
      return m && String(m.name || "").trim();
    });

    var title = F["Awaiting Section Title"] || "Eagerly Awaiting Your Presence";
    text($("awaitingTitle"), title);

    /* The paragraph ships with the `hidden` attribute, so writing text into it
       is not enough — it has to be un-hidden the way the family note above
       is, or anything typed into "Awaiting Section Note" renders nowhere,
       on screen or in print, without a single error. */
    var awaitNote = $("awaitingNote");
    text(awaitNote, F["Awaiting Section Note"] || "");
    if (awaitNote) awaitNote.hidden = !String(F["Awaiting Section Note"] || "").trim();

    if (!members.length) {
      sec.hidden = true;
      return;
    }

    /* Keep the workbook's side order stable and put anything unrecognised last
       under no heading, rather than sorting names around and hiding which
       family a relative belongs to.

       The leading family comes first, so the bride's version prints her own
       relatives before his. Falls back to the old fixed order if the build
       somehow left data.order out. */
    var sideOrder = (order && order.sides && order.sides.length)
      ? order.sides
      : ["Groom's side", "Bride's side", "Both"];
    var groups = [];
    sideOrder.forEach(function (side) {
      var inGroup = members.filter(function (m) {
        return String(m.side || "").trim() === side;
      });
      if (inGroup.length) groups.push({ side: side, people: inGroup });
    });
    var other = members.filter(function (m) {
      return sideOrder.indexOf(String(m.side || "").trim()) === -1;
    });
    if (other.length) groups.push({ side: "", people: other });

    /* Only label the groups when there is more than one. The card is printed
       by one family, so a single list headed by that family's own side tells
       the reader nothing they do not already know from who sent it. If
       relatives are later added to the other side, both headings come back. */
    var showSide = groups.length > 1;

    groups.forEach(function (group) {
      var wrap = document.createElement("div");
      wrap.className = "awaiting__group";

      if (showSide && group.side) {
        var h = document.createElement("h3");
        h.className = "awaiting__side";
        h.textContent = group.side;
        wrap.appendChild(h);
      }

      var ul = document.createElement("ul");
      ul.className = "awaiting__names";
      group.people.forEach(function (m) {
        var li = document.createElement("li");
        li.className = "awaiting__person";

        var nameEl = document.createElement("span");
        nameEl.className = "awaiting__name";
        nameEl.textContent = m.name;
        li.appendChild(nameEl);

        /* Relation and hometown are both optional; only the bits actually
           filled in are appended, so a row with only a name stays a clean
           single line. */
        var detail = [m.relation, m.from].filter(Boolean).join("  ·  ");
        if (detail) {
          var d = document.createElement("span");
          d.className = "awaiting__detail";
          d.textContent = detail;
          li.appendChild(d);
        }
        ul.appendChild(li);
      });
      wrap.appendChild(ul);
      list.appendChild(wrap);
    });

    sec.hidden = false;
  })();

  /* The "Who To Call" section moved to app.js. It is on both cards now, so it
     is rendered once, in the shared script, and this file no longer owns it.
     That is also why `family.contacts` is not read here any more. */

  /* The "Practical Details" blocks moved into the Venue & Directions section,
     and with them into app.js. They were a section of their own, and therefore
     family-page only, which hid the parking and the hotel from anyone reading
     the friends link. Same reason the "Who To Call" section moved: it is on
     both cards now, so the shared script owns it and this file does not. */
})();
