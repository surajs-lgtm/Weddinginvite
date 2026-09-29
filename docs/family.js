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
   edit happens in wedding-details.xlsx followed by build_site.py.

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

  /* The one phrase this file writes that app.js does not already own: the note
     under the rituals heading, which is rewritten to match however many
     verified rituals are actually shown. English unless data.ui says else. */
  function ritualsSub(shown, total) {
    var ui = data.ui || {};
    if (ui.ritualsExplained) {
      return ui.ritualsExplained.replace("{shown}", shown).replace("{total}", total);
    }
    return shown + " of the " + total +
      " functions are explained here. The rest are listed in the schedule below.";
  }
  var events = data.events || [];
  var family = data.family || {};
  var F = family.details || {};
  var contacts = family.contacts || [];
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

  function el(tag, className, content) {
    var n = document.createElement(tag);
    if (className) n.className = className;
    if (content != null) n.textContent = content;
    return n;
  }

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
       rituals, the family list, the contacts and the logistics after it. */
    var any = blessing || (order.names || []).length || ask || dates;
    sec.hidden = !any;
  })();

  /* ------------------------------------------------------------ rituals */
  /* One card per function that has a verified explanation, in schedule order.
     Unverified rows never reach here: build_site.py drops them when it reads
     the Rituals sheet, so a guess cannot be shown to a relative. */
  (function rituals() {
    var sec = $("rituals");
    var list = $("ritualList");
    if (!sec || !list) return;

    var shown = 0;

    events.forEach(function (ev) {
      var r = ev.ritual || {};
      /* Only the verified ritual text opens a card. This used to read
         `r.familyNote` as well, but familyNote lives on the event, not the
         ritual, so that term was always undefined and the check rested
         entirely on what/why. Left as the explicit pair it actually is: the
         guarantee that nothing unverified renders is build_site.py dropping
         those rows at the Rituals sheet, and `ritual` is `{}` without one. */
      if (!r.what && !r.why) return;

      var card = el("article", "ritual");

      if (ev.iconSvg) {
        var ic = el("div", "ritual__icon");
        ic.setAttribute("aria-hidden", "true");
        ic.innerHTML = ev.iconSvg;
        card.appendChild(ic);
      }

      var head = el("div", "ritual__head");
      head.appendChild(el("h3", "ritual__name", ev.event));

      var when = [niceDate(ev.date), ev.timeStart].filter(Boolean).join(" · ");
      head.appendChild(el("p", "ritual__when", when));
      card.appendChild(head);

      if (r.what) card.appendChild(el("p", "ritual__what", r.what));
      if (r.why) card.appendChild(el("p", "ritual__why", r.why));

      /* The family note is the part a relative cannot look up: who does this,
         what to bring, what it means for them specifically. */
      var note = r.note || ev.familyNote;
      if (note) card.appendChild(el("p", "ritual__note", note));

      list.appendChild(card);
      shown++;
    });

    /* The heading promises every function, but only the verified ones are
       listed. A relative counting twelve rows against seven cards would think
       five had gone missing, so the count is rewritten to what is shown. */
    if (shown && shown !== events.length) {
      var sub = sec.querySelector(".section__sub");
      if (sub) {
        sub.textContent = ritualsSub(shown, events.length);
      }
    }

    sec.hidden = shown === 0;
  })();

  /* ---------------------------------------------------- eagerly awaiting */
  /* The named relatives, grouped by side. Grouping is what makes this read
     like a printed card instead of a guest list: the two families sit side by
     side, each name with its relation underneath. */
  (function awaitingSection() {
    var sec = $("awaiting");
    var list = $("awaitingList");
    if (!sec || !list) return;

    var members = (family.members || []).filter(function (m) {
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

  /* ---------------------------------------------------------- contacts */
  /* Rows from the Family Contacts sheet, then the two named contacts from
     Family Details, deduped so the same person does not appear twice. */
  (function contactsSection() {
    var sec = $("contactsSec");
    var list = $("contactList");
    if (!sec || !list) return;

    var people = [];
    var seen = {};

    function add(name, relation, phone, about) {
      name = (name || "").trim();
      phone = (phone || "").trim();
      if (!name && !phone) return;
      var key = (name + "|" + phone).toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      people.push({ name: name, relation: relation, phone: phone, about: about });
    }

    contacts.forEach(function (c) { add(c.name, c.relation, c.phone, c.about); });

    [1, 2].forEach(function (i) {
      add(
        F["Family Contact " + i + " Name"],
        F["Family Contact " + i + " Role"],
        F["Family Contact " + i + " Phone"],
        ""
      );
    });

    if (!people.length) { sec.hidden = true; return; }

    people.forEach(function (p) {
      var card = el("div", "contact");
      if (p.name) card.appendChild(el("p", "contact__name", p.name));
      if (p.relation) card.appendChild(el("p", "contact__role", p.relation));
      if (p.about) card.appendChild(el("p", "contact__about", p.about));
      if (p.phone) {
        var a = el("a", "contact__phone", p.phone);
        a.href = telHref(p.phone);
        card.appendChild(a);
      }
      list.appendChild(card);
    });

    sec.hidden = false;
  })();

  /* ---------------------------------------------------------- logistics */
  (function logistics() {
    var sec = $("logistics");
    var list = $("logisticsList");
    if (!sec || !list) return;

    var blocks = [
      ["Baraat Route", "Where the baraat gathers, and where it goes"],
      ["Parking", "Where to park"],
      ["Accommodation", "Places to stay for out-of-town relatives"],
      ["Gift / Shagun Note", "A note on gifts"]
    ];

    var shown = 0;
    blocks.forEach(function (pair) {
      var value = (F[pair[0]] || "").trim();
      if (!value) return;

      var block = el("div", "logistic");
      block.appendChild(el("h3", "logistic__title", pair[0]));
      /* the second element is the default hint; the written value replaces it */
      block.appendChild(el("p", "logistic__text", value || pair[1]));
      list.appendChild(block);
      shown++;
    });

    sec.hidden = shown === 0;
  })();
})();
