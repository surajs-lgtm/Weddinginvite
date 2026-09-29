/* ==========================================================================
   Wedding invitation - reads window.WEDDING_DATA (site/data.js)
   ========================================================================== */
(function () {
  "use strict";

  var RSVPS = "wedding-rsvps-v1";

  /* ------------------------------------------------------------ helpers */
  function $(id) { return document.getElementById(id); }
  function esc(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }
  function digits(value) {
    return String(value || "").replace(/[^\d]/g, "");
  }
  function niceDate(iso, opts) {
    if (!iso) return "";
    var d = new Date(iso + "T00:00:00");
    if (isNaN(d)) return iso;
    return d.toLocaleDateString(DATE_LOCALE, opts || { weekday: "long", day: "numeric", month: "long", year: "numeric" });
  }
  function todayISO() {
    var n = new Date();
    return n.getFullYear() + "-" + String(n.getMonth() + 1).padStart(2, "0") + "-" + String(n.getDate()).padStart(2, "0");
  }

  /* ------------------------------------------------------------ data */
  /* Read straight from data.js. There is deliberately no localStorage override
     and no in-page editor: this page is opened by guests, so nothing here may
     let a visitor rewrite the couple's details. Edit wedding-details.xlsx and
     run `python3 build_site.py` instead. */
  var data = JSON.parse(JSON.stringify(window.WEDDING_DATA || { details: {}, events: [], days: [], venues: [] }));
  var D = data.details || {};

  /* Weekdays and month names come from the browser's own date data rather than
     a table kept in this file, so a translated card gets the names, and the
     date order, of the language it is actually in. en-GB is the default, which
     is what the English card has always used. */
  var DATE_LOCALE = data.lang === "hi" ? "hi-IN" : "en-GB";

  /* The handful of phrases this file composes itself, rather than reading whole
     from the workbook -- the countdown label, the schedule subtitle, the venue
     card's caption and its two buttons. They are English here, and a translated
     card supplies its own in data.ui. Reading them from one table means a
     translation does not have to be chased through the script. */
  var UI_DEFAULT = {
    countingDownTo: "Counting down to",
    countingDownToWedding: "Counting down to the wedding",
    allCelebrations: "All celebrations,",
    venueCaption: "Wedding / main function",
    openInMaps: "Open in Maps",
    copyAddress: "Copy address"
  };
  var UI = data.ui || {};
  function ui(key) { return UI[key] || UI_DEFAULT[key]; }

  /* Which side of the couple this card was printed for. build_site.py works
     this out once from the workbook's "Card Flank" and hands it over in
     data.order, so the names, the parents, the relations list and the
     hashtag all read the same answer instead of each deciding for itself. */
  var ORDER = data.order || { flank: "groom", names: [], sides: [] };
  var LEAD = ORDER.flank === "bride" ? "bride" : "groom";
  var FOLLOW = LEAD === "bride" ? "groom" : "bride";

  /* Put the two names in the order the card leads.

     These move real nodes, not CSS `order`. Ordering alone looks correct but
     leaves the document order alone, and this card is read in that order by a
     screen reader, printed in it, and indexed in it -- so a card whose whole
     point is that Priyanka comes first would still announce "Suraj and
     Priyanka" to anyone not looking at the screen. The ampersand is the middle
     element and has to stay the middle element, so it is put back between the
     two after the pair has been swapped. Idempotent: doing it twice is a no-op. */
  function placeNames(groomId, brideId, ampId) {
    var groom = $(groomId), bride = $(brideId);
    if (!groom || !bride) return;
    groom.textContent = D["Groom Name"] || "";
    bride.textContent = D["Bride Name"] || "";
    var lead = LEAD === "bride" ? bride : groom;
    var follow = LEAD === "bride" ? groom : bride;
    var parent = groom.parentNode;
    if (!parent) return;
    parent.insertBefore(lead, follow);
    var amp = ampId && $(ampId);
    if (amp && amp.parentNode === parent) parent.insertBefore(amp, follow);
  }

  /* The two family blocks are siblings in a flex row, with a rule between
     them that is a sibling too. All three move, so the rule stays between the
     two families instead of being pushed outside them. */
  function placeFamilies(families) {
    if (!families) return;
    var lead = families.querySelector('[data-side="' + LEAD + '"]');
    var follow = families.querySelector('[data-side="' + FOLLOW + '"]');
    var rule = families.querySelector(".family__divider");
    if (!lead || !follow) return;
    families.insertBefore(lead, follow);
    if (rule) families.insertBefore(rule, follow);
  }
  var events = data.events || [];
  var days = data.days || [];
  var venues = data.venues || [];

  /* ============================================================ meta */
  var coupleLine = D["Couple Line (short)"] ||
    [D["Groom Name"], D["Bride Name"]].filter(Boolean).join(" weds ") || "We're Getting Married";
  var title = D["Website Title"] || coupleLine;
  document.title = title;
  var summary = [D["Primary Venue Name"], D["City"], D["Function Dates"]].filter(Boolean).join(" · ") || "Wedding invitation and schedule";
  $("metaDesc").setAttribute("content", summary);
  $("ogTitle").setAttribute("content", title);
  $("ogDesc").setAttribute("content", summary);

  if (D["Accent Style"]) document.body.dataset.style = D["Accent Style"];

  /* ============================================================ hero */
  function setText(id, value, fallback) {
    var el = $(id);
    if (!el) return;
    el.textContent = value || fallback || "";
  }
  setText("heroBlessing", D["Opening Blessing Line"], "With the blessings of our families");
  placeNames("heroGroom", "heroBride", "heroAmp");
    /* Invitation Line, not Couple Line (short). The latter renders as "Suraj
       weds Priyanka" directly beneath the names, which already read "Suraj &
       Priyanka" — the same two names twice in three lines. The invitation
       phrase completes the classic request structure instead: blessing above
       the names, request below them. Falls back to the couple line, and then
       to a generic phrase, so an empty workbook still renders a sensible line. */
    setText("heroLine", D["Invitation Line"] || D["Couple Line (short)"], "Together with their families");

  /* build_site.py emits weddingDateISO alongside the workbook's human string
     because "11 Dec 2026" + "T00:00:00" is not a date any engine is obliged to
     parse, and whether the hero showed its date came down to which fallback
     parser the browser reached for first. The ISO value is the same string
     every time.

     The separators are static markup, so a failed parse would leave two bare
     gold dots sitting where the date should be. Drop the whole group instead. */
  var heroDate = data.weddingDateISO || "";
  if (heroDate) {
    var hd = new Date(heroDate + "T00:00:00");
    if (!isNaN(hd.getTime())) {
      setText("heroDay", hd.toLocaleDateString(DATE_LOCALE, { day: "numeric" }));
      setText("heroMonth", hd.toLocaleDateString(DATE_LOCALE, { month: "long" }));
      setText("heroYear", hd.getFullYear());
    } else {
      heroDate = "";
    }
  } else {
    heroDate = "";
  }
  if (!heroDate) {
    var hdGroup = $("heroDate");
    if (hdGroup) hdGroup.hidden = true;
  }
  setText("topHashtag", D["Hashtag"]);
  setText("footHashtag", D["Hashtag"]);
  setText("footContact", [D["Contact Name"], D["Contact Phone"]].filter(Boolean).join(" · "));

  /* The families block lists father and mother on their own labelled lines
     rather than one joined string, so a reader can tell the two apart. A line
     with no name is removed, and a whole side is removed when both its lines
     go, so a half-filled block still reads as deliberate rather than broken. */
  var anyParent = false;
  ["groom", "bride"].forEach(function (side) {
    var sideName = side === "groom" ? "Groom" : "Bride";
    var block = $(side + "Parents");
    if (!block) return;
    ["Father", "Mother"].forEach(function (role) {
      var nameEl = $(side + role);
      var line = nameEl && nameEl.parentNode;
      if (nameEl) nameEl.textContent = D[sideName + "'s " + role] || "";
      if (line && !nameEl.textContent.trim()) line.parentNode.removeChild(line);
    });
    if (block.textContent.trim()) {
      anyParent = true;
      return;
    }
    /* Nothing on this side, so drop the whole column rather than leaving a
       "Groom's Family" heading with nothing under it. */
    var column = block.closest && block.closest(".family");
    if (column && column.parentNode) column.parentNode.removeChild(column);
  });
  $("families").hidden = !anyParent;
  /* After the empty columns are gone, put whichever family printed the card
     on the left. The divider between them is symmetric, so it needs no move. */
  placeFamilies($("families"));

  if (D["Quote / Verse"]) {
    setText("verseText", D["Quote / Verse"]);
    $("verse").hidden = false;
  }

  /* ============================================================ countdown */
  var target = data.countdown && data.countdown.date ? data.countdown.date : heroDate;
  var targetTime = null;
  if (target) {
    var firstStart = events.find(function (e) { return e.date === target; });
    var mins = firstStart && firstStart.startMinutes != null ? firstStart.startMinutes : 9 * 60;
    targetTime = new Date(target + "T00:00:00");
    targetTime.setMinutes(mins);
  }
  setText("countdownLabel", (data.countdown && data.countdown.event) ? ui("countingDownTo") + " " + data.countdown.event : ui("countingDownToWedding"));

  // The Baraat feature card is captioned from the same countdown object, so it
  // can never drift out of step with the countdown above it.
  (function () {
    var box = $("feature");
    if (!box || !data.countdown) return;
    var c = data.countdown;
    setText("featureName", c.event || "");
    var when = niceDate(c.date, { weekday: "long", day: "numeric", month: "long" });
    if (c.time) when += " · " + c.time;
    setText("featureDate", when);
    box.hidden = false;
  })();

  function tick() {
    if (!targetTime) { $("countdown").hidden = true; return; }
    var diff = targetTime - new Date();
    if (diff <= 0) {
      $("countdownGrid").hidden = true;
      $("countdownDone").hidden = false;
      return;
    }
    var s = Math.floor(diff / 1000);
    $("cDays").textContent = Math.floor(s / 86400);
    $("cHours").textContent = String(Math.floor(s / 3600) % 24).padStart(2, "0");
    $("cMins").textContent = String(Math.floor(s / 60) % 60).padStart(2, "0");
    $("cSecs").textContent = String(s % 60).padStart(2, "0");
  }
  tick();
  setInterval(tick, 1000);

  /* ============================================================ schedule */
  $("scheduleSub").textContent = D["Function Dates"] ? ui("allCelebrations") + " " + D["Function Dates"] : "";

  var tabs = $("dayTabs"), panels = $("dayPanels");
  if (!days.length) {
    $("schedule").hidden = true;
  } else {
    var today = todayISO();
    var activeIndex = Math.max(0, days.findIndex(function (d) { return d.date === today; }));

    days.forEach(function (day, i) {
      var d = new Date(day.date + "T00:00:00");
      var tab = document.createElement("button");
      tab.className = "tab";
      tab.type = "button";
      tab.setAttribute("role", "tab");
      tab.setAttribute("aria-selected", i === activeIndex ? "true" : "false");
      tab.innerHTML = d.toLocaleDateString(DATE_LOCALE, { day: "numeric", month: "short" }) +
        '<span class="tab__dow">' + esc(d.toLocaleDateString(DATE_LOCALE, { weekday: "long" })) + "</span>";
      tab.addEventListener("click", function () { selectDay(i); });
      tabs.appendChild(tab);
    });

    days.forEach(function (day, i) {
      var panel = document.createElement("div");
      panel.className = "day-panel";
      panel.id = "day-" + i;
      panel.setAttribute("role", "tabpanel");
      panel.hidden = i !== activeIndex;

      var heading = document.createElement("h3");
      heading.className = "section__title";
      heading.style.fontSize = "1.2rem";
      heading.style.marginBottom = "18px";
      heading.textContent = niceDate(day.date);
      panel.appendChild(heading);

      var list = document.createElement("div");
      list.className = "events";
      day.events.forEach(function (ev) { list.appendChild(eventRow(ev)); });
      panel.appendChild(list);
      panels.appendChild(panel);
    });

    function selectDay(index) {
      Array.prototype.forEach.call(tabs.children, function (b, i) {
        b.setAttribute("aria-selected", i === index ? "true" : "false");
      });
      Array.prototype.forEach.call(panels.children, function (p, i) {
        p.hidden = i !== index;
      });
    }
    selectDay(activeIndex);
  }

  function eventRow(ev) {
    var row = document.createElement("div");
    row.className = "event";
    row.dataset.date = ev.date;
    row.dataset.start = ev.startMinutes == null ? "" : ev.startMinutes;
    row.dataset.end = ev.endMinutes == null ? "" : ev.endMinutes;

    var time = ev.timeStart || ev.time;
    var icon =
      '<div class="event__icon" aria-hidden="true">' + iconMarkup(ev) + "</div>";

    row.innerHTML =
      icon +
      '<div class="event__time">' + esc(time || "Time to be announced") +
        (ev.timeEnd ? '<small>to ' + esc(ev.timeEnd) + "</small>" : "") + "</div>" +
      "<div><h4 class='event__name'>" + esc(ev.event) +
        (ev.note ? "<span class='event__tag'>" + esc(ev.note) + "</span>" : "") + "</h4></div>" +
      (ev.venue ? "<span class='event__venue'>" + esc(ev.venue) + "</span>" : "");
    return row;
  }

  // highlight whatever is running right now
  function markLive() {
    var now = new Date();
    var nowMin = now.getHours() * 60 + now.getMinutes();
    var today = todayISO();
    document.querySelectorAll(".event").forEach(function (row) {
      var start = row.dataset.start === "" ? null : Number(row.dataset.start);
      var end = row.dataset.end === "" ? null : Number(row.dataset.end);
      var live = row.dataset.date === today && start != null && nowMin >= start && (end == null || nowMin < end);
      row.dataset.now = live ? "true" : "false";
      var flag = row.querySelector(".event__happen");
      if (live && !flag) {
        var span = document.createElement("span");
        span.className = "event__happen";
        span.textContent = "Happening now";
        row.querySelector(".event__name").appendChild(span);
      } else if (!live && flag) {
        flag.remove();
      }
    });
  }
  markLive();
  setInterval(markLive, 60000);

  function iconMarkup(ev) {
    return ev && ev.iconSvg ? ev.iconSvg : emptyIcon();
  }

  /* The neutral frame for a function with no drawing of its own. A function,
     not a var holding a string, because the schedule is rendered further up
     this same scope (the day panels, ~line 171) and a `var` is hoisted without
     its value: reading it from there returned undefined, and the literal text
     "undefined" landed in the icon slot for every event without an icon. Four
     of the thirteen functions had no drawing, so it was visible. Returning the
     markup keeps the fallback independent of statement order. */
  function emptyIcon() {
    /* Generic enough to stand in for any function, so a row added to the
       workbook without a matching drawing still gets a mark rather than a gap.
       The loop of petals around a centre is deliberately not tied to one
       ritual: it reads as a flower, a thali and a lamp at the size it draws. */
    var petals = "";
    for (var i = 0; i < 8; i++) {
      petals += '<ellipse cx="32" cy="19" rx="4" ry="8" transform="rotate(' +
        (i * 45) + ' 32 32)" opacity=".45"/>';
    }
    return '<svg class="icon" viewBox="0 0 64 64" fill="none" stroke="currentColor" ' +
      'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" ' +
      'aria-hidden="true" focusable="false">' +
      petals +
      '<circle cx="32" cy="32" r="6" fill="currentColor" opacity=".7"/>' +
      '<circle cx="32" cy="32" r="24" opacity=".25"/></svg>';
  }

  /* ============================================================ rituals */
  /* One card per ritual, so a guest who does not know the customs can read
     what the function is about before they arrive. */
  /* ============================================================ blessings
     Replaces renderRituals. The old grid rendered one card per function —
     a name and a date, twelve of them — which repeated the Schedule and
     never actually explained the rituals as its own copy claimed to. This
     renders a single ashirwad line instead.

     Stays hidden when the line is empty, so a workbook with the field left
     blank does not leave an "Ashirwad" heading over nothing. */
  (function renderBlessings() {
    var sec = $("blessings");
    var el = $("ashirwadText");
    if (!sec || !el) return;
    var line = (D["Ashirwad Line"] || "").trim();
    el.textContent = line;
    sec.hidden = !line;
  })();

  /* ============================================================ venues */
  var vgrid = $("venueGrid");
  vgrid.addEventListener("click", function (e) {
    var btn = e.target.closest("[data-copy]");
    if (!btn) return;
    navigator.clipboard.writeText(btn.dataset.copy).then(function () {
      var was = btn.textContent;
      btn.textContent = "Copied";
      setTimeout(function () { btn.textContent = was; }, 1400);
    });
  });

  function mapHrefFor(v) {
    return v.maps || ("https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(v.name + " " + (v.address || "")));
  }

  function renderVenues() {
    vgrid.innerHTML = "";
    var list = (data.venues || []).filter(function (v) { return v.name; });
    if (!list.length && D["Primary Venue Name"]) {
      list = [{
        name: D["Primary Venue Name"],
        address: D["Primary Venue Address"],
        maps: D["Google Maps Link"],
        functions: ui("venueCaption")
      }];
    }
    $("venue").hidden = !list.length;
    $("mapFloat").hidden = !list.length;
    if (!list.length) { $("mapFloat").removeAttribute("href"); return; }

    list.forEach(function (v) {
      var card = document.createElement("div");
      card.className = "venue";
      card.innerHTML =
        '<p class="venue__what">' + esc(v.functions || "Venue") + "</p>" +
        '<h3 class="venue__name">' + esc(v.name) + "</h3>" +
        (v.address ? '<p class="venue__addr">' + esc(v.address) + "</p>" : "") +
        '<div class="venue__links">' +
          '<a class="btn btn--gold btn--sm" target="_blank" rel="noopener" href="' + esc(mapHrefFor(v)) + '">' + esc(ui("openInMaps")) + "</a>" +
          '<button class="btn btn--ghost btn--sm" data-copy="' + esc(v.address || v.name) + '" type="button">' + esc(ui("copyAddress")) + "</button>" +
        "</div>";
      vgrid.appendChild(card);
    });
    $("mapFloat").href = mapHrefFor(list[0]);
  }
  renderVenues();

  /* ============================================================ calendar */
  function buildICS() {
    var now = new Date().toISOString().replace(/[-:]/g, "").split(".")[0] + "Z";
    function icsDate(iso, minutes) {
      var d = new Date(iso + "T00:00:00");
      d.setMinutes(minutes == null ? 0 : minutes);
      return d.toISOString().replace(/[-:]/g, "").split(".")[0] + "Z";
    }
    var lines = [
      "BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Wedding Invitation//EN", "CALSCALE:GREGORIAN"
    ];
    events.forEach(function (ev) {
      var place = ev.venue || D["Primary Venue Name"] || "";
      var addr = D["Primary Venue Address"] || "";
      var where = [place, addr, D["City"]].filter(Boolean).join(", ");
      lines.push(
        "BEGIN:VEVENT",
        "UID:" + ev.date + "-" + ev.order + "@wedding",
        "DTSTAMP:" + now,
        "DTSTART:" + icsDate(ev.date, ev.startMinutes),
        "DTEND:" + icsDate(ev.date, ev.endMinutes == null ? (ev.startMinutes == null ? null : ev.startMinutes + 60) : ev.endMinutes),
        "SUMMARY:" + coupleLine + " - " + ev.event,
        "LOCATION:" + where,
        "DESCRIPTION:" + (ev.note || coupleLine) +
          (where ? "\\nMaps: " + (D["Google Maps Link"] || "") : ""),
        "END:VEVENT"
      );
    });
    lines.push("END:VCALENDAR");
    return lines.join("\r\n");
  }
  $("addCalendar").addEventListener("click", function () {
    if (!events.length) { alert("No functions to add yet."); return; }
    var blob = new Blob([buildICS()], { type: "text/calendar" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "wedding-schedule.ics";
    a.click();
    URL.revokeObjectURL(url);
  });

  /* ============================================================ share */
  $("shareBtn").addEventListener("click", async function () {
    var shareData = {
      title: title,
      text: coupleLine + (D["Function Dates"] ? " · " + D["Function Dates"] : ""),
      url: location.href
    };
    if (navigator.share) {
      try { await navigator.share(shareData); return; } catch (e) { /* user cancelled */ }
    }
    try {
      await navigator.clipboard.writeText(location.href);
      var btn = this;
      var label = btn.querySelector("span");
      label.textContent = "Copied";
      setTimeout(function () { label.textContent = "Share"; }, 1600);
    } catch (e) {
      prompt("Copy this link to share the invitation:", location.href);
    }
  });

  /* ============================================================ RSVP */
  var form = $("rsvpForm");
  var waBtn = $("whatsappBtn");
  setText("rsvpDeadline", D["RSVP Deadline"] ? "Please respond by " + D["RSVP Deadline"] : "");

  function waLink(text) {
    var phone = digits(D["WhatsApp Number"]);
    if (!phone) return "";
    return "https://wa.me/" + phone + "?text=" + encodeURIComponent(text);
  }
  var waFloatHref = waLink(coupleLine + " — I'd love to attend!");
  /* Both branches are required. The element ships with `hidden` in the markup,
     so setting the href alone leaves it invisible forever. */
  if (waFloatHref) {
    $("waFloat").href = waFloatHref;
    $("waFloat").hidden = false;
  } else {
    $("waFloat").hidden = true;
  }

  function readRSVPs() {
    try { return JSON.parse(localStorage.getItem(RSVPS) || "[]"); } catch (e) { return []; }
  }
  function writeRSVPs(list) {
    try { localStorage.setItem(RSVPS, JSON.stringify(list)); } catch (e) { /* ignore */ }
  }
  function status(text, tone) {
    var el = $("formStatus");
    el.textContent = text;
    if (tone) el.dataset.tone = tone; else delete el.dataset.tone;
  }
  function statusWord(value) {
    return value === "Attending" ? "joyfully accepts" : value === "Declined" ? "regretfully declines" : "is unsure yet";
  }
  function refreshSaved() {
    var list = readRSVPs();
    $("savedBox").hidden = list.length === 0;
    $("savedCount").textContent = list.length;
    $("savedList").innerHTML = list.map(function (r) {
      return '<div class="saved-list-item"><b>' + esc(r.name) + "</b><span>" +
        esc(r.guests) + (r.guests === 1 ? " guest" : " guests") + " · " +
        esc(r.status) + (r.meal ? " · " + esc(r.meal) : "") + "</span></div>";
    }).join("");
  }
  refreshSaved();

  form.addEventListener("submit", function (e) {
    e.preventDefault();
    var name = $("rName").value.trim();
    if (!name) { status("Please tell us your name.", "err"); $("rName").focus(); return; }

    var record = {
      name: name,
      relation: $("rRelation").value.trim(),
      phone: $("rPhone").value.trim(),
      guests: Number($("rGuests").value) || 1,
      status: (form.querySelector('input[name="status"]:checked') || {}).value || "Attending",
      meal: $("rMeal").value,
      note: $("rNote").value.trim(),
      at: new Date().toISOString()
    };

    var list = readRSVPs().filter(function (r) {
      var samePerson = record.phone ? r.phone === record.phone : r.name === record.name;
      return !samePerson;
    });
    list.push(record);
    writeRSVPs(list);
    refreshSaved();

    var message = coupleLine + "\n" +
      name + " " + statusWord(record.status) + ".\n" +
      "Guests: " + record.guests + "\n" +
      (record.relation ? "Relation: " + record.relation + "\n" : "") +
      (record.meal ? "Meal: " + record.meal + "\n" : "") +
      (record.note ? "Message: " + record.note + "\n" : "") +
      (D["Primary Venue Name"] ? "Venue: " + D["Primary Venue Name"] + "\n" : "");

    var href = waLink(message);
    if (href) {
      waBtn.href = href;
      waBtn.hidden = false;
      waBtn.textContent = "Send on WhatsApp";
    }
    status("Thank you, " + name + "! Your RSVP is saved on this device" +
      (href ? " — tap WhatsApp to send it to us." : "."), "ok");
    form.reset();
    $("rName").focus();
  });

  $("clearSaved").addEventListener("click", function () {
    if (!confirm("Remove all saved RSVPs from this device?")) return;
    writeRSVPs([]);
    refreshSaved();
    status("Cleared.", "ok");
  });

  /* Re-check which function is running once the page has finished loading. */
  window.addEventListener("load", function () { markLive(); });

  /* ============================================================ scroll reveal
     Runs last, so every card the renderers created is in the DOM before it
     looks for them. Two deliberate choices:

     - The `data-reveal` attribute is added *here* and only after confirming
       IntersectionObserver exists. The stylesheet hides `[data-reveal]`, so if
       the observer were unavailable and the attribute were still applied, the
       page would go permanently blank. Gating the attribute on the feature
       means any other environment simply never gets the hiding rule.
     - Each target is unobserved once it has been revealed, so scrolling back
       up re-runs nothing and there is no long-lived observer cost.

     This is deferred to DOMContentLoaded rather than run inline. app.js is a
     blocking script at the foot of the body, so anything it does here happens
     the moment it is parsed — which on family.html is *before* family.js, a
     deferred script, has created a single ritual card, contact or name. The
     comment above used to claim the opposite. Deferred scripts all run before
     DOMContentLoaded, so waiting for that event is what puts the family
     page's cards in the DOM first; on the friends page it is simply a few
     milliseconds later and nothing else changes. */
  function initReveal() {
    if (!("IntersectionObserver" in window)) return;
    if (window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;

    var blocks = [
      ".hero__card", ".hero__cta", ".countdown", ".banner", ".families",
      "#schedule", "#feature", "#venue", "#blessings", "#rsvp", "#verse", ".footer"
    ].join(",");
    var nodes = Array.prototype.slice.call(document.querySelectorAll(blocks))
      .concat(Array.prototype.slice.call(document.querySelectorAll(".event, .ritual, .venue, .count")))
      /* Anything still hidden — an unpopulated venue, a verse nobody filled
         in — is skipped rather than left stranded at opacity 0. */
      .filter(function (n) { return n.offsetParent !== null || n.getClientRects().length; });

    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) return;
        entry.target.classList.add("is-in");
        io.unobserve(entry.target);
      });
    }, { rootMargin: "0px 0px -6% 0px", threshold: 0.05 });

    nodes.forEach(function (n) {
      n.setAttribute("data-reveal", "");
      /* Stagger siblings so a grid of cards arrives in sequence rather than
         as one block. Capped, so a twelve-item ritual grid does not take
         most of a second to finish. */
      var index = Array.prototype.indexOf.call(n.parentNode.children, n);
      n.style.setProperty("--reveal-delay", Math.min(Math.max(index, 0), 7) * 65 + "ms");
      io.observe(n);
    });

    /* Safety net. The observer is reliable in every browser this site will
       meet, but the failure it guards against is not a wrong animation, it is
       a permanently invisible section — which on a wedding invitation is the
       worst possible outcome. Generous enough that a guest who scrolls at
       human speed still sees every reveal fire normally, short enough that
       anything genuinely stranded self-heals. */
    setTimeout(function () {
      nodes.forEach(function (n) { n.classList.add("is-in"); });
      io.disconnect();
    }, 8000);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initReveal);
  } else {
    initReveal();
  }
})();

/* ============================================================ card intro
   Two beats. Tap one lifts the cover; tap two reveals the site. In between,
   the inner face is given time to settle — arriving all at once reads as a
   glitch rather than a card opening.

   Both beats are optional. The CSS timings carry the identical sequence on
   their own, and the Continue control is display:none without script, so a
   guest who never taps anything still gets to the site and never sees a
   control that cannot be pressed. That matters because the old intro's
   `pointer-events: none` guarantee is gone the moment the card is tappable.

   The tap that opens the cover is also the only thing that can start the
   music, because browsers refuse to play sound that did not begin inside a
   real user gesture. There is no autoplay path to guard against. */
(function () {
  "use strict";

  var intro = document.getElementById("intro");
  if (!intro) return;

  var D = (window.WEDDING_DATA || {}).details || {};
  /* Same order block the hero read, re-read from the global because this is a
     separate IIFE. Both come from one value, so they cannot disagree. */
  var INTRO_ORDER = (window.WEDDING_DATA || {}).order || {};
  var INTRO_LEAD = INTRO_ORDER.flank === "bride" ? "bride" : "groom";
  var INTRO_FOLLOW = INTRO_LEAD === "bride" ? "groom" : "bride";
  /* Declared again rather than borrowed from the IIFE above. That one is
     closed by the time this runs, so a reference to it here would not be a
     compile error -- it would be `undefined` at runtime, and
     toLocaleDateString(undefined) quietly falls back to the browser's own
     locale. The card date would then read in whatever language the guest's
     device is set to instead of the card's. */
  var INTRO_LANG = (window.WEDDING_DATA || {}).lang;
  var INTRO_DATE_LOCALE = INTRO_LANG === "hi" ? "hi-IN" : "en-GB";
  function field(name) { return String(D[name] || "").trim(); }
  function put(id, value) {
    var el = document.getElementById(id);
    if (el && value) el.textContent = value;
    return !!value;
  }

  /* Names and parents both lead with whoever printed the card. The two spans
     keep their role-named ids and are physically moved, not just restyled, so
     the order a screen reader reads is the order the card shows. */
  var INTRO_GROOM = { name: "introGroom", parents: "introParentsGroom" };
  var INTRO_BRIDE = { name: "introBride", parents: "introParentsBride" };
  var leadPair = INTRO_LEAD === "bride" ? INTRO_BRIDE : INTRO_GROOM;
  var followPair = INTRO_LEAD === "bride" ? INTRO_GROOM : INTRO_BRIDE;

  put(leadPair.name, field(INTRO_LEAD === "bride" ? "Bride Name" : "Groom Name"));
  put(followPair.name, field(INTRO_FOLLOW === "bride" ? "Bride Name" : "Groom Name"));
  put(leadPair.parents, field((INTRO_LEAD === "bride" ? "Bride" : "Groom") + "'s Parents"));
  put(followPair.parents, field((INTRO_FOLLOW === "bride" ? "Bride" : "Groom") + "'s Parents"));

  /* Names first: the ampersand has to stay between the two, so it is put back
     in the middle after the pair is swapped. */
  (function orderIntroNames() {
    var lead = document.getElementById(leadPair.name);
    var follow = document.getElementById(followPair.name);
    if (!lead || !follow) return;
    var parent = lead.parentNode;
    if (!parent) return;
    parent.insertBefore(lead, follow);
    var amp = document.getElementById("introAmp");
    if (amp && amp.parentNode === parent) parent.insertBefore(amp, follow);
  })();

  /* Then the parents. The two `.card__parent` blocks are what sit side by
     side, so those move -- the name span inside one of them is not a flex item
     and ordering it would do nothing. */
  (function orderIntroParents() {
    var leadSpan = document.getElementById(leadPair.parents);
    var followSpan = document.getElementById(followPair.parents);
    if (!leadSpan || !followSpan || !leadSpan.closest || !followSpan.closest) return;
    var lead = leadSpan.closest(".card__parent");
    var follow = followSpan.closest(".card__parent");
    if (lead && follow && lead.parentNode) lead.parentNode.insertBefore(lead, follow);
  })();

  put("introBlessing", field("Opening Blessing Line") || "With the blessings of our families");
  put("introInvite", field("Invitation Line") || "request the honour of your presence");

  /* Date only: "Friday 11 Dec". The year and the start time are left off so
     the card carries the one thing a guest checks first, and nothing else.

     Derived from the ISO date build_site.py emits rather than typed in. It was
     the only date on the page not coming from the workbook, so moving the
     wedding in the spreadsheet moved the hero, the countdown and the four day
     tabs while this line stayed on 11 Dec.

     Formatted here rather than with niceDate() because that is in the other
     IIFE and this one is deliberately standalone. */
  var introISO = (window.WEDDING_DATA || {}).weddingDateISO;
  var introWhen = "";
  if (introISO) {
    var id = new Date(introISO + "T00:00:00");
    if (!isNaN(id.getTime())) {
      introWhen = id.toLocaleDateString(INTRO_DATE_LOCALE, { weekday: "long", day: "numeric", month: "short" });
    }
  }
  put("introDate", introWhen);

  /* Drop a line rather than print an empty one. A wedding card with a blank
     ruled space looks like a bug; a shorter card looks intentional. */
  ["introParentsGroom", "introParentsBride"].forEach(function (id) {
    var el = document.getElementById(id);
    if (el && !el.textContent.trim() && el.parentNode) el.parentNode.remove();
  });
  if (!intro.querySelector(".card__parent")) {
    var parents = intro.querySelector(".card__parents");
    if (parents) parents.remove();
  }

  /* Reduced motion means no card at all — the page is simply the page. */
  if (window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches) {
    intro.remove();
    return;
  }

  var openBtn = document.getElementById("introOpen");
  var goBtn = document.getElementById("introGo");
  var done = false;
  var started = false;
  /* Every safety-net timer, so the first real interaction can cancel all of
     them. Without this the no-interaction fallback keeps counting from page
     load and cuts a guest off mid-reveal: someone who reads the card for ten
     seconds and then taps gets the site pulled out from under the opening
     cover, 3.1s later, while the inner face is still staggering in. */
  var nets = [];

  function net(fn, ms) {
    var id = setTimeout(fn, ms);
    nets.push(id);
    return id;
  }
  function clearNets() {
    nets.forEach(function (id) { clearTimeout(id); });
    nets = [];
  }

  function dismiss() {
    if (done) return;
    done = true;
    clearNets();
    intro.remove();
  }

  /* Beat two: hand the site over. */
  function proceed() {
    if (done) return;
    intro.classList.add("is-going");
    if (goBtn) goBtn.disabled = true;
    setTimeout(dismiss, 950);
  }

  /* Beat one: lift the cover and start the music. Driven entirely by a click —
     the card, the Open control, or anywhere on the overlay. */
  function open() {
    if (done || started) return;
    started = true;
    clearNets();
    intro.classList.add("is-open");
    if (openBtn) openBtn.disabled = true;
    playFlute();
    /* From here the invitation only moves on because the guest asks it to. The
       one long timer left in the file is a bail-out, not a UX path: if the
       Continue control's handler is ever broken, the overlay lets go after a
       minute instead of trapping the guest on it. */
    net(dismiss, 60000);
    /* Move focus off the now-disabled first control and onto the second, so a
        keyboard guest is not dropped back to the top of the document. Deferred
        to match the 2.4s reveal in the stylesheet — focusing a control that is
        still at opacity 0 would move the caret somewhere the guest cannot see.
        focus-visible keeps the ring off for anyone who got here with a
        pointer. */
    if (goBtn) {
      setTimeout(function () {
        if (done) return;
        goBtn.removeAttribute("tabindex");
        goBtn.focus({ preventScroll: true });
      }, 2400);
    }
  }

  if (openBtn) openBtn.addEventListener("click", open);
  if (goBtn) goBtn.addEventListener("click", proceed);
  /* "or anywhere on the page": a click on the overlay background — not just the
     card itself — lifts the cover. The guard inside open() makes this a no-op
     once the cover is up, so it can never race the Continue control. */
  intro.addEventListener("click", open);

  /* Click-driven, so there is no auto-advance left to cancel. The only timer
     left is one long bail-out, armed at load and re-armed on the way back from
     a backgrounded tab, so a guest who never touches anything still reaches the
     site instead of being stranded on the cover. It sits far beyond any real
     reading time and the first tap clears it. */
  net(dismiss, 60000);
  window.addEventListener("load", function () { net(dismiss, 60000); });
  document.addEventListener("visibilitychange", function () {
    /* A tab backgrounded at load can leave the very first animation unstarted,
       which is the one case CSS cannot rescue. Re-arm on the way back rather
       than cutting the card short: a guest who glanced at another tab should
       still find the card waiting, and still be able to open it — which
       matters, because the tap is the only thing that can start the music. */
    if (document.hidden || done || started) return;
    net(dismiss, 60000);
  });

  /* ============================================================ music
     Deliberately minimal, because the one thing that must not happen is a
     broken card. A missing file, a rejected play(), or a browser with no
     support all resolve to the same outcome: the card opens, and the sound
     control simply never appears.

     Muted preference is remembered, so a guest who turned it off is not
     ambushed by it on the next visit. */
  var audio = document.getElementById("flute");
  var soundBtn = document.getElementById("soundBtn");
  var KEY = "wedding-muted-v1";
  var muted = false;
  try { muted = localStorage.getItem(KEY) === "1"; } catch (e) { /* private mode */ }

  function showSound() {
    if (soundBtn) { soundBtn.hidden = false; paint(); }
  }
  function paint() {
    if (!soundBtn) return;
    soundBtn.setAttribute("aria-pressed", muted ? "true" : "false");
    soundBtn.title = muted ? "Play music" : "Mute music";
    var label = soundBtn.querySelector("span");
    if (label) label.textContent = muted ? "Muted" : "Sound";
  }

  function playFlute() {
    /* The one gate. Sound cannot start without a gesture, and the only
       gesture that counts is opening the card. */
    if (!audio || muted || !started) return;
    var p;
    try { p = audio.play(); } catch (e) { return; }
    if (p && typeof p.then === "function") {
      p.then(showSound).catch(function () { /* no file, or blocked */ });
    }
  }

  if (soundBtn) {
    paint();
    soundBtn.addEventListener("click", function () {
      muted = !muted;
      try { localStorage.setItem(KEY, muted ? "1" : "0"); } catch (e) { /* ignore */ }
      if (muted) {
        try { audio.pause(); } catch (e) { /* ignore */ }
      } else {
        playFlute();
      }
      paint();
    });
  }

  /* Never keep playing into a backgrounded tab, and never start it there
     either — resuming is gated on `started`, the same gesture flag that gates
     the first note. */
  document.addEventListener("visibilitychange", function () {
    if (!audio) return;
    if (document.hidden) { try { audio.pause(); } catch (e) { /* ignore */ } }
    else if (started && !muted && !audio.ended) {
      try { audio.play().catch(function () {}); } catch (e) { /* ignore */ }
    }
  });
})();
