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
    return d.toLocaleDateString("en-GB", opts || { weekday: "long", day: "numeric", month: "long", year: "numeric" });
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
  setText("heroGroom", D["Groom Name"], "Groom");
  setText("heroBride", D["Bride Name"], "Bride");
  setText("heroLine", D["Couple Line (short)"], "Together with their families");

  var heroDate = D["Wedding Date"] || (events[0] && events[0].date) || "";
  if (heroDate) {
    var hd = new Date(heroDate + "T00:00:00");
    if (!isNaN(hd)) {
      setText("heroDay", hd.toLocaleDateString("en-GB", { day: "numeric" }));
      setText("heroMonth", hd.toLocaleDateString("en-GB", { month: "long" }));
      setText("heroYear", hd.getFullYear());
    }
  }
  setText("heroVenue", D["Primary Venue Name"] ? [D["Primary Venue Name"], D["City"]].filter(Boolean).join(", ") : "");
  setText("topHashtag", D["Hashtag"]);
  setText("footHashtag", D["Hashtag"]);
  setText("footContact", [D["Contact Name"], D["Contact Phone"]].filter(Boolean).join(" · "));
  setText("groomParents", D["Groom's Parents"]);
  setText("brideParents", D["Bride's Parents"]);
  $("families").hidden = !D["Groom's Parents"] && !D["Bride's Parents"];

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
  setText("countdownLabel", (data.countdown && data.countdown.event) ? "Counting down to " + data.countdown.event : "Counting down to the wedding");

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
  $("scheduleSub").textContent = D["Function Dates"] ? "All celebrations, " + D["Function Dates"] : "";

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
      tab.innerHTML = d.toLocaleDateString("en-GB", { day: "numeric", month: "short" }) +
        '<span class="tab__dow">' + esc(d.toLocaleDateString("en-GB", { weekday: "long" })) + "</span>";
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

  /* ============================================================ icons */
  /* Icon markup is inlined by build_site.py. The drawings are stroked with
     currentColor, so an <img> would render them black instead of themed
     maroon. If a row somehow has no icon, fall back to a plain frame. */
  var EMPTY_ICON =
    '<svg class="icon" viewBox="0 0 64 64" fill="none" stroke="currentColor" ' +
    'stroke-width="1.8" aria-hidden="true" focusable="false">' +
    '<circle cx="32" cy="32" r="20" opacity=".45"/>' +
    '<path d="M32 20v24M20 32h24" opacity=".45"/></svg>';

  function iconMarkup(ev) {
    return ev && ev.iconSvg ? ev.iconSvg : EMPTY_ICON;
  }

  /* ============================================================ rituals */
  /* One card per ritual, so a guest who does not know the customs can read
     what the function is about before they arrive. */
  function renderRituals() {
    var grid = $("ritualGrid");
    if (!grid) return;
    var seen = {};
    var cards = [];
    events.forEach(function (ev) {
      if (!ev.icon || seen[ev.icon]) return;
      seen[ev.icon] = true;
      cards.push({
        svg: ev.iconSvg,
        name: ev.event,
        when: niceDate(ev.date, { day: "numeric", month: "short" }),
      });
    });
    $("rituals").hidden = cards.length === 0;
    grid.innerHTML = cards.map(function (c) {
      return '<div class="ritual">' +
        '<div class="ritual__icon" aria-hidden="true">' + iconMarkup(c) + "</div>" +
        '<p class="ritual__name">' + esc(c.name) + "</p>" +
        '<p class="ritual__when">' + esc(c.when) + "</p>" +
        "</div>";
    }).join("");
  }

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
        functions: "Wedding / main function"
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
          '<a class="btn btn--gold btn--sm" target="_blank" rel="noopener" href="' + esc(mapHrefFor(v)) + '">Open in Maps</a>' +
          '<button class="btn btn--ghost btn--sm" data-copy="' + esc(v.address || v.name) + '" type="button">Copy address</button>' +
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

  /* Rendered before initReveal below, so the ritual cards exist in the DOM by
     the time it goes looking for targets to animate. */
  renderRituals();

  /* ============================================================ scroll reveal
     Runs last, so every card the renderers created is in the DOM before it
     looks for them. Two deliberate choices:

     - The `data-reveal` attribute is added *here* and only after confirming
       IntersectionObserver exists. The stylesheet hides `[data-reveal]`, so if
       the observer were unavailable and the attribute were still applied, the
       page would go permanently blank. Gating the attribute on the feature
       means any other environment simply never gets the hiding rule.
     - Each target is unobserved once it has been revealed, so scrolling back
       up re-runs nothing and there is no long-lived observer cost. */
  (function initReveal() {
    if (!("IntersectionObserver" in window)) return;
    if (window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;

    var blocks = [
      ".hero__card", ".countdown", ".banner", ".families",
      "#schedule", "#feature", "#venue", "#rituals", "#rsvp", "#verse", ".footer"
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
  })();
})();

/* ============================================================ ceremonial intro
   Two independent mechanisms retire the overlay, so a failure in either is
   harmless. The CSS outro holds `opacity: 0; visibility: hidden` on its own
   (fill-mode forwards), and this removes the node outright a beat later. The
   timer is the backstop for the one case CSS cannot cover: a tab backgrounded
   at load, where some browsers never start the first animation at all.

   Nothing inside the overlay is focusable, so it cannot trap a keyboard user,
   and a click anywhere skips it. */
(function () {
  "use strict";

  var intro = document.getElementById("intro");
  if (!intro) return;

  var D = (window.WEDDING_DATA || {}).details || {};
  var groom = String(D["Groom Name"] || "").trim();
  var bride = String(D["Bride Name"] || "").trim();

  var line = intro.querySelector(".intro__names");
  if (groom && bride) {
    document.getElementById("introGroom").textContent = groom;
    document.getElementById("introBride").textContent = bride;
  } else if (line) {
    /* Never show a lone ampersand if a name is missing from the workbook. */
    line.remove();
  }

  /* Plays on every load, including a refresh. It used to be suppressed for
     the rest of the session via sessionStorage, which meant a guest who
     reloaded or followed an #rsvp link never saw it again. */
  var reduce = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  if (reduce) { intro.remove(); return; }

  var done = false;
  function dismiss() {
    if (done) return;
    done = true;
    intro.remove();
  }

  /* No click handler: the overlay is `pointer-events: none` so it can never
     intercept a tap meant for the page underneath. The timer is the real
     guarantee, and it is re-armed on the way back into the foreground because
     a tab that was hidden at load often never starts its first animation at
     all, which would otherwise leave the card sitting on screen. */
  setTimeout(dismiss, 2600);
  document.addEventListener("visibilitychange", function () {
    if (!document.hidden) setTimeout(dismiss, 120);
  });
  window.addEventListener("load", function () { setTimeout(dismiss, 2600); });
})();
