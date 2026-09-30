# Wedding Invitation Website

A single-page wedding invitation built from the printed card. The Excel file is
the source of truth; the site is generated from it.

```
wedding planning/
├── groom-details.xlsx        <- edit this
├── bride-details.xlsx        <- her copy, edited independently
├── build_site.py            <- turns the workbook into docs/data.js
├── make_workbook.py         <- recreates the workbook from scratch
├── add_family_sheets.py     <- the family-only sheets' content, and a
│                                script to add those sheets to an existing file
├── make_art.py              <- draws the SVG artwork
├── check_art.py             <- checks the SVGs for broken geometry
└── docs/                    <- this is what gets hosted, and what Pages serves
    ├── index.html           <- the short version, for friends
    ├── family.html          <- the long version, for relatives
    ├── styles.css
    ├── app.js               <- shared by both pages
    ├── family.js            <- family.html only
    ├── data.js              <- generated, do not hand-edit
    └── assets/
        ├── *.svg            <- decoration (mandala, jali, toran, ...)
        ├── favicon.svg
        └── icons/*.svg      <- one line drawing per function
```

`groom-details.xlsx` is deliberately **not** committed. It is your private
copy, and once filled in it will hold real phone numbers, plus an RSVP List sheet
meant for guests' names and replies. It does not need to be published: `docs/`
is the finished site, and `docs/data.js` is tracked, so editing the spreadsheet
and running `build_site.py` still updates the live site on the next push.

## 1. Fill in the details

Open **`groom-details.xlsx`**. Every gold `-- FILL IN --` cell needs your
answer. These came off the card by machine reading and still need checking:

**Wedding Details sheet** — Groom, Bride, city, venue name and
address, Google Maps link, contact phone, WhatsApp number, RSVP deadline.

The parents are four separate fields: `Groom's Father`, `Groom's Mother`,
`Bride's Father`, `Bride's Mother`. The card used to have one field per side
and the name typed into it was a father's, so that name was moved into the
Father field and the Mother fields were left for you. Leave a mother blank if
you would rather not list her — the label is dropped rather than shown empty.

**Wedding Functions sheet** — all 12 functions, their dates and times are filled
in. Confirm each one against the card, then flip **Verified?** from `No` to
`Yes`. The two to look at hardest are `Kalra / Devpuji` on 10 Dec and the two
separate 3:00 PM functions on 11 Dec. Anything in **Note** stays hidden on the
website until you set Verified to `Yes`, so a half-checked row never shows
guests a note you are still unsure about.

Put a venue name in the **Venue** column only for the days that are not at the
main venue, and it renders as a small tag on that row.

**Family Details sheet** — the long formal wording and the message to
relatives are used only by `docs/family.html`. The rest of the sheet is not:
parking, accommodation, its map link, the baraat route, the gift note and the
two named contacts all reach both cards, because they are practical answers
rather than formal ones. `Accommodation Link` is optional: fill it in with a
Google Maps URL and the accommodation block gets an **Open in Maps** button
under the hotel names, so a guest can tap through instead of reading an
address out loud. **Family Contacts sheet** — one row per person, for the
"who do I call" section.

**Family Members sheet** — the "Eagerly Awaiting Your Presence" list on the
family page. One row per relative, as many rows as you want: `Name` (required),
`Relation to the couple`, `Side` (Groom's side / Bride's side / Both, which
groups the page into two lists), and `From / location`. Set **Verified?** to
`Yes` for a row to appear. Until at least one row is verified the whole section
hides itself, so a half-filled sheet never shows a bare heading.

**Wedding Functions sheet**, column H, **Family Detail** — a short note per
function for the family page only, e.g. which side hosts it. Unlike the Note
column this one is not gated on Verified?, because it is written by you for
relatives rather than read off the card.

**RSVP List sheet** — optional, and not the same as the web form. Leave it alone
unless you are tracking responses yourself.

The year is not printed on the card, so `2026` is a placeholder. Change it in
one place, `Wedding Year`, and the function dates follow.

### Publish and Audience

Four of the sheets carry records rather than settings: **Wedding Functions**,
**Family Members**, **Family Contacts** and **RSVP List**. Each one ends with
two dropdown columns.

**Publish** decides whether a row exists on the site at all.

| Value | Meaning |
| --- | --- |
| `Yes` | the row is published |
| `No` | the row is dropped at build time |

`No` is stronger than hiding something: the row is never written into the
generated `data.js`, so it cannot be reached on either page, not even by
viewing the page source. Anything left blank counts as `No`, which means a row
someone has just typed in stays private until they have decided to publish it.

**Audience** decides which of the two pages a published row may appear on.

| Value | Meaning |
| --- | --- |
| `Both` | friends and family (the default) |
| `Friends` | the friends page only |
| `Family` | the family page only |

This is the column to reach for when something is finished but not for
everyone — a family member's name is fine on the family card but not on the one
going out to guests, for instance. A blank or unrecognised value counts as
`Both`, so a typo widens the audience rather than silently hiding a row.

This replaces the old `Verified?` column. That column only ever gated the
Family Members and the guest-facing note on Wedding Functions, and it did two
different jobs depending on the sheet. `Publish` is the same idea applied
consistently, and `Audience` covers the cases it could not express.

## 2. Rebuild the site

```bash
python3 build_site.py
```

That rebuilds the groom's side from `groom-details.xlsx`. Her side is a
separate workbook with a separate command, and it reads nothing from his:

```bash
python3 build_site.py --book bride-details.xlsx --out docs/data-bride.js
```

The two never touch each other's file, so a change in one cannot reach the
other's pages. Run whichever side you changed; neither one overwrites the
other's output.

It prints how many functions and days it found, and lists anything still
unfilled. Then open or refresh the page.

The same script is safe to run every time you change the spreadsheet. If it
prints `Missing groom-details.xlsx`, run `python3 make_workbook.py` first.

**`make_workbook.py` recreates the whole file and overwrites whatever is
there.** It has no `--dry-run` and no `--help`, and it does not merge, so do not
run it against the workbook you have been editing. Everything you have typed
goes. It is only for starting over.

**`add_family_sheets.py` overwrites three sheets** — Family Details, Family
Contacts and Family Members — with the defaults held in that file, and
rewrites the four parent rows in Wedding Details. It exists so a workbook
predating these sheets can be given them, and it is safe to re-run: it will not
overwrite a parent name or a family member you have already typed. To change
what the family page says, type into the workbook and leave that script alone.

## 3. Run it locally

```bash
python3 -m http.server 8347 --directory docs
```

Then visit http://localhost:8347

A server is needed because the site loads `data.js` as a separate file; opening
`index.html` by double-clicking it will not work.

## 4. What the site does

- Hero with the couple's names, date, venue, and a live countdown to the first
  function
- Tabbed schedule, one tab per day, with a line drawing for each function, and the
  running function is highlighted
- Venue cards with Maps links and a copy-address button
- Download the whole schedule as a `.ics` file, so all 12 functions land in the
  guest's phone calendar
- RSVP form. Responses are saved in the browser and can be sent on WhatsApp
- Share button, which uses the native share sheet on phones

### Two pages

`index.html` is the short one, for friends. It assumes the reader either knows
the customs or does not need to: the couple's names, the countdown, the
schedule, the venue, RSVP.

`family.html` is the long one, for relatives, and adds five things:

- **The formal invitation** in printed-card order — blessing, both sets of
  parents, the couple, the request, the dates
- **Eagerly Awaiting Your Presence**, the named relatives from the Family
  Members sheet, grouped into the groom's side and the bride's side
- **Who To Call**, from the Family Contacts sheet

**Practical Details** is not on that list any more. Baraat route, parking,
accommodation and the gift note moved inside the **Venue & Directions**
section, so they are on the friends card too. Parking and the hotel answer the
same "where do I go" question as the venue, and a guest holding the short
friends link needs them as much as a relative does. `Accommodation Link` adds
an **Open in Maps** button under the hotel names.

Each of those blocks hides itself when its data is blank, so a half-filled
workbook produces a shorter page rather than a page of empty headings.

They are two URLs, not one page with a toggle: you send family
`.../family.html` and friends `.../`, and there is no control on the page that a
guest can press to reach the other one. `app.js` is shared between them
unchanged; only `family.html` loads `family.js`.

There is no **Edit details** button, and that is deliberate. The page is opened by
guests, so anything on it has to be read-only: an in-page editor would let any
visitor rewrite the names, the venue, or the contact number, and its *Download
JSON* button would hand them a file shaped exactly like `data.js`. For the same
reason `app.js` reads `data.js` directly and ignores anything in
`localStorage`.

To change a detail, edit `groom-details.xlsx`, run `python3 build_site.py`, and
reload. That is the only path, and it stays on your machine.

## 5. The look

The page is soft pastel romance on a warm ivory base: sage and olive for the
structure, blush for the names, hairline rules instead of heavy borders, and
Cormorant Garamond over Jost. The sage is lifted from
[theweddingwebsite.in](https://theweddingwebsite.in), whose homepage runs on
`#d0dfb9` and `#4e5e32`; the rest is original.

There are no photographs on the page. If you ever want one, drop a file into
`docs/assets/` and add an `<img>` inside `.hero__card` in `index.html`, above
the blessing line.

Everything else visual is original vector art drawn by `make_art.py`: a Hawa
Mahal jali, a toran of marigolds and mango leaves, a kalash, a mandap, peacocks,
diyas, and one line drawing per function.

```bash
python3 make_art.py    # rewrites the 24 SVGs in docs/assets/
python3 check_art.py   # verifies paths, viewBoxes and framing
```

`check_art.py` should report `0 problem(s)`. It catches the mistakes that are easy
to make when drawing SVG by hand: a path command with the wrong number of
arguments, a curve that escapes its viewBox, an icon that is not centred, and so
on. Run it after any change to `make_art.py`.

Two things worth knowing if you edit the drawings:

- **The function icons are stroked with `currentColor`.** They are inlined into
  `data.js` by `build_site.py` rather than loaded with `<img>`, so they inherit
  the theme colour. As an `<img>`, a `currentColor` icon renders black.
- **An icon's viewBox is fitted to its own drawing.** That is why the numbers are
  odd, like `viewBox="4.97 9.00 54.02 54.02"`. It is deliberate, and it is what
  keeps all twelve icons the same visual weight.

### Recolouring it

Every colour is a custom property at the top of `styles.css`. Change `--sage-700`
and `--rose` and the whole page follows, buttons, tabs and focus rings included.
Keep body text at 4.5:1 or better against the ivory base; `--ink-mute` is the
one that fails most easily, because it is used for the small letterspaced labels.

## 6. Put it online, free

The `docs/` folder is a static site, so any static host works. All three below
are free and take about five minutes.

### Cloudflare Pages (easiest, no card needed)

1. Push this folder to a new GitHub repo.
2. Sign in at dash.cloudflare.com, then **Workers & Pages → Create → Pages**.
3. Connect the GitHub repo. Build command: leave empty. Output directory: `docs`.
4. Deploy. You get a `*.pages.dev` address; add a custom domain later if you
   want one.

### Netlify (best drag-and-drop)

1. Sign in at netlify.com.
2. Drag the `docs` folder onto the deploys page.
3. The site is live immediately. Later, connect the repo and set the publish
   directory to `docs` so re-pushes go live on their own.

### GitHub Pages

```bash
cd docs
git init && git add . && git commit -m "wedding invitation"
git branch -M main
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```

Then in the repo, **Settings → Pages**, source **Deploy from a branch**, branch
`main`, folder `/ (root)`. The address becomes
`https://<you>.github.io/<repo>/`.

### A custom name

Worth it for a wedding. On any of the three hosts, **Custom domains** lets you
point a domain you already own. Names like `abhaywedsanjana.com` are widely
available. Budget roughly $10-15 for a `.com` for one year.

## 7. Sharing it

The **Share** button on the page uses WhatsApp, or the native share sheet on a
phone. To put a link in the printed card or a WhatsApp broadcast, just share the
site's address.

You have two links to hand out:

- `https://surajsingh81.github.io/Weddinginvite/` — friends
- `https://surajsingh81.github.io/Weddinginvite/family.html` — relatives

Neither is secret. Anyone with the URL can read it, so `family.html` is only
more detailed, not private. It does not carry anything you would not want a
guest to see, and there is no password on it.

## A note on the details

The printed card is written in a decorative script that OCR cannot read
reliably, so names, parents, venue, and address came off the image by eye and
are the values to check most carefully. The 12 function rows came off the
printed table and are more trustworthy, but still worth a look.

How to run
Edit groom-details.xlsx (or bride-details.xlsx for her side).
Run python3 build_site.py
  python3 build_site.py --book bride-details.xlsx --out docs/data-bride.js
git add -A && git commit -m "…" && git push origin main.
GitHub Pages rebuilds; the live site now shows your changes.

