#!/usr/bin/env python3
"""
Make the bride's copy of the workbook.

The wedding is one wedding: the dates, the venue, the twelve functions and
their explanations are all shared, and neither family should be asked to
retype them or be trusted to type them the same way. So this does not build a
blank file. It copies the real workbook and then does exactly three things to
the bride's half of it:

  1. Sets "Card Flank" to Bride, which is the only thing that makes her card
     different. That cell is what puts her name first, her parents above the
     names, and her relatives first in the list.
  2. Empties her parents' names, so she writes them herself.
  3. Empties the bride's-side rows in Family Members, so she writes her own
     relatives. The groom's-side rows are left alone and stay as a reminder of
     the format to follow.

Nothing here is destructive to the groom's workbook: it reads wedding-details
and writes a separate file. Run it again after the groom's family adds more
details and the bride's copy picks those up too, still with her half blank.

Run:  python3 make_bride_workbook.py
      python3 make_bride_workbook.py --out bride-details.xlsx
"""
import argparse
import os
import shutil
import sys
from copy import copy

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font
except ImportError:
    sys.exit("openpyxl is required:  pip3 install openpyxl")

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "wedding-details.xlsx")

NOTE = ("The bride's copy. Fill in your parents' names and your own relatives, "
        "then send this file back. The dates, venue and functions are already "
        "filled in and are the same for both families -- please leave them "
        "alone, and tell us if something is wrong rather than changing it.")

# Rows emptied on the bride's side. Everything shared stays exactly as it is.
BLANK_DETAILS = ("Bride's Father", "Bride's Mother")

# The placeholder the generated sheet uses for a row nobody has filled in yet.
FILL_MARK = "-- FILL IN --"

# How many blank rows to leave her below the groom's list.
BLANK_MEMBER_ROWS = 18


def blank_cell(ws, row, column):
    """Clear a value and leave the cell looking like the ones around it."""
    src = ws.cell(row=row, column=1)
    dst = ws.cell(row=row, column=column)
    dst.value = ""
    dst.font = copy(src.font)
    dst.alignment = copy(src.alignment)
    dst.border = copy(src.border)


def set_value(ws, row, column, value):
    src = ws.cell(row=row, column=1)
    dst = ws.cell(row=row, column=column)
    dst.value = value
    dst.font = copy(src.font)
    dst.alignment = copy(src.alignment)
    dst.border = copy(src.border)


def find_row(ws, label, column=1):
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(row=r, column=column).value or "").strip() == label:
            return r
    return None


def blank_bride_members(wb):
    """Clear the bride's-side relatives, keep the rows so the format is visible.

    The rows are emptied rather than deleted, so she can see a groom's-side row
    directly above the first one she fills in. Anything set to Both stays, since
    a relative who is blood to both families belongs on the card either way.

    The empty rows are pre-labelled "Bride's side" rather than left as the
    generic placeholder: the Side cell has a dropdown, but the only value she
    should ever pick on those rows is her own, and making her choose it every
    time is the easiest way to get a list that quietly drops half the family.
    """
    if "Family Members" not in wb.sheetnames:
        return 0
    ws = wb["Family Members"]
    cleared = 0

    # Collect first, delete after. Removing a row while walking downward shifts
    # every later row up by one, and max_row moves under the loop, so the
    # placeholder row survived and the new block started one row too low.
    drop = []
    bride_rows = []
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(row=r, column=1).value or "").strip()
        side = str(ws.cell(row=r, column=3).value or "").strip().lower()
        if name.startswith(FILL_MARK):
            drop.append(r)
        elif side.startswith("bride"):
            bride_rows.append(r)
    for r in reversed(drop):
        ws.delete_rows(r, 1)

    for r in bride_rows:
        for c in range(1, 6):
            blank_cell(ws, r, c)
        set_value(ws, r, 3, "Bride's side")
        cleared += 1

    # A fresh block of empty rows for her, each already labelled as her side.
    # The Name cell carries the fill-in marker rather than sitting blank: a
    # blank cell in a column of blank cells is easy to scroll straight past,
    # and the instructions on the How To Use sheet point at this marker.
    start = ws.max_row + 1
    for i in range(BLANK_MEMBER_ROWS):
        r = start + i
        set_value(ws, r, 1, FILL_MARK)
        set_value(ws, r, 2, FILL_MARK)
        set_value(ws, r, 3, "Bride's side")
        set_value(ws, r, 4, "")
        set_value(ws, r, 5, "No")
    return cleared


def rewrite_instructions(wb):
    """Replace the 'How To Use' sheet with instructions for the bride.

    The sheet in the original is written for whoever maintains the site, and
    it says to run `python3 build_site.py`. That is not what the bride is being
    asked to do, and leaving it there is the kind of thing that reads as though
    the file were broken. She is being asked to fill in names and send the file
    back, so the sheet says exactly that.
    """
    if "How To Use" not in wb.sheetnames:
        return

    ws = wb["How To Use"]
    # Clear the sheet rather than appending: the original instructions are
    # about the build, and they would still be sitting further down otherwise.
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None

    lines = [
        ("YOUR COPY OF THE WEDDING DETAILS", ""),
        ("", ""),
        ("Please fill in the blanks marked", FILL_MARK),
        ("", ""),
        ("What to fill in", ""),
        ("1. Wedding Details sheet:", "your father's and mother's names."),
        ("2. Family Members sheet:", "your relatives, one per row."),
        ("   Set Side to", "'Bride's side' for each of them."),
        ("", ""),
        ("What is already filled in", ""),
        ("The dates, times, venue and the twelve functions are the same for",
         "both families and are already done."),
        ("Please leave them as they are.", "If something looks wrong, tell us"),
        ("", "rather than changing it."),
        ("", ""),
        ("When you are done", ""),
        ("Save the file and send it back to us.", "That is all."),
        ("", ""),
        ("One more thing", ""),
        ("Your name comes first on your version of the card, and your parents'",
         "names come first as well."),
        ("If you would rather the names stayed as they were on ours, change",
         "'Card Flank' on the Wedding Details sheet back to 'Groom'."),
    ]

    for i, (left, right) in enumerate(lines, start=1):
        if left:
            ws.cell(row=i, column=1, value=left)
        if right:
            ws.cell(row=i, column=2, value=right)

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 62
    ws.sheet_view.showGridLines = False

    # The first line is the title; make it stand out without a colour scheme
    # that might not survive being opened somewhere else.
    title = ws.cell(row=1, column=1)
    title.font = Font(bold=True, size=14)


def make(source=SOURCE, out=None):
    out = out or os.path.join(HERE, "bride-details.xlsx")
    if not os.path.exists(source):
        sys.exit(f"Missing {source}\nRun:  python3 make_workbook.py  (once), then edit it.")

    shutil.copyfile(source, out)
    wb = load_workbook(out)

    if "Wedding Details" not in wb.sheetnames:
        sys.exit("That workbook has no 'Wedding Details' sheet; it is not the wedding workbook.")

    ws = wb["Wedding Details"]

    row = find_row(ws, "Card Flank")
    if row is None:
        # An older copy saved before the field existed. Adding it at the bottom
        # rather than failing, so the couple are never blocked from making the
        # bride's card.
        row = ws.max_row + 1
        for c, value in ((1, "Card Flank"), (2, "Bride"),
                         (3, "Groom | Bride - whose name comes first on the card")):
            set_value(ws, row, c, value)
    else:
        set_value(ws, row, 2, "Bride")

    cleared_parents = 0
    for label in BLANK_DETAILS:
        r = find_row(ws, label)
        if r is not None:
            # The marker, not an empty cell, for the same reason as the
            # relatives below: build_site.py treats the marker as blank, so it
            # never reaches the site, but it is visible in the spreadsheet.
            set_value(ws, r, 2, FILL_MARK)
            cleared_parents += 1

    # The couple line and the hashtag read "Suraj weds Priyanka" and
    # "#SurajwedsPriyanka" in the groom's copy. They are the same two names in
    # the other order on her card, so they are rewritten here rather than left
    # as the one place her version still said his first. She can change either
    # in the spreadsheet afterwards; they are ordinary fields.
    bride_first = str(ws.cell(row=find_row(ws, "Bride Name") or 1, column=2).value or "").strip()
    groom_first = str(ws.cell(row=find_row(ws, "Groom Name") or 1, column=2).value or "").strip()
    if bride_first and groom_first:
        r = find_row(ws, "Couple Line (short)")
        if r is not None:
            set_value(ws, r, 2, "%s weds %s" % (bride_first, groom_first))
        r = find_row(ws, "Hashtag")
        if r is not None:
            set_value(ws, r, 2, "#%sWeds%s" % (bride_first, groom_first))

    cleared_members = blank_bride_members(wb)
    rewrite_instructions(wb)

    wb.save(out)
    print(f"wrote {out}")
    print(f"  Card Flank set to Bride")
    print(f"  cleared {cleared_parents} parent fields for her to fill in")
    print(f"  cleared {cleared_members} bride's-side relative rows")
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Make the bride's copy of the workbook.")
    ap.add_argument("--source", default=SOURCE)
    ap.add_argument("--out", default=None)
    args = ap.parse_args()
    make(args.source, args.out)
