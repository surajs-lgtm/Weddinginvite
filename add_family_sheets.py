#!/usr/bin/env python3
"""
The family-only sheets, as data.

This module holds the literal contents of the three family sheets plus the
Family Detail column, and one function that writes them into an open workbook.
Two things import it:

  - add_family_sheets.py_add_family (this file's __main__) adds the sheets to
    the workbook you already have
  - make_workbook.py calls write_family_sheets() when building from scratch

So the family content exists in exactly one place. Editing a contact's wording
means editing it here and re-running add_family_sheets.py, and a workbook
rebuilt from scratch gets the same text.

Run:  python3 add_family_sheets.py
"""
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
BOOK = os.path.join(HERE, "groom-details.xlsx")

FILLIN = "-- FILL IN --"
MAROON = "7B1E2B"
CREAM = "FDF6E9"

head_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
head_fill = PatternFill("solid", fgColor=MAROON)
title_font = Font(name="Calibri", size=11, bold=True, color=MAROON)
fill_font = Font(name="Calibri", size=11, color="B8860B", bold=True)
fill_fill = PatternFill("solid", fgColor=CREAM)
set_font = Font(name="Calibri", size=11, bold=True, color="1F6F43")
set_fill = PatternFill("solid", fgColor="EAF6EE")
thin = Side(style="thin", color="D0D0D0")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center")


def style_table(ws, widths):
    for c in ws[1]:
        c.font = head_font
        c.fill = head_fill
        c.border = border
        c.alignment = center
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = border
            c.alignment = wrap
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def mark_field_value(ws, col, first_row, last_row):
    """Gold when still a placeholder, green once typed in. Mirrors sheet 1."""
    for r in range(first_row, last_row + 1):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, str) and v.strip() == FILLIN:
            ws.cell(row=r, column=col).font = fill_font
            ws.cell(row=r, column=col).fill = fill_fill
        elif v:
            ws.cell(row=r, column=col).font = set_font
            ws.cell(row=r, column=col).fill = set_fill


# ------------------------------------------------------------ Parent fields
# The card carries both parents' names, but the workbook had one field per side
# and the name typed into it ("Randhir Prasad Singh") is a father's. So the two
# old fields are split into four. The old name moves to the Father field rather
# than being copied into both, because putting a father's name in as the mother
# would be worse than leaving the mother blank.
PARENT_FIELDS = [
    ("Groom's Father", FILLIN, "Father of the groom. Goes first on the card."),
    ("Groom's Mother", FILLIN, "Mother of the groom. Leave blank if you would rather not list her."),
    ("Bride's Father", FILLIN, "Father of the bride"),
    ("Bride's Mother", FILLIN, "Mother of the bride"),
]

# The old single-parent labels, in the order they appear on the sheet. Kept as a
# list rather than a dict because the split writes two rows per old label; see
# migrate_parent_fields, which pairs them with PARENT_FIELDS explicitly.
PARENT_MIGRATION = [
    ("Groom's Parents", "Groom's Father", "Groom's Mother"),
    ("Bride's Parents", "Bride's Father", "Bride's Mother"),
]


def migrate_parent_fields(wb):
    """Split "Groom's Parents" into a father and a mother field.

    Moves the existing value into the Father field and clears the old one, so
    the name is never printed in both places. Both new rows are written
    explicitly rather than derived from PARENT_MIGRATION alone, because a
    migration map naming only the father target silently produced a Father row
    and no Mother row.
    """
    if "Wedding Details" not in wb.sheetnames:
        return
    ws = wb["Wedding Details"]

    def row_for(label):
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value or "").strip() == label:
                return r
        return None

    # old label -> (father label, mother label)
    splits = PARENT_MIGRATION
    details = {lbl: (val, nte) for lbl, val, nte in PARENT_FIELDS}

    for old_label, father_label, mother_label in splits:
        old_row = row_for(old_label)
        if old_row is None:
            continue

        old_value = ws.cell(row=old_row, column=2).value
        moving = None if is_placeholder(old_value) else old_value

        # Once a side has been split, the old row stays on the sheet with an
        # empty value as a tombstone. Without this guard the second run sees
        # that tombstone, decides there is nothing to move, and overwrites the
        # father with the placeholder — silently deleting the name.
        already_split = all(
            row_for(label) is not None for label in (father_label, mother_label)
        )
        if already_split and moving is None:
            continue

        # Father first, then Mother directly under it, so the sheet keeps
        # reading groom then bride.
        for offset, label in ((0, father_label), (1, mother_label)):
            if row_for(label) is None:
                ws.insert_rows(old_row + 1 + offset)
                r = old_row + 1 + offset
                ws.cell(row=r, column=1).value = label
                ws.cell(row=r, column=1).font = title_font
                ws.cell(row=r, column=1).border = border
                ws.cell(row=r, column=1).alignment = wrap

            r = row_for(label)
            # the father inherits the old name, the mother starts empty
            value = moving if label == father_label else FILLIN
            cell = ws.cell(row=r, column=2)
            cell.value = value
            cell.border = border
            cell.alignment = wrap
            if is_placeholder(value):
                cell.font = fill_font
                cell.fill = fill_fill
            else:
                cell.font = set_font
                cell.fill = set_fill

            _, notes = details.get(label, ("", ""))
            ncell = ws.cell(row=r, column=3)
            ncell.value = notes
            ncell.border = border
            ncell.alignment = wrap
            ncell.font = Font(name="Calibri", size=11, color="7A756C", italic=True)

        # clear the old field so the name is not printed in both places
        old_value_cell = ws.cell(row=old_row, column=2)
        old_value_cell.value = ""
        old_value_cell.font = Font(name="Calibri", size=11, color="7A756C", italic=True)
        old_notes_cell = ws.cell(row=old_row, column=3)
        old_notes_cell.value = (
            "Split into '%s' and '%s'. The name above moved." % (father_label, mother_label)
        )
        old_notes_cell.font = Font(name="Calibri", size=11, color="7A756C", italic=True)


def is_placeholder(value):
    return value is None or str(value).strip() in ("", FILLIN)


# ------------------------------------------------------------ Family Details
FAMILY_ROWS = [
    ("Field", "Value", "Notes"),
    ("Family Greeting", "With the blessings of our families", "Overrides the short line on the family page only"),
    ("Family Invitation Line",
     "together with their respective families, request the honour of your presence",
     "The long formal wording, e.g. 'Shri ... and Smt ... beg to invite you'"),
    ("Family Note", FILLIN, "A message to relatives. Leave blank for none."),
    ("Baraat Route", FILLIN, "Where the baraat gathers and where it goes. Leave blank for none."),
    ("Parking", FILLIN, "Where relatives should park. Leave blank for none."),
    ("Accommodation", FILLIN, "Nearby hotels for out-of-town relatives. Leave blank for none."),
    ("Accommodation Link", FILLIN, "Google Maps link for the hotel. Leave blank for none."),
    ("Family Contact 1 Name", FILLIN, "Who a relative should call first"),
    ("Family Contact 1 Phone", FILLIN, "With country code"),
    ("Family Contact 1 Role", FILLIN, "e.g. Groom's uncle, coordinating on the day"),
    ("Family Contact 2 Name", FILLIN, "Second person, e.g. for the ladies' functions"),
    ("Family Contact 2 Phone", FILLIN, "With country code"),
    ("Family Contact 2 Role", FILLIN, ""),
    ("Gift / Shagun Note", FILLIN, "Optional. Left blank, nothing appears."),
    ("Awaiting Section Title", "Eagerly Awaiting Your Presence",
     "Heading above the list of family members. Leave blank to use the default."),
    ("Awaiting Section Note", FILLIN, "One line under the heading, e.g. 'From both our families'. Optional."),
]

# ----------------------------------------------------------- Family Members
# The "Eagerly Awaiting Your Presence" list. One row per person, name required,
# and as many rows as you want — this is the sheet to keep adding to. Grouped on
# the page by Side, so the two families read as two lists rather than one long
# column of names.
MEMBER_ROWS = [
    ("Name", "Relation to the couple", "Side", "From / location", "Verified?"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "No"),
]

# Side is a closed list because the page groups on it; DataValidation is applied
# to the column in write_family_sheets. Anything unrecognised falls back to a
# single untitled group rather than disappearing.
MEMBER_SIDES = ("Groom's side", "Bride's side", "Both")
# ------------------------------------------------------------ Family Contacts
CONTACT_ROWS = [
    ("Name", "Relation to the couple", "Phone", "Reaches them about", "Notes"),
    (FILLIN, FILLIN, FILLIN, FILLIN, "One row per person family may need to reach."),
    (FILLIN, FILLIN, FILLIN, FILLIN, ""),
    (FILLIN, FILLIN, FILLIN, FILLIN, ""),
    (FILLIN, FILLIN, FILLIN, FILLIN, ""),
    (FILLIN, FILLIN, FILLIN, FILLIN, ""),
    (FILLIN, FILLIN, FILLIN, FILLIN, ""),
]

# Per-function family notes, keyed by the Event name already in the sheet.
FUNCTION_NOTES = [
    ("Matkor", "Groom's side. Traditionally very early."),
    ("Madwa", "The canopy is raised and stays up for all four days."),
    ("Haldi Kalsa", "Groom's side. Come in old clothes."),
    ("Kalra / Devpuji", "Confirm with the pandit whether this is one function or two."),
    ("Dhid Hari / Ghee Dhari", FILLIN),
    ("Janau", "Groom's side, midday."),
    ("Bhunga Lawa", FILLIN),
    ("Baraat Prasthan", "THE main function. Everyone is welcome."),
    ("Tilak", "Groom is welcomed by the bride's family."),
    ("Jaimala", "Part of the baraat, immediately after."),
    ("Sindoor Daan", FILLIN),
    ("Vidai (Milap)", "The farewell. Traditional for close family."),
]


def write_family_sheets(wb, help_sheet=True):
    """Write the three family sheets and the Family Detail column into `wb`.

    Takes and returns a Workbook, and does not save it, so the caller decides
    where it goes. make_workbook.py calls this on a fresh workbook; running
    this file as a script calls it on the one already on disk.

    Existing sheets of the same name are dropped first, so running this twice
    in a row is a no-op rather than a duplicate-sheet error.
    """
    for name in ("Family Details", "Family Contacts", "Family Members"):
        if name in wb.sheetnames:
            del wb[name]

    ws = wb.create_sheet("Family Details")
    for r in FAMILY_ROWS:
        ws.append(list(r))
    style_table(ws, {"A": 26, "B": 52, "C": 46})
    for r in range(2, len(FAMILY_ROWS) + 1):
        ws.cell(row=r, column=1).font = title_font
    mark_field_value(ws, 2, 2, len(FAMILY_ROWS))

    ws = wb.create_sheet("Family Contacts")
    for r in CONTACT_ROWS:
        ws.append(list(r))
    style_table(ws, {"A": 20, "B": 28, "C": 18, "D": 40, "E": 40})
    for r in range(2, len(CONTACT_ROWS) + 1):
        for col in (1, 2, 3):
            v = ws.cell(row=r, column=col).value
            if isinstance(v, str) and v.strip() == FILLIN:
                ws.cell(row=r, column=col).font = fill_font
                ws.cell(row=r, column=col).fill = fill_fill

    ws = wb.create_sheet("Family Members")
    for r in MEMBER_ROWS:
        ws.append(list(r))
    style_table(ws, {"A": 24, "B": 30, "C": 16, "D": 26, "E": 12})
    for r in range(2, len(MEMBER_ROWS) + 1):
        ws.cell(row=r, column=1).font = title_font
        ws.cell(row=r, column=5).alignment = center
        for col in (1, 2, 3, 4):
            v = ws.cell(row=r, column=col).value
            if isinstance(v, str) and v.strip() == FILLIN:
                ws.cell(row=r, column=col).font = fill_font
                ws.cell(row=r, column=col).fill = fill_fill
    dv_side = DataValidation(
        type="list", formula1='"%s"' % ",".join(MEMBER_SIDES), allow_blank=True
    )
    ws.add_data_validation(dv_side)
    dv_side.add("C2:C300")
    dv_ok = DataValidation(type="list", formula1='"No,Yes"', allow_blank=True)
    ws.add_data_validation(dv_ok)
    dv_ok.add("E2:E300")

    migrate_parent_fields(wb)
    write_function_family_column(wb)

    if help_sheet and "How To Use" in wb.sheetnames:
        how = wb["How To Use"]
        existing = {str(c.value).strip() for row in how.iter_rows() for c in row if c.value}
        if "6. THE FAMILY PAGE" not in existing:
            how.append(["", ""])
            how.append(("6. THE FAMILY PAGE", ""))
            how.append(("docs/family.html is a longer version for relatives: who to call, parking and baraat.", ""))
            how.append(("It reads the Family Details and Family Contacts sheets, plus the Family Detail column.", ""))
            how.append(("Every record has a Publish column: No keeps it off the site entirely.", ""))


    return wb


def write_function_family_column(wb):
    """Add column H, 'Family Detail', to Wedding Functions, matched by name.

    Separated from write_family_sheets because the function names here have to
    agree with the Wedding Functions sheet, and that sheet is the one a person
    is most likely to have edited.
    """
    funcs = wb["Wedding Functions"]
    header = funcs.cell(row=1, column=8)
    header.value = "Family Detail"
    header.font = head_font
    header.fill = head_fill
    header.border = border
    header.alignment = center
    funcs.column_dimensions["H"].width = 44

    by_name = dict(FUNCTION_NOTES)
    for r in range(2, funcs.max_row + 1):
        name = funcs.cell(row=r, column=3).value
        note = by_name.get(str(name).strip() if name else "", "")
        cell = funcs.cell(row=r, column=8)
        cell.value = note if note else ""
        cell.border = border
        cell.alignment = wrap
        if note == FILLIN:
            cell.font = fill_font
            cell.fill = fill_fill


def main():
    if not os.path.exists(BOOK):
        raise SystemExit("Missing groom-details.xlsx")

    wb = write_family_sheets(load_workbook(BOOK))
    wb.save(BOOK)
    print("added Family Details, Family Contacts")
    print("added the Family Detail column to Wedding Functions")
    print("appended a family-page note to How To Use")
    print()
    print("This overwrote those three sheets with the defaults above.")
    print("Anything you had typed into them is gone. To edit the family")
    print("content, edit FAMILY_ROWS / CONTACT_ROWS / MEMBER_ROWS in this file instead,",1)
    print("or just type into the workbook and leave this script alone.")


if __name__ == "__main__":
    main()
