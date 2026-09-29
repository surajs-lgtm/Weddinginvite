#!/usr/bin/env python3
"""
The family-only sheets, as data.

This module holds the literal contents of the three family sheets plus the
Family Detail column, and one function that writes them into an open workbook.
Two things import it:

  - add_family_sheets.py_add_family (this file's __main__) adds the sheets to
    the workbook you already have
  - make_workbook.py calls write_family_sheets() when building from scratch

So the family content exists in exactly one place. Editing a ritual's wording
means editing it here and re-running add_family_sheets.py, and a workbook
rebuilt from scratch gets the same text.

Run:  python3 add_family_sheets.py
"""
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
BOOK = os.path.join(HERE, "wedding-details.xlsx")

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
    ("Family Contact 1 Name", FILLIN, "Who a relative should call first"),
    ("Family Contact 1 Phone", FILLIN, "With country code"),
    ("Family Contact 1 Role", FILLIN, "e.g. Groom's uncle, coordinating on the day"),
    ("Family Contact 2 Name", FILLIN, "Second person, e.g. for the ladies' functions"),
    ("Family Contact 2 Phone", FILLIN, "With country code"),
    ("Family Contact 2 Role", FILLIN, ""),
    ("Gift / Shagun Note", FILLIN, "Optional. Left blank, nothing appears."),
]

# ------------------------------------------------------------------- Rituals
# Every line here is researched but not all of it confirmed, so each row has a
# Verified? column that gates it, exactly like Wedding Functions. A row stays
# hidden until that is Yes, so a wrong guess never reaches a relative.
RITUAL_ROWS = [
    ("Function", "What it is", "Why it happens", "Family note", "Verified?"),
    ("Matkor",
     "The groom digs and lifts a handful of clean earth from the ground.",
     "The wedding is built on that earth. The mandap, the fire and the meal are all understood to stand on it.",
     "Traditionally done by the groom alone, early, before anyone else arrives.", "No"),
    ("Madwa",
     "A temporary canopy of bamboo is raised and dressed with leaves and cloth.",
     "It is the roof over the wedding itself: the mandap stands under it for four days and comes down after.",
     "Also called chaura or mandwa. The canopy is raised before the first function.", "Yes"),
    ("Haldi Kalsa",
     "Turmeric paste is applied to the couple, and a pool of haldi is kept for the guests to join in.",
     "Turmeric is auspicious and purifying; the yellow is said to soften the eye and set the marriage on a bright note.",
     "Wear old clothes. The paste is deliberately messy, and guests are meant to get some on them.", "Yes"),
    ("Devpuji",
     "The family deity is worshipped before the ceremony, with the pandit leading.",
     "No auspicious beginning is made without asking, so the gods are invited first.",
     "Locally also called Kalra. Confirm with your pandit whether it is one function or two.", "No"),
    ("Ghee Dhari",
     FILLIN, FILLIN, FILLIN, "No"),
    ("Janeu",
     "The sacred thread is tied on the groom, in the presence of the family.",
     "It marks him as a householder carrying the line of his ancestors, and is a rite of passage in its own right.",
     "Held on the groom's father's hand, from his paternal grandmother's generation.", "Yes"),
    ("Lawa Bhujai",
     FILLIN, FILLIN, FILLIN, "No"),
    ("Baraat Prasthan",
     "The groom rides out with music and a decorated vehicle, and is brought to the mandap.",
     "The procession announces the wedding to the neighbourhood, and is the last time the groom arrives as an unmarried man.",
     "The main function. This is the one the countdown runs to.", "Yes"),
    ("Tilak",
     "The groom's forehead is marked with a tilak as he joins the bride's family.",
     "It is the groom being welcomed and accepted, and is the point the two families formally receive each other.",
     "Tilak is applied by the bride's side.", "Yes"),
    ("Jaimala",
     "The couple exchange garlands.",
     "The garland is a test: each lifts it and tries to place it over the other's neck, and the crowd cheers or heckles.",
     "An open invitation to the baraat to play along.", "Yes"),
    ("Sindoor Daan",
     "The bride applies vermillion in the parting of the groom's hair.",
     "It is the point the marriage is sealed, and from it the couple begin life together.",
     "Traditionally the first sindoor of the marriage is applied by the bride alone.", "Yes"),
    ("Vidai",
     "The farewell: the bride leaves her natal home with the groom's party.",
     "The name means 'leaving'. It is the emotional close of four days of functions and the start of the new home.",
     "Traditionally the couple are given a bowl of rice and dal, to eat as they leave.", "No"),
]

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
    for name in ("Family Details", "Rituals", "Family Contacts"):
        if name in wb.sheetnames:
            del wb[name]

    ws = wb.create_sheet("Family Details")
    for r in FAMILY_ROWS:
        ws.append(list(r))
    style_table(ws, {"A": 26, "B": 52, "C": 46})
    for r in range(2, len(FAMILY_ROWS) + 1):
        ws.cell(row=r, column=1).font = title_font
    mark_field_value(ws, 2, 2, len(FAMILY_ROWS))

    ws = wb.create_sheet("Rituals")
    for r in RITUAL_ROWS:
        ws.append(list(r))
    style_table(ws, {"A": 20, "B": 46, "C": 52, "D": 44, "E": 12})
    for r in range(2, len(RITUAL_ROWS) + 1):
        ws.cell(row=r, column=1).font = title_font
        ws.cell(row=r, column=5).alignment = center
        for col in (2, 3, 4):
            v = ws.cell(row=r, column=col).value
            if isinstance(v, str) and v.strip() == FILLIN:
                ws.cell(row=r, column=col).font = fill_font
                ws.cell(row=r, column=col).fill = fill_fill
    dv = DataValidation(type="list", formula1='"No,Yes"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("E2:E200")

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

    write_function_family_column(wb)

    if help_sheet and "How To Use" in wb.sheetnames:
        how = wb["How To Use"]
        how.append(["", ""])
        how.append(("6. THE FAMILY PAGE", ""))
        how.append(("docs/family.html is a longer version for relatives: what each ritual means, who to call, parking and baraat.", ""))
        how.append(("It reads the Family Details, Rituals and Family Contacts sheets, plus the Family Detail column.", ""))
        how.append(("Rituals stay hidden until Verified? is Yes, so an unconfirmed line is never shown.", ""))

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
        raise SystemExit("Missing wedding-details.xlsx")

    wb = write_family_sheets(load_workbook(BOOK))
    wb.save(BOOK)
    print("added Family Details, Rituals, Family Contacts")
    print("added the Family Detail column to Wedding Functions")
    print("appended a family-page note to How To Use")
    print()
    print("This overwrote those three sheets with the defaults above.")
    print("Anything you had typed into them is gone. To edit the family")
    print("content, edit FAMILY_ROWS / RITUAL_ROWS in this file instead,")
    print("or just type into the workbook and leave this script alone.")


if __name__ == "__main__":
    main()
