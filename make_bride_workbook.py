#!/usr/bin/env python3
"""
Retired. Do not run this.

This used to build bride-details.xlsx by copying groom-details.xlsx and then
clearing her half. That was right when the two sides shared one set of dates
and one list of twelve functions, so there was nothing to retype and nothing
to disagree about.

That is no longer the case. The two workbooks are now independent by decision:
they hold their own schedules, their own relatives and their own contacts, and
neither is derived from the other. The old script cannot survive that, because
the thing it was built to do -- copy his functions into her file -- is now
exactly the wrong thing:

  * it would replace her thirteen functions with his twelve, dropping the ones
    that only her side has and silently rescheduling the rest;
  * it would overwrite her contacts, which are hers and were typed by hand;
  * it would reset the Rituals sheet, including any row she has written;
  * and it still spoke of a "Verified?" column that no longer exists.

Running it would lose work, so it refuses to do anything at all. The file is
kept rather than deleted so that the old command fails with an explanation
instead of a stack trace or, worse, appearing to work.

What to do instead
------------------
Edit bride-details.xlsx directly. It is already a real, independent workbook,
with her schedule and her contacts in it.

Then rebuild her side:

    python3 build_site.py --book bride-details.xlsx --out docs/data-bride.js

And the groom's, which is unaffected by anything done in her file:

    python3 build_site.py
"""
import sys

MESSAGE = __doc__.strip()


def main():
    sys.stderr.write(MESSAGE + "\n")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())