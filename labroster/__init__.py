"""Lab Shift Roster — a secure and intelligent workforce planning tool
for diagnostic laboratories.

The package is deliberately made of small, independently testable modules, but
kept flat and dependency-light so the whole thing can be loaded into Pyodide and
run in a browser tab with no build step.

Module map (modules are added as the migration proceeds):

    timeutils   shift durations, midnight crossing, rest intervals, week patterns
"""

__version__ = "1.1.0-beta"
"""Single source of truth for the version.  The interface, the exported workbooks
and the command line all read this rather than carrying their own copy."""

VERSION_LABEL = "1.1 Beta"
PRODUCT_NAME = "Lab Shift Roster"
COMPANY_NAME = "Optymum SS"
CONTACT_EMAIL = "projects@optymumss.com"
TAGLINE = "A secure and intelligent workforce planning tool for diagnostic laboratories."
TAGLINE_FREE_PLAN = ("A free, secure and intelligent workforce planning tool for "
                     "diagnostic laboratories.")


def export_tagline(paid: bool = False) -> str:
    """The tagline written into an exported rota.

    A roster made on a paid plan (Pro or Enterprise) does not call the tool free;
    one made on the free plan keeps saying so.
    """
    return TAGLINE if paid else TAGLINE_FREE_PLAN

SCHEMA_VERSION = 3
"""Workbook schema version.  Written into the Settings sheet so an older
workbook can be recognised and explained rather than crashing the app.

3 added the Shifts sheet's Active column and the Staff sheet's Initials column.
Both are optional and a version 2 workbook still reads correctly, so the bump
exists to tell somebody their template predates two useful options rather than to
reject their file."""
