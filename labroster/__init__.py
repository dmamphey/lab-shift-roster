"""LabRoster — competency-aware workforce planning for diagnostic laboratories.

The package is deliberately made of small, independently testable modules, but
kept flat and dependency-light so the whole thing can be loaded into Pyodide and
run in a browser tab with no build step.

Module map (modules are added as the migration proceeds):

    timeutils   shift durations, midnight crossing, rest intervals, week patterns
"""

__version__ = "2.0.0-dev"

SCHEMA_VERSION = 2
"""Workbook schema version.  Written into the Settings sheet so an older
workbook can be recognised and explained rather than crashing the app."""
