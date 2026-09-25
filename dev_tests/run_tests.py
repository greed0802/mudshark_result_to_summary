#!/usr/bin/env python3
"""Run mudshark_summary.py on the fixture and diff against the expected file.

- Reads BOTH files with the tool's own stdlib reader (consistency).
- ALSO opens the generated file with openpyxl to prove the writer produces
  a valid, Excel-compatible .xlsx.
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from mudshark_summary import read_workbook  # noqa: E402

FIXTURE = os.path.join(HERE, "fixture_result.xlsx")
EXPECTED = os.path.join(HERE, "expected_summary.xlsx")
ACTUAL = os.path.join(HERE, "actual_summary.xlsx")


def sheet_matrix(path):
    book = read_workbook(path)
    rows = next(iter(book.values()))  # first (and only) sheet
    return {rn: cells for rn, _hidden, _lvl, cells in rows}


def nearly(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) \
            and not isinstance(a, bool) and not isinstance(b, bool):
        return abs(float(a) - float(b)) <= 1e-9 * max(1.0, abs(float(a)))
    return a == b


def compare():
    exp = sheet_matrix(EXPECTED)
    act = sheet_matrix(ACTUAL)
    problems = 0
    for i in sorted(set(exp) | set(act)):
        e = exp.get(i, {})
        a = act.get(i, {})
        cols = sorted(set(e) | set(a))
        for c in cols:
            ev = e.get(c)
            av = a.get(c)
            if ev in (" ", ""):
                ev = None
            if av in (" ", ""):
                av = None
            if ev is None and av is None:
                continue
            if isinstance(ev, str):
                ev = ev.strip()
            if isinstance(av, str):
                av = av.strip()
            if not nearly(ev, av):
                print("  MISMATCH row %-3d col %-2d expected=%r actual=%r"
                      % (i, c, ev, av))
                problems += 1
    return problems


def trap_checks():
    """Verify grouped/detail rows from the export are NOT picked up."""
    matrix = sheet_matrix(ACTUAL)
    bad_tokens = ("TrenchRun :", "TrenchSegment", "A - B", "UNDER STRUCTURE",
                  "TRENCH RUN 1", "Trench Run")
    problems = 0
    for rn, cells in sorted(matrix.items()):
        for col, v in cells.items():
            if isinstance(v, str) and any(tok in v for tok in bad_tokens):
                print("  TRAP-FAIL row %d col %d: %r" % (rn, col, v))
                problems += 1
    labels = {cells.get(1): rn for rn, cells in matrix.items()
              if isinstance(cells.get(1), str)}
    ttf = matrix[labels["TOTAL TRENCH FILL"]]
    if not nearly(ttf.get(6), 458.608016833591):
        print("  TRAP-FAIL: TOTAL TRENCH FILL imported = %r (dup summed?)"
              % ttf.get(6))
        problems += 1
    so = matrix[labels["SITE CUT ONLY"]]
    if so.get(2, 1) < 0:
        print("  TRAP-FAIL: SITE CUT ONLY exported negative: %r" % so.get(2))
        problems += 1
    mats = [l for l in labels if l in ("CLASS 2", "CLASS 3", "SITE DIRT")]
    if len(mats) != 3:
        print("  TRAP-FAIL: material class rows = %r" % mats)
        problems += 1
    return problems


def negative_test():
    """Inflated trench cut (as seen on the Bulla Road export) must raise the
    red CHECK note in the workbook and a console warning."""
    try:
        import openpyxl
    except ImportError:
        print("(openpyxl not available - negative test skipped)")
        return 0
    variant = os.path.join(HERE, "fixture_inflated.xlsx")
    out = os.path.join(HERE, "actual_inflated.xlsx")
    wb = openpyxl.load_workbook(FIXTURE)
    ws = wb["Trench Run Strata Operations"]
    ws.cell(row=3, column=2, value=1925.763693)   # exported
    ws.cell(row=3, column=3, value=None)          # reused
    ws.cell(row=3, column=4, value=1925.763693)   # cut
    # the deeper-level per-run rows should be ignored by the level filter
    for r in (4, 5):
        for c in (2, 3, 4):
            ws.cell(row=r, column=c, value=None)
    wb.save(variant)
    r = subprocess.run([sys.executable,
                        os.path.join(ROOT, "mudshark_summary.py"),
                        variant, "-o", out],
                       capture_output=True, text=True)
    problems = 0
    if "bigger than TOTAL SITE CUT" not in r.stdout:
        print("  NEG-FAIL: console warning missing:\n" + r.stdout)
        problems += 1
    matrix = sheet_matrix(out)
    notes = [v for rn, cells in matrix.items() for v in cells.values()
             if isinstance(v, str) and v.startswith("!! CHECK:")]
    if not notes:
        print("  NEG-FAIL: red CHECK note row missing in workbook")
        problems += 1
    elif "row" not in notes[0]:
        print("  NEG-FAIL: note should list the rows used: %r" % notes[0][:120])
        problems += 1
    so_label = next(rn for rn, cells in matrix.items()
                    if cells.get(1) == "SITE CUT ONLY")
    so = matrix[so_label]
    if not (so.get(2, 1) < 0 and so.get(4, 1) < 0):
        print("  NEG-FAIL: SITE CUT ONLY should be negative here: %r" % so)
        problems += 1
    for f in (variant, out):
        try:
            os.remove(f)
        except OSError:
            pass
    return problems


def main():
    print("== running mudshark_summary.py on fixture ==")
    r = subprocess.run([sys.executable,
                        os.path.join(ROOT, "mudshark_summary.py"),
                        FIXTURE, "-o", ACTUAL],
                       capture_output=True, text=True)
    print(r.stdout)
    if r.returncode != 0:
        print(r.stderr)
        return 1

    print("== diff actual vs expected ==")
    problems = compare()
    if problems:
        print("FAILED: %d mismatch(es)" % problems)
        return 1
    print("OK - generated Summary matches the hand-made one cell-for-cell.")

    print("== trap-row checks ==")
    problems = trap_checks()
    if problems:
        print("FAILED: %d trap check(s)" % problems)
        return 1
    print("OK - grouped detail rows correctly ignored.")

    print("== inflated trench cut sanity warning ==")
    problems = negative_test()
    if problems:
        print("FAILED: %d negative check(s)" % problems)
        return 1
    print("OK - anomaly produces red CHECK note + console warning.")

    print("== validating output with openpyxl ==")
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ACTUAL)
        ws = wb.active
        print("openpyxl opened '%s' fine: sheet=%r size=%dx%d"
              % (os.path.basename(ACTUAL), ws.title,
                 ws.max_row, ws.max_column))
        print("merged ranges: %s" % [str(r) for r in ws.merged_cells.ranges])
        a2 = ws["A2"].fill.fgColor.rgb
        a3 = ws["A3"].fill.fgColor.rgb
        f13 = ws["A13"].font.name
        print("A2 fill=%s A3 fill=%s font=%s" % (a2, a3, f13))
        if a2 != "FFFFFF00" or a3 != "FF9DC3E6":
            print("FAILED: unexpected band colours")
            return 1
        if not any(str(r).startswith("B13") for r in ws.merged_cells.ranges):
            print("FAILED: SUMMARY merge missing")
            return 1
    except ImportError:
        print("(openpyxl not available here - skipped)")
    except Exception as exc:  # noqa: BLE001
        print("FAILED: openpyxl could not open the generated file:", exc)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
