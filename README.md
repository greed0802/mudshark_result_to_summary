# mudshark_result_to_summary

Turns a **Mudshark "Results" Excel export** (`.xlsx`) into a compact
**Summary workbook** automatically — no more manual copying every time
a new export lands.

## Files

| File | What it is |
|---|---|
| `mudshark_summary.py` | The program (pure Python standard library — **nothing to install** via pip) |
| `Run_Mudshark_Summary.bat` | Windows helper — drag & drop the export onto it |
| `dev_tests/` | Test fixtures & verification (only used for development) |

## Quick start (company laptop, **no admin rights needed**)

You need Python 3.8+ only. Everything else is already inside Python.

**Don't have Python yet?** Either of these works without admin:

1. **Microsoft Store** — open the Store app, search **"Python 3.12"**
   (published by the Python Software Foundation), click **Get**.
   It installs per-user, no admin prompt.
2. **python.org installer** — during setup, *untick*
   "Install launcher for all users" and leave "Install for all users"
   unticked, so it installs just for you.

**Then run it — two ways:**

* **Easy:** drag the Mudshark export `.xlsx` onto
  `Run_Mudshark_Summary.bat`.
* **Command line:**
  ```
  py mudshark_summary.py "C:\path\to\Mudshark Export.xlsx"
  py mudshark_summary.py "C:\path\to\Mudshark Export.xlsx" -o Summary.xlsx
  ```

The Summary is written next to the export as
`<export name>_SUMMARY.xlsx` (or the name you give with `-o`).

## What the Summary contains (and where it comes from)

| Summary row | Source inside the Mudshark export |
|---|---|
| **TOTAL SITE CUT** | The `Cut` row of sheet **All Strata Operations** (Exported / Reused / Cut / Site Balance) |
| **SITE CUT ONLY** | `TOTAL SITE CUT` − `TRENCH CUT` |
| **TRENCH CUT** | Cut-bearing row(s) on sheet **Trench Run Strata Operations** — these rows are often *collapsed/hidden* in the export; the script reads them anyway |
| **TOTAL TRENCH FILL** | The `…Trenching site` row of **All Strata Operations** (Imported / Fill / Site Balance) |
| **CLASS 2, CLASS 3, SITE DIRT, …** | Every material row of sheet **Trench Run Materials** (labels are upper-cased) |
| **TOTAL TRENCH LENGTH** schedule | The `Category : …` and `TrenchNetwork : …` lines of sheet **Trenches**, split into `PIPE/PIT NAME`, `COUNTS`, `LENGTHS` |

Number states (Bulked / Compressed / Banked) are copied exactly as
Mudshark reports them; nothing is re-calculated except
`SITE CUT ONLY` = `TOTAL SITE CUT` − `TRENCH CUT` (same formula used in
the hand-made Summary).

## If something doesn't line up

* The script expects Mudshark's standard sheet names. If your export
  differs, adjust the **CONFIG** block at the top of
  `mudshark_summary.py` (`SHEET_ALL_STRATA`, `SHEET_TRENCH_STRATA`,
  `SHEET_TRENCH_MATERIALS`, `SHEET_TRENCHES`, `CUT_ROW_LABEL`,
  `TRENCH_SITE_HINT`).
* If the **TRENCH CUT** volumes can't be located, those two rows are
  left blank and the script prints a warning plus writes a
  `*_DEBUG_DUMP.csv` (every row of every sheet) next to the export —
  send that CSV to the person maintaining the script so the rule can be
  tuned.
* You can also force the dump any time with `--debug`:
  ```
  py mudshark_summary.py export.xlsx --debug
  ```

Only `.xlsx` files are supported (that is what Mudshark's Excel export
produces).

## For developers

```
cd dev_tests
python build_fixtures.py   # needs openpyxl (dev only)
python run_tests.py        # regenerates the summary and diffs it
```

The shipped program itself has **zero third-party dependencies**;
`openpyxl` is only used to build the test fixture.
