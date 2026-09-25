# mudshark_result_to_summary

Turns a **Mudshark "Results" Excel export** (`.xlsx`) into a compact
**Summary workbook** automatically — no more manual copying every time
a new export lands.

## Files

| File | What it is |
|---|---|
| `mudshark_summary.py` | The program (pure Python standard library — **nothing to install** via pip) |
| `SETUP_Python.bat` | **One-time** helper: downloads a *portable* Python (plain ZIP — no admin, no Store, no installer) |
| `Run_Mudshark_Summary.bat` | Windows helper — drag & drop the export onto it |
| `dev_tests/` | Test fixtures & verification (only used for development) |

## Quick start (company laptop — **no admin, no Microsoft Store needed**)

1. **One time only:** double-click `SETUP_Python.bat`.
   It downloads a *portable* Python ZIP (~11 MB) from python.org and
   extracts it into a `python` folder next to itself. Nothing is
   installed — it's just extracted files, so no admin rights are needed
   and IT policy is untouched. (If Python is already on the machine,
   skip this step entirely.)

   *If the download is blocked by the company network, the script shows
   a link — download it in your browser, right-click → Extract All…
   into a folder called `python` next to the BAT files. Same result.*

2. **Every export:** drag the Mudshark export `.xlsx` onto
   `Run_Mudshark_Summary.bat`.
   Or from the command line:
   ```
   python\python.exe mudshark_summary.py "C:\path\to\Mudshark Export.xlsx"
   python\python.exe mudshark_summary.py "C:\path\to\Mudshark Export.xlsx" -o Summary.xlsx
   ```

3. The Summary is written next to the export as
   `<export name>_SUMMARY.xlsx` (or the name you give with `-o`).

## What the Summary contains (and where it comes from)

| Summary row | Source inside the Mudshark export |
|---|---|
| **TOTAL SITE CUT** | The `Cut` row of sheet **All Strata Operations** (top outline level — grouped detail rows are ignored) |
| **SITE CUT ONLY** | `TOTAL SITE CUT` − `TRENCH CUT` |
| **TRENCH CUT** | Top-level cut-bearing row(s) on sheet **Trench Run Strata Operations** (grouped per-run duplicates are ignored — prevents double counting) |
| **TOTAL TRENCH FILL** | The top-level `…Trenching site` row(s) of **All Strata Operations** (Imported / Fill / Site Balance) |
| **CLASS 2, CLASS 3, SITE DIRT, …** | The **class-level** material rows of sheet **Trench Run Materials** — the per-trench-run breakdown rows grouped underneath are *not* copied |
| **TOTAL TRENCH LENGTH** schedule | Only the `Category : …` and `TrenchNetwork : …` lines of sheet **Trenches**, split into `PIPE/PIT NAME`, `COUNTS`, `LENGTHS` (grouped `TrenchRun :` / `TrenchSegment` / `A - B` detail lines are *not* copied) |

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
* If your company blocks **running** the portable `python.exe`
  (unapproved-software policy — rare, but it happens on very strict
  builds), say so: the whole converter can also be delivered as an
  Excel macro (`.xlsm`) that runs inside Excel itself, with no Python
  at all.
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
