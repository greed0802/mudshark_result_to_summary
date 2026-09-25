#!/usr/bin/env python3
"""Build test fixtures (requires openpyxl - DEV ONLY, not needed by the tool).

Reconstructs a Mudshark-style Result workbook from the values parsed out of
the user's sample file, and the expected Summary workbook the tool must
produce.  Hidden rows are included deliberately: in the real export the
"TrenchNetwork : ..." schedule rows and the trench-run cut row are
collapsed/hidden, so the tool must never rely on row visibility.
"""
import openpyxl

HDRS = ["Operation Group", "Exported (Bulked m\u00b3)", "Reused (Bulked m\u00b3)",
        "Cut (Bulked m\u00b3)", "From Site (Compressed m\u00b3)",
        "Imported (Banked m\u00b3)", "Fill (Compressed m\u00b3)",
        "Site Balance (Bulked m\u00b3)"]

# ---------------------------------------------------------------- Result book
def build_result(path):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "All Strata Operations"
    ws.append(HDRS)
    ws.append(["Built Structures", None, None, None, None,
               3180.354722891, 3180.354722891, 3180.354722891])
    ws.append([])
    ws.append(["Stripping (0.15m)", 1416.45377977016, None, 1416.45377977016,
               None, None, None, -1416.45377977016])
    ws.append([])
    ws.append(["Cut", 258.22757067882, 1505.32531478296, 1763.55288546177,
               None, None, None, -258.22757067882])
    ws.append([])
    ws.append(["Fill", None, None, None, 1505.32531478295, None,
               1505.32531478295, None])
    ws.append([])
    ws.append(["All Stormwater Drainage in Trenching site", None, None, None,
               None, 458.608016833591, 458.608016833591, 458.608016833591])
    ws.append([])
    ws.append([None, 1674.681, 1505.325, 3180.007, 1505.325, 3638.963,
               5144.288, 1964.281])

    ws = wb.create_sheet("Ground Layer Operations")
    ws.append(HDRS)
    ws.append(["Stripping (0.15m)", 1416.45377977016, None, 1416.45377977016,
               None, None, None, -1416.45377977016])
    ws.append([])
    ws.append(["Cut", 258.22757067882, 1505.32531478296, 1763.55288546177,
               None, None, None, -258.22757067882])
    ws.append([])
    ws.append(["Fill", None, None, None, 1505.32531478295, None,
               1505.32531478295, None])
    ws.append([])
    ws.append([None, 1674.681, 1505.325, 3180.007, 1505.325, 0.0, 1505.325,
               -1674.681])

    ws = wb.create_sheet("Structure Strata Operations")
    ws.append(HDRS)
    ws.append(["Built Structures", None, None, None, None, 3180.354722891,
               3180.354722891, 3180.354722891])
    ws.append([])
    ws.append([None, 0.0, 0.0, 0.0, 0.0, 3180.355, 3180.355, 3180.355])

    ws = wb.create_sheet("Trench Run Strata Operations")
    ws.append(HDRS)
    ws.append(["All Stormwater Drainage in Trenching site", None, None, None,
               None, 458.608016833591, 458.608016833591, 458.608016833591])
    # hidden trench-run cut row (hypothesised source of TRENCH CUT)
    ws.append(["Cut", 258.22757067882, 232.201633976261, 490.429204655081,
               None, None, None, -258.22757067882])
    ws.row_dimensions[3].hidden = True
    # hidden per-network fill rows, like a collapsed outline group
    ws.append(["TrenchNetwork : Stormwater Drainage", None, None, None, None,
               458.608016833591, 458.608016833591, 458.608016833591])
    ws.row_dimensions[4].hidden = True
    ws.append([None, 0.0, 0.0, 0.0, 0.0, 458.608, 458.608, 458.608])

    ws = wb.create_sheet("All Materials")
    ws.append(["Material"] + HDRS[1:])
    mats = [
        ("Concrete slab", 488.052191219668), ("Sand", 142.524296605108),
        ("Asphalt Type H", 72.7609702501979),
        ("Class 2", 876.184806741428), ("Capping Layer", 786.664290032471),
        ("Pavers", 189.460459760529),
        ("Mortar Bed with Anti-leachate Additive", 94.7302298802792),
        ("N32 Concrete Pavement", 473.651149401043),
        ("Class 3", 405.410184227484), ("25MPa Concrete", 70.2256337226629),
    ]
    for name, vol in mats:
        ws.append([name, None, None, None, None, vol, vol, vol])
        ws.append([])
    for name, e, r, c in [
            ("Fill: Admixed Silt, Sand, Clay, Gravel, Crushed Rock",
             1436.00284289913, 777.111534422031, 2213.11437732117),
            ("Fill: Admixed Silt, Sand, Clay, Rubble",
             38.8883403570581, 330.417450027128, 369.305790384187),
            ("Fill: Admixed Basalt, Rock and Clay",
             4.39759857874705, 2.6586662357779, 7.05626481452495),
            ("Silty Clay", 135.767515326358, 360.803310574691,
             496.570825901049),
            ("Basalt", 59.6250532876846, 34.3343535233266,
             93.9594068110112)]:
        ws.append([name, e, r, c, r, None, r, -e])
        ws.append([])
    ws.append(["Site Dirt", None, None, None, None, 39.2985278837208,
               39.2985278837208, 39.2985278837208])
    ws.append([])
    ws.append([None, 1674.681, 1505.325, 3180.007, 1505.325, 3638.963,
               5144.288, 1964.281])

    ws = wb.create_sheet("Structure Materials")
    ws.append(["Material"] + HDRS[1:])
    ws.append(["Class 2", None, None, None, None, 525.882144965585,
               525.882144965585, 525.882144965585])
    ws.append([None, 0.0, 0.0, 0.0, 0.0, 3180.355, 3180.355, 3180.355])

    ws = wb.create_sheet("Trench Run Materials")
    ws.append(["Material"] + HDRS[1:])
    ws.append(["Class 2", None, None, None, None, 350.302661775843,
               350.302661775843, 350.302661775843])
    ws.append([])
    ws.append(["Class 3", None, None, None, None, 69.0068271740277,
               69.0068271740277, 69.0068271740277])
    ws.append([])
    ws.append(["Site Dirt", None, None, None, None, 39.2985278837208,
               39.2985278837208, 39.2985278837208])
    ws.append([])
    ws.append([None, 0.0, 0.0, 0.0, 0.0, 458.608, 458.608, 458.608])

    ws = wb.create_sheet("Ground Layer Materials")
    ws.append(["Material"] + HDRS[1:])
    ws.append(["Fill: Admixed Silt, Sand, Clay, Gravel, Crushed Rock",
               1436.00284289913, 777.111534422031, 2213.11437732117,
               777.111534422031, None, 777.111534422031, -1436.00284289913])
    ws.append([None, 1674.681, 1505.325, 3180.007, 1505.325, 0.0, 1505.325,
               -1674.681])

    ws = wb.create_sheet("Trenches")
    lines = [
        ("Category : Stormwater Drainage (5 items) Quantity Total length : "
         "622.796m", False),
        ("TrenchNetwork : 150mm UPVC Pipe (32 items) Quantity Total length : "
         "120.851m", True),
        ("TrenchNetwork : 300mm RCP Pipes (19 items) Quantity Total length : "
         "405.433m", True),
        ("TrenchNetwork : 375mm RCP Pipes (1 item) Quantity Total length : "
         "12.091m", True),
        ("TrenchNetwork : 450mm RCP Pipe (2 items) Quantity Total length : "
         "32.373m", True),
        ("TrenchNetwork : 600mm RCP Pipe (3 items) Quantity Total length : "
         "52.048m", True),
        ("Category : Stormwater Pit (8 items) Quantity Total length : "
         "23.135m", False),
        ("TrenchNetwork : 1200 x 1200mm Grated Side Entry Pit (2 items) "
         "Quantity Total length : 2.698m", True),
        ("TrenchNetwork : 1500 x 1500mm Grated Entry Side Pit (1 item) "
         "Quantity Total length : 1.651m", True),
        ("TrenchNetwork : 1875 x 1800mm Grated Side Entry Pit Fitted with "
         "Altan Flow Filter (1 item) Quantity Total length : 2.265m", True),
        ("TrenchNetwork : 900 x 600mm Grated Side Entry Pit (1 item) "
         "Quantity Total length : 0.753m", True),
        ("TrenchNetwork : 900 x 900mm Grated Inlet Pit (4 items) Quantity "
         "Total length : 4.207m", True),
        ("TrenchNetwork : 900 x 900mm Grated Side Entry Pit (5 items) "
         "Quantity Total length : 5.260m", True),
        ("TrenchNetwork : 900 x 900mm Junction Pit (3 items) Quantity Total "
         "length : 3.153m", True),
        ("TrenchNetwork : 900 x 900mm Side Entry Pit (3 items) Quantity "
         "Total length : 3.148m", True),
    ]
    for i, (text, hidden) in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=text)
        if hidden:
            ws.row_dimensions[i].hidden = True

    ws = wb.create_sheet("Trench Depth Categories")
    ws.cell(row=1, column=1,
            value="Category : Stormwater Drainage (2 items) Quantity Total "
                  "length : 622.796m")
    ws.cell(row=2, column=1,
            value="Category : Stormwater Pit (2 items) Quantity Total "
                  "length : 23.135m")

    ws = wb.create_sheet("Areas and Perimeters")
    ws.append([None, "Name", "True Area (m\u00b2)", "Top Down Area (m\u00b2)",
               "Outside Perimeter (m)", "All Edges Length (m)",
               "Total Thickness (m)", "Note"])
    ws.append(["Built Structures"])
    ws.append([None, "BUILDING FOOTPRINT", 2850.48593210401, None,
               225.670395462282, 225.670395462282, 0.2, None])

    ws = wb.create_sheet("Structure Thickness Breakdown")
    ws.append([None, None, "Material", "Thickness (m)", "True Area (m\u00b2)",
               "Volume (m\u00b3)"])
    ws.append(["Built Structures"])

    wb.save(path)


# ------------------------------------------------------------- Expected book
def build_expected(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    rows = [
        HDRS,
        ["TOTAL SITE CUT", 258.22757067882, 1505.32531478296,
         1763.55288546177, None, None, None, -258.22757067882],
        ["SITE CUT ONLY", 0, 1273.12368080669, 1273.12368080669, 0, 0, 0, 0],
        ["TRENCH CUT", 258.22757067882, 232.201633976261, 490.429204655081,
         0, 0, 0, -258.22757067882],
        [], [],
        ["TOTAL TRENCH FILL", None, None, None, None, 458.608016833591,
         458.608016833591, 458.608016833591],
        ["CLASS 2", None, None, None, None, 350.302661775843,
         350.302661775843, 350.302661775843],
        ["CLASS 3", None, None, None, None, 69.0068271740277,
         69.0068271740277, 69.0068271740277],
        ["SITE DIRT", None, None, None, None, 39.2985278837208,
         39.2985278837208, 39.2985278837208],
        [], [],
        ["TOTAL TRENCH LENGTH", "SUMMARY"],
        ["Category : Stormwater Drainage (5 items) Quantity Total length : "
         "622.796m", "PIPE/PIT NAME", "COUNTS", "LENGTHS"],
        ["TrenchNetwork : 150mm UPVC Pipe (32 items) Quantity Total length : "
         "120.851m", "150mm UPVC Pipe", "32 items", "120.851m"],
        ["TrenchNetwork : 300mm RCP Pipes (19 items) Quantity Total length : "
         "405.433m", "300mm RCP Pipes", "19 items", "405.433m"],
        ["TrenchNetwork : 375mm RCP Pipes (1 item) Quantity Total length : "
         "12.091m", "375mm RCP Pipes", "1 item", "12.091m"],
        ["TrenchNetwork : 450mm RCP Pipe (2 items) Quantity Total length : "
         "32.373m", "450mm RCP Pipe", "2 items", "32.373m"],
        ["TrenchNetwork : 600mm RCP Pipe (3 items) Quantity Total length : "
         "52.048m", "600mm RCP Pipe", "3 items", "52.048m"],
        ["Category : Stormwater Pit (8 items) Quantity Total length : "
         "23.135m"],
        ["TrenchNetwork : 1200 x 1200mm Grated Side Entry Pit (2 items) "
         "Quantity Total length : 2.698m", "1200 x 1200mm Grated Side Entry "
         "Pit", "2 items", "2.698m"],
        ["TrenchNetwork : 1500 x 1500mm Grated Entry Side Pit (1 item) "
         "Quantity Total length : 1.651m", "1500 x 1500mm Grated Entry Side "
         "Pit", "1 item", "1.651m"],
        ["TrenchNetwork : 1875 x 1800mm Grated Side Entry Pit Fitted with "
         "Altan Flow Filter (1 item) Quantity Total length : 2.265m",
         "1875 x 1800mm Grated Side Entry Pit Fitted with Altan Flow Filter",
         "1 item", "2.265m"],
        ["TrenchNetwork : 900 x 600mm Grated Side Entry Pit (1 item) "
         "Quantity Total length : 0.753m", "900 x 600mm Grated Side Entry "
         "Pit", "1 item", "0.753m"],
        ["TrenchNetwork : 900 x 900mm Grated Inlet Pit (4 items) Quantity "
         "Total length : 4.207m", "900 x 900mm Grated Inlet Pit", "4 items",
         "4.207m"],
        ["TrenchNetwork : 900 x 900mm Grated Side Entry Pit (5 items) "
         "Quantity Total length : 5.260m", "900 x 900mm Grated Side Entry "
         "Pit", "5 items", "5.260m"],
        ["TrenchNetwork : 900 x 900mm Junction Pit (3 items) Quantity Total "
         "length : 3.153m", "900 x 900mm Junction Pit", "3 items", "3.153m"],
        ["TrenchNetwork : 900 x 900mm Side Entry Pit (3 items) Quantity "
         "Total length : 3.148m", "900 x 900mm Side Entry Pit", "3 items",
         "3.148m"],
    ]
    for r in rows:
        ws.append(r)
    wb.save(path)


if __name__ == "__main__":
    import os
    here = os.path.dirname(os.path.abspath(__file__))
    build_result(os.path.join(here, "fixture_result.xlsx"))
    build_expected(os.path.join(here, "expected_summary.xlsx"))
    print("fixtures written")
