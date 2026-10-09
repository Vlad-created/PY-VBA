import win32com.client as win32
import os
import shutil

excel = win32.Dispatch("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

# Base file path
base_path = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Krachev_ Crak\NBS_Target_South.xlsm"

# Output folder
output_folder = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Krachev_ Crak"

# Region groups (rows matched in column E are the ones kept untouched)
region_groups = {
    "Пловдив 1": ["Пловдив 1", "Регион Пловдив 1"],
    "Пловдив 2": ["Пловдив 2", "Регион Пловдив 2"],
    "Пловдив 3": ["Пловдив 3", "Регион Пловдив 3"],
    "Стара Загора": ["Стара Загора", "Регион Стара Загора"]
}

# Sheets to process (same order)
sheets_to_process = [
    "Upsell Revenue RSDN",
    "Voice",
    "Fix",
    "Devices",
    "Other Stacks",
    "Target per FTE",
    "Bonus&Checks",
    "Retention Stack"
]

# (column D value, column E value) -> None means "any value / not checked"

# Rules that are ALWAYS applied
mandatory_rules = [
    ("Запад", "Total"),
    ("Юг", "Total"),
    ("Изток", "Total"),
]

# Rules that are applied ONLY if at least one matching row exists in the file
optional_rules = [
    ("Total Own Shops", None),
    (None, "Франчайз магазини"),
    (None, "Регион Франчайз Дилъри"),
]

START_ROW = 21
xlUp = -4162
xlPasteValues = -4163
xlCalculationAutomatic = -4105
xlCalculationManual = -4135


def clean(v):
    return "" if v is None else str(v).strip()


def read_rows(ws):
    """Return (last_row, list of (row_number, D, E)) for rows from START_ROW down."""
    last_row = max(
        ws.Cells(ws.Rows.Count, "D").End(xlUp).Row,
        ws.Cells(ws.Rows.Count, "E").End(xlUp).Row,
    )
    if last_row < START_ROW:
        return last_row, []

    data = ws.Range(f"D{START_ROW}:E{last_row}").Value  # tuple of (D, E)
    rows = [(START_ROW + i, clean(d), clean(e)) for i, (d, e) in enumerate(data)]
    return last_row, rows


def rule_matches(rule, d, e):
    rule_d, rule_e = rule
    return (rule_d is None or d == rule_d) and (rule_e is None or e == rule_e)


def matches_any_rule(rules, d, e):
    return any(rule_matches(rule, d, e) for rule in rules)


def find_active_rules(wb):
    """Mandatory rules + only those optional rules that exist in at least one sheet."""
    found = set()
    for sheet_name in sheets_to_process:
        ws = wb.Sheets(sheet_name)
        _, rows = read_rows(ws)
        for _, d, e in rows:
            for i, rule in enumerate(optional_rules):
                if i not in found and rule_matches(rule, d, e):
                    found.add(i)

    active = list(mandatory_rules)
    for i, rule in enumerate(optional_rules):
        if i in found:
            active.append(rule)
            print(f"  Optional rule FOUND:   {rule}")
        else:
            print(f"  Optional rule SKIPPED (not in file): {rule}")
    return active


def to_runs(row_numbers):
    """Group sorted row numbers into consecutive (start, end) blocks."""
    runs = []
    for r in row_numbers:
        if runs and r == runs[-1][1] + 1:
            runs[-1][1] = r
        else:
            runs.append([r, r])
    return runs


def paste_values(ws, start, end, last_col):
    rng = ws.Range(ws.Cells(start, 1), ws.Cells(end, last_col))
    try:
        rng.Value = rng.Value  # fast, no clipboard
    except Exception:
        # fallback (e.g. merged cells)
        rng.Copy()
        rng.PasteSpecial(Paste=xlPasteValues)
        excel.CutCopyMode = False


for region_name, keep_values in region_groups.items():
    print(f"\n=== Creating file for region: {region_name} ===")

    region_file = os.path.join(output_folder, f"{region_name}.xlsm")
    shutil.copy2(base_path, region_file)

    wb = excel.Workbooks.Open(region_file)
    excel.ScreenUpdating = False
    excel.Calculation = xlCalculationManual

    # ------------------------------------------------------------
    # STEP 0: check which optional rules exist in the file
    # ------------------------------------------------------------
    print("STEP 0: Check which keep_rules exist in the file")
    keep_rules = find_active_rules(wb)

    # ------------------------------------------------------------
    # STEP 1: in ALL sheets, region rows stay untouched,
    #         every other row (from row 21 down) -> values
    # ------------------------------------------------------------
    print("STEP 1: Copy all non-region rows as values")
    for sheet_name in sheets_to_process:
        ws = wb.Sheets(sheet_name)
        print(f"  [Step 1] Sheet: {sheet_name}")

        last_row, rows = read_rows(ws)
        if not rows:
            continue

        last_col = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

        other_rows = [r for r, d, e in rows if e not in keep_values]
        for start, end in to_runs(other_rows):
            paste_values(ws, start, end, last_col)

    # ------------------------------------------------------------
    # STEP 2: only after ALL sheets are converted, delete rows that are
    #         neither region rows nor one of the active keep_rules rows
    # ------------------------------------------------------------
    print("STEP 2: Delete rows not matching region / keep rules")
    for sheet_name in sheets_to_process:
        ws = wb.Sheets(sheet_name)
        print(f"  [Step 2] Sheet: {sheet_name}")

        last_row, rows = read_rows(ws)
        if not rows:
            continue

        to_delete = [
            r for r, d, e in rows
            if e not in keep_values and not matches_any_rule(keep_rules, d, e)
        ]

        # Bottom-up so row numbers above never shift
        for start, end in reversed(to_runs(to_delete)):
            ws.Rows(f"{start}:{end}").Delete()

    # ------------------------------------------------------------
    # STEP 3: hide the rows that match the active keep_rules (in every sheet)
    #         Re-read the rows, because Step 2 changed the row numbers
    # ------------------------------------------------------------
    print("STEP 3: Hide rows matching keep_rules")
    for sheet_name in sheets_to_process:
        ws = wb.Sheets(sheet_name)
        print(f"  [Step 3] Sheet: {sheet_name}")

        last_row, rows = read_rows(ws)
        if not rows:
            continue

        to_hide = [
            r for r, d, e in rows
            if e not in keep_values and matches_any_rule(keep_rules, d, e)
        ]

        for start, end in to_runs(to_hide):
            ws.Rows(f"{start}:{end}").EntireRow.Hidden = True

    excel.Calculation = xlCalculationAutomatic
    excel.ScreenUpdating = True

    wb.Save()
    wb.Close()

    print(f"Saved: {region_file}")

print("\n✅ All region files created successfully.")