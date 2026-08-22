#!/usr/bin/env python3
"""End-to-end formula evaluation: computes the real workbook with the `formulas` engine
and asserts the business math on sample data."""
import formulas, os

xl = formulas.ExcelModel().loads("dist/NegosyoSheet.xlsx").finish()
sol = xl.calculate()

def get(sheet, cell):
    key = f"'[NegosyoSheet.xlsx]{sheet.upper()}'!{cell}"
    v = sol[key].value
    try: return float(v[0][0])
    except Exception:
        try: return float(v)
        except Exception: return v

ok = 0; bad = 0
def check(name, cond):
    global ok, bad
    if cond: ok += 1; print(f"  ✓ {name}")
    else: bad += 1; print(f"  ✗ FAIL: {name}")

# INVENTORY samples: Noka 48-12=36 stock, value 36×12=432; Softdrinks 24-5=19, value 19×68=1292
check("Noka current stock = 45 (48 in − 3 auto from sales)", get("INVENTORY", "G2") == 45)
check("Noka stock value = 540", get("INVENTORY", "H2") == 540)
check("Softdrinks stock value = 1564", get("INVENTORY", "H3") == 1564)
# Noka reorder: 36 > 12 → OK; Sabon: 30-8=22 > 10 → OK; make a reorder case? sample3 reorder=10, stock=22 → OK
check("Noka status OK", get("INVENTORY", "J2") == "OK")

# SALES LOG: noka x3 → price 18 (VLOOKUP), total 54, cost 36, profit 18; softdrinks x1 → total 90, profit 22
check("Noka unit price looked up = 18", get("SALES LOG", "D2") == 18)
check("Noka sale total = 54", get("SALES LOG", "E2") == 54)
check("Noka sale cost = 36", get("SALES LOG", "F2") == 36)
check("Noka sale profit = 18", get("SALES LOG", "G2") == 18)
check("Softdrinks profit = 22", get("SALES LOG", "G3") == 22)

# DASHBOARD: total sales 54+90=144; gross profit 18+22=40; net = 40 − 0 expenses = 40
check("dashboard total sales = 144", get("DASHBOARD", "B6") == 144)
check("dashboard gross profit = 40", get("DASHBOARD", "E6") == 40)
check("dashboard net = 40", get("DASHBOARD", "E9") == 40)
check("dashboard inventory value = 540+1564+660", abs(get("DASHBOARD", "B12") - (540+1564+660)) < 0.001)
check("dashboard utang unpaid = 0", get("DASHBOARD", "E12") == 0)
check("dashboard reorder count = 0", get("DASHBOARD", "B15") == 0)
check("GCash fee income card = 0 (empty ledger)", get("DASHBOARD", "B18") == 0)
check("today sales card = 144 (samples dated today)", get("DASHBOARD", "E18") == 144)
check("MTD sales card = 144", get("DASHBOARD", "B21") == 144)

print(f"\nRESULT: {ok} passed, {bad} failed")
raise SystemExit(1 if bad else 0)
