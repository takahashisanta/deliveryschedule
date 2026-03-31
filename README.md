# PSD Driver Schedule – Setup Guide

Automated delivery scheduling system for PSD orders.
Reads from **MASTER DATA**, assigns vehicles, sorts stops, and outputs a printable driver schedule.

---

## Files

| File | Purpose |
|------|---------|
| `create_workbook.py` | One-time script — creates the Excel workbook with all sheets |
| `DriverSchedule.bas` | VBA module — contains all macro logic |
| `README.md` | This file |

---

## Step 1 – Generate the Excel Workbook

Requires Python 3.8+ and `openpyxl`:

```bash
pip install openpyxl
python create_workbook.py
```

This creates **`PSD_Driver_Schedule.xlsx`** with three sheets:

| Sheet | Purpose |
|-------|---------|
| `DRIVER_RULES` | Configuration — day/area schedule, customer vehicle preferences, capacity |
| `DRIVER_ENGINE` | Hidden calculation table (populated by macro, don't edit manually) |
| `DRIVER_SCHEDULE` | Printable output — Truck Route + Van Route |

---

## Step 2 – Add the DRIVER_RULES sheet to your PSD Workbook

The DRIVER_RULES sheet is the only new sheet you need to add. Everything else
(MASTER DATA, Driver Schedule Sheet) already lives in your PSD Order Process Sheet.xlsm.

1. Open both files in Excel simultaneously
2. In `PSD_Driver_Schedule.xlsx`, right-click the **DRIVER_RULES** tab → **Move or Copy…**
   - "To book:" → `PSD Order Process Sheet.xlsm`
   - Check **Create a copy** → OK
3. Optionally also copy **DRIVER_ENGINE** (hidden sheet used for audit/debug — not required)

---

## Step 3 – Import the VBA Module

1. Open the workbook in Excel
2. Press **Alt + F11** to open the VBA Editor
3. In the menu: **File → Import File…**
4. Select `DriverSchedule.bas` → click Open
5. Close the VBA Editor
6. Save the file as **`.xlsm`** (macro-enabled): File → Save As → Excel Macro-Enabled Workbook

---

## Step 4 – Add the "Generate Schedule" Button

1. Go to the **Driver Schedule Sheet**
2. **Developer tab → Insert → Button (Form Control)**
   - If Developer tab is hidden: File → Options → Customize Ribbon → check Developer
3. Draw a button near cell F2 (where the placeholder text says)
4. When prompted "Assign Macro" → select **`GenerateSchedule`** → OK
5. Right-click the button → **Edit Text** → type `Generate Schedule`

Optional: Add a second button assigned to **`ClearSchedule`**.

---

## Step 5 – Add a Pallets Column to MASTER DATA (if not already there)

The macro needs a numeric **Pallets** column in MASTER DATA to decide Truck vs Van.

**Suggested setup:**
1. In MASTER DATA, insert a new column at **col J** and label it `Pallets`
2. Enter the pallet count per order row (can be decimals, e.g. `0.5` for half pallet)
3. The VBA constant `MD_COL_PALLETS = 10` (col J) already points to this column

If you place it in a different column, update the constant in the VBA module:
```vba
Private Const MD_COL_PALLETS  As Long = 10  ' J – change to match your column
```

---

## Step 6 – Verify All MASTER DATA Column Positions

Open `DRIVER_RULES` sheet and check **TABLE 5 – MASTER DATA COLUMN MAP**.
Compare against your actual MASTER DATA columns and update the `MD_COL_*` constants
at the top of the VBA module (`Alt+F11` → DriverScheduleMod) if anything differs:

```vba
Private Const MD_COL_SOLD_TO    As Long = 3   ' C – Sold To
Private Const MD_COL_ADDRESS    As Long = 5   ' E – Full delivery address
Private Const MD_COL_AREA       As Long = 6   ' F – Area
Private Const MD_COL_SHIP_MTH   As Long = 7   ' G – Shipping Method
Private Const MD_COL_PRODUCT    As Long = 8   ' H – Product
Private Const MD_COL_QTY        As Long = 9   ' I – Quantity / Cases
Private Const MD_COL_PALLETS    As Long = 10  ' J – Pallets (add this column)
' ... etc.
```

---

## Daily Usage

1. Open **PSD Order Process Sheet.xlsm**
2. Go to **Driver Schedule Sheet**
3. Click cell **B1** — enter the delivery date (e.g. `4/7/2026`)
4. Click **Generate Schedule**
5. The macro will:
   - Filter MASTER DATA for that date + matching day areas + Delivery/Pick-up method
   - Skip superseded versions (keeps latest version per Request ID)
   - Assign vehicles (Truck / Van) per customer rules + pallet count
   - Calculate time priority (urgent first)
   - Sort: Vehicle → Time Priority → Area → Customer
   - Write Section 1 (Truck) starting at **row 5**
   - Write Section 2 (Van) starting at **row 74**
   - Groups multiple products for the same customer+address as one numbered stop
   - Puts one blank row between stops
   - Pink background on Pick-up rows
   - Time shown as: **Anytime** / **By HH:MM** / **Between HH:MM and HH:MM**
6. Print with **Ctrl + P**

---

## Business Rules Reference

### Day → Area Schedule

| Day | Areas |
|-----|-------|
| Mon | JFC/IDC, SOUTH BAY, CENTRAL LA, SAN GABRIEL, EAST LA |
| Tue | LAX, NAX, JFC/IDC, CENTRAL LA |
| Wed | LAX, NAX, JFC/IDC, SAN GABRIEL, NORTH OC, SOUTH LA |
| Thu | JFC/IDC, SOUTH BAY, CENTRAL LA, EAST LA |
| Fri | LAX, NAX, WEST LA, CENTRAL LA, NORTH OC |

### Vehicle Assignment

| Customer Group | Assignment |
|---------------|-----------|
| Cal Hono, Kpac, Broad Leaf, AJ Foods LLC, SJ, JK Trucking, Khee Trading, Horn Foods, Han Seafood, Maruzen, Honolulu Freight, World Food, HPL Apollo, Recycling, NEU | **Truck** |
| Mitsuwa (all), Turuhashi, Hey Den, CRC, Sushi Zanmai | **Van** |
| LA Cold, JFC LA, JFC IDC, Wismettac, Capital Meat, Central Boeki, NAX, WCPM | **Flexible** → Truck if ≥4 pallets, else Van |

### Time Priority

| Window End | Priority | Color |
|-----------|---------|-------|
| By 8:00 AM | 1 – Urgent | Orange highlight |
| 8:01 – 12:00 | 2 – Morning | — |
| 12:01 – 15:00 | 3 – Afternoon | — |
| No time | 9 – Open | — |

### Capacity Limits

| Vehicle | Max Pallets |
|---------|------------|
| Truck | 8 |
| Van | 3 |

> Overflow handling: flexible customers with fewer pallets are automatically bumped to Van to stay within truck capacity.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "Sheet 'MASTER DATA' not found" | Rename your sheet to exactly `MASTER DATA` |
| "No matching deliveries found" | Check: date is correct, area matches day schedule, Shipping Method column contains "Delivery" |
| Wrong vehicle assignments | Update TABLE 2 in DRIVER_RULES — add/edit customer names |
| Wrong columns read | Update MD_COL_* constants in VBA module |
| Superseded orders still showing | Ensure Version column (T) has numeric values; latest = highest number |

---

## Customizing Rules (No VBA Required)

All rule changes can be made directly in the **DRIVER_RULES** sheet:

- **Add/change areas per day** → Edit TABLE 1
- **Add/change customer vehicle preference** → Edit TABLE 2 (Truck / Van / Flexible)
- **Column map reference** → TABLE 5

After editing DRIVER_RULES, the next **Generate Schedule** run uses the new rules automatically.
