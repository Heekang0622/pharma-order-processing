# Pharma Order Processing Pipeline (Excel VBA)

An Excel VBA macro pipeline that automates a weekly pharmaceutical import/export order-processing workflow. It replaces manual transcription of purchase orders into the production-request format used by the manufacturer.

> **Note:** This is a sanitized/anonymized version of a tool used in a real workflow. Original client names, agency codes, and partner names have been replaced with generic placeholders (`PARTNER_CO`, `CLIENT01`-`CLIENT31`, etc.), and sheet names have been translated to English. Business logic and structure are preserved as-is. The weekday labels (`일,토,금,목,수,화,월` = Sun-Mon) stay in Korean because they are data values that the sorting depends on.

## Pipeline

All steps write to one target sheet (`Distribution`). Each step clears its own target range before writing, so every macro can be re-run safely.

| Step | Macro | Source sheet | What it does |
|---|---|---|---|
| 1 | `Step1_AutoFill_Smart_Optimized` | `OrderData_Intl` | Clears the previous data and expands each international order into **one row per unit** |
| 2 | `Step2_Final_Merge_Resolved` | `Distribution` | Sorts the rows and merges them into **package groups** |
| 3 | `Step3_Domestic_Only_Final_Fixed` | `OrderData_Domestic` | Appends domestic orders below the international rows, sorted the same way |
| 4 | `Step4_Solution_Mapping` | `SolutionData` | Maps the third order channel, normalizes client names through a lookup sheet, and sorts |

The helper `ToggleEvents` switches screen updating, alerts, events, and calculation mode on or off.

### Step 1 - Expand orders to rows
- Clears the previous run (the old separate "Step 0" is now part of this step): unmerges `L15:P981`, then clears `B:C`, `E:G`, `J`, and `L:Q` from row 15 down.
- Reads each order row. It accepts two input formats: a single cell holding `Agency Day mCi Qty` separated by spaces, or the same four values in columns A-D.
- Writes one row per unit of quantity: client (`C`), weekday (`E`), mCi (`F`), and the channel marker `PARTNER_CO` (`J`). Column `D` is filled down from the first row.
- Stops with a message if the order sheet is empty.

### Step 2 - Sort and merge into packages
- Finds the last `PARTNER_CO` row and works on `A15:Q` down to it.
- **Sort order:** mCi descending → weekday (Sun → Mon) → client priority (a fixed list of 31 client codes).
- **Package rule by dose:** 50, 30, 25, 5 and 2 mCi group 4 rows per package, 100 mCi groups 2 rows, any other dose stays 1 row. A group only forms while weekday and client are identical.
- For each package it writes the numbering (`No.` plus `=A<row>` references), a merged label such as `1 Package (100mCi)`, and a thin outer border.
- Rows of the priority client group (`CLIENT26`, `CLIENT30`, `CLIENT18`, `CLIENT27`, `CLIENT17`, `CLIENT_SPECIAL`) are marked pink in column `Q`.

### Step 3 - Domestic orders
- Starts right below the last `PARTNER_CO` row and clears everything below it first.
- Expands each domestic order into one row per unit (client, weekday, mCi, and the channel value taken from the domestic sheet).
- Sorts by mCi descending, then weekday (Sat → Mon), merges `L:P` for each row, and draws borders.

### Step 4 - Solution channel
- Reads `SolutionData` and writes the rows starting at row 987 of `Distribution`.
- **Client names are standardized** through the `ClientNameMap` sheet using a wildcard `MATCH`; if no match is found, the original name is kept.
- Sorts by mCi descending, then weekday (Sat → Mon).

## Design notes

- Row expansion works on arrays and ranges in bulk, with screen updating, alerts, and calculation switched off while running.
- Package grouping follows the dose-denomination rules above.
- Name normalization uses a lookup sheet with fuzzy `MATCH`.
- Each step clears its target range first, so macros can be re-run without duplicating rows.

## How to use

1. Prepare a workbook with the sheets `OrderData_Intl`, `OrderData_Domestic`, `SolutionData`, `ClientNameMap`, and `Distribution`.
2. Press `Alt+F11` → File → Import File → select `OrderPipeline_Anonymized.bas`.
3. Press `Alt+F8` and run the steps in order, 1 → 4.

## Known limitations

- Business logic is split between VBA and manual PDF reading upstream. A Python PDF-parsing layer (pdfplumber) was added to reduce transcription errors.
- The client priority list and the package-size mapping are hardcoded.
- Row positions in the target sheet are fixed (data starts at row 15, the solution block at row 987, clear ranges end around row 981), so a very large week would need these limits raised.
- There are no automated tests. Validation compares totals against the printed PO totals.

## Stack

Excel VBA with no external dependencies, designed for a locked-down corporate environment.
