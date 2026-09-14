# BYD / Reference Model Contract

The reference model's revenue split sheet is a contract, not a loose inspiration.

## Must Preserve

Preserve these from the BYD/reference revenue split sheet:

- sheet name style and final sheet name requested by the user
- row and section order
- header rows and year columns
- unit display
- revenue bridge logic
- percentage-of-revenue rows
- detailed row grouping
- payroll, employee-count, supplier/cost, bank, government/tax, depreciation, amortization, net-profit sections
- formulas for calculated rows
- red/bold/fill conventions
- bright-yellow forecast cell behavior
- number formats and percentage formats
- validation/check block logic
- formula structure, but not BYD-specific unit conversion factors
- navigation intent, but not stale copied hyperlink locations

The output must look and behave like the reference model after substituting target-company data and disclosure-specific detail rows.

## Allowed Changes

Allowed changes are narrow:

- Replace the company name and source sheet references.
- Replace BYD-specific unit conversion factors after reading the target three-statement workbook unit and revenue-split output unit.
- Replace internal hyperlinks after renaming/copying the sheet so they point to the actual output sheet and correct target rows.
- Reset sheet-view/freeze-pane settings if the copied view prevents normal scrolling.
- Replace detailed expense rows with the target company's disclosed note items.
- Insert or delete detail rows inside a section when disclosure requires it.
- Update SUM ranges, ratio formulas, and check formulas after row insertions/deletions.
- Leave forecast/yellow cells blank.
- Add source-note or validation rows when needed for auditability, preferably near the check block.

## Forbidden Changes

Do not:

- generate a generic table from scratch when a reference model is available
- drop the reference model's bridge rows or percentage rows
- simplify the workbook into only extracted notes
- replace formula rows with hardcoded values
- remove validation rows
- remove styles or visual conventions
- keep BYD-only detail items that the target company does not disclose
- create residual "其他" rows unless disclosed by the target company
- keep BYD's `/100`, `*100`, `*0.01`, or other conversion factors without checking the target source/output units
- keep hyperlinks pointing to `Sheet1` or any old template sheet name
- keep a deep freeze pane that makes the generated output difficult or impossible to scroll

## Pre-Edit Checklist

Before writing data, inspect and record:

- reference sheet name
- used range dimensions
- year/header row
- unit row
- source statement unit and expected conversion factor into the output unit
- start/end rows for each major section
- rows with formulas
- cells with bright-yellow fill
- cells with red/bold/fill formatting
- freeze pane location and whether it leaves a usable scrollable pane
- A1/A2 hyperlink locations after the final sheet name is set

After editing, compare the output against this checklist. Any structural change must be explained by a target-company disclosure difference.

## View And Navigation

When copying a template sheet, Excel view state is copied too. Rebuild it after all row edits:

- Prefer no freeze pane for smoke tests and generated deliverables unless the user asks for one.
- If freezing panes, freeze only a shallow header area such as the year/header rows; do not freeze 30+ rows by default.
- Set the active selection to a valid visible cell such as `A1`.
- Recreate A1/A2 as internal hyperlinks using the final sheet name. For example:
  - A1 -> `'<final sheet name>'!C3`
  - A2 -> `'<final sheet name>'!C37`
- Do not use external-style `#Sheet!Cell` hyperlinks when an internal `location` hyperlink is available.
