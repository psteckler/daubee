# Daubee User Guide

Daubee checks that the *units of measure* in your Excel formulas are
used consistently — the way spell-check catches typos, Daubee catches errors
like adding euros to dollars, or adding meters to seconds. You tell
it the unit a cell represents (`kg`, `USD`, `meters/second`, ...), and
Daubee tracks that unit through every formula that touches it, flagging the
formulas where there's unit inconsistency.

## Contents

- [Opening the pane](#opening-the-pane)
- [Annotating a cell with a unit](#annotating-a-cell-with-a-unit)
- [Writing a unit formula](#writing-a-unit-formula)
- [Browsing available units](#browsing-available-units)
- [Infer Units](#infer-units)
- [Automation](#automation)
- [Check Units](#check-units)
- [Synthesizing units for CONVERT](#synthesizing-units-for-convert)
- [Reading the borders](#reading-the-borders)
- [Errors for selection](#errors-for-selection)
- [Consolidate](#consolidate)
- [Assert](#assert)
- [WITHUNIT](#withunit)
- [Currencies](#currencies)
- [Settings](#settings)
- [Unit dictionaries](#unit-dictionaries)
- [Audit report](#audit-report)
- [Tips](#tips)

## Opening the pane

Daubee adds its own **Daubee** tab to the Excel ribbon. Click **Show
pane** to open the task pane; it stays open alongside your worksheet like
any other Excel add-in pane. The pane itself is split into two parts: a
scrollable upper section with your annotation controls, and a fixed
status line at the bottom that shows the result of your last action.

## Annotating a cell with a unit

An **annotation** is how you tell Daubee what unit a cell represents.
Select a cell, a range, an entire row, or an entire column, then:

1. Type a unit formula into the **Unit formula** box (see below).
2. Click **Annotate selection**.

Whatever you select is what gets annotated — annotate a whole column and
every cell in that column is treated as carrying that unit (this is the
normal way to annotate a long column of data at once, rather than
annotating cell by cell).

The selection doesn't have to be contiguous: Ctrl-click (Cmd-click on a
Mac) several cells, ranges, rows, or columns and annotate them all at
once. Each selected area gets its own annotation, and the pane's list,
**Delete all units within selection**, Consolidate, Assert, and **Errors
for selection** all work across every selected area. An existing
annotation counts as "within" the selection only if a single selected
area covers all of it.

Once a selection has an annotation, the pane's **Units for selection**
list shows it, along with a **Delete** button to remove it. If your
selection contains several different annotations, they all show up in
the list; a **Delete all units within selection** button appears
above it whenever your selection covers more than one annotation, so you
can clear them out in one step.
Each unit in the list has a copy icon next to it, which copies the unit
formula to the clipboard.

## Writing a unit formula

A unit formula is built from unit names combined with a few operators:

| Symbol | Meaning | Example |
| --- | --- | --- |
| `-` | multiplication | `kg-m` (kilogram-meters) |
| `/` | division | `m/s` (meters per second) |
| `^` | exponent | `m^2` (square meters) |
| `(` `)` | grouping | `kg-m/(s^2)` |

A unit name is a letter followed by any mix of letters, digits,
underscores, and `$` — either a unit's own canonical name (`kilogram`) or one of
its aliases (`kg`). An exponent is a plain non-negative number (`m^2`),
or a negative number in parentheses (`m^(-1)`, meaning "per meter").
Spaces are fine anywhere you'd naturally put one — `kg - m / s^2` and
`kg-m/s^2` mean exactly the same thing; Daubee strips them out before
reading the formula.

A few real examples:

- `USD` — US dollars
- `kg-m/s^2` — a force (this is dimensionally the same as `newton`)
- `mi/hr` — miles per hour

If a unit some formula computes matches a *named* unit exactly (like
`newton` for `kg-m/s^2`), Daubee still checks it correctly even if you
never mention `newton` yourself — units are compared by their actual
dimensions, not by spelling.

## Browsing available units

Rather than memorize every unit's exact name, click the **Browse**
button next to the Unit formula box. This opens the **Available units**
dialog, which lists every unit Daubee currently knows about, organized
by category (Currencies, Excel-native units, and any custom dictionaries
you've loaded — see [Unit dictionaries](#unit-dictionaries)). Click any
unit's canonical name or alias to insert it at your cursor position in
the formula box; click one of the syntax tokens (`-`, `/`, `^`, `(`,
`)`) or a digit to insert that instead. This makes building a formula
like `kg-m/s^2` a matter of clicking rather than typing. The **⌫** button removes the last character of the
formula box, and **Clear** empties it.

## Infer Units

**Infer** looks at every formula on the current sheet and tries to fill
in units for cells you haven't annotated yet, based on cells around them
that *are* annotated. For example, if `A1` is annotated `USD` and a
formula elsewhere computes `=A1+B1`, Daubee knows `B1` must also be
`USD` for the addition to make sense — if `B1` has no annotation of its
own, Infer offers to add one.

Each offer shows up as a dialog: *"Infer 'USD' as the unit annotation
for cell B1?"* with **Add** and **Skip** buttons. The dialog also offers
**Add all** / **Skip all**, which answer every remaining offer in that
Infer run the same way, so you don't have to click through each one.

Infer only ever proposes annotations for cells that don't have a formula
of their own (a plain value cell) — it never guesses at what a *formula*
cell's own result should be. That's Check Units' job.

## Automation

The **Unit automation** section of the pane, above **Units for selection**,
holds actions that work out annotations for you. Each is switched on in
[Settings](#settings) (Automation section), and its button stays disabled
until it is.

### Scan headers

With *Scan column headers for units* and/or *Scan row headers for units*
ticked (both off by default), **Scan headers** looks for headers that name
a unit **inside parentheses or square brackets**, such as `Weight (kg)`,
`Price [USD]` or `Speed ( mi / hr )` (spaces inside the brackets are
ignored). For each one it offers to annotate the plain numbers below a
column header, or to the right of a row header, with that unit. The offer
names the header that prompted it and the range it would annotate, for
example: *Column header "Weight (kg)" (C1) suggests the unit "kg" for
C2:C50 (49 cells). Add this annotation?* — with the usual **Add all** /
**Skip all** buttons.

- **Any header counts, not just the first row or column.** A sheet with
  several sections, each with its own header row, gets an offer for each.
  The numbers under a header run until a text cell, or a gap of more than
  three blank cells in a row. Blank cells before the first number, or
  between numbers, don't end the run, and the blank cells between numbers
  are included in the annotated range.
- **Merged headers** (a header centered over several columns, or a row
  header spanning several rows) cover everything they span. One or two
  sub-header cells (such as "Min" and "Max") between the header and the
  numbers are skipped. A header that is only visually centered, with the
  neighbouring cells left empty rather than merged, covers just its own
  column.
- **Only plain numbers** are annotated: formulas are left to Check Units
  and Infer, and dates are skipped.
- **Nothing is overwritten.** Cells that already have an annotation are
  left alone, and the offer says how many. A cell that a column header and a
  row header assign *different* units is left out altogether, and the status
  message says how many such cells there were.
- Text such as `Item(s)` (a plural marker attached to a word) is not
  taken for a unit.
- **Number formats follow the usual rules.** For a currency header such as
  `(USD)`, Daubee offers a matching currency format for the cells still in
  the "General" format, exactly as when you annotate by hand (including the
  choice between the default and three-letter-code format, or applying it
  automatically if that Setting is on). For any other unit, the *Autoformat
  with* Setting applies its unit format, but only when every cell is still
  in "General". A cell already formatted for a *different* currency than
  its header is left out and mentioned in the offer, rather than being
  reformatted.

### Wrap dimensionless arguments

Some functions need an argument in one particular unit: `SIN` and the other
trig functions want radians, `DEGREES` wants radians, `RADIANS` wants
degrees, and the securities-pricing functions want the currency chosen in
Settings. Giving one a plain number is flagged by Check Units.
**Wrap dimensionless arguments** fixes that by rewriting just the
offending argument with [WITHUNIT](#withunit):

```
=SIN(A1*B1)   →   =SIN(WITHUNIT(A1*B1,"radian"))
```

Tick *Wrap dimensionless arguments with WITHUNIT* in [Settings](#settings)
(Automation section; off by default). The button then becomes available
once Check Units has flagged at least one such argument on the sheet. For
each one it shows the formula before and after and offers **Add** /
**Skip** (or **Add all** / **Skip all**). WITHUNIT must be defined in the
workbook, and Excel's Undo can't reverse the change, just as for
[Wrap dimensionless results](#wrap-dimensionless-results). Where a function
accepts "any currency", or a value of some unspecified unit, there is no
single unit to suggest, so nothing is offered.

- The same button also wraps a **plain number written in a formula next to
  a quantity**: in `=IF(A1>42, …)` with `A1` in miles, the `42` becomes
  `WITHUNIT(42,"mile")`. This applies to a comparison (including the
  conditions of `IFS`), to a number added to or subtracted from a quantity
  (`=A1+5`), to the values of `IF`, `IFS`, `IFNA`/`IFERROR` and `SWITCH`,
  and to the case values of `SWITCH`. The unit comes from the other side.
  Only numbers typed into the formula are wrapped; a cell with no
  annotation is better fixed by annotating it.
- A formula with several problems is fixed one at a time: run Check Units
  again after wrapping to see whether another remains.
- If you edit a formula after Check, the offer for it is skipped (the
  status message says so); run Check again.
- If the number really is in degrees, say, rather than radians, wrap it
  yourself with the right unit, or use `RADIANS`.

### Wrap dimensionless results

Some formulas give a plain number — `=RAND()`, say — in a cell you have
annotated with a unit such as `kg`. Check Units flags that as a mismatch,
because the value is dimensionless. **Wrap dimensionless results** fixes it
by rewriting the formula with [WITHUNIT](#withunit):

```
=RAND()   →   =WITHUNIT(RAND(),"kg")
```

Tick *Wrap dimensionless results with WITHUNIT* in [Settings](#settings)
(Automation section; off by default). The button then becomes available
once Check Units has flagged at least one such cell on the sheet. For each
one it shows the formula before and after and offers **Add** / **Skip**
(or **Add all** / **Skip all**). The unit used is the cell's own annotation.

- WITHUNIT must be defined in the workbook. Ticking the Settings option
  above defines it for you (it ticks *Define the WITHUNIT, CONVERTUNIT
  functions in this workbook* too); if the workbook has a differently
  defined `WITHUNIT`, Daubee leaves it alone, keeps the option off, and says
  so.
- This rewrites the formulas in your cells, and Excel's Undo can't reverse
  it.
- A formula that already starts with WITHUNIT is left alone, as is a
  formula that spills over several cells.
- If the number is dimensionless only because the cells it uses have no
  annotations, annotating *those* cells is usually the better fix.

## Check Units

**Check** walks every formula on the current sheet and verifies that
every operation is dimensionally consistent: you can't add feet to
kilograms, a `SUM` range has to share one unit throughout, and so on.

For any formula cell whose unit Check Units can work out for certain but
that doesn't yet carry that annotation, it offers to add it: *"Add
'USD' as the unit annotation for cell B4?"* with **Add** / **Skip**.
Accepting these offers is how a spreadsheet gradually ends up with its
formula cells annotated too, not just its raw inputs.

Cells that don't check out get a colored border instead of an offer —
see the next section.

Run Check Units again any time after editing formulas or annotations;
it only looks at the sheet's current state, so there's no harm in
running it as often as you like.

## Synthesizing units for CONVERT

Excel's `CONVERT` function accepts "from" and "to" unit codes with a
metric or binary multiplier prefix attached — `"mm"` (millimeters),
`"kPa"` (kilopascals), `"Kibyte"` (kibibytes), and so on. Daubee
recognizes these by decomposing the prefix and the base unit on the
fly — `mm` is understood as `m` (meter) with a milli prefix — rather
than treating every possible prefixed form as its own separately
listed unit. You can type a prefixed code like `mm` directly into the
Unit formula box and it resolves correctly, the same as typing `m`
itself; the only difference is that a prefixed unit never appears in
the [Available units](#browsing-available-units) dialog, and it's
always treated as its own distinct unit — `mm` is never considered
equal to `m`, the same way `kilogram` and `gram` are already two
separate units in Daubee.

A prefix only applies to a *metric* unit — matching Excel's own
documentation — so `mile`, `foot`, `year`, and other non-metric units
never take one, even though they're perfectly usable unprefixed. If you
type a prefix+unit combination Excel doesn't actually recognize, Excel
itself will reject it (the cell shows `#N/A`) the same as it would
without Daubee involved.

For example, `=CONVERT(B3,"mm","cm")` converts the millimeter value in
cell B3 to centimeters. If B3 isn't already annotated, running **Infer**
offers to annotate it `mm`; once B3 is annotated (`mm` or otherwise),
running **Check** offers to annotate the cell containing the `CONVERT`
formula with `cm`.

## Reading the borders

Daubee never changes a cell's value — it only draws a border around a
cell to show what it found. The border color tells you why:

| Color | Meaning |
| --- | --- |
| 🟦 Blue | A normal unit annotation — this is what the cell represents. |
| 🟥 Red | A genuine mismatch — the formula's operands don't agree on a unit (e.g. adding meters to seconds). |
| 🟧 Amber | An argument to a function needed a specific kind of unit (like a currency, or a dimensionless number) and didn't get one. |
| 🟨 Yellow | Unverifiable — Daubee can't determine a unit here at all (often because it depends on something else that's also unverifiable, or the formula uses a live spreadsheet value Daubee can't simulate). |
| 🟪 Purple | The cell's number format expects a specific currency that isn't a registered unit yet — see [Unit dictionaries](#unit-dictionaries). |

A cell with no border at all simply has nothing to check — it's a plain
number, or a formula whose result is genuinely dimensionless (a count, a
ratio, a boolean, ...).

Annotation borders are dashed and error/warning borders are double-line by
default; **Annotation borders** and **Error/warning borders** in
[Settings](#settings) let you pick *Dashed*, *Line* (a thick solid line),
or *Double-line* for each, and every border in the workbook is redrawn
right away.

If the blue borders are distracting, tick **Hide annotation borders**
(just below the **Annotate selection** button). The blue borders on
annotated cells disappear; untick it and they come back. Hiding them
doesn't touch the annotations themselves, and the error borders — red,
amber, yellow, and purple — always stay visible, since they point at
something you may need to fix. The setting is saved with the workbook.

## Errors for selection

Below the annotation list, the **Errors for selection** section explains,
in plain language, exactly why a selected cell has a red, amber, yellow,
or purple border — for example: *"Unit mismatch: expected kilogram, got
second."* or *"Unverifiable: depends on cell C4, which is itself
unverifiable."* This is usually the fastest way to understand what
Daubee is flagging without hunting through the formula yourself.

## Consolidate

If a selection contains several individually-annotated cells that all
share the same unit and sit in one contiguous block, **Consolidate**
replaces them with a single annotation covering the whole block. This is
just housekeeping — it doesn't change what Daubee thinks the cells'
units are, only how many separate annotation records exist behind the
scenes. Only annotations that already agree are merged; nothing is
silently changed to make them agree.

## Assert

Sometimes Daubee genuinely can't determine a cell's unit (a yellow
border) even though *you* know what it should be — maybe the formula
uses a function Daubee doesn't simulate, or a live spreadsheet value it
can't read ahead of time. If you've reviewed a yellow-bordered cell and
you're confident it's correct, select it (or a range of them) and click
**Assert**.

Assert doesn't create or change any annotation — it vouches for whatever
is already there: a cell that already has a manual annotation is trusted
as exactly that unit; a cell with no annotation at all is trusted as
plain dimensionless. The next time you run Check Units, an asserted
cell's yellow border clears instead of reappearing, unless something
about the formula actually changes.

## WITHUNIT

A formula that produces a plain number — `RANDBETWEEN(10,100)`, for
example — is dimensionless as far as Daubee can tell, so annotating it
with a unit is flagged as a mismatch. **WITHUNIT** gives such a value a
unit explicitly:

```
=WITHUNIT(RANDBETWEEN(10,100), "units")
```

The second argument is the unit, written as text — a single unit
(`"kg"`) or a unit formula (`"USD/units"`) — or a cell containing that
text. The first argument must be dimensionless; if it already has a unit,
Daubee flags an error rather than silently overriding it. The result then
carries the stated unit everywhere it's used, and Check Units offers to
add a matching annotation.

WITHUNIT isn't built into Excel. It's a named LAMBDA you define once per
workbook, which just returns its first argument, so your numbers don't
change. The easiest way is **Define the WITHUNIT, CONVERTUNIT functions in this
workbook** in [Settings](#settings) (off by default); unticking it removes
the names again. To define it yourself, open Formulas > Name Manager > New,
name it `WITHUNIT`, and set "Refers to" to
`=LAMBDA(value, unit, value)`. It needs a version of Excel with LAMBDA
(Microsoft 365 or Excel 2021 and later).

If you use WITHUNIT with a particular Excel function frequently, it may
be worth defining your own LAMBDA in the Name Manager that wraps it. For
example, define `SINRADIAN` as `=LAMBDA(t, SIN(WITHUNIT(t, "radian")))`.
Then `=SINRADIAN(A1)` does the same as `=SIN(WITHUNIT(A1, "radian"))`,
without writing WITHUNIT each time or putting the angle in an annotated
cell. As with WITHUNIT itself, the argument must be a plain, dimensionless
number. [CONVERTUNIT](#convertunit), below, is another example of this
pattern.

### CONVERTUNIT

**CONVERTUNIT(value, unit1, unit2)** converts a dimensionless number from
`unit1` to `unit2`, with Daubee checking the units:

```
=CONVERTUNIT(A1, "kg", "lbm")
```

It is shorthand for `CONVERT(WITHUNIT(A1, "kg"), "kg", "lbm")`: the result
carries `unit2`, and Daubee flags an error if the two units can't be
converted into one another, or if `value` already has a unit. (For a cell
that is already annotated, use `CONVERT` directly.) Like WITHUNIT it's a
named LAMBDA, and it's defined by the same Settings option. To define it
yourself, add a name `CONVERTUNIT` that refers to
`=LAMBDA(value, unit1, unit2, CONVERT(WITHUNIT(value, unit1), unit1, unit2))`,
after defining `WITHUNIT`.

## Currencies

Daubee understands currency units (`USD`, `EUR`, `Canadian_dollar`, ...)
as a special case, because Excel's own Currency and Accounting number
formats display a currency symbol that Daubee can cross-check against
your annotation:

- **Annotating a cell that's still in Excel's default "General" format**
  with a currency unit offers to apply a matching Currency number format
  for you (e.g. formatting the cell as `$42.99`). The offer lets you
  choose between the **Default currency format** (`$42.99`) and the
  **Three-letter currency code format** (`USD 42.99`). Declining just
  skips the formatting — your annotation is saved either way.
  Deleting the annotation later does **not** undo a currency format
  Daubee applied; the cell keeps it until you change it yourself (Daubee
  can't tell its own formatting from yours).
- **Annotating a cell whose format already shows a *different* currency**
  offers to rewrite the format to match your annotation instead of
  blocking the save outright. Declining this one *does* cancel the save,
  since saving without fixing it would leave a real mismatch behind.
- **A cell's number format naming a currency Daubee doesn't have seeded**
  gets a purple border and shows up in Errors for selection, rather than
  a silent block — see [Unit dictionaries](#unit-dictionaries) for how to
  add it.
- **If you later reformat an already-annotated cell to a different
  currency**, selecting that cell shows a small "Cell's format now shows
  ⟨currency⟩ — Update annotation" row with a button, so your annotation
  can be kept in sync without a surprise pop-up.
- **The same applies to an annotated range.** If every cell in the range
  is reformatted to a different currency, the pane offers **Update
  annotation**. If only some cells are, it says how many (for example, "2 of
  6 cells' formats show EUR, not USD") and offers **Reformat to match**,
  which sets just those cells back to the annotation's currency. The check
  covers the whole annotated range, not only the cells you selected.

The dollar currencies also have symbol aliases, so `US$`, `C$`, `A$`,
`NZ$`, `S$`, `HK$`, and `NT$` can be used for US, Canadian, Australian,
New Zealand, Singapore, Hong Kong, and New Taiwan dollars.
When "shortest alias" is in effect (see [Settings](#settings)), these are
the names Daubee prefers for those currencies.

Two related Settings (see below) control whether currency formatting is
applied automatically, and whether it uses a currency's symbol (`$`) or
its three-letter ISO code (`USD`).

**A formula's currency symbol doesn't prove its units are right.** When you
type a formula into a cell in the "General" format, Excel itself gives the
result the number format of the first cell the formula refers to. So
`=A1+A2`, with `A1` in US dollars and `A2` in euros, displays its sum with a
`$` — even though it adds two different currencies. Daubee doesn't do this
and doesn't hide it: run Check Units, which flags the cell.

## Settings

Click **Edit** (in the Daubee tab's Settings group) to open Daubee's
settings:

- **Apply currency format when adding a currency unit** — when checked,
  the "no format yet" currency offer above happens automatically instead
  of asking every time.
- **Use three-letter currency code formats by default** — when checked,
  a currency format shows, for example, `USD 42.99` instead of `$42.99`.
  This is the format used when currency formats are applied
  automatically, and the choice preselected in the offer dialog (where you
  can pick either format each time).
- **Save unit aliases** — when you annotate using an alias (`kg` instead
  of `kilogram`), this controls whether the alias itself or the
  canonical name is what actually gets saved.
- **Save inferred and checked units using the shortest alias** — when
  checked, a unit that Infer Units or Check Units offers to save (and the
  offer's own text) uses the unit's shortest alias (`kg`) instead of its
  canonical name (`kilogram`). If several names are equally short, one
  starting with a capital letter wins. Only names that can be typed into
  an annotation are used, so a symbol such as `Ω` falls back to the next
  shortest name.
- **Define the WITHUNIT, CONVERTUNIT functions in this workbook** —
  adds (or, when unticked, removes) the `WITHUNIT` and `CONVERTUNIT` named
  LAMBDAs described in [WITHUNIT](#withunit). Off by default. If the
  workbook already has a differently-defined name called `WITHUNIT` or
  `CONVERTUNIT`, Daubee leaves it alone.
- **Simplify using unit definitions** — when checked, saving an
  annotation that matches a named unit's own definition (e.g.
  `kilogram-meter/second^2`) automatically rewrites it to that unit's
  name (`newton`) instead of keeping the expanded form.
- **Autoformat with** — when checked, saving a unit annotation on a
  still-"General"-formatted, non-currency cell also applies a matching
  Custom number format showing the unit alongside the value (e.g. `0.00
  "kg"`). Choose whether the shown unit uses its canonical name or its
  shortest alias (a capitalized one, if several tie), and how many decimal
  digits to show.
- **Use a specific currency for securities-pricing functions** — controls
  what currency functions like `PRICE`, `PV`, and `YIELD` (click
  "securities-pricing functions" to see the full list) expect when
  nothing in the formula itself pins down a currency. Uncheck the box to
  let each use *any* currency instead (as long as everything involved
  agrees with itself); check it and pick a currency from the dropdown to
  pin them all to one specific currency, or leave it on "Locale default"
  to follow your computer's own regional settings.
- **Annotation borders** and **Error/warning borders** — the line style
  for the blue annotation borders and for the red, amber, yellow, and
  purple error/warning borders: *Dashed*, *Line*, or *Double-line*.
  Annotation borders default to dashed and error/warning borders to
  double-line.
- **Automation: Scan column headers for units** and **Scan row headers
  for units** — enable the pane's **Scan headers** button, which offers
  annotations taken from headers that name a unit in parentheses or square
  brackets; see [Scan headers](#scan-headers). Both are off by default.
- **Automation: Wrap dimensionless arguments with WITHUNIT** and **Wrap
  dimensionless results with WITHUNIT** — make the pane's **Wrap
  dimensionless arguments** and **Wrap dimensionless results** buttons
  available; see [Wrap dimensionless
  arguments](#wrap-dimensionless-arguments) and [Wrap dimensionless
  results](#wrap-dimensionless-results). Both are off by default. Ticking
  either also ticks *Define the WITHUNIT, CONVERTUNIT functions in this
  workbook* (they need `WITHUNIT`), and unticking that option unticks both.

## Unit dictionaries

Click **Select dictionaries** (Daubee tab, Units group) to control which
built-in unit groups Daubee makes available (for example, splitting
currencies into "major" and "other" groups, or scientific units into
several groups), and to add your own dictionaries:

- **Add from file...** loads a YAML file of custom unit
  definitions from your computer.
- **Add from URL** loads one from a web address instead — useful for
  sharing a shop- or team-wide dictionary without emailing files around.

The built-in **Trigonometric** group holds the angle units `degree`
(`deg`, `degrees`) and `radian` (`rad`, `radians`). Excel's trig functions
work in radians, so leave it on if you use `SIN`, `COS` and friends with
units (including [Wrap dimensionless
arguments](#wrap-dimensionless-arguments)); switching it off just
declutters the Browse list, and any annotation using those units is then
rejected as unknown.

A dictionary file is just a set of category names, each listing the
units that belong to it. Here's a small one, defining a new "shipping"
category with three units:

```yaml
shipping:
  - canonicalName: pallet
    aliases: [plt, pallets]
  - canonicalName: container
    aliases: [cntr, containers]
  - canonicalName: containers_per_hour
    aliases: [cph]
    expansion: container/hr
```

Each unit needs a `canonicalName` and an `aliases` list (use `[]` if it
has no aliases). Unit names are matched exactly: Daubee doesn't treat
`pallet` and `pallets`, or `plt` and `PLT`, as the same unit. List every
variant you expect to use — singular and plural, lowercase and uppercase —
as its own alias, as `pallets` is above. The optional `expansion` field
gives a unit its own definition in terms of others already registered —
built-in units work too, like `hr` above (Daubee's own built-in `newton`
is defined the same way, as `kg-m/s^2`). Once loaded, `pallet`/`plt`,
`container`/`cntr`, and `containers_per_hour`/`cph` are all usable in
annotations and show up in the
[Available units](#browsing-available-units) dialog like any built-in
unit.

**Names must be unique.** A unit name or alias can belong to only one unit.
If any name in your file is already used by a built-in unit, by a dictionary
you added earlier, or by another unit in the same file, Daubee doesn't load
the dictionary at all and says which names clash (for example: *"m" (an alias
of "furlong") is already used for "meter"*). Fix the file and add it again.
It never loads part of a dictionary. A dictionary already added that later
turns out to clash (say, after a built-in group is switched on) is left
out the same way, and the pane shows the message when you open Browse.

If a category in your file has the same name as one that already exists —
ignoring capitalization, spaces, hyphens, and underscores, so `Distance`,
`distance`, and `Weight and mass` / `weight_and_mass` each match — its units
are added to that category rather than starting a second one.

An `expansion` also makes the unit interchangeable with its definition
when Infer Units and Check Units compare units: `containers_per_hour`
and `container/hr` count as the same unit, and so do the built-in `watt`
and `joule/second`, or `kWh` (defined as `kW-hour`) and `kW-hour`. This is
about dimensions only — Daubee doesn't track scale factors, so a bare
number like the `100` in a cents-to-dollars conversion is invisible to it.
To state a conversion factor, put it in a cell and annotate that cell with
the unit ratio (a `100` annotated `cents/US_dollar`), or use
[WITHUNIT](#withunit) to give a number its unit inside a formula.

This is also where you'd add a currency Daubee doesn't already know
about, if you ever hit a purple "currency not found" border.

## Audit report

Click **Generate report** (Daubee tab, Audit group) to produce a PDF
summarizing how much Daubee has actually been used in this workbook:
how many unit annotations exist (broken down by whether they were typed
by hand, inferred, or added via Check Units), how many errors have been
detected, and how many cells have needed an Assert. These are
**cumulative, lifetime counts** — a running tally of everything Daubee
has ever found or created in this workbook, not a snapshot of what
exists right now, so the report reflects how much value the tool has
actually delivered over time even after annotations get cleaned up or
consolidated.

Choose whether to report on just the current sheet or the whole
workbook; a whole-workbook report includes both the combined totals and
a sheet-by-sheet breakdown table.

## Tips

- **A cell with no border isn't a problem** — it just means there's
  nothing for Daubee to check there (a plain value, or a genuinely
  dimensionless result).
- **Circular references can be analyzed.** If a set of formulas refer to
  one another, but something in the loop pins down a real unit,
  Check Units can work it out — it only falls back to "unverifiable" when
  nothing in the loop actually determines an answer.
- **Powers: `^` and `POWER`** — `=A1^2` and `=POWER(A1, 2)` mean the
  same thing, and raise the cell's unit to that power (`kg` becomes
  `kg^2`; `SUM(A1:A7)^2` works too). The exponent can be a number, a
  cell, or arithmetic on those (`A1^(B1/2)`), but not a function call;
  Daubee reads the cell's current value, so the result can change if that
  value does. The result has to have whole-number unit exponents:
  `(m^2)^0.5` is `m`, but `m^0.5` is flagged, since a unit formula can't
  express it. The exponent itself must be dimensionless, and a
  dimensionless value stays dimensionless whatever the power.
- **Named LAMBDA formulas from the Name Manager are understood too** —
  calling a named LAMBDA directly (`AddTax(A1)`), or passing one by
  reference into `REDUCE`, `SCAN`, `MAKEARRAY`, `MAP`, `BYCOL`, or
  `BYROW`, works the same as writing the LAMBDA inline.
- **Copying a sheet copies its annotations — with two conditions.**
  Excel doesn't tell Daubee which sheet a new one was copied from, so
  Daubee infers it from Excel's default copy name: a new sheet called
  `Sales (2)` is treated as a copy of `Sales`, and gets its own copy of
  that sheet's annotations, errors, and assertions. This only works if the
  Daubee pane is open when you make the copy, and if the copy still has
  its default name when it's created (renaming it afterward is fine). A
  copy made with the pane closed, or given a different name in the Move or
  Copy dialog, keeps the borders but not the annotations — re-annotate it,
  or copy the sheet again with the pane open.
- **When in doubt, run Check Units.** It's the fastest way to see whether
  your latest edit introduced a problem.
