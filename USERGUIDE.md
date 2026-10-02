# Daubee User Guide

Daubee checks that the *units of measure* in your Excel formulas are
used consistently — the way spell-check catches typos, Daubee catches errors
like adding pounds to dollars, or a formula that stops meaning
what you think it means partway down a column. You tell it the unit a
cell represents (`kg`, `USD`, `meters/second`, ...), and Daubee tracks
that unit through every formula that touches it, flagging the formulas
where there's unit inconsistency.

## Contents

- [Opening the pane](#opening-the-pane)
- [Annotating a cell with a unit](#annotating-a-cell-with-a-unit)
- [Writing a unit formula](#writing-a-unit-formula)
- [Browsing available units](#browsing-available-units)
- [Infer Units](#infer-units)
- [Check Units](#check-units)
- [Synthesizing units for CONVERT](#synthesizing-units-for-convert)
- [Reading the borders](#reading-the-borders)
- [Errors for selection](#errors-for-selection)
- [Consolidate](#consolidate)
- [Assert](#assert)
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

Once a selection has an annotation, the pane's **Units for selection**
list shows it, along with a **Delete** button to remove it. If your
selection contains several different annotations, they all show up in
the list; a **Delete all units within selection** button appears
whenever your selection covers more than one annotation, so you can
clear them out in one step.

## Writing a unit formula

A unit formula is built from unit names combined with a few operators:

| Symbol | Meaning | Example |
| --- | --- | --- |
| `-` | multiplication | `kg-m` (kilogram-meters) |
| `/` | division | `m/s` (meters per second) |
| `^` | exponent | `m^2` (square meters) |
| `(` `)` | grouping | `kg-m/(s^2)` |

A unit name is a letter followed by any mix of letters, digits, and
underscores — either a unit's own canonical name (`kilogram`) or one of
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
like `kg-m/s^2` a matter of clicking rather than typing.

## Infer Units

**Infer** looks at every formula on the current sheet and tries to fill
in units for cells you haven't annotated yet, based on cells around them
that *are* annotated. For example, if `A1` is annotated `USD` and a
formula elsewhere computes `=A1+B1`, Daubee knows `B1` must also be
`USD` for the addition to make sense — if `B1` has no annotation of its
own, Infer offers to add one.

Each offer shows up as a dialog: *"Infer 'USD' as the unit annotation
for cell B1?"* with **Add** and **Skip** buttons. If more than one cell
qualifies for the same inferred unit, the dialog also offers **Add all**
/ **Skip all**, so you don't have to click through every one individually.

Infer only ever proposes annotations for cells that don't have a formula
of their own (a plain value cell) — it never guesses at what a *formula*
cell's own result should be. That's Check Units' job.

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

Daubee never changes a cell's value — it only draws a double-line border
around a cell to show what it found. The border color tells you why:

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

## Currencies

Daubee understands currency units (`USD`, `EUR`, `Canadian_dollar`, ...)
as a special case, because Excel's own Currency and Accounting number
formats display a currency symbol that Daubee can cross-check against
your annotation:

- **Annotating a cell that's still in Excel's default "General" format**
  with a currency unit offers to apply a matching Currency number format
  for you (e.g. formatting the cell as `$42.99`). Declining just skips
  the formatting — your annotation is saved either way.
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

Two related Settings (see below) control whether currency formatting is
applied automatically, and whether it uses a currency's symbol (`$`) or
its three-letter ISO code (`USD`).

## Settings

Click **Edit** (in the Daubee tab's Settings group) to open Daubee's
settings:

- **Apply currency format when adding a currency unit** — when checked,
  the "no format yet" currency offer above happens automatically instead
  of asking every time.
- **Use three-letter currency codes for applied currency formats** —
  when checked, an automatically-applied currency format shows `USD
  42.99` instead of `$42.99`.
- **Save unit aliases** — when you annotate using an alias (`kg` instead
  of `kilogram`), this controls whether the alias itself or the
  canonical name is what actually gets saved.
- **Simplify using unit definitions** — when checked, saving an
  annotation that matches a named unit's own definition (e.g.
  `kilogram-meter/second^2`) automatically rewrites it to that unit's
  name (`newton`) instead of keeping the expanded form.
- **Autoformat with** — when checked, saving a unit annotation on a
  still-"General"-formatted, non-currency cell also applies a matching
  Custom number format showing the unit alongside the value (e.g. `0.00
  "kg"`). Choose whether the shown unit uses its canonical name or its
  shortest alias, and how many decimal digits to show.
- **Use a specific currency for securities-pricing functions** — controls
  what currency functions like `PRICE`, `PV`, and `YIELD` (click
  "securities-pricing functions" to see the full list) expect when
  nothing in the formula itself pins down a currency. Uncheck the box to
  let each use *any* currency instead (as long as everything involved
  agrees with itself); check it and pick a currency from the dropdown to
  pin them all to one specific currency, or leave it on "Locale default"
  to follow your computer's own regional settings.

## Unit dictionaries

Click **Select dictionaries** (Daubee tab, Units group) to control which
built-in unit groups Daubee makes available (for example, splitting
currencies into "major" and "other" groups, or scientific units into
several groups), and to add your own dictionaries:

- **Add from file...** loads a YAML file of custom unit
  definitions from your computer.
- **Add from URL** loads one from a web address instead — useful for
  sharing a shop- or team-wide dictionary without emailing files around.

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
has no aliases). The optional `expansion` field
gives a unit its own definition in terms of others already registered —
built-in units work too, like `hr` above (Daubee's own built-in `newton`
is defined the same way, as `kg-m/s^2`). Once loaded, `pallet`/`plt`,
`container`/`cntr`, and `containers_per_hour`/`cph` are all usable in
annotations and show up in the
[Available units](#browsing-available-units) dialog like any built-in
unit.

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
- **Circular references aren't automatically hopeless.** If a pair of
  formulas refer to each other (`A1=B1+1`, `B1=A1+1`) but something else
  in the loop pins down a real unit, Check Units can still work it out —
  it only falls back to "unverifiable" when nothing in the loop actually
  determines an answer.
- **Named LAMBDA formulas from the Name Manager are understood too** —
  calling a named LAMBDA directly (`AddTax(A1)`), or passing one by
  reference into `REDUCE`, `SCAN`, `MAKEARRAY`, `MAP`, `BYCOL`, or
  `BYROW`, works the same as writing the LAMBDA inline.
- **When in doubt, run Check Units.** It's safe to run as often as you
  like, and it's the fastest way to see whether your latest edit
  introduced a real problem.
