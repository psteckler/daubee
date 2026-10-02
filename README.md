# Daubee

Daubee is an Excel add-in that checks the *units of measure* in your
formulas for consistency — the way spell-check catches typos, Daubee
catches things like adding pounds to dollars, or a formula that quietly
stops meaning what you think it means partway down a column.

- **[User guide](USERGUIDE.md)** — how to annotate cells, Infer Units vs.
  Check Units, currencies, Settings, and everything else.
- **[Issues](../../issues)** — found a bug, or have a feature request?
  File it here.

## Installing the add-in

1. Download [`manifest.xml`](manifest.xml) from this repo.
2. In Excel: **Insert > My Add-ins > Upload My Add-in**, and select the
   file you downloaded.
3. A new **Daubee** tab appears on the ribbon.

This works on Excel for Windows, Mac, and the web.

## About this repo

This repo holds the built add-in and its documentation — the published,
user-facing side of the project. The add-in's own source isn't kept
here.
