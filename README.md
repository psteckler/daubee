# Daubee

Daubee is an Excel add-in that checks the *units of measure* in your
formulas for consistency — the way spell-check catches typos, Daubee
catches things like adding euros to dollars, or adding meters to
seconds.

- **[User guide](USERGUIDE.md)** — how to annotate cells, Infer Units vs.
  Check Units, currencies, Settings, and everything else.
- **[Issues](../../issues)** — found a bug, or have a feature request?
  File it here.
- **[Screenshots](screenshots/)** — what Daubee looks like in Excel.
- **[Unit dictionaries](#unit-dictionaries)** — ready-made sets of units you
  can add to Daubee: SI derived, imperial, US customary, electrical, time and
  calendar, chemistry, particle physics, astronomy, medicine, radiation and
  photometry, typography, data rates and communications, and counting and
  paper, energy and fuel, and cryptocurrencies.

## Screenshots

- [Blue borders mark annotated cells; a red border marks a unit mismatch](screenshots/USD-CAD-mismatch.png)
- [The pane explains the mismatch under "Errors for selection"](screenshots/USD-CAD-error.png)
- [Browsing the available units](screenshots/browse-units.png)
- [Choosing unit dictionaries](screenshots/unit-dictionaries.png)
- [An audit report](screenshots/audit-report.png)
- [Editing settings](screenshots/settings.png)

## Installation

On all platforms, you'll first need the Daubee **[manifest file](manifest.xml)**.
To download it, open the link, then click the download icon (or right-click
the **Raw** button and choose Save As).

- **Windows**

  Create a folder on a local drive, for example `C:\Daubee`, or on a network
  drive. Save the manifest file in that folder.

  In Excel, go to File > Options > Trust Center > Trust Center Settings >
  Trusted Add-in Catalogs. Enter the UNC path to the folder in "Catalog Url".
  For the example above, the path would be `\\LOCALHOST\C$\Daubee`. Click
  "Add catalog", tick "Show in Menu", and click OK.

  If that path doesn't work on your machine (for example, if the `C$` admin
  share is disabled), share the folder instead: right-click it in File
  Explorer, choose Properties > Sharing > Share, and share it with yourself.
  Then use the shared folder's network path, such as `\\YOURPC\Daubee`, in
  "Catalog Url".

  Close and restart Excel. Go to Home > Add-ins > More Add-ins, choose the
  Shared Folder tab, and add the Daubee add-in.

- **Mac**

  Open Finder. From the menu bar, choose Go > Go to Folder... (or press
  Cmd + Shift + G). Paste this path, replacing `<username>` with your macOS
  short username:

  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`

  If the `wef` folder doesn't exist, create it, naming it exactly `wef`.
  Copy the manifest file into the `wef` folder.

  Open or restart Excel and open any document. Go to the Home or Insert tab,
  click Add-ins, and choose the Daubee add-in to load it.

- **Excel on the Web**

  Open Excel online, go to Home > Add-ins > More Add-ins, select
  MY ADD-INS, then Manage My Add-ins, click Upload My Add-in, and select the
  manifest file.

## Unit dictionaries

Daubee comes with built-in units from Excel's `CONVERT` list and built-in
currency units. This repository's [`unit-dictionaries`](unit-dictionaries/)
folder has more, as ready-made **unit dictionaries** you can add when you need them:

|  |  |  |
| --- | --- | --- |
| [astronomy-units.yaml](unit-dictionaries/astronomy-units.yaml) | [chemistry-units.yaml](unit-dictionaries/chemistry-units.yaml) | [counting-and-paper-units.yaml](unit-dictionaries/counting-and-paper-units.yaml) |
| [cryptocurrency-units.yaml](unit-dictionaries/cryptocurrency-units.yaml) | [data-rates-and-communications-units.yaml](unit-dictionaries/data-rates-and-communications-units.yaml) | [electrical-units.yaml](unit-dictionaries/electrical-units.yaml) |
| [energy-and-fuel-units.yaml](unit-dictionaries/energy-and-fuel-units.yaml) | [imperial-units.yaml](unit-dictionaries/imperial-units.yaml) | [medicine-units.yaml](unit-dictionaries/medicine-units.yaml) |
| [particle-physics-units.yaml](unit-dictionaries/particle-physics-units.yaml) | [radiation-and-photometry-units.yaml](unit-dictionaries/radiation-and-photometry-units.yaml) | [SI-derived-units.yaml](unit-dictionaries/SI-derived-units.yaml) |
| [time-and-calendar-units.yaml](unit-dictionaries/time-and-calendar-units.yaml) | [typography-units.yaml](unit-dictionaries/typography-units.yaml) | [US-customary-units.yaml](unit-dictionaries/US-customary-units.yaml) |

**Load a dictionary from its URL.** In Excel, in the Daubee tab, click **Select dictionaries**,
then **Add from URL**, paste the address of the dictionary, and press the button. The URL is
`https://psteckler.github.io/daubee/unit-dictionaries/` followed by the file's
name.

**Or download it and load it as a file.** Open the dictionary's link in the
table above, then click the download icon (or right-click the **Raw** button
and choose Save As). In the Unit dictionaries dialog, click
**Add from file...** and choose the saved file.

## Research background

Daubee relies on techniques developed for XeLda, a research prototype
for checking units of measure in spreadsheets, as described in this paper:
[Antoniu, Steckler, Krishnamurthi, Neuwirth, Felleisen. "Validating the Unit Correctness of Spreadsheet Programs." ICSE 2004.](https://cs.brown.edu/people/sk/Publications/Papers/Published/asknf-valid-unit-sprdsht/paper.pdf)

## About this repo

This repo holds the built Daubee add-in and its documentation.
