# Daubee

Daubee is an Excel add-in that checks the *units of measure* in your
formulas for consistency — the way spell-check catches typos, Daubee
catches things like adding euros to dollars, or adding meters to
seconds.

- **[User guide](USERGUIDE.md)** — how to annotate cells, Infer Units vs.
  Check Units, currencies, Settings, and everything else.
- **[Issues](../../issues)** — found a bug, or have a feature request?
  File it here.

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
  MY ADD-INS, click Upload My Add-in, and select the manifest file.

## About this repo

This repo holds the built Daubee add-in and its documentation.
