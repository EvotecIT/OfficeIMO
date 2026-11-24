---
topic: about_PSWriteOffice_Notebook
schema: 1.0.0
---
# about_PSWriteOffice_Notebook

## Short Description

Open, edit, run, and export notebooks and their saved results.

## Long Description

Open-OfficeNotebook returns a live notebook. Use Get-OfficeNotebookCell and
Get-OfficeNotebookOutput to inspect its ordered cells and saved outputs.
Cell and output indexes start at zero. Move-OfficeNotebookCell takes the
destination's final index. Close-OfficeNotebook releases the live document.

Use Save-OfficeNotebook to write .ipynb. For an imported source, give the
destination explicitly. Use Export-OfficeNotebook with a live notebook to
write PDF, HTML, Markdown, Word, or PowerShell source with a conversion
report. Use -FailOnLoss to reject an export with fidelity losses. File
publishing commands support -WhatIf and -NoClobber.

Open-OfficeNotebook recognizes .ipynb, .qmd, .rmd, .nb.html, .deepnote, and
.zpln. Use -Format for Polyglot .dib, Jupytext percent scripts, ordinary
Markdown, PowerShell scripts and Zeppelin JSON. Supply -CompanionPath for
Jupytext if saved results should be retained from a verified .ipynb.
Supply -DeepnoteNotebookId to select a notebook in a multi-notebook file.

For Markdown export, -PreserveMarkdownSource keeps the spelling,
whitespace, HTML and links in Markdown cells. This mode reports omitted
notebook-local attachment data. Use the default export to embed it.

For HTML export, -IncludeInteractiveWidgets enables supported saved
controls and client-side links. Notebook cells and custom widget code
stay inert. The conversion report identifies unsupported saved views.

Use -IncludeInteractiveCharts with an HTML destination to zoom and pan
supported saved charts and inspect their data values. The charts work
offline. Unsupported specifications use an available static fallback.

The PowerShell script profile opens an ordinary .ps1 as one code cell.
Scripts exported by OfficeIMO include comment markers for code and
Markdown cells. The export report identifies content that a script cannot
carry, including saved results and cells in other languages. Notebook
import and export operate on stored data without executing cells.

Use Invoke-OfficeNotebook to run trusted PowerShell cells. Each notebook
runs in one child PowerShell session, so later cells can use variables and
functions defined by earlier cells. Select zero-based indexes with
-CellIndex. Use -IncludeUnspecifiedLanguage for cells without a language.

The report includes the updated notebook. Use -OutputPath to write its
saved outputs and execution counts to .ipynb. -StreamOutput emits live
records ahead of the report. By default, each cell has 60 seconds to run.
Use -TimeoutSeconds to change the limit. A timeout ends the session and
keeps the output captured before it stopped.

## Examples


```powershell
$report = Invoke-OfficeNotebook -LiteralPath ./lab.ipynb -OutputPath ./executed.ipynb -NoClobber
try {
$report.Cells
```

} finally {

```powershell
$report.Notebook | Close-OfficeNotebook
}
```

```powershell
$source = Open-OfficeNotebook -Path ./lab.py -Format JupytextPercent -CompanionPath ./lab.ipynb -FailOnLoss
try {
$source | Save-OfficeNotebook -Path ./reviewed.ipynb
```

} finally {

```powershell
$source | Close-OfficeNotebook
}
```

```powershell
$notebook = New-OfficeNotebook
try {
$notebook | Add-OfficeNotebookCell -Kind Markdown -Source '# Lab'
$notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source 'Get-Date'
$notebook | Export-OfficeNotebook -Path ./lab.ps1 -FailOnLoss
```

} finally {

```powershell
$notebook | Close-OfficeNotebook
}
```

```powershell
$notebook = Open-OfficeNotebook -Path ./lab.ipynb
try {
$notebook | Get-OfficeNotebookCell -Index 2 |
Get-OfficeNotebookOutput -Index 0 |
Export-OfficeNotebook -Path ./results.xlsx
```

} finally {

```powershell
$notebook | Close-OfficeNotebook
}
```

That last example exports a selected saved output containing an inline
structured data resource. Use a .csv or .xlsx destination.

## See Also

Open-OfficeNotebook
Invoke-OfficeNotebook
Get-OfficeNotebookCell
Get-OfficeNotebookOutput
Save-OfficeNotebook
Export-OfficeNotebook
Convert-OfficeDocument
