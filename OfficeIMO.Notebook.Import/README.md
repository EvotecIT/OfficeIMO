# OfficeIMO.Notebook.Import

Import notebook source and saved results into the editable Jupyter model.
Select the input profile explicitly when opening a path, stream, or bytes.

```csharp
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Import;

var imported = NotebookImporter.Load("report.qmd", NotebookImportFormat.Quarto);
using var notebook = imported.Value;
foreach (var finding in imported.Report.Diagnostics) {
    Console.WriteLine(finding.Message);
}
notebook.Save("report.ipynb", new NotebookSaveOptions { AllowLoss = true });
```

The returned document can also be used with the focused notebook PDF, HTML,
Markdown, Word, and spreadsheet adapters. The input stream remains open. The
example prints findings and explicitly accepts them when saving.

## Install

```sh
dotnet add package OfficeIMO.Notebook.Import
```

## Input profiles

| Profile | Input | Imported content |
| --- | --- | --- |
| `PolyglotDib` | .NET Polyglot `.dib` | Code and Markdown cells with declared kernels and aliases. The first `#!meta` block may follow earlier submissions and declares aliases for subsequent cells. Unsupported directives remain in source. |
| `JupytextPercent` | Explicit Python, R, or Julia percent script | Code, Markdown, raw cells, supported cell markers and metadata. Optional saved results from a verified Jupyter companion. |
| `PowerShellScript` | Explicit PowerShell `.ps1` | Plain script source as one code cell. Scripts written by OfficeIMO also restore their marked code and Markdown cells. |
| `Quarto` | `.qmd` | Markdown, front matter and top-level executable fences. Chunk options are stored as metadata. |
| `MySTMarkdown` | Explicit MyST notebook Markdown | Prose, front matter and top-level `{code-cell}` fences. Illustrative code fences remain prose. |
| `RMarkdown` | `.Rmd` | Prose, front matter and top-level executable chunks. Chunk options are stored as metadata. |
| `RNotebookHtml` | Rendered `.nb.html` with embedded R Markdown source | Source plus saved console text, supported inline PNG/JPEG and static HTML tables. Unmatched output is kept in a separate cell and reported. |
| `Deepnote` | `.deepnote` or `.snapshot.deepnote` version 1.0.0 | Supported text blocks, Python and SQL source, embedded images, and saved Jupyter MIME results. Select a notebook by exact ID when the project contains multiple notebooks. |
| `Zeppelin` | `.zpln` or explicitly selected Zeppelin JSON | Ordered paragraphs and interpreter names, saved text, errors, HTML, images and rectangular tables. Interactive forms, Angular data and graph state may be omitted with diagnostics. |

Use `NotebookDocument.Load` for an existing `.ipynb`. All other profiles
produce a normalized `NotebookDocument`. `NotebookImporter.TryGetFormatFromPath`
can recognize notebook-specific extensions such as `.qmd` and `.nb.html`.
Scripts, ordinary Markdown, `.dib`, and vendor JSON require an explicit
profile. Conversion reads local input and
embedded results as data. Kernel execution, remote resource lookup, publisher
rendering and native source writing are outside the import profiles. Exporters
apply their own static MIME and resource policies.

`NotebookImportOptions.Limits` controls input size, cell and output counts and
parser limits. Review `imported.Report.Diagnostics` for unsupported structures
and approximations. The imported model passes those diagnostics to the standard
exporters. Call `imported.ThrowIfLoss()` to reject a lossy import before
writing an artifact.

## Pair a Jupytext script with saved results

```csharp
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Import;

using var companion = await NotebookDocument.LoadAsync("analysis.ipynb");
var imported = await NotebookImporter.LoadAsync(
    "analysis.py",
    NotebookImportFormat.JupytextPercent,
    new NotebookImportOptions { PairedNotebook = companion });
using var notebook = imported.Value;
imported.ThrowIfLoss();
notebook.Save("updated.ipynb");
```

Pairing checks cell order, exact source after line-ending normalization,
cell types, effective code languages, and explicit IDs where supplied.
Recognized leading language magics such as `%%R` identify mixed-language cells.
Repeated code with saved results requires matching explicit cell IDs. A mismatch
returns the script cells and a finding. The caller selects and owns the
companion. OfficeIMO does not search for sidecars.

## Exchange PowerShell source

```csharp
var imported = NotebookImporter.Load("report.ps1", NotebookImportFormat.PowerShellScript);
using var notebook = imported.Value;
notebook.Save("report.ipynb");
NotebookPowerShellScript.Save(notebook, "roundtrip.ps1", failOnLoss: true);
```

The script writer keeps PowerShell source in code cells and writes Markdown
cells as line comments. Its versioned comment boundaries preserve the IDs and
order of supported cells. It reports saved output, execution state, unsupported
metadata, or non-PowerShell cells that the script cannot carry. Pass
`failOnLoss: true` to reject those omissions before file publication. The
exchange reads and writes source without executing code. Execution as a
combined script depends on how the original cells were authored.

## Shared file conversion

The shared converter, workflows, and tool recognize `.qmd`, `.rmd`,
`.deepnote`, and `.zpln` as notebook sources for `.ipynb` and PDF output.

```sh
officeimo convert report.qmd report.ipynb
officeimo convert report.qmd report.pdf
```

The direct PDF route uses the imported title when neither notebook title nor
PDF metadata title was supplied.
Use the explicit importer for format choices involving ordinary Markdown,
script, bitmap, HTML or JSON extensions, paired outputs, and notebook selection
inside a Deepnote project.

See the [notebook support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Notebook/SUPPORT.md) for
available exporters and the [benchmark suite](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Notebook.Benchmarks/README.md)
for reproducible import and readback checks.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Notebook.Import` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
