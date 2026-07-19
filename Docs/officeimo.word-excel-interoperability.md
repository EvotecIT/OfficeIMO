# Word and Excel interoperability

OfficeIMO uses one workflow for current Open XML packages and legacy binary
documents: constrain the input, inspect what is present, preflight the intended
operation, and let guarded save or conversion APIs reject known loss.

This guide covers `.docx`, `.docm`, `.doc`, `.xlsx`, `.xlsm`, `.xls`, and
`.xlsb`. It describes the current contract, including the places where OfficeIMO
preserves content but does not offer a complete editing model.

## Format contract

| Format | Load | Edit and save | Important boundary |
| --- | --- | --- | --- |
| DOCX/DOCM | Supported | Supported for the documented Word model | Advanced package parts may be preserved rather than editable. Signed packages are blocked on save unless signature invalidation is explicitly allowed. |
| DOC | Supported subset | Native first-party writer for the supported subset; DOCX conversion is available | Import reports identify unsupported and preserve-only binary features. Loss is blocked by default. |
| XLSX/XLSM | Supported | Supported for the documented Excel model | Feature preflight distinguishes editable, partially editable, preserved, and unsupported package content. |
| XLS | Supported BIFF8 subset | Native first-party writer for the supported subset; XLSX conversion is available | Import reports identify non-projected BIFF records and legacy-only features. Loss is blocked by default. |
| XLSB | Supported BIFF12 subset | New-workbook writing and preservation-aware native rewrite for supported mutations; XLSX conversion is available | Unknown records and unmodified package parts are preserved. Unsupported mutations and save-time transforms fail before output is written. |

"Preserved" is intentionally different from "editable." For example, a macro
project, embedded package, control, external relationship, or unknown binary
record can survive a round trip even when OfficeIMO does not expose its complete
application-level object model.

## Open untrusted files with an explicit policy

Normal loads retain compatibility behavior when `PackageSecurity` is omitted.
For files received from users, mail, or external systems, select a package policy
explicitly:

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Word;

var security = OfficePackageSecurityOptions.UntrustedDefaults;

using ExcelDocument workbook = ExcelDocument.Load(
    "incoming.xlsb",
    new ExcelLoadOptions {
        AccessMode = DocumentAccessMode.ReadOnly,
        PackageSecurity = security
    });

using WordDocument document = WordDocument.Load(
    "incoming.docx",
    new WordLoadOptions {
        AccessMode = DocumentAccessMode.ReadOnly,
        PackageSecurity = security
    });
```

`UntrustedDefaults` enforces package, part, aggregate-size, and compression-ratio
limits and rejects macros, embedded payloads, ActiveX, and external
relationships. `SecureDefaults` applies the structural limits while allowing
those content classes. Both Word and Excel use the same policy before parsing
Open XML ZIP packages or legacy compound binary files.

Use `OfficePackageSecurityInspector.Inspect(...)` when you need an inventory
without opening a Word or Excel model. `Validate(...)` throws a typed
`OfficePackageSecurityException`; its `Rule`, `PartName`, `ObservedValue`, and
`Limit` properties are suitable for logs and rejection responses.

## Preflight the operation, not just the file

A file can be safe to read while still being unsafe to restructure, bind as a
template, render, or save without loss. Feature reports express that distinction:

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Word;

using ExcelDocument workbook = ExcelDocument.Load("model.xlsm");
ExcelFeatureReport excelReport = workbook.InspectFeatures();
excelReport.EnsureCan(ExcelPreflightCapability.ReadWorkbookData);

using WordDocument document = WordDocument.Load("contract.docm");
WordFeatureReport wordReport = document.InspectFeatures();
wordReport.EnsureCan(WordPreflightCapability.RenderFixedLayout);

File.WriteAllText("excel-features.md", excelReport.ToMarkdown());
File.WriteAllText("word-features.md", wordReport.ToMarkdown());
```

Use `Can(...)` for routing, `GetCapabilityDiagnostics(...)` for a user-facing
decision, `GetRepairHints(...)` for actionable alternatives, and `EnsureCan(...)`
at the execution boundary. The reports keep signature-safe reads separate from
edits that would invalidate signatures and distinguish content edits from
structure or template operations that can disturb preserve-only parts.

## Work with XLSB without disguising XLSX bytes

XLSB detection uses package content, not only the file extension. The importer
projects supported workbook, worksheet, value, formula, style, date, geometry,
view, merge, calculation, defined-name, and hyperlink records into the normal
Excel model. It records unsupported BIFF12 records in
`XlsbImportDiagnostics` and `XlsbPreservedRecords`.

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb;

using ExcelDocument workbook = ExcelDocument.Load(
    "source.xlsb",
    new ExcelLoadOptions {
        XlsbImportOptions = new XlsbImportOptions {
            MaxCells = 2_000_000,
            MaxSharedStrings = 500_000
        }
    });

workbook["Data"].CellValue(2, 2, 1250m);
workbook.Save("edited.xlsb");

ExcelDocument.Convert("source.xlsb", "editable.xlsx");
```

An unchanged XLSB source can be copied byte-for-byte. A supported cell edit uses
the native BIFF12 rewrite path while retaining other package parts. New
workbooks can also be saved as `.xlsb` through the first-party writer. If the
requested mutation or save-time transform is outside that writer's contract,
OfficeIMO throws before creating mislabeled or partially rewritten output.

## Import or convert DOC and XLS with an approved loss decision

Use the normal `Load(...)` surface when you only need the projected model. Use
the report-bearing surface when conversion or archival fidelity matters:

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;

using LegacyXlsLoadResult xls = ExcelDocument.LoadLegacyXlsWithReport("book.xls");
xls.EnsureNoImportErrors();
File.WriteAllText("book.import-report.md", xls.CreateImportReport().ToMarkdown());
xls.Document.Save("book.xlsx");

using LegacyDocLoadResult doc = WordDocument.LoadLegacyDocWithReport("letter.doc");
doc.EnsureNoImportErrors();
File.WriteAllText("letter.import-report.md", doc.ImportReport.ToMarkdown());
doc.Document.Save("letter.docx");
```

Native DOC/XLS saves and `WordDocument.Convert(...)` /
`ExcelDocument.Convert(...)` block known conversion loss by default. Set the
corresponding `LossPolicy` to `Allow` only after recording and accepting the
report. This opt-in is a policy decision; it does not make an unsupported feature
convertible.

## Inventory macros and embedded payloads

Macro and embedded-object APIs operate on bytes without executing VBA, OLE, or
ActiveX content:

```csharp
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("automation.docm");

OfficeVbaProjectInfo? vba = document.InspectVbaProject(includeSha256: true);
IReadOnlyList<OfficeEmbeddedPayloadInfo> payloads =
    document.GetEmbeddedPayloads(includeSha256: true);

if (payloads.Count > 0) {
    document.SaveEmbeddedPayload(payloads[0].Id, "payload.bin", maxBytes: 16 * 1024 * 1024);
}
```

Word and Excel can inspect, extract, replace, and remove embedded package/OLE/
ActiveX payloads by package-local id. They can attach, inspect, extract, and
remove VBA projects. OfficeIMO does not execute VBA, edit VBA source modules, or
sign VBA projects.

## Extend formula calculation deliberately

Excel preserves unsupported formulas for Excel-compatible applications. For
application-owned functions, register a bounded evaluator instead of rewriting
the workbook formula:

```csharp
using OfficeIMO.Excel;

using ExcelDocument workbook = ExcelDocument.Load("model.xlsx");
workbook.Calculation.RegisterCustomFunction("DOUBLEVALUE", (_, arguments) =>
    arguments.Count == 1 && arguments[0].Kind == ExcelFormulaValueKind.Number
        ? ExcelFormulaValue.FromNumber(arguments[0].Number * 2d)
        : null);

int calculatedCells = workbook.Calculate();
```

The evaluator writes cached results for supported functions and leaves the
original formula text intact. Inspection reports unsupported formulas and
dependency issues; save options can request Office to recalculate the remainder
on open.

## Render Word pages in batches

Word image export produces dependency-free PNG or SVG previews. Preflight fixed
layout first when the document may contain linked images, unmaterialized
alternative content, SmartArt, or ActiveX:

```csharp
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("report.docx");
document.InspectFeatures().EnsureCan(WordPreflightCapability.RenderFixedLayout);

IReadOnlyList<OfficeImageExportResult> pages = document.ToImages()
    .FromPage(0)
    .TakePages(3)
    .AsPng()
    .Save("previews");
```

The renderer estimates pagination; it is not Microsoft Word's layout engine.
Use its diagnostics and approved visual baselines for workflows where exact page
fidelity matters.

## Executable compatibility evidence

The checked-in corpus currently inventories 30 binary artifacts across
Microsoft Word, Microsoft Excel, Apache POI, and OpenPreserve sources:

- 24 XLS compatibility and diagnostic files with approved import reports;
- five Microsoft Excel-authored XLSB workbooks;
- one Microsoft Word COM-authored DOC file, complemented by generated binary DOC
  parser/writer contracts.

[`corpus-manifest.json`](../OfficeIMO.TestAssets/Documents/OfficeInteroperabilityCorpus/corpus-manifest.json)
locks artifact identity, provenance, purpose, and required reports. The tests
reject changed, missing, duplicate, and untracked binaries. Run the same focused
gate used by CI with:

```powershell
./Build/Test-OfficeInteroperabilityGate.ps1 -Suite Full
```

The gate covers corpus identity and load behavior, approved DOC/XLS report drift,
Excel-authored XLSB projection/native rewrite/XLSX conversion, shared package
security, preservation-sensitive package parts, macros and embedded payloads,
threaded comments, custom formula functions, feature preflight, and Word batch
rendering.
