# OfficeIMO.Pdf.Testing

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Pdf.Testing)](https://www.nuget.org/packages/OfficeIMO.Pdf.Testing)

Snapshot tests for PDFs. `PdfSnapshot.Verify` renders every page with OfficeIMO's managed renderer, reads the logical text and structure, and compares both with an approved baseline. It works in xUnit, NUnit, MSTest or any other framework: a mismatch throws `PdfSnapshotMismatchException`, which the framework reports as a failed test.

## Install

```sh
dotnet add package OfficeIMO.Pdf.Testing
```

Targets `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.

## Example

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Testing;

[Fact]
public void InvoiceLayout() {
    var options = new PdfCreateOptions(); // the same font file on every machine
    options.RegisterFontFamily(PdfStandardFont.Helvetica, PdfEmbeddedFontFamily.FromFiles("Inter", "Fonts/Inter-Regular.ttf"));
    PdfDocument invoice = PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Invoice 2026-118")
        .AddTable(new[] {
            new[] { "Item", "Amount" },
            new[] { "Support", "900.00" },
            new[] { "Total", "1,250.00" }
        })), options);

    PdfSnapshot.Verify(invoice, "invoice");
}
```

`Verify` also accepts `byte[]` and `Stream`, so PDFs produced by Word, Excel, HTML or Markdown conversions can be checked the same way.

## Approve a snapshot

The first run fails because nothing has been approved yet. It writes the received files to a `__pdfsnapshots__` folder next to the test source file:

```text
__pdfsnapshots__/
  invoice.received.structure.json
  invoice.received.page-001.png
```

Check them, then approve them in one of two ways:

- run the tests again with `OFFICEIMO_PDF_SNAPSHOT_ACCEPT=1`, which approves every verified snapshot instead of comparing it;
- call `PdfSnapshot.Accept(directory)`, which approves the received files waiting in that folder.

Approved files drop the `.received` part (`invoice.structure.json`, `invoice.page-001.png`). Commit them with the tests and add `*.received.*` to `.gitignore`.

## What a mismatch reports

Each run compares, in this order:

1. the PDF bytes, when `Pdf` is `Store` (identical bytes pass without rendering) or `Exact` (any byte change fails);
2. `name.structure.json`: page count, page sizes and rotation, logical text per page as Markdown, embedded fonts, outline, form field names and render diagnostics. This check is exact;
3. page pixels, using `PdfVisualComparer.CompareImages` with the tolerances and ignored regions from `Visual`.

When anything differs, the exception message shows a line diff of the structure file, the pages whose pixels changed and where, the received files, and `name.received.html`, a self-contained gallery with approved, received and highlighted pages:

```text
PDF snapshot 'invoice' does not match the approved baseline.
Structure differs (invoice.structure.json):
@@ approved line 16 @@
          "| --- | ---: |",
          "| Support | 900.00 |",
-         "| Total | 1,250.00 |"
+         "| Total | 1,256.00 |"
        ]
      }
Page 1 pixels differ: 19 changed (0.001%) within x=1047 y=298 w=9 h=7.
Review: /src/Tests/__pdfsnapshots__/invoice.received.html
```

`PdfSnapshotMismatchException.ReceivedFiles` and `ReviewPath` hold the same paths, for example to attach them to a CI test result. A later passing run removes stale received files.

## Options

```csharp
var options = new PdfSnapshotOptions {
    Directory = "Snapshots",          // default: __pdfsnapshots__ beside the test file
    Pdf = PdfSnapshotPdfMode.Store    // also keep the approved PDF
};
options.Visual.IgnoredRegions.Add(new PdfPixelRegion(0, 0, 1224, 120)); // a page header with today's date
PdfSnapshot.Verify(pdfBytes, "statement", options);
```

`Visual.Scale` sets the stored image resolution (default `2`, 144 DPI); ignored regions are pixel rectangles at that scale. By default `Visual.IgnoreAntialiasing` skips edges that moved by less than a pixel, `ChannelTolerance` (48) skips small colour differences, and `AllowedDifferenceRatio` (0.0002 %, about four pixels on a Letter page) still catches one changed digit in table text.

Builds that map source paths, such as deterministic CI builds, hide the test file's folder. Set `Directory` there.

## Keep snapshots portable

A snapshot approved on one machine passes on another only when both machines write the same PDF. Fonts are the usual difference. By default OfficeIMO.Pdf embeds a font chosen from those installed on the machine, so the same code embeds Aptos on one computer and Liberation Sans or Arial on another. Register font files you ship with the tests through `PdfCreateOptions.RegisterFontFamily`, as in the example, and converters can use `PdfResourcePolicy.CreatePortableDeterministic()`.

The structure file lists the embedded fonts, so this case fails with a font diff and an explanation rather than as unexplained pixel changes. When a PDF uses fonts that are not embedded at all, `Verify` fails with their names. Set `AllowHostFonts = true` only when the tests always run on the same machine image.

The structure file follows OfficeIMO's reader. A newer OfficeIMO version can read text slightly differently; review those changes like any other snapshot update.
