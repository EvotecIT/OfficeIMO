---
title: "Convert Publisher (.pub) to PDF after Microsoft Publisher retirement"
description: "Open and convert Microsoft Publisher .pub files to PDF without Publisher. Use the OfficeIMO CLI, PowerShell, or .NET to preserve and review your archive."
meta.seo_title: "Convert Publisher (.pub) to PDF after retirement | OfficeIMO"
meta.eyebrow: "Microsoft Publisher retirement"
meta.source_format: "Publisher (.pub)"
meta.destination_format: "PDF"
meta.package: "OfficeIMO.Publisher.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Publisher.Pdf"
meta.runtime: ".NET, OfficeIMO CLI, and PowerShell"
meta.primary_label: "Get the OfficeIMO CLI"
meta.primary_url: "/tool/"
meta.secondary_label: "Check Publisher format support"
meta.secondary_url: "/compatibility/"
meta.summary_title: "Publisher to PDF summary"
meta.limit: "Supported Publisher 2002-and-later files. Check fonts, graphics, and page layout before relying on the PDF."
meta.related_label: "Explore all conversion guides"
meta.related_url: "/convert/guides/"
meta.howto.name: "Convert a Publisher file to PDF"
meta.howto.description: "Convert a local .pub file with OfficeIMO.Tool and keep the original."
meta.howto.steps:
  - name: "Install the CLI"
    text: "Install OfficeIMO.Tool with the .NET SDK."
  - name: "Convert a copy"
    text: "Run officeimo convert brochure.pub brochure.pdf from the folder containing your publication."
  - name: "Review the result"
    text: "Read the conversion diagnostics and open the PDF to check the pages before archiving it."
---

## What changed with Publisher retirement

Microsoft began removing Publisher from Microsoft 365 on October 1, 2026. Support for the perpetual edition ended on October 13, 2026. A licensed perpetual edition can still open files after support ends. [Microsoft Publisher retirement guidance](https://support.microsoft.com/en-us/publisher/microsoft-publisher-will-no-longer-be-supported-after-october-2026)

If you have a folder of newsletters, flyers, or brochures, keep the original `.pub` files and make PDF copies for viewing, sharing, and printing. OfficeIMO reads supported Publisher files locally without a Microsoft Office installation.

## Convert one .pub file from the command line

Install the .NET tool, then run the conversion from the folder containing the publication.

```shell
dotnet tool install --global OfficeIMO.Tool
officeimo convert brochure.pub brochure.pdf
```

Omit the destination to write `brochure.pdf` beside the source file. The CLI writes diagnostics to standard error and keeps existing PDFs unless you pass `--force`. Add `--require-no-loss` to reject a conversion that reports fidelity loss before it writes the PDF.

## Convert a Publisher archive with PowerShell

Install PSWriteOffice in PowerShell 7 or Windows PowerShell 5.1, then use `Convert-OfficeDocument` for a single file.

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
Import-Module PSWriteOffice

Convert-OfficeDocument -LiteralPath ./brochure.pub -OutputPath ./brochure.pdf -NoClobber
```

For a folder, select only `.pub` files and write the PDFs to a separate directory. Use `-Plan` to inspect the selected files and destinations before writing. `-NoClobber` skips existing PDFs, and `-SummaryVariable` retains the final batch counts. Use paths without symbolic links or junctions for batch conversion.

```powershell
$archive = @{
    LiteralPath = './publications'
    Recurse = $true
    Extension = '.pub'
    To = 'Pdf'
    OutputDirectory = './pdf'
    NoClobber = $true
}

Convert-OfficeDocument @archive -Plan
Convert-OfficeDocument @archive -SummaryVariable summary
$summary
```

Check the returned results for failures, skipped files, and fidelity diagnostics. See the [PowerShell command reference](/api/powershell/).

## Convert from .NET

Add the `OfficeIMO.Publisher.Pdf` package to your application. The adapter uses `OfficeIMO.Publisher` to read the file and returns a report with source and PDF diagnostics.

```shell
dotnet add package OfficeIMO.Publisher.Pdf
```

```csharp
using OfficeIMO.Publisher;
using OfficeIMO.Publisher.Pdf;

var publication = PublisherDocument.Load("brochure.pub");
var report = publication.SaveAsPdf("brochure.pdf");

foreach (var diagnostic in report.Diagnostics) {
    Console.WriteLine(diagnostic.Message);
}
```

## Check the PDFs before relying on them

OfficeIMO preserves page order and physical dimensions, and supported text remains searchable. Compare a few representative files with a known good copy where possible. Check page count and size, font substitutions, line breaks, pictures, and graphics. Earlier Publisher formats, encrypted files, and damaged files outside the reader's supported profile require a separate migration path. [Publisher support contract](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Publisher/SUPPORT.md)

The PDF is a fixed-layout copy. If you need to keep editing, save the original and rebuild the document in the target application. You can also [convert the PDF to Word](/convert/pdf-to-word/) to recover editable content, then review the resulting layout.
