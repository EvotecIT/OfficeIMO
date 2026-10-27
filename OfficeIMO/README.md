# OfficeIMO

```sh
dotnet add package OfficeIMO
```

Convert files through the format libraries and their adapters:

```csharp
using OfficeIMO;

OfficeConversionReport report = OfficeConverter.Convert("Report.docx", "Report.pdf");
report.ThrowIfLoss();
```

The front door and batch engine support .NET Standard 2.0, .NET 8 and .NET 10. Windows builds also target .NET Framework 4.7.2. Format and adapter packages keep their existing framework coverage.

`OfficeConversionOptions.WordSave`, `ExcelSave` and `PowerPointSave` control native destination settings, including passwords, for cross-format conversions as well as native saves. `FailOnLoss` rejects conversion loss before destination writes. If final publication fails, `OfficeConversionException.Report` retains the completed findings. Analysis leaves image and SVG sidecars unpublished; file conversion applies its publication guard and conflict policy to those assets too.

Convert batches with the same options and report vocabulary:

```csharp
using OfficeIMO;

OfficeBatchReport batch = await OfficeConverter.ConvertAsync(
    Directory.EnumerateFiles("Reports", "*.docx"),
    outputDirectory: "Pdf",
    target: OfficeDocumentFormat.Pdf,
    new OfficeBatchOptions { MaxParallelism = 4, ResumeFrom = "Pdf/.officeimo-batch.json" });
Console.WriteLine(batch.Succeeded);
```

Use the source model for in-memory conversion. Async conversions return the model and its completed report together:

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

using var source = MarkdownDocument.Parse("# Quarterly review");
OfficeConversionResult<OfficeIMO.Word.WordDocument> converted = await source.ToWordAsync();
using var word = converted.Value;
converted.Report.ThrowIfLoss();
word.Save("review.docx");
```

Inspect native content or select provenance and signature reports without extracting content:

```csharp
using OfficeIMO;
using OfficeIMO.Provenance;

OfficeDocumentInfo info = OfficeDocument.Inspect("review.docx", new OfficeInspectOptions {
    IncludeContent = false,
    Provenance = new OfficeProvenanceOptions()
});
Console.WriteLine(info.Source);
```

`OfficeIMO.Workflows` and `OfficeIMO.Workflows.IWork` register their additional conversion routes with `OfficeConverter` through `RegisterConversions()`. The batch engine, checkpoint format and saved options remain the same on every supported framework.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 27 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
