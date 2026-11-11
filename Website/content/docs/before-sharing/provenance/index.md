---
title: "Content Credentials and AI labels"
description: "Find C2PA Content Credentials and IPTC AI source labels in Office, Visio, OpenDocument and EPUB files and their pictures, and remove them selectively with OfficeIMO."
layout: docs
---

Provenance records say how content was made. OfficeIMO reads two kinds: **Content Credentials** (C2PA manifests, embedded or referenced from outside the file) and **IPTC digital source type** labels, such as a picture marked as made by a generative AI model. They can sit on the whole document package and on each picture inside it.

These records are disclosures, not personal data, so OfficeIMO keeps them unless you ask. Decide per document: an AI label can be required by your publishing policy, or it can reveal a tool you are not allowed to name.

## Inspect

```csharp
using OfficeIMO;
using OfficeIMO.Provenance;

OfficeDocumentInfo info = OfficeDocument.Inspect("Report.docx", new OfficeInspectOptions {
    Provenance = new OfficeProvenanceOptions()
});

OfficeProvenanceReport provenance = info.Provenance!;
foreach (OfficeProvenanceEvidence evidence in provenance.Evidence) {
    Console.WriteLine($"{evidence.Carrier} at {evidence.Location}: {evidence.DigitalSourceKind}");
    if (evidence.Manifest is { } manifest) {
        Console.WriteLine($"  claimed by {manifest.ClaimGenerator}, AI declared: {manifest.DeclaresGenerativeAi}");
    }
}
Console.WriteLine($"Generative AI declared: {provenance.HasGenerativeAiDeclaration}");
```

`OfficeDocument.Inspect` comes with the `OfficeIMO` package and covers Word, Excel and PowerPoint files. `VisioDocument.InspectProvenance`, `OdfDocument.InspectProvenance` and `EpubDocument.InspectProvenance` inspect Visio, OpenDocument and EPUB files.

`Manifest` summarises what a Content Credentials record claims: the app that wrote it, its steps, the AI model named for a step, its ingredients and the certificate subject. These are the record's own statements. Inspection checks the structure; it does not verify the signature or the signer. Use `OfficeIMO.Provenance.C2pa` with a `c2patool` you trust for verification.

## Remove only provenance

```csharp
using OfficeIMO;
using OfficeIMO.Provenance;
using OfficeIMO.Word;

OfficeProvenanceRemovalResult result = WordDocument.RemoveProvenance("Report.docx", "Report-clean.docx",
    new OfficeProvenanceRemovalOptions {
        RemoveC2paManifests = true,
        RemoveExternalC2paReferences = true,
        RemoveAiSourceMetadata = true,
        SignatureMutationPolicy = OfficeSignatureMutationPolicy.BlockSave
    });

foreach (OfficeProvenanceChange change in result.Changes) {
    Console.WriteLine($"Removed {change.Carrier} at {change.Location} ({change.RemovedBytes} bytes)");
}
Console.WriteLine($"Records left: {result.After.Evidence.Count}");
```

`RemoveProvenance` is static: it reads a file or bytes, removes the selected records and leaves every other byte alone. It writes the output atomically and inspects it again; `Before`, `After` and `Changes` describe the result. `ExcelDocument`, `PowerPointDocument`, `VisioDocument`, `OdfDocument` and `EpubDocument` have the same method.

Removing a record from a signed package breaks the package signature. `BlockSave`, the default, refuses. `RemoveInvalidatedSignatures` removes the signatures the change would break and reports it in `WereInvalidatedSignaturesRemoved`.

Only AI-specific source labels are removed by `RemoveAiSourceMetadata`; a label that says a picture was captured by a camera stays. Malformed or ambiguous records are left in place unless you set `RequireStructurallyValidCarrier = false`.

## With hidden data

`RemoveHiddenData` reports provenance as `OfficeHiddenDataKind.Provenance` findings and keeps them unless you select that kind:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Load("Report.docx");
OfficeHiddenDataReport report = document.RemoveHiddenData(new OfficeHiddenDataOptions {
    Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.Provenance
});
foreach (OfficeHiddenDataFinding finding in report.GetFindings(OfficeHiddenDataKind.Provenance)) {
    Console.WriteLine($"{finding.Outcome}: {finding.Description} {finding.Note}");
}
document.Save("Report-shared.docx");
```

| You want to | Use |
| --- | --- |
| Remove personal data and keep origin records | `RemoveHiddenData()` |
| Remove personal data and origin records | `RemoveHiddenData` with `Default | Provenance` |
| Remove origin records and change nothing else | `RemoveProvenance(input, output)` |

While provenance is kept, picture metadata cleanup keeps each picture's digital source type and leaves pictures with Content Credentials untouched, because changing their bytes would invalidate the credential. Content Credentials for the whole package are the exception: they cover every byte, so `RemoveHiddenData` always removes them and says so in the finding's `Note`. If you need them intact, don't change the document.

## Command line

```powershell
officeimo provenance inspect .\report.docx --format text
officeimo provenance remove .\report.docx --output .\report.cleaned.docx
officeimo provenance batch remove .\one.docx .\two.pdf --output-directory .\cleaned
officeimo provenance audit .\documents --format sarif > provenance.sarif
```

The commands cover the same formats as the libraries plus images, PDF, HTML and Markdown. Results use the `officeimo.provenance.result.v2` schema, shared with Studio and the browser tool. The [command guide](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Tool#provenance-workflows) lists every option, including verification with `c2patool`.

## In the browser

[Check file origin](/browser/file-origin/) inspects images, PDFs and Office files in your browser, explains what each Content Credentials record claims, and saves a copy without the records you select. The [provenance page](/provenance/) explains what the tool reads and what it cannot prove.

The [provenance support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.provenance-support-matrix.md) lists every carrier, format and limit.
