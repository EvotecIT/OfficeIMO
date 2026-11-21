---
title: "Content Credentials and AI labels"
description: "Find C2PA Content Credentials and IPTC AI source labels in Office, Visio, OpenDocument and EPUB files and their pictures, and remove them selectively."
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

`OfficeDocument.Inspect` comes with the `OfficeIMO` package and covers Word, Excel, PowerPoint, Visio, OpenDocument, EPUB and PDF files.

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

`RemoveHiddenData` reports provenance as `OfficeHiddenDataKind.Provenance` findings and keeps them unless you select that kind. `options.Provenance` chooses which kinds of record go; this example removes the AI label and keeps the Content Credentials:

```csharp
using OfficeIMO;

var options = new OfficeHiddenDataOptions {
    Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.Provenance
};
options.Provenance.RemoveC2paManifests = false;
options.Provenance.RemoveExternalC2paReferences = false;
options.Provenance.RemoveAiSourceMetadata = true;

OfficeHiddenDataReport report = OfficeDocument.RemoveHiddenData("Report.docx", "Report-shared.docx", options);
foreach (OfficeHiddenDataFinding finding in report.GetFindings(OfficeHiddenDataKind.Provenance)) {
    Console.WriteLine($"{finding.Outcome} in {finding.Part ?? "the file"} ({finding.Reason}): {finding.Description}");
}
```

| You want to | Use |
| --- | --- |
| Remove personal data and keep origin records | `RemoveHiddenData()` |
| Remove personal data and origin records | `RemoveHiddenData` with `Default | Provenance` |
| Remove personal data and only some origin records | `Default | Provenance` and the `Provenance` options |
| Remove origin records and change nothing else | `RemoveProvenance(input, output)` |

Each provenance finding says where it sits: `Scope` is `EmbeddedImage` for a picture inside the document, with the picture's path in `Part`, or `Package` for a record that covers the whole file. While provenance is kept, picture metadata cleanup keeps each picture's digital source type and leaves pictures with Content Credentials untouched, because changing their bytes would invalidate the credential; those findings carry `Reason` `SourceTypeKept` or `ProtectsContentCredentials`. Content Credentials for the whole package are the exception: they cover every byte, so `RemoveHiddenData` always removes them, with `Reason` `WholeFileCredentialsCannotSurvive`. The same holds for Content Credentials embedded in a PDF. If you need them intact, don't change the document.

## Command line

```powershell
officeimo provenance inspect .\report.docx --format text
officeimo provenance remove .\report.docx --output .\report.cleaned.docx
officeimo provenance batch remove .\one.docx .\two.pdf --output-directory .\cleaned
officeimo provenance audit .\documents --format sarif > provenance.sarif
```

The commands cover the same formats as the libraries plus images, PDF, HTML and Markdown. Results use the `officeimo.provenance.result.v2` schema, shared with Studio. The [command guide](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Tool#provenance-workflows) lists every option, including verification with `c2patool`.

In PowerShell, `Get-OfficeDocumentProvenance` reports the same evidence, and `Remove-OfficeHiddenData` with `New-OfficeHiddenDataOptions -Kind Default, Provenance` removes it together with other hidden data. See [Clean files before sharing](/docs/pswriteoffice/before-sharing/#content-credentials-and-ai-labels).

## In the browser

[Clean a file before sharing](/browser/clean-before-sharing/) shows Content Credentials and AI labels in the same report as the rest of the hidden data, for Word, Excel, PowerPoint, PDF, JPEG, PNG and WebP files up to 25 MB. For each Content Credentials record it explains what the record claims:

- **Made with:** the app that wrote the record, and the AI model named for each step when one is recorded.
- **AI-generated:** whether a step declares a generative AI source (the IPTC "trained algorithmic media" type) or a mix of AI and other content.
- **Steps:** what the record says happened, such as created, edited, converted or cropped.
- **Based on:** other files the record lists as ingredients.
- **Certificate subject:** the organization named in the signing certificate, and who issued it.

These are the record's own statements. The tool reads the certificate names but does not verify the signature, so it cannot tell whether the record was changed after it was written. Origin records are unticked by default, and Content Credentials, links to external credentials and AI labels are separate choices, so you can keep a credential and drop an AI label. Labels that say content was captured by a camera are shown but never removed.

Removing a record does not remove a visible or invisible watermark in the content itself; when a record says one was added, the result says so. The tool does not fetch external credentials, prove that a file is authentic, or decide whether a person or an AI made it. A file with no records found is just that; the absence is not proof of origin.

The [provenance support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.provenance-support-matrix.md) lists every carrier, format and limit.
