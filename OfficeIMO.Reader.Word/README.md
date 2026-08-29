# OfficeIMO.Reader.Word

Word ingestion for `OfficeIMO.Reader.Core` with DOCX, DOCM, and legacy DOC support.

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .Build();
```

To add best-effort page locations for search and citations:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler(new ReaderWordOptions {
        IncludePageLocations = true
    })
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("policy.docx");
OfficeDocumentSearchResult matches = document.Search("retention period");

foreach (int page in matches.PageNumbers) {
    Console.WriteLine($"Found on page {page} of {matches.TotalPageCount}");
}
```

Word does not store stable physical pages. Enabling `IncludePageLocations` runs the OfficeIMO.Word layout engine
with the configured fonts and resources, maps visible body-block fragments to computed pages, and reports
`OfficeDocumentPageProvenance.Computed`. Results can differ from Microsoft Word when fonts, metrics, or unsupported
layout features differ, so the option is disabled by default.

Selected WordPerfect, WordStar, Ami Pro, Lotus Word Pro, Works/Write, and Word for DOS sources are opt-in through the owning legacy package:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .AddLegacyWordHandler()
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("archive.wpd");
```

Legacy warnings include the detected profile, structured-versus-salvage quality, and feature-level losses. The handler never executes source macros or embedded code.
