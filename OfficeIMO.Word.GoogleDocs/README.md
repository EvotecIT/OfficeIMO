# OfficeIMO.Word.GoogleDocs - Word and Google Docs translation

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word.GoogleDocs)](https://www.nuget.org/packages/OfficeIMO.Word.GoogleDocs)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word.GoogleDocs?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word.GoogleDocs)

`OfficeIMO.Word.GoogleDocs` provides bidirectional Word and Google Docs translation with tab-aware authoring, revision-guarded replacement, fidelity preflight, native import, and Drive DOCX fallback.

## Install

```powershell
dotnet add package OfficeIMO.Word.GoogleDocs
```

## Quick start

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;

using var document = WordDocument.Create("proposal.docx");
document.AddParagraph("Quarterly business review").SetStyle(WordParagraphStyles.Heading1);
document.AddParagraph("Highlights and next steps.");

var session = new GoogleWorkspaceSession(
    new StaticAccessTokenCredentialSource("<google-access-token>"),
    new GoogleWorkspaceSessionOptions {
        DefaultFolderId = "documents-folder-id",
        MaxRetryCount = 5,
        DiagnosticSink = entry => Console.WriteLine($"{entry.Severity}: {entry.Feature} - {entry.Message}")
    });

var options = new GoogleDocsSaveOptions {
    Title = "Quarterly business review"
};

var plan = document.BuildGoogleDocsPlan(options);
var result = await document.ExportToGoogleDocsAsync(session, options);

Console.WriteLine(result.DocumentId);
Console.WriteLine(result.WebViewLink);
```

## What it does

- Builds a translation plan before network export.
- Exports to a new Google Docs document or replaces an existing document through `Location.ExistingFileId`.
- Imports through native Docs resources or Drive-exported DOCX.
- Handles document tabs explicitly and requires observed revision evidence before replacing an existing document.
- Replaces selected/all tab bodies and stale headers, footers, named ranges, and related segments according to the chosen reset policy.
- Executes comments through Drive and renderer-owned flatten/rasterize fallbacks where configured.
- Uses session-level default Drive and folder placement.
- Preserves retry, warning, and failure detail through `TranslationReport`.
- Throws Google Workspace export exceptions that retain failure category and diagnostics.

## Boundaries

- Word document modeling belongs in `OfficeIMO.Word`.
- Credentials, sessions, retry, Drive placement, and report primitives belong in `OfficeIMO.GoogleWorkspace`.
- This package owns Word/Google Docs mapping, import, safe replacement, and format-specific diff planning.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, plus `net472` on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** `System.Text.Json` plus platform HTTP/cryptography through `OfficeIMO.GoogleWorkspace`; no Google client SDK.
- **OfficeIMO:** `OfficeIMO.Word` and `OfficeIMO.GoogleWorkspace` own the document model, session, translation plan, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
