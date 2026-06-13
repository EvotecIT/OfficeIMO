# OfficeIMO.Word.GoogleDocs - Word to Google Docs export

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word.GoogleDocs)](https://www.nuget.org/packages/OfficeIMO.Word.GoogleDocs)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word.GoogleDocs?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word.GoogleDocs)

`OfficeIMO.Word.GoogleDocs` builds translation plans and export requests for sending `OfficeIMO.Word` documents to Google Docs through `OfficeIMO.GoogleWorkspace`.

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

var plan = document.CreateGoogleDocsTranslationPlan(options);
var result = await document.ExportToGoogleDocsAsync(session, options);

Console.WriteLine(result.DocumentId);
Console.WriteLine(result.WebViewLink);
```

## What it does

- Builds a translation plan before network export.
- Exports to a new Google Docs document or replaces an existing document through `Location.ExistingFileId`.
- Uses session-level default Drive and folder placement.
- Preserves retry, warning, and failure detail through `TranslationReport`.
- Throws Google Workspace export exceptions that retain failure category and diagnostics.

## Boundaries

- Word document modeling belongs in `OfficeIMO.Word`.
- Credentials, sessions, retry, Drive placement, and report primitives belong in `OfficeIMO.GoogleWorkspace`.
- This package owns Word-to-Google-Docs mapping and export request construction.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
