---
title: "Export and Import Google Docs"
description: "Plan a Word-to-Google-Docs export offline, create a Google Doc from a WordDocument, import a Google Doc, and replace an existing one safely."
order: 20
packages:
  - OfficeIMO.Word.GoogleDocs
  - OfficeIMO.GoogleWorkspace
cmdlets:
  - Export-OfficeWordGoogleDocument
  - New-OfficeWordGoogleDocsOptions
  - New-OfficeGoogleWorkspaceSession
limits: "Word comments become unanchored Drive comments, because Google Docs has no equivalent custom anchor."
search_terms:
  - word to google docs
  - docx to google docs
  - import google doc
  - google docs api
verified: 2026-11-14
---

`OfficeIMO.Word.GoogleDocs` translates a `WordDocument` into Google Docs API requests and imports Google Docs back into a `WordDocument`. Build the plan first when you need a review step or a strict fidelity check: planning needs no credentials.

## Check what an export would change

### C#

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;

using WordDocument document = WordDocument.Load("report.docx");

GoogleDocsTranslationPlan plan = document.BuildGoogleDocsPlan(new GoogleDocsSaveOptions { Title = "Quarterly review" });
Console.WriteLine($"{plan.ParagraphCount} paragraphs, {plan.TableCount} tables, {plan.ImageCount} images");

foreach (TranslationNotice notice in plan.Report.Notices) {
    Console.WriteLine($"{notice.Severity} {notice.Code} {notice.Path}: {notice.Message}");
}
```

### PowerShell

```powershell
Export-OfficeWordGoogleDocument -Path .\report.docx -Mode Plan
```

## Create a Google Doc from Word

### C#

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials, new GoogleWorkspaceSessionOptions { DefaultFolderId = "folder-id" });

using WordDocument document = WordDocument.Create();
document.AddParagraph("Quarterly review").SetStyle(WordParagraphStyles.Heading1);
document.AddParagraph("Results and next steps.");

var options = new GoogleDocsSaveOptions {
    Title = "Quarterly review",
    Tabs = new GoogleDocsTabOptions { Strategy = GoogleDocsTabStrategy.FirstTab },
    Comments = GoogleDocsCommentMode.UnanchoredDriveComments
};

GoogleDocumentReference created = await document.ExportToGoogleDocsAsync(session, options);
Console.WriteLine(created.WebViewLink);
```

### PowerShell

```powershell
$session = New-OfficeGoogleWorkspaceSession -KeyPath .\Secrets\reporting.json -Scope Documents, DriveFile
try {
    $options = New-OfficeWordGoogleDocsOptions -Title 'Quarterly review' -FolderId 'folder-id'
    Export-OfficeWordGoogleDocument -Path .\report.docx -Session $session -Options $options
} finally {
    $session.CredentialSource.Dispose()
}
```

## Import a Google Doc and replace it safely

Native import projects supported tabs, body content, styles, lists, tables, links, and images. `DriveExport` mode converts to DOCX and loads it through `OfficeIMO.Word` when you need broader Word fidelity. For a replacement, pass the revision you imported; the export fails if someone changed the document in between.

### C#

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json");
var session = new GoogleWorkspaceSession(credentials);

GoogleDocsImportResult current = await session.ImportGoogleDocAsync(
    "document-id",
    new GoogleDocsImportOptions { Mode = GoogleWorkspaceImportMode.Native });

using (current.Document) {
    current.Document.AddParagraph("Reviewed on " + DateTime.Today.ToString("yyyy-MM-dd"));

    var replace = new GoogleDocsSaveOptions {
        Location = new GoogleDriveFileLocation { ExistingFileId = current.Source.FileId },
        Replace = new GoogleDocsReplaceOptions { ExpectedRevisionId = current.Source.RevisionId },
        Tabs = new GoogleDocsTabOptions { Strategy = GoogleDocsTabStrategy.ReplaceEveryTab }
    };
    await current.Document.ExportToGoogleDocsAsync(session, replace);
}
```

`FirstTab`, `SelectedTab`, and `ReplaceEveryTab` must be chosen explicitly. Replacing clears the selected scope, including stale header, footer, and named-range structures. `OverwriteLatest` and target-revision merge are alternatives you have to select; neither is used silently.

### PowerShell

```powershell
$options = New-OfficeWordGoogleDocsOptions -ExistingFileId $fileId -ExpectedRevisionId $revisionId -TabStrategy ReplaceEveryTab
Export-OfficeWordGoogleDocument -Path .\report.docx -Session $session -Options $options
```
