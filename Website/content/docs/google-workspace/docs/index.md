---
title: Word and Google Docs
description: Plan, export, import, and safely replace Google Docs with OfficeIMO.Word.
order: 20
---

Install `OfficeIMO.Word.GoogleDocs`. Build the plan before acquiring credentials when you need a review or strict fidelity gate.

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;

using WordDocument document = WordDocument.Create();
document.AddParagraph("Quarterly review").SetStyle(WordParagraphStyles.Heading1);
document.AddParagraph("Results and next steps.");

var options = new GoogleDocsSaveOptions {
    Title = "Quarterly review",
    Tabs = new GoogleDocsTabOptions { Strategy = GoogleDocsTabStrategy.FirstTab },
    Comments = GoogleDocsCommentMode.UnanchoredDriveComments
};
GoogleDocsTranslationPlan plan = document.BuildGoogleDocsPlan(options);
GoogleDocumentReference created = await document.ExportToGoogleDocsAsync(session, options);
```

For a safe replacement, native-import the target and pass the observed revision:

```csharp
GoogleDocsImportResult current = await session.ImportGoogleDocAsync(
    "document-id",
    new GoogleDocsImportOptions { Mode = GoogleDocsImportMode.Native });

using (current.Document) {
    var replace = new GoogleDocsSaveOptions {
        Location = new GoogleDriveFileLocation { ExistingFileId = current.Source.FileId },
        Replace = new GoogleDocsReplaceOptions { ExpectedRevisionId = current.Source.RevisionId },
        Tabs = new GoogleDocsTabOptions { Strategy = GoogleDocsTabStrategy.ReplaceEveryTab }
    };
    await document.ExportToGoogleDocsAsync(session, replace);
}
```

`FirstTab`, `SelectedTab`, and `ReplaceEveryTab` are explicit. Existing replacement cleans the selected scope, including stale header/footer and named-range structures. `OverwriteLatest` and target-revision merge are deliberate alternatives; neither is silently selected.

Native import projects supported tabs, body content, styles, lists, tables, links, and images. `DriveExport` converts to DOCX and loads through `OfficeIMO.Word` when broader Word fidelity matters. Word comments can become unanchored Drive comments because Google editors do not expose an equivalent custom anchor contract.
