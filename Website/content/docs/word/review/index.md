---
title: "Track Changes and Compare Word Documents"
description: "Add and read comments and tracked changes, accept or reject revisions by author or type, compare two DOCX versions, and create a redline document."
order: 18
packages:
  - OfficeIMO.Word
cmdlets:
  - Get-OfficeWordReview
  - Resolve-OfficeWordRevision
  - New-OfficeWordRevisionFilter
  - Compare-OfficeDocument
  - New-OfficeWordComparisonOptions
limits: Newer review extensions that OfficeIMO detects but does not parse are kept in the file and listed in UnsupportedMetadata rather than exposed in the review model.
search_terms:
  - track changes
  - word comments
  - compare docx
  - redline
  - accept revisions
verified: 2026-11-14
---

`OfficeIMO.Word` reads the review state of a document without changing it: comments, replies, resolved state, the text they point at, and tracked insertions, deletions, and moves. It can accept or reject revisions selectively, and compare two versions either as structured findings for code or as a redline document for people.

## Add comments and turn on track changes

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("proposal.docx");

WordParagraph scope = document.AddParagraph("The project covers 12 sites.");
scope.AddComment("Alice Smith", "AS", "Confirm the site count with operations.");

document.TrackChanges = true;
document.Save();
```

## Read comments and tracked changes

`InspectReview` returns a snapshot. Keep `UnsupportedMetadata` in mind: it lists review data in the file that the model does not map.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("proposal-reviewed.docx");
WordReviewInfo review = document.InspectReview();

Console.WriteLine($"{review.CommentCount} comments, {review.UnresolvedCommentCount} open, {review.RevisionCount} revisions");
foreach (WordCommentInfo comment in review.Comments) {
    Console.WriteLine($"{comment.Author}: {comment.Text} (on \"{comment.TargetText}\")");
}
```

### PowerShell

```powershell
Get-OfficeWordReview -Path .\proposal-reviewed.docx
```

## Accept or reject tracked changes

Accept or reject everything, or pass a `WordRevisionFilter` to limit the operation by author, revision type, date, or location. The filtered overload returns a report of what changed.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("contract-draft.docx");

WordRevisionOperationReport report = document.AcceptRevisions(new WordRevisionFilter { Author = "Legal" });
Console.WriteLine(report);

document.RejectRevisions("External reviewer");
document.Save("contract-clean.docx");
```

### PowerShell

```powershell
$filter = New-OfficeWordRevisionFilter -Author 'Legal'
Resolve-OfficeWordRevision -Path .\contract-draft.docx -OutputPath .\contract-clean.docx -Action Accept -Filter $filter
```

## Compare two versions

`CompareStructure` compares text and order, tables, images, formatting, fields, content controls, bookmarks, hyperlinks, lists, comments, and revisions. Each finding has a stable code, a location, and the old and new text, which suits CI checks and audit reports.

### C#

```csharp
using OfficeIMO.Word;

WordComparisonResult result = WordDocumentComparer.CompareStructure("policy-v1.docx", "policy-v2.docx");

foreach (WordComparisonFinding finding in result.Findings) {
    Console.WriteLine($"{finding.ChangeKind} at {finding.DetailedLocation}: {finding.Message}");
}
File.WriteAllText("policy-changes.md", result.ToMarkdown());
```

### PowerShell

```powershell
$options = New-OfficeWordComparisonOptions -IgnoreWhitespace -NoVolatileMetadata
Compare-OfficeDocument -ReferencePath .\policy-v1.docx -DifferencePath .\policy-v2.docx -Options $options
```

## Create a redline document

A redline is a DOCX for reviewers who work in Word. The default `ReportArtifact` mode writes a review document with a summary, a findings table, and tracked snippets. `InPlaceTarget` copies the new version and applies the supported differences to it as tracked changes.

### C#

```csharp
using OfficeIMO.Word;

WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
    "policy-v1.docx", "policy-v2.docx", "policy-redline.docx",
    new WordComparisonRedlineOptions {
        Mode = WordComparisonRedlineMode.InPlaceTarget,
        Author = "Policy review"
    });
Console.WriteLine($"{result.Findings.Count} differences");
```

### PowerShell

```powershell
Compare-OfficeDocument -ReferencePath .\policy-v1.docx -DifferencePath .\policy-v2.docx -OutputPath .\policy-redline.docx
```
