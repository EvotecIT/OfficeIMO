---
title: "Review, Revisions, and Security"
description: "Inspect comments and tracked changes, compare document structure, produce redlines, accept or reject revisions, and protect signed packages."
layout: docs
---

OfficeIMO.Word includes a structured review model for documents that move through approval workflows. It can inspect comments, replies, resolved state, targets, tracked revisions, and unsupported review metadata without mutating the source.

## Inspect review state

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Load("proposal-reviewed.docx");
WordReviewInfo review = document.InspectReview();

Console.WriteLine($"Comments: {review.CommentCount}");
Console.WriteLine($"Open comments: {review.UnresolvedCommentCount}");
Console.WriteLine($"Revisions: {review.Revisions.Count}");
```

`InspectReviewReport` packages inspection results with optional accept/reject operation reports. Preserve its unsupported-metadata list: Office documents can contain newer review extensions that are retained but not fully mapped into the public model.

## Compare two versions

`WordDocumentComparer.CompareStructure` compares text, ordering, tables, images, effective formatting, fields, content controls, bookmarks, hyperlinks, lists, comments, and revisions. `WordComparisonOptions` lets you exclude volatile identifiers or a feature family when it is irrelevant to your approval policy.

Use the structured result for CI, audit, or custom reporting. Use the redline path when reviewers need a document representation of inserted, deleted, moved, or changed content.

## Accept and reject revisions

`AcceptRevisions` and `RejectRevisions` support whole-document and filtered operations, including paragraph-scoped variants. Capture the returned operation report when a workflow must show what changed. Save to a new file until the result has passed semantic readback.

## Protection and package signatures

Document editing protection and package signatures solve different problems. Protection requests an editing mode and can be password-backed; it is not a cryptographic proof of origin. Package signing uses an X.509 certificate and should be applied after all document mutations.

The default signed-document save policy blocks edits that would invalidate a signature. Inspect and validate signature coverage before relying on a signed artifact, and define an explicit re-signing policy for approved modifications.
