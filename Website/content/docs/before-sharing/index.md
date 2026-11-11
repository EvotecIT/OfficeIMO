---
title: "Prepare a document for sharing"
description: "Remove author names, comments, hidden content, scripts and photo locations from Office, PDF and image files, decide what happens to Content Credentials and AI labels, and check accessibility before you send a document."
layout: docs
---

A document that leaves your team carries more than its visible text. Author and company names, reviewer comments, tracked changes, hidden rows, speaker notes, the GPS location of a photo and Content Credentials can all travel with it. OfficeIMO finds them, removes what you select and reports what happened to each item, without Microsoft Office.

## Three checks

- **[Hidden data](/docs/before-sharing/hidden-data/):** what personal or internal information leaves with the file. `OfficeDocument.RemoveHiddenData(input, output, options)`, `InspectHiddenData()` and `RemoveHiddenData(options)` on a loaded document, or `officeimo hidden-data` on the command line.
- **[Content Credentials and AI labels](/docs/before-sharing/provenance/):** does the file say how it was made, for example that a picture is AI-generated? Reported and removed as the `Provenance` category of hidden data, or on their own with `RemoveProvenance` and `officeimo provenance`.
- **[Accessibility](/docs/before-sharing/accessibility/):** can everyone read it, including people who use a screen reader? `InspectAccessibility()` and `FixAccessibility(options)`.

Hidden data covers Word, Excel, PowerPoint, Visio, OpenDocument, EPUB, PDF and JPEG, PNG, WebP, GIF and TIFF images, and the accessibility checker covers Word and PowerPoint. `RemoveProvenance` removes origin records and leaves every other byte alone.

## One pass in code

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Load("Proposal.docx");

OfficeHiddenDataReport hidden = document.RemoveHiddenData(new OfficeHiddenDataOptions {
    TrackedChanges = OfficeTrackedChangesAction.AcceptAll
});
OfficeAccessibilityReport accessibility = document.FixAccessibility(new WordAccessibilityFixOptions {
    Title = "Proposal",
    Language = "en-US",
    MarkTableHeaderRows = true
});

foreach (OfficeHiddenDataFinding finding in hidden.Findings) {
    Console.WriteLine($"{finding.Outcome,-10} {finding.Kind}: {finding.Description}");
}
foreach (OfficeAccessibilityFinding finding in accessibility.Findings) {
    Console.WriteLine($"{finding.Severity,-10} {finding.Code} at {finding.Location}");
}

document.Save("Proposal-shared.docx");
```

Remove hidden data first and fix accessibility second: the scrub removes the author and other identity properties but keeps the title, so the title you add for screen readers stays. Sign the document last. Any change after signing breaks the signature, so OfficeIMO refuses to save a signed document unless you allow it.

## Kept unless you ask

`RemoveHiddenData()` with no options removes the categories that never change what readers see, including scripts and other actions that run when the file is opened. These stay until you select them, because removing them changes the content or what the file discloses:

- **Tracked changes.** Accepting or rejecting them changes the text. Set `TrackedChanges` to `OfficeTrackedChangesAction.AcceptAll` or `RejectAll`.
- **Hidden text, sheets, rows, columns, slides and shapes.** Readers would get different content, and deleting hidden rows or sheets can break formulas. Add `OfficeHiddenDataKind.HiddenContent`.
- **Attached files and PDF bookmarks.** The reader may need them. Add `OfficeHiddenDataKind.Attachments` or `Bookmarks`.
- **Content Credentials and AI labels.** They tell the reader how content was made, for example that a picture is AI-generated. Add `OfficeHiddenDataKind.Provenance`, and use `options.Provenance` to keep some kinds of record.
- **Power Query data in a workbook.** The workbook's queries live there. Delete the queries in Excel.

Two items are reported but always kept, because removing them would break the file: links that formulas or pivot refreshes depend on, and linked images and OLE links in Word. One item is always removed: Content Credentials that cover the whole file, such as `META-INF/content_credential.c2pa` in a package or a credential embedded in a PDF. They cover every byte, so any change invalidates them; the finding's `Reason` is `WholeFileCredentialsCannotSurvive`. To remove provenance and leave every other byte alone, use `RemoveProvenance` instead of `RemoveHiddenData`.

Every finding has an `Outcome`: `Found` after inspection, then `Removed`, `Anonymized`, `Accepted`, `Rejected` or `Retained`. A retained finding carries a `Reason` code and a `Note` sentence that say why.

## In the browser

[Clean a file before sharing](/browser/clean-before-sharing/) is the same check without code. Drop a file and it lists everything that would leave with it, grouped by category:

- **Personal and hidden data:** properties, names, comments, tracked changes, hidden content, speaker notes, photo details and the rest of the categories above, with the defaults above already ticked.
- **Content Credentials and AI labels:** what each record claims, such as the app and AI model that made a picture, and who the certificate names. Credentials, links to credentials and AI labels are separate choices, all unticked.
- **Accessibility**, for Word and PowerPoint: what a screen reader user would miss. Marking table header rows is offered as a fix; everything else needs a person.

One button saves a clean copy. The tool checks the copy again and lists what was removed, what was kept and why, and offers the report as JSON. The file stays in your browser and the original is never changed. It takes Word, Excel and PowerPoint files, PDFs, and JPEG, PNG and WebP images.

## Command line

```powershell
officeimo hidden-data inspect proposal.docx
officeimo hidden-data remove proposal.docx --kinds default,provenance
officeimo hidden-data remove photo.jpg --kinds default,provenance --keep-provenance c2pa
```

Install [OfficeIMO.Tool](/tool/) first. The [hidden data page](/docs/before-sharing/hidden-data/#command-line) lists every option.

## For AI agents

The [OfficeIMO MCP server](/agents/) offers `officeimo_inspect_hidden_data` and `officeimo_remove_hidden_data`, with the same choices as the command line. An agent can report what would leave with a file, then write a cleaned copy beside it.

## Hidden characters in text

Text can carry invisible characters too: zero-width spaces, direction overrides that make part of a line read backwards, such as a file name that hides its real extension, and Unicode tags. [Find hidden characters in text](/browser/hidden-characters/) checks pasted text or a UTF-8 file, or a UTF-16/32 file with a byte-order mark, up to 1 MiB and 512 findings. Nothing is selected for you, because language joiners, emoji selectors and typographic spaces are often intentional, and no finding proves AI authorship. The copy keeps the original encoding, byte-order mark and line endings. Reports use the `officeimo.text-integrity.result.v1` schema; in .NET, use `OfficeTextIntegrityInspector` and `OfficeTextIntegrityCleaner`.

## Reference

- [Hidden data support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.hidden-data-support-matrix.md): every category per format.
- [Provenance support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.provenance-support-matrix.md): every carrier and format.
- [Word accessibility rules](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.word-guide.md#accessibility-checks): every rule code and severity.
