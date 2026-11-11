---
title: "Check accessibility"
description: "Check Word documents and PowerPoint presentations for missing alternative text, headings, table headers, unclear links, low contrast, title and language, and apply safe fixes with OfficeIMO."
layout: docs
---

A shared document should work for everyone who receives it, including people who use a screen reader or high contrast. `InspectAccessibility()` checks a loaded Word document or PowerPoint presentation and returns findings with a severity, a stable rule code and a location. Word also has `FixAccessibility`, which applies the fixes that can't change meaning.

## Word

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Load("Proposal.docx");

WordAccessibilityReport report = document.InspectAccessibility();
foreach (WordAccessibilityFinding finding in report.Findings) {
    Console.WriteLine($"{finding.Severity} {finding.Code} at {finding.Location}: {finding.Message}");
}
Console.WriteLine($"{report.ErrorCount} errors, {report.WarningCount} warnings");
```

The checker covers the document body:

- a document title and a default language;
- alternative text on pictures, shapes, charts and SmartArt, unless they are marked decorative;
- heading levels that skip a level, and long documents with no headings;
- a header row on data tables, and tables that look like layout tables;
- links with no text or text such as "click here" or a bare URL;
- text below 4.5:1 contrast (3:1 for large text) against its background;
- floating pictures and runs of empty paragraphs used for spacing.

`new WordAccessibilityOptions { Profile = WordAccessibilityPolicyProfile.Strict }` turns heading, floating-object and spacing warnings into errors. Each check can be switched off on the same options object. The [Word guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.word-guide.md#accessibility-checks) lists every rule code and severity.

### Safe fixes

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Load("Proposal.docx");

WordAccessibilityReport remaining = document.FixAccessibility(new WordAccessibilityFixOptions {
    Title = "Office move proposal",
    Language = "en-US",
    MarkTableHeaderRows = true
});
remaining.EnsureCompliant();

document.Save("Proposal-accessible.docx");
```

`FixAccessibility` sets a missing title and default language and marks the first row of data tables as a header row. It skips tables that look like layout tables and never overwrites a title or language that is already there. It returns a fresh report; `EnsureCompliant()` throws a `WordAccessibilityException` while errors remain, so a pipeline can stop before it sends the document. Alternative text, headings and link wording need a person, so they are reported, not guessed.

## PowerPoint

```csharp
using OfficeIMO.PowerPoint;

using var presentation = PowerPointDocument.Load("Pitch.pptx");

PowerPointAccessibilityReport report = presentation.InspectAccessibility();
foreach (PowerPointAccessibilityFinding finding in report.Findings) {
    string where = finding.SlideIndex is int index ? $"slide {index + 1}" : "presentation";
    Console.WriteLine($"{finding.Severity} {finding.Code} on {where}: {finding.Message}");
}
File.WriteAllText("Pitch-accessibility.json", report.ToJson());
```

The presentation checker covers the title, slide titles, alternative text, table headers, link labels and tooltips, contrast and charts that rely on colour alone. Its report lists every slide's shapes in reading order and exports JSON with `ToJson()`.

## Before you send

Run the accessibility check after [removing hidden data](/docs/before-sharing/hidden-data/). The scrub removes the author and other identity properties but keeps the title, so a title you add for screen readers stays. The [overview](/docs/before-sharing/) shows both steps in one pass.
