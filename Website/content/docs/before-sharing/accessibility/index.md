---
title: "Check accessibility"
description: "Check Word and PowerPoint files for missing alt text, headings, table headers, unclear links, low contrast, title and language, and apply safe fixes."
layout: docs
---

A shared document should work for everyone who receives it, including people who use a screen reader or high contrast. `InspectAccessibility()` checks a loaded Word document or PowerPoint presentation and returns findings with a severity, a stable rule code and a location. Word also has `FixAccessibility`, which applies the fixes that can't change meaning.

## Word

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Load("Proposal.docx");

OfficeAccessibilityReport report = document.InspectAccessibility();
foreach (OfficeAccessibilityFinding finding in report.Findings) {
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

`new WordAccessibilityOptions { Profile = OfficeAccessibilityPolicyProfile.Strict }` turns heading, floating-object and spacing warnings into errors. Each check can be switched off on the same options object. The [Word guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.word-guide.md#accessibility-checks) lists every rule code and severity.

### Safe fixes

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Load("Proposal.docx");

OfficeAccessibilityReport remaining = document.FixAccessibility(new WordAccessibilityFixOptions {
    Title = "Office move proposal",
    Language = "en-US",
    MarkTableHeaderRows = true
});
remaining.EnsureCompliant();

document.Save("Proposal-accessible.docx");
```

`FixAccessibility` sets a missing title and default language and marks the first row of data tables as a header row. It skips tables that look like layout tables and never overwrites a title or language that is already there. It returns a fresh report; `EnsureCompliant()` throws an `OfficeAccessibilityException` while errors remain, so a pipeline can stop before it sends the document. Alternative text, headings and link wording need a person, so they are reported, not guessed.

## PowerPoint

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using var presentation = PowerPointDocument.Load("Pitch.pptx");

OfficeAccessibilityReport report = presentation.InspectAccessibility();
foreach (OfficeAccessibilityFinding finding in report.Findings) {
    Console.WriteLine($"{finding.Severity} {finding.Code} at {finding.Location}: {finding.Message}");
}
report.SaveJson("Pitch-accessibility.json");
```

The presentation checker covers the title, slide titles, alternative text, table headers, link labels and tooltips, contrast and charts that rely on colour alone. `InspectReadingOrder()` lists every slide's shapes in the order a screen reader reads them.

Word and PowerPoint return the same `OfficeAccessibilityReport`: each finding has a `Severity`, a rule `Code`, a `Message` and a `Location`, and `LocationIndex` holds the slide or paragraph number where there is one. `ToJson()` and `SaveJson(path)` write it as JSON.

## Before you send

Run the accessibility check after [removing hidden data](/docs/before-sharing/hidden-data/). The scrub removes the author and other identity properties but keeps the title, so a title you add for screen readers stays. The [overview](/docs/before-sharing/) shows both steps in one pass.
