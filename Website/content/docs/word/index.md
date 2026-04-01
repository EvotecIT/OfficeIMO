---
title: Word Documents
description: Overview of the OfficeIMO.Word package for creating and manipulating Microsoft Word documents.
order: 10
---

# Word Documents

The `OfficeIMO.Word` package provides a higher-level API for creating, reading, and modifying Microsoft Word documents (`.docx`) without requiring Microsoft Office. It wraps the Open XML SDK with a more approachable object model, while the deeper examples and tests in the repo cover a broader surface than this page alone.

## Key Classes

| Class | Description |
|-------|-------------|
| `WordDocument` | The root class representing a Word document. Provides factory methods for creating and loading documents. |
| `WordSection` | Represents a section within a document. Controls page layout, headers, footers, and contains paragraphs and tables. |
| `WordParagraph` | Represents a paragraph with text, formatting, images, hyperlinks, bookmarks, and field codes. |
| `WordTable` | Represents a table with rows, cells, and 105+ built-in table styles. |
| `WordImage` | Represents an inline or anchored image with sizing, cropping, and positioning controls. |
| `WordHeader` / `WordFooter` | Header and footer instances attached to sections. |
| `WordChart` | Embeds bar, line, area, pie, scatter, and radar charts. |
| `WordList` | Manages numbered and bulleted lists with customizable styles. |
| `WordBookmark` | Represents a named bookmark within the document. |
| `WordTableOfContent` | Inserts and manages a table of contents. |
| `WordFluentDocument` | Fluent builder API for composing documents with chained method calls. |

## Creating a Document

```csharp
using OfficeIMO.Word;

// Create and save to a file
using var document = WordDocument.Create("output.docx");

document.AddParagraph("Welcome to OfficeIMO");
document.Save();
```

To create a document in memory (no file path):

```csharp
using var document = WordDocument.Create();
document.AddParagraph("In-memory document");

// Save to a stream later
using var stream = new MemoryStream();
document.Save(stream);
```

## Loading an Existing Document

```csharp
using var document = WordDocument.Load("existing.docx");

// Read all paragraphs
foreach (var paragraph in document.Paragraphs) {
    Console.WriteLine(paragraph.Text);
}

// Modify and save
document.AddParagraph("Appended paragraph");
document.Save();
```

## Document Structure

A `WordDocument` contains one or more `WordSection` instances. Each section can have its own page layout (size, orientation, margins), headers, and footers. Sections contain paragraphs, tables, images, and other block-level elements.

```
WordDocument
  +-- Sections[]
  |     +-- Paragraphs[]
  |     +-- Tables[]
  |     +-- Images[]
  |     +-- Header (Default, First, Even)
  |     +-- Footer (Default, First, Even)
  +-- Paragraphs      (all paragraphs across sections)
  +-- Tables           (all tables across sections)
  +-- Lists            (all list definitions)
  +-- Bookmarks        (all bookmark paragraphs)
```

## Document Properties

```csharp
using var document = WordDocument.Create("props.docx");

document.BuiltinDocumentProperties.Title = "Quarterly Report";
document.BuiltinDocumentProperties.Creator = "OfficeIMO";
document.BuiltinDocumentProperties.Description = "Q4 2025 financial summary";

document.ApplicationProperties.Company = "Evotec";

document.Save();
```

## Sections

Add a new section to change page layout mid-document:

```csharp
using var document = WordDocument.Create("sections.docx");

document.AddParagraph("Portrait section content");

// Start a new section in landscape
var section = document.AddSection();
section.PageOrientation = Orientation.Landscape;
section.PageSettings.Width = 15840;
section.PageSettings.Height = 12240;

document.AddParagraph("Landscape section content");

document.Save();
```

## Further Reading

- [Paragraphs](/docs/word/paragraphs) -- Text formatting, alignment, spacing, and fonts.
- [Tables](/docs/word/tables) -- Creating and styling tables.
- [Images](/docs/word/images) -- Adding images from files, streams, URLs, and base64.
- [Headers and Footers](/docs/word/headers-footers) -- Page numbers, different first page headers, and section-specific content.
