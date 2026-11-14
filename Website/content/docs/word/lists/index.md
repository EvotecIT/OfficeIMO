---
title: Add Bulleted and Numbered Lists to Word Documents
description: Create bulleted, numbered, multi-level, and custom-numbered lists in DOCX files, start numbering at any value, and read list items from existing documents with OfficeIMO.Word.
order: 15
packages:
  - OfficeIMO.Word
cmdlets:
  - Add-OfficeWordList
  - Add-OfficeWordListItem
  - Get-OfficeWordList
search_terms:
  - word bullet list
  - numbered list
  - docx list
  - multilevel list
  - outline numbering
verified: 2026-11-14
---

A `WordList` is a native Word list backed by a numbering definition, so Word renumbers it as items change. `WordListStyle` selects a built-in definition: `Bulleted`, `Numbered`, `LowerLetterWithDot`, `UpperLetterWithBracket`, `Headings111`, `Chapters`, and more. Items take a level from 0 to 8 for nesting.

## Add a bulleted or numbered list

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("lists.docx");

document.AddParagraph("Highlights");
WordList bullets = document.AddListBulleted();
bullets.AddItem("Release completed on schedule");
bullets.AddItem("Customer guide updated");

document.AddParagraph("Next steps");
WordList steps = document.AddListNumbered();
steps.AddItem("Review adoption metrics");
steps.AddItem("Plan the next milestone");

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\lists.docx {
    Add-OfficeWordParagraph -Text 'Highlights'
    Add-OfficeWordList -Items 'Release completed on schedule', 'Customer guide updated'
    Add-OfficeWordParagraph -Text 'Next steps'
    Add-OfficeWordList -Numbered {
        Add-OfficeWordListItem -Text 'Review adoption metrics'
        Add-OfficeWordListItem -Text 'Plan the next milestone'
    }
}
```

## Nest items in a multi-level list

The second argument of `AddItem` is the level. Each level gets its own bullet or number format from the list style.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("outline.docx");

WordList outline = document.AddList(WordListStyle.Headings111);
outline.AddItem("Scope");
outline.AddItem("In scope", 1);
outline.AddItem("Out of scope", 1);
outline.AddItem("Delivery");
outline.AddItem("Pilot", 1);
outline.AddItem("Pilot sites", 2);

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\outline.docx {
    Add-OfficeWordList -Numbered {
        Add-OfficeWordListItem -Text 'Scope'
        Add-OfficeWordListItem -Text 'In scope' -Level 1
        Add-OfficeWordListItem -Text 'Out of scope' -Level 1
        Add-OfficeWordListItem -Text 'Delivery'
    }
}
```

## Format list items

`AddItem` returns the item's `WordParagraph`, so you can format it like any paragraph. Settings on the list itself (`Bold`, `Color`, `FontSize`) apply to the numbers or bullets.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("formatted-list.docx");

WordList risks = document.AddList(WordListStyle.Numbered);
risks.Color = OfficeColor.Parse("C00000");
risks.Bold = true;

risks.AddItem("Supplier delay").AddText(" - high impact").SetItalic();
risks.AddItem("Budget overrun");

document.Save();
```

## Start numbering at a specific value

A custom list lets you define the level format and the start value.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("continued.docx");

WordList list = document.AddCustomList();
list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.Decimal).SetStartNumberingValue(4));
list.AddItem("Fourth step");
list.AddItem("Fifth step");

document.Save();
```

## Read lists from an existing document

`document.Lists` returns every list. `ListItems` gives the paragraphs in order, and each item reports its level.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("lists.docx");

foreach (WordList list in document.Lists) {
    Console.WriteLine($"{list.Style} list with {list.ListItems.Count} items");
    foreach (WordParagraph item in list.ListItems) {
        Console.WriteLine($"{new string(' ', (item.ListItemLevel ?? 0) * 2)}- {item.Text}");
    }
}
```

### PowerShell

```powershell
Get-OfficeWordList -Path .\lists.docx
```
