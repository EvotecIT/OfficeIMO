---
title: Add Images to Word Documents
description: Insert pictures into DOCX files from files, streams, or byte arrays, size them, wrap text around them, add alt text, and read or crop existing images with OfficeIMO.Word.
order: 13
packages:
  - OfficeIMO.Word
cmdlets:
  - Add-OfficeWordImage
  - Get-OfficeWordImage
  - Set-OfficeWordImage
search_terms:
  - word image
  - docx picture
  - insert image
  - text wrapping
  - alt text
verified: 2026-11-14
---

Images belong to a paragraph. `paragraph.AddImage` embeds the picture in that paragraph, inline with the text by default. Width and height are `OfficeLength` values, so the unit is stated (`Pixels`, `Points`, `Centimeters`, `Inches`); leave them out to use the image's own size. `WordImageOptions` sets wrapping, alt text, alignment, and a maximum width.

## Add an image from a file

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("images.docx");

WordParagraph paragraph = document.AddParagraph();
paragraph.AddImage("photo.jpg", OfficeLength.Pixels(300), OfficeLength.Pixels(200),
    new WordImageOptions { AltText = "Team photo at the 2026 kickoff" });

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\images.docx {
    Add-OfficeWordParagraph {
        Add-OfficeWordImage -Path .\photo.jpg -Width 8cm -Description 'Team photo at the 2026 kickoff'
    }
}
```

## Add an image from a stream or bytes

Pass a file name with a stream so the image part gets the right extension and content type.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

byte[] logoBytes = File.ReadAllBytes("logo.png");

using WordDocument document = WordDocument.Create("logo.docx");
using var stream = new MemoryStream(logoBytes);
document.AddParagraph().AddImage(stream, "logo.png", OfficeLength.Pixels(150), OfficeLength.Pixels(50));
document.Save();
```

## Wrap text around an image

Any wrapping other than `InLineWithText` anchors the image so text flows around it. Available modes are `InLineWithText`, `Square`, `Tight`, `Through`, `TopAndBottom`, `BehindText`, and `InFrontOfText`.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("wrapped.docx");

WordParagraph paragraph = document.AddParagraph();
WordImage chart = paragraph.AddImage("chart.png", OfficeLength.Centimeters(8), OfficeLength.Centimeters(6),
    new WordImageOptions { Wrapping = WordImageTextWrapping.Square, AltText = "Revenue by quarter" });
paragraph.AddText("Revenue grew in every quarter. The chart shows the trend by region, " +
    "and the table on the next page lists the figures.");

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\wrapped.docx {
    Add-OfficeWordParagraph {
        Add-OfficeWordImage -Path .\chart.png -Width 8cm -Height 6cm -Wrap Square -Description 'Revenue by quarter'
        Add-OfficeWordText 'Revenue grew in every quarter.'
    }
}
```

## Center an image and limit its width

`MaximumWidth` caps the displayed width when the source image is wider than the page area you have, which helps when the source size is unknown.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("banner.docx");
document.AddParagraph().AddImage("banner.png", options: new WordImageOptions {
    Alignment = WordParagraphAlignment.Center,
    MaximumWidth = OfficeLength.Centimeters(16),
    AltText = "Product banner"
});
document.Save();
```

## Crop and adjust an existing image

`document.Images` lists every image in the body. Crop ratios select a fraction of the source picture while the frame size stays the same.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("images.docx");

foreach (WordImage image in document.Images) {
    Console.WriteLine($"{image.FileName}: {image.Width} x {image.Height}");
}

WordImage first = document.Images[0];
first.CropTopRatio = 0.1;
first.CropBottomRatio = 0.05;
first.GrayScale = true;

document.Save();
```

### PowerShell

```powershell
Get-OfficeWordImage -Path .\images.docx
```
