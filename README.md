
<div align="right">
  <details>
    <summary >🌐 Language</summary>
    <div>
      <div align="center">
        <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=en">English</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=zh-CN">简体中文</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=zh-TW">繁體中文</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=ja">日本語</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=ko">한국어</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=hi">हिन्दी</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=th">ไทย</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=fr">Français</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=de">Deutsch</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=es">Español</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=it">Italiano</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=ru">Русский</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=pt">Português</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=nl">Nederlands</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=pl">Polski</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=ar">العربية</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=fa">فارسی</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=tr">Türkçe</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=vi">Tiếng Việt</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=id">Bahasa Indonesia</a>
        | <a href="https://openaitx.github.io/view.html?user=EvotecIT&project=OfficeIMO&lang=as">অসমীয়া</
      </div>
    </div>
  </details>
</div>

# OfficeIMO - Microsoft Word .NET Library

OfficeIMO is available as a NuGet package from the gallery and is the preferred installation method.

[![nuget downloads](https://img.shields.io/nuget/dt/officeIMO.Word?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word)
[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word)](https://www.nuget.org/packages/OfficeIMO.Word)
[![license](https://img.shields.io/github/license/EvotecIT/OfficeIMO.svg)](#)
[![CI](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml/badge.svg?branch=master)](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml)
[![codecov](https://codecov.io/gh/EvotecIT/OfficeIMO/branch/master/graph/badge.svg)](https://codecov.io/gh/EvotecIT/OfficeIMO)

If you would like to contact me you can do so via Twitter or LinkedIn.

[![twitter](https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social)](https://twitter.com/PrzemyslawKlys)
[![blog](https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg)](https://evotec.xyz/hub)
[![linked](https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn)](https://www.linkedin.com/in/pklys)
[![discord](https://img.shields.io/discord/508328927853281280?style=flat-square&label=discord%20chat)](https://evo.yt/discord)

## What it's all about

This is a small project (under development) that allows to create Microsoft Word documents (.docx) using .NET.
Underneath it uses [OpenXML SDK](https://github.com/OfficeDev/Open-XML-SDK) but heavily simplifies it.
It was created because working with OpenXML is way too hard for me and time-consuming.
I created it for use within the PowerShell module called [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice),
but thought it may be useful for others to use in the .NET community.
This repository also includes an experimental **OfficeIMO.Excel** component for creating simple spreadsheets.

If you want to see what it can do take a look at this [blog post](https://evotec.xyz/officeimo-free-cross-platform-microsoft-word-net-library/).

I used to use the DocX library (which I co-authored/maintained before it was taken over by Xceed) to create Microsoft Word documents,
but it only supports .NET Framework, and their newest community license makes the project unusable.

*As I am not really a developer, and I hardly know what I'm doing if you know how to help out - please do.*

- If you see bad practice, please open an issue/submit PR.
- If you know how to do something in OpenXML that could help this project - please open an issue/submit PR
- If you see something that could work better - please open an issue/submit PR
- If you see something that I made a fool of myself - please open an issue/submit PR
- If you see something that works not the way I think it works - please open an issue/submit PR

If you notice any issues or have suggestions for improvement, please open an issue or submit a pull request.
The main thing is - it has to work with .NET Framework 4.7.2, .NET Standard 2.0 and so on.

**This project is under development and as such there's a lot of things that can and will change, especially if some people help out.**


| Platform | Status | Code Coverage | .NET |
| -------- | ------ | ------------- | ---- |
| Windows  | ![Windows](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml/badge.svg?branch=master) | [Codecov](https://codecov.io/gh/EvotecIT/OfficeIMO) | OfficeIMO.Word: `netstandard2.0`, `net472`, `net8.0`, `net9.0`; OfficeIMO.Excel: `netstandard2.0`, `net472`, `net48`, `net8.0`, `net9.0` |
| Linux    | ![Linux](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml/badge.svg?branch=master) | [Codecov](https://codecov.io/gh/EvotecIT/OfficeIMO) | OfficeIMO.Word: `net8.0`; OfficeIMO.Excel: `net8.0`, `net9.0` |
| MacOs    | ![macOS](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml/badge.svg?branch=master) | [Codecov](https://codecov.io/gh/EvotecIT/OfficeIMO) | OfficeIMO.Word: `net8.0`; OfficeIMO.Excel: `net8.0`, `net9.0` |
## Support This Project

If you find this project helpful, please consider supporting its development.
Your sponsorship will help the maintainers dedicate more time to maintenance and new feature development for everyone.

It takes a lot of time and effort to create and maintain this project.
By becoming a sponsor, you can help ensure that it stays free and accessible to everyone who needs it.

To become a sponsor, you can choose from the following options:

- [Become a sponsor via GitHub Sponsors :heart:](https://github.com/sponsors/PrzemyslawKlys)
- [Become a sponsor via PayPal :heart:](https://paypal.me/PrzemyslawKlys)

Your sponsorship is completely optional and not required for using this project.
We want this project to remain open-source and available for anyone to use for free,
regardless of whether they choose to sponsor it or not.

If you work for a company that uses our .NET libraries or PowerShell Modules,
please consider asking your manager or marketing team if your company would be interested in supporting this project.
Your company's support can help us continue to maintain and improve this project for the benefit of everyone.

Thank you for considering supporting this project!

## Please share with the community

Please consider sharing a post about OfficeIMO and the value it provides. It really does help!

[![Share on reddit](https://img.shields.io/badge/share%20on-reddit-red?logo=reddit)](https://reddit.com/submit?url=https://github.com/EvotecIT/OfficeIMO&title=OfficeIMO)
[![Share on hacker news](https://img.shields.io/badge/share%20on-hacker%20news-orange?logo=ycombinator)](https://news.ycombinator.com/submitlink?u=https://github.com/EvotecIT/OfficeIMO)
[![Share on twitter](https://img.shields.io/badge/share%20on-twitter-03A9F4?logo=twitter)](https://twitter.com/share?url=https://github.com/EvotecIT/OfficeIMO&t=OfficeIMO)
[![Share on facebook](https://img.shields.io/badge/share%20on-facebook-1976D2?logo=facebook)](https://www.facebook.com/sharer/sharer.php?u=https://github.com/EvotecIT/OfficeIMO)
[![Share on linkedin](https://img.shields.io/badge/share%20on-linkedin-3949AB?logo=linkedin)](https://www.linkedin.com/shareArticle?url=https://github.com/EvotecIT/OfficeIMO&title=OfficeIMO)

## Features

Here's a list of features currently supported (and probably a lot I forgot) and those that are planned. It's not a closed list, more of TODO, and I'm sure there's more:

- ☑️ Word basics
    - ☑️ Create
    - ☑️ Load
    - ☑️ Save (auto open on save as an option)
    - ☑️ SaveAs (auto open on save as an option)
- ☑️ Word properties
    - ☑️ Reading basic and custom properties
    - ☑️ Setting basic and custom properties
- ☑️ Sections
    - ☑️ Add Paragraphs
    - ☑️ Add Headers and Footers (Odd/Even/First)
    - ☑️ Remove Headers and Footers (Odd/Even/First)
    - ☑️ Remove Paragraphs
    - ☑️ Remove Sections
- ☑️ Headers and Footers in the document (not including sections)
    - ☑️ Add Default, Odd, Even, First
    - ☑️ Remove Default, Odd, Even, First
- ☑️ Paragraphs/Text and make it bold, underlined, colored and so on
    - ☑️ Custom paragraph styles via `WordParagraphStyle`
- ☑️ Paragraphs and change alignment
- ☑️ Paragraph indentation (before, after, first line, hanging)
- ☑️ Line spacing with support for twips and points
- ☑️ Tables
    - ☑️ [Add and modify table styles (one of 105 built-in styles)](https://evotec.xyz/docs/adding-tables-with-built-in-styles-managing-borders/)
    - ☑️ Add rows and columns
    - ☑️ Add cells
    - ☑️ Add cell properties
    - ☑️ [Add and modify table cell borders](https://evotec.xyz/docs/adding-tables-with-built-in-styles-managing-borders/)
    - ☑️ Remove rows
    - ☑️ Remove cells
    - ☑️ Others
        - ☑️ Merge cells (vertically, horizontally)
        - ☑️ Split cells (vertically)
        - ☑️ Split cells (horizontally)
        - ☑️ Detect merged cells (vertically, horizontally)
        - ☑️ Nested tables
        - ☑️ Repeat header row on each page
        - ☑️ Control row page breaks
        - ☑️ Set row height and table width
- ☑️ Images/Pictures (supports BMP, GIF, JPEG, PNG, TIFF, EMF with various wrapping options)
    - ☑️ Add images from file to Word
    - ☑️ Add images from Base64 strings
    - ☑️ Save image from Word to File
    - ☑️ Crop images and set transparency
    - ☑️ Image positioning and location retrieval
    - ◼️ Other location types
- ☑️ Hyperlinks
    - ☑️ Add HyperLink
    - ☑️ Read HyperLink
    - ☑️ Remove HyperLink
    - ☑️ Change HyperLink
- ☑️ PageBreaks
    - ☑️ Add PageBreak
    - ☑️ Read PageBreak
    - ☑️ Remove PageBreak
    - ☑️ Change PageBreak
- ☑️ Page numbering
    - ☑️ Insert page numbers in headers or footers
    - ☑️ Choose style with `WordPageNumberStyle`
- ☑️ Bookmarks
    - ☑️ Add Bookmark
    - ☑️ Read Bookmark
    - ☑️ Remove Bookmark
    - ☑️ Change Bookmark
- ☑️ Find and replace text
  - ☑️ Comments
      - ☑️ Add comments
      - ☑️ Read comments
      - ☑️ Remove comments (single or all)
      - ☑️ Track comments
      - ☑️ Threaded replies
- ☑️ Fields
    - ☑️ Add Field
    - ☑️ Read Field
    - ☑️ Remove Field
    - ☑️ Change Field
- ☑️ Footnotes
    - ☑️ Add new footnotes
    - ☑️ Read footnotes
    - ☑️ Remove footnotes
- ☑️ Endnotes
    - ☑️Add new endnotes
    - ☑️Read endnotes
    - ☑️Remove endnotes
- ☑️ Document variables
    - ☑️ Set and read variables stored in the document
    - ☑️ Remove variables by name or index
- ☑️ Macros
    - ☑️ Add or extract VBA projects
    - ☑️ Remove macro modules
- ☑️ Mail merge
    - ☑️ Replace `MERGEFIELD` values
    - ☑️ Optionally keep field codes
- ☑️ Content Controls
    - ☑️ Add controls
    - ☑️ Read controls
    - ☑️ Update control text
    - ☑️ Remove controls
    - ☑️ Checkbox form controls
    - ☑️ Date picker form controls
    - ☑️ Dropdown list form controls
    - ☑️ Combo box form controls
    - ☑️ Picture form controls
    - ☑️ Repeating section form controls
- ☑️ Shapes
    - ☑️ Add rectangles
    - ☑️ Add ellipses
    - ☑️ Add lines
    - ☑️ Add polygons
    - ☑️ Set fill and stroke color
    - ☑️ Remove shapes
- ☑️ Charts
    - ☑️ Add charts
        - ☑️ Pie and Pie 3D
        - ☑️ Bar and Bar 3D
        - ☑️ Line and Line 3D
        - ☑️ Combo (Bar + Line)
        - ☑️ Area and Area 3D
        - ☑️ Scatter
        - ☑️ Radar
    - ☑️ Add categories and legends
    - ☑️ Configure axes
    - ☑️ Add multiple series
    - ⚠️ When mixing bar and line series call `AddChartAxisX` before adding
      any data so both chart types share the same category axis.
- ☑️ Equations
    - ☑️ Insert Office Math equations from OMML
    - ☑️ Remove equations when needed
- ☑️ Lists
    - ☑️ Add lists
    - ☑️ Remove lists
    - ☑️ Clone lists preserving numbering settings
    - ☑️ Add picture bullet lists
    - ☑️ Create custom bullet and numbered lists
    - ☑️ Detect list style from existing paragraphs
- ☑️ Table of contents
    - ☑️ Add TOC
    - ☑️ Update TOC fields on open
- ☑️ Borders
    - ☑️ Built-in styles or custom settings
    - ☑️ Change size, color, style and spacing
- ☑️ Background
    - ☑️ Set background color
- ☑️ Watermarks
    - ☑️ Add text or image watermark
    - ☑️ Choose text or image style via `WordWatermarkStyle`
    - ☑️ Set rotation, width and height
    - ☑️ Remove watermark

- ☑️ Cover pages
    - ☑️ Add built-in cover pages

- ☑️ Fonts
    - ☑️ Embed fonts via `WordDocument.EmbedFont`

- ☑️ Embedded content
    - ☑️ Add embedded documents (RTF, HTML, TXT)
    - ☑️ Add HTML fragments
    - ☑️ Insert HTML fragment after a paragraph
    - ☑️ Replace text with an HTML fragment
    - ☑️ Remove embedded documents
    - ☑️ Embed objects with custom icons and sizes
 - ☑️ [Digital signatures and document security](OfficeIMO.Tests/Word.DigitalSignature.cs)
 - ☑️ [Document protection options](OfficeIMO.Tests/Word.Settings.cs) (final document, read-only recommended, read-only enforced)
 - ☑️ [Accepting/rejecting revisions](OfficeIMO.Tests/Word.Revisions.cs)
 - ☑️ [Async save/load APIs](OfficeIMO.Tests/Word.Async.cs)
 - ☑️ [Merging multiple documents](OfficeIMO.Tests/Word.MergeDocuments.cs)
 - ☑️ [Text boxes with positioning options](OfficeIMO.Tests/Word.TextBox.cs)
 - ☑️ [Page orientation, page size, and margin presets](OfficeIMO.Tests/Word.PageSettings.cs) ([margins](OfficeIMO.Tests/Word.Sections.cs))
 - ☑️ [Tab characters and custom tab stops](OfficeIMO.Tests/Word.TabStops.cs)
 - ☑️ [Document validation utilities](OfficeIMO.Tests/Word.Validation.cs)
 - ☑️ [CleanupDocument method](OfficeIMO.Tests/Word.Cleanup.cs) merges identical runs
 - ☑️ [Paragraph XML serialization](OfficeIMO.Examples/Word/XmlSerialization/XmlSerialization.Basic.cs)
 - ☑️ [Measurement unit conversion helpers](OfficeIMO.Tests/Word.HelpersConversions.cs)

- ☑️ Experimental Excel component
    - ☑️ Create and load workbooks
    - ☑️ Add worksheets
    - ☑️ Async save and load APIs


## Features (oneliners):

This list of features is for times when you want to quickly fix something rather than playing with full features.
These features are available as part of `WordHelpers` class.

- ☑️ Remove Headers and Footers from a file
- ☑️ Convert DOTX template to DOCX

## Examples

### Basic Document with few document properties and paragraph

This short example shows how to create a Word document with just one paragraph of text and a few document properties.

```csharp
string filePath = System.IO.Path.Combine(
    "Support",
    "GitHub",
    "PSWriteOffice",
    "Examples",
    "Documents",
    "BasicDocument.docx");

using (WordDocument document = WordDocument.Create(filePath)) {
    document.BuiltinDocumentProperties.Title = "This is my title";
    document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
    document.BuiltinDocumentProperties.Keywords = "word, docx, test";

    var paragraph = document.AddParagraph("Basic paragraph");
    paragraph.ParagraphAlignment = JustificationValues.Center;
    paragraph.Color = SixLabors.ImageSharp.Color.Red;

document.Save(true);
}
```

### Creating documents directly in a stream

This overload allows generating a document entirely in memory or on any provided stream.

```csharp
using var stream = new MemoryStream();
using (var document = WordDocument.Create(stream)) {
    document.AddParagraph("Stream based document");
    document.Save(stream);
}
stream.Position = 0;
using (var loaded = WordDocument.Load(stream)) {
    Console.WriteLine(loaded.Paragraphs[0].Text);
}
```

### Saving as a new document

`SaveAs` clones the current document to a new path and returns a new `WordDocument` instance without changing the original `FilePath`.

```csharp
using (WordDocument document = WordDocument.Create()) {
    document.AddParagraph("Some text");
    using var copy = document.SaveAs(filePath);
    // document.FilePath is still null
    // copy.FilePath equals filePath
}
```

### Saving to byte arrays and streams

`SaveAsByteArray` and `SaveAsMemoryStream` let you keep everything in memory. `SaveAs(Stream)` clones the document to a provided stream and returns a new instance loaded from it.

```csharp
using (WordDocument document = WordDocument.Create()) {
    document.AddParagraph("In memory");
    byte[] data = document.SaveAsByteArray();
    using MemoryStream stream = document.SaveAsMemoryStream();
    using var clone = document.SaveAs(stream);
}
```

### Embedding a font

`EmbedFont` adds a TrueType or OpenType font file to the document so it can be used on any machine.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.EmbedFont(fontPath);
    document.AddParagraph("This document uses an embedded font.");
    document.Save();
}
```

`EmbedFont` has an overload that can register a paragraph style automatically:

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.EmbedFont(fontPath, "DejaVuStyle", "DejaVu Style");
    document.AddParagraph("Styled paragraph").SetStyleId("DejaVuStyle");
    document.Save();
}
```

### Mixing builtin and embedded fonts

After embedding a font you can reference it by name along with standard fonts.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.EmbedFont(fontPath);
    document.AddParagraph("Builtin Arial").SetFontFamily("Arial");
    document.AddParagraph("Embedded DejaVu Sans").SetFontFamily("DejaVu Sans");
    var mixed = document.AddParagraph("Mix: ");
    mixed.AddText("Arial, ").SetFontFamily("Arial");
    mixed.AddText("DejaVu").SetFontFamily("DejaVu Sans");
    document.Save();
}
```

You can also create a paragraph style that uses the embedded font.

```csharp
var style = new Style { Type = StyleValues.Paragraph, StyleId = "EmbeddedStyle" };
style.Append(new StyleName { Val = "EmbeddedStyle" });
style.Append(new StyleRunProperties(new RunFonts { Ascii = "DejaVu Sans" }));
WordParagraphStyle.RegisterCustomStyle("EmbeddedStyle", style);

using (WordDocument document = WordDocument.Create(filePath)) {
    document.EmbedFont(fontPath);
    document.AddParagraph("Paragraph with embedded style").SetStyleId("EmbeddedStyle");
    document.AddParagraph("Built-in style example").Style = WordParagraphStyles.Normal;
    document.Save();
}
```

### Basic Document with Headers/Footers (first, odd, even)

This short example shows how to add headers and footers to a Word document.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
    document.AddParagraph("Test Section0");
    document.AddHeadersAndFooters();
    document.DifferentFirstPage = true;
    document.DifferentOddAndEvenPages = true;

    document.Sections[0].Header.First.AddParagraph().SetText("Test Section 0 - First Header");
    document.Sections[0].Header.Default.AddParagraph().SetText("Test Section 0 - Header");
    document.Sections[0].Header.Even.AddParagraph().SetText("Test Section 0 - Even");

    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();

    var section1 = document.AddSection();
    section1.PageOrientation = PageOrientationValues.Portrait;
    section1.AddParagraph("Test Section1");
    section1.AddHeadersAndFooters();
    section1.Header.Default.AddParagraph().SetText("Test Section 1 - Header");
    section1.DifferentFirstPage = true;
    section1.Header.First.AddParagraph().SetText("Test Section 1 - First Header");

    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();
    document.AddPageBreak();

    var section2 = document.AddSection();
    section2.AddParagraph("Test Section2");
    section2.PageOrientation = PageOrientationValues.Landscape;
    section2.AddHeadersAndFooters();
    section2.Header.Default.AddParagraph().SetText("Test Section 2 - Header");

    document.AddParagraph("Test Section2 - Paragraph 1");

    var section3 = document.AddSection();
    section3.AddParagraph("Test Section3");
    section3.AddHeadersAndFooters();
    section3.Header.Default.AddParagraph().SetText("Test Section 3 - Header");

    Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
    Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
    Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
    Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
    Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);

    Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Header.Default.Paragraphs[0].Text);
    Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Header.Default.Paragraphs[0].Text);
    Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Header.Default.Paragraphs[0].Text);
    Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Header.Default.Paragraphs[0].Text);
    document.Save(true);
}
```

### Adding a Content Control

This example shows how to add and update a simple content control and then retrieve it by tag.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var sdt = document.AddStructuredDocumentTag("Hello", "MyAlias", "MyTag");
    sdt.Text = "Changed";
    document.Save(true);
}

using (WordDocument document = WordDocument.Load(filePath)) {
    var tag = document.GetStructuredDocumentTagByTag("MyTag");
Console.WriteLine(tag.Text);
}
```

### Multiple Content Controls

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddStructuredDocumentTag("First", "Alias1", "Tag1");
    document.AddStructuredDocumentTag("Second", "Alias2", "Tag2");
    document.AddStructuredDocumentTag("Third", "Alias3", "Tag3");
    document.Save(true);
}

using (WordDocument document = WordDocument.Load(filePath)) {
    foreach (var control in document.StructuredDocumentTags) {
        Console.WriteLine(control.Tag + ": " + control.Text);
    }
}
```

### Advanced Content Control Usage

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddStructuredDocumentTag("First", "Alias1", "Tag1");
    document.AddStructuredDocumentTag("Second", "Alias2", "Tag2");
    document.Save(true);
}

using (WordDocument document = WordDocument.Load(filePath)) {
    var alias = document.GetStructuredDocumentTagByAlias("Alias2");
    alias.Text = "Updated";
    var tag = document.GetStructuredDocumentTagByTag("Tag1");
    Console.WriteLine(tag.Text);
}
```

### Advanced usage of OfficeIMO

This short example shows multiple features of `OfficeIMO.Word`

```csharp
string filePath = System.IO.Path.Combine(folderPath, "AdvancedDocument.docx");
using (WordDocument document = WordDocument.Create(filePath)) {
    // lets add some properties to the document
    document.BuiltinDocumentProperties.Title = "Cover Page Templates";
    document.BuiltinDocumentProperties.Subject = "How to use Cover Pages with TOC";
    document.ApplicationProperties.Company = "Evotec Services";

    // we force document to update fields on open, this will be used by TOC
    document.Settings.UpdateFieldsOnOpen = true;

    // lets add one of multiple added Cover Pages
    document.AddCoverPage(CoverPageTemplate.IonDark);

    // lets add Table of Content (1 of 2)
    document.AddTableOfContent(TableOfContentStyle.Template1);

    // lets add page break
    document.AddPageBreak();

    // lets create a list that will be binded to TOC
    var wordListToc = document.AddTableOfContentList(WordListStyle.Headings111);

    wordListToc.AddItem("How to add a table to document?");

    document.AddParagraph("In the first paragraph I would like to show you how to add a table to the document using one of the 105 built-in styles:");

    // adding a table and modifying content
    var table = document.AddTable(5, 4, WordTableStyle.GridTable5DarkAccent5);
    table.Rows[3].Cells[2].Paragraphs[0].Text = "Adding text to cell";
    table.Rows[3].Cells[2].Paragraphs[0].Color = Color.Blue; ;
    table.Rows[3].Cells[3].Paragraphs[0].Text = "Different cell";

    document.AddParagraph("As you can see adding a table with some style, and adding content to it ").SetBold().SetUnderline(UnderlineValues.Dotted).AddText("is not really complicated").SetColor(Color.OrangeRed);

    wordListToc.AddItem("How to add a list to document?");

    var paragraph = document.AddParagraph("Adding lists is similar to ading a table. Just define a list and add list items to it. ").SetText("Remember that you can add anything between list items! ");
    paragraph.SetColor(Color.Blue).SetText("For example TOC List is just another list, but defining a specific style.");

    var list = document.AddList(WordListStyle.Bulleted);
    list.AddItem("First element of list", 0);
    list.AddItem("Second element of list", 1);

    var paragraphWithHyperlink = document.AddHyperLink("Go to Evotec Blogs", new Uri("https://evotec.xyz"), true, "URL with tooltip");
    // you can also change the hyperlink text, uri later on using properties
    paragraphWithHyperlink.Hyperlink.Uri = new Uri("https://evotec.xyz/hub");
    paragraphWithHyperlink.ParagraphAlignment = JustificationValues.Center;

    list.AddItem("3rd element of list, but added after hyperlink", 0);
    list.AddItem("4th element with hyperlink ").AddHyperLink("included.", new Uri("https://evotec.xyz/hub"), addStyle: true);

    document.AddParagraph();

    // create a custom bullet list
    var custom = document.AddCustomBulletList(WordListLevelKind.BulletSquareSymbol, "Courier New", SixLabors.ImageSharp.Color.Red, fontSize: 16);
    custom.AddItem("Custom bullet item");

    // create a list using an image as the bullet
    var pictureList = document.AddPictureBulletList(Path.Combine(folderPath, "Images", "Kulek.jpg"));
    pictureList.AddItem("Image bullet 1");
    pictureList.AddItem("Image bullet 2");

    // create a multi-level custom list
    var builder = document.AddCustomList()
        .AddListLevel(1, WordListLevelKind.BulletSquareSymbol, "Courier New", SixLabors.ImageSharp.Color.Red, fontSize: 14)
        .AddListLevel(5, WordListLevelKind.BulletBlackCircle, "Arial", colorHex: "#00ff00", fontSize: 10);
    builder.AddItem("First");
    builder.AddItem("Fifth", 4);

    // Note: use AddCustomList() rather than AddList(WordListStyle.Custom)
    // when you want to build lists with your own levels.
    // See [Custom Lists](Docs/custom-lists.md) for details on configuring levels.

    var listNumbered = document.AddList(WordListStyle.Heading1ai);
    listNumbered.AddItem("Different list number 1");
    listNumbered.AddItem("Different list number 2", 1);
    listNumbered.AddItem("Different list number 3", 1);
    listNumbered.AddItem("Different list number 4", 1);

    var section = document.AddSection();
    section.PageOrientation = PageOrientationValues.Landscape;
    section.PageSettings.PageSize = WordPageSize.A4;

    wordListToc.AddItem("Adding headers / footers");

    // lets add headers and footers
    document.AddHeadersAndFooters();

    // adding text to default header
    document.Header.Default.AddParagraph("Text added to header - Default");

    var section1 = document.AddSection();
    section1.PageOrientation = PageOrientationValues.Portrait;
    section1.PageSettings.PageSize = WordPageSize.A5;

    wordListToc.AddItem("Adding custom properties to document");

    document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty { Value = DateTime.Today });
    document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Some text"));
    document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

    // document variables available via DocVariable fields
    document.SetDocumentVariable("Project", "OfficeIMO");
    document.SetDocumentVariable("Year", DateTime.Now.Year.ToString());

    if (document.HasDocumentVariables) {
        foreach (var pair in document.DocumentVariables) {
            Console.WriteLine($"{pair.Key}: {pair.Value}");
        }
    }

    document.Save(openWord);
}
```

## Tests

In addition to the fact that `OfficeIMO.Word` uses Unit Tests, [Characterization Tests](https://en.wikipedia.org/wiki/Characterization_test) are also used.
Characterization test were added in order to not overlook a change that breaks the behavior. These tests are based on [Verify](https://github.com/VerifyTests/Verify) (["Snapshot Testing in .NET with Verify"](https://youtu.be/wA7oJDyvn4c)).
if you need to add or update a verified snapshot, you can use the powershell script:
```bash
$ pwsh -c ./Build/ApproveVerifyTests.ps1
```
To show a graphical diff instead of console output when Verify tests fail, set
the environment variable `DiffEngine_Disabled=false` before running the tests.
