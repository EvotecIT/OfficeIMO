## 0.5.0

- Add `DistributeColumnsEvenly` and `SetWidthPercentage` methods to Tables

```cs
internal static void Example_DifferentTableSizes(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with tables of different sizes");
    string filePath = System.IO.Path.Combine(folderPath, "Document with Tables of different sizes.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        document.AddParagraph("Table 1");
        WordTable wordTable = document.AddTable(3, 4, WordTableStyle.PlainTable1);
        wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

        document.AddParagraph();
        document.AddParagraph();
        document.AddParagraph("Table 2 - Sized for 2000 width / Centered");
        WordTable wordTable1 = document.AddTable(2, 6, WordTableStyle.PlainTable1);
        wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
        wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 1 - ok longer text, no autosize right?";
        wordTable1.WidthType = TableWidthUnitValues.Pct;
        wordTable1.Width = 2000;
        wordTable1.Alignment = TableRowAlignmentValues.Center;

        document.AddParagraph();
        document.AddParagraph();
        document.AddParagraph("Table 3 - By default the table is autosized for full width");
        WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
        wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

        document.AddParagraph();
        document.AddParagraph();
        document.AddParagraph("Table 4 - Magic number 5000 (full width)");
        WordTable wordTable3 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
        wordTable3.WidthType = TableWidthUnitValues.Pct;
        wordTable3.Width = 5000;
        wordTable3.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

        document.AddParagraph();
        document.AddParagraph();
        document.AddParagraph("Table 5 - 50% by using 2500 width (pct)");
        WordTable wordTable4 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
        wordTable4.WidthType = TableWidthUnitValues.Pct;
        wordTable4.Width = 2500;
        wordTable4.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

        document.AddParagraph();
        document.AddParagraph();
        document.AddParagraph("Table 6 - 50% by using 2500 width (pct), that we fix to full width");
        WordTable wordTable5 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
        // this data is temporary just to prove things work
        wordTable5.WidthType = TableWidthUnitValues.Pct;
        wordTable5.Width = 2500;
        // lets fix it for full width
        wordTable5.DistributeColumnsEvenly();

        document.AddParagraph();
        document.AddParagraph();
        document.AddParagraph("Table 6 - 50%");
        WordTable wordTable6 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
        wordTable6.SetWidthPercentage(50);

        document.AddParagraph();
        document.AddParagraph();
        document.AddParagraph("Table 6 - 75%");
        WordTable wordTable7 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
        wordTable7.SetWidthPercentage(75);

        document.Save(openWord);
    }
}
```

## 0.4.6 - 2023.01.11

### What's Changed
* Added tests for fragmented instruction parsing by @byteSamurai in https://github.com/EvotecIT/OfficeIMO/pull/97
* Add paragraphs property to start on a new page by @byteSamurai in https://github.com/EvotecIT/OfficeIMO/pull/98

## 0.4.5 - 2023.01.09

### What's Changed
* Merge instructions distributed over runs to field string by @byteSamurai in https://github.com/EvotecIT/OfficeIMO/pull/95

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.4...v0.4.5

## 0.4.4 - 2023.01.09

### What's Changed
* Fixes HyperLink AddStyle not being applied by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/93

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.3...v0.4.4

## 0.4.3 - 2023.01.09

### What's Changed
* Improve tests, fix HyperLink order, add method ValidateDocument() by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/87
* Fix badges by @rstm-sf in https://github.com/EvotecIT/OfficeIMO/pull/88
* Ability to insert table before/after paragraph by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/92
  * Paragraph.AddTableAfter()
  * Paragraph.AddTableBefore()

```csharp
internal static void Example_TablesAddedAfterParagraph(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with width and alignment");
    string filePath = System.IO.Path.Combine(folderPath, "Document with Table Alignment.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        var paragraph = document.AddParagraph("Lets add table with some alignment ");
        paragraph.ParagraphAlignment = JustificationValues.Center;
        paragraph.Bold = true;
        paragraph.Underline = UnderlineValues.DotDash;

        WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
        wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
        wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
        wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
        wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

        var paragraph1 = document.AddParagraph("Lets add another table showing text wrapping around, but notice table before and after it anyways, that we just added at the end of the document.");

        WordTable wordTable1 = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
        wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
        wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
        wordTable1.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
        wordTable1.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

        wordTable1.WidthType = TableWidthUnitValues.Pct;
        wordTable1.Width = 3000;

        wordTable1.AllowTextWrap = true;

        var paragraph2 = document.AddParagraph("This paragraph should continue but next to to the table");

        document.AddParagraph();
        document.AddParagraph();

        var paragraph3 = document.AddParagraph("Lets add another table showing AutoFit");

        WordTable wordTable2 = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
        wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
        wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
        wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
        wordTable2.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";


        paragraph1.AddParagraphBeforeSelf();
        paragraph1.AddParagraphAfterSelf();

        var table3 = paragraph1.AddTableAfter(4, 4, WordTableStyle.GridTable1LightAccent1);
        table3.Rows[0].Cells[0].Paragraphs[0].Text = "Inserted in the middle of the document after paragraph";

        var table4 = paragraph1.AddTableBefore(4, 4, WordTableStyle.GridTable1LightAccent1);
        table4.Rows[0].Cells[0].Paragraphs[0].Text = "Inserted in the middle of the document before paragraph";

        document.Save(openWord);
    }
}
```

## 0.4.2 - 2022.11.25

### What's Changed
* The order of paragraph properties and runs does matter by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/83
* This release fixes LISTS. Sorry for that!

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.1...v0.4.2

## 0.4.1 - 2022.11.20

### What's Changed
* Extending available FieldCodes and allow setting parameters by @byteSamurai in https://github.com/EvotecIT/OfficeIMO/pull/75
* Fix check IsToc for Word.Lists by @rstm-sf in https://github.com/EvotecIT/OfficeIMO/pull/76
* Provide multi platform project config by @byteSamurai in https://github.com/EvotecIT/OfficeIMO/pull/79

### New Contributors
* @byteSamurai made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/75

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.4.0...v0.4.1


## 0.4.0 - 2022.11.13

## What's Changed
* Improve settings for document with PT-BR and HighAnsi chars by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/56
  - FontFamilyHighAnsi for document.Settings which is required for special chars

```csharp
public static void Example_BasicWordWithDefaultFontChange(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with different default style (PT/BR)");
    string filePath = System.IO.Path.Combine(folderPath, "BasicWordWithDefaultStyleChangeBR.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        document.Settings.FontSize = 30;
        //document.Settings.FontSizeComplexScript = 30;
        document.Settings.FontFamily = "Calibri Light";
        document.Settings.FontFamilyHighAnsi = "Calibri Light";
        document.Settings.Language = "pt-Br";

        string title = "INSTRUMENTO PARTICULAR DE CONSTITUIÇÃO DE GARANTIA DE ALIENAÇÃO FIDUCIÁRIA DE IMÓVEL";

        document.AddParagraph(title).SetBold().ParagraphAlignment = JustificationValues.Center;

        document.Save(openWord);
    }
}
```

* Add Compatibility Mode, set default compatibility to the highest version by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/58
  - Adds default settings, including compatibility settings, math properties, and few other things that are automatically added when Microsoft Word creates a document
  - Adds default web settings that are the same to what Microsoft Word does
  - Sample of usage for compatiblity mode

```csharp
public static void Example_BasicWordWithDefaultFontChange(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with different default style (PT/BR)");
    string filePath = System.IO.Path.Combine(folderPath, "BasicWordWithDefaultStyleChangeBR.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        document.Settings.FontSize = 30;
        //document.Settings.FontSizeComplexScript = 30;
        document.Settings.FontFamily = "Calibri Light";
        document.Settings.FontFamilyHighAnsi = "Calibri Light";
        document.Settings.Language = "pt-Br";

        document.Settings.ZoomPreset = PresetZoomValues.BestFit;

        Console.WriteLine(document.CompatibilitySettings.CompatibilityMode);

        document.CompatibilitySettings.CompatibilityMode = CompatibilityMode.Word2013;

        Console.WriteLine(document.CompatibilitySettings.CompatibilityMode);

        document.CompatibilitySettings.CompatibilityMode = CompatibilityMode.None;

        Console.WriteLine(document.CompatibilitySettings.CompatibilityMode);

        string title = "INSTRUMENTO PARTICULAR DE CONSTITUIÇÃO DE GARANTIA DE ALIENAÇÃO FIDUCIÁRIA DE IMÓVEL";

        document.AddParagraph(title).SetBold().ParagraphAlignment = JustificationValues.Center;

        document.Save(openWord);
    }
}
```

* Bump nuggets for tests and excel by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/59
* Fix not saving list when saving to Stream by @rstm-sf in https://github.com/EvotecIT/OfficeIMO/pull/62
* Add an image via stream by @rstm-sf in https://github.com/EvotecIT/OfficeIMO/pull/51
* Add images to headers/footers/tables and other image improvements by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/53

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.BuiltinDocumentProperties.Title = "This is sparta";
    document.BuiltinDocumentProperties.Creator = "Przemek";
    var filePathImage = System.IO.Path.Combine(imagePaths, "Kulek.jpg");

    document.AddHeadersAndFooters();

    var header = document.Header.Default;
    var paragraphHeader = header.AddParagraph("This is header");

    // add image to header, directly to paragraph
    header.AddParagraph().AddImage(filePathImage, 100, 100);

    // add image to footer, directly to paragraph
    document.Footer.Default.AddParagraph().AddImage(filePathImage, 100, 100);

    // add image to header, but to a table
    var table = header.AddTable(2, 2);
    table.Rows[1].Cells[1].Paragraphs[0].Text = "Test123";
    table.Rows[1].Cells[0].Paragraphs[0].AddImage(filePathImage, 50, 50);
    table.Alignment = TableRowAlignmentValues.Right;

    var paragraph = document.AddParagraph("This paragraph starts with some text");
    paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";
    paragraph.Bold = true;

    // add table with an image, but to document
    var table1 = document.AddTable(2, 2);
    table1.Rows[1].Cells[1].Paragraphs[0].Text = "Test - In document";
    table1.Rows[1].Cells[0].Paragraphs[0].AddImage(filePathImage, 50, 50);


    // lets add image to paragraph
    paragraph.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22);
}
```

- Rotation / VerticalFlip / HorizontalFlip / Shape work for images

```csharp
internal static void Example_AddingImagesInline(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with inline images");
    string filePath = System.IO.Path.Combine(folderPath, "DocumentWithInlineImages2.docx");
    string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

    using (WordDocument document = WordDocument.Create(filePath)) {
        var file = System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg");
        var paragraph = document.AddParagraph();
        var pargraphWithImage = paragraph.AddImage(file, 100, 100);

        // Console.WriteLine("Image is inline: " + pargraphWithImage.Image.Rotation);

        pargraphWithImage.Image.VerticalFlip = false;
        pargraphWithImage.Image.HorizontalFlip = false;
        pargraphWithImage.Image.Rotation = 270;
        pargraphWithImage.Image.Shape = ShapeTypeValues.Cloud;


        document.Save(openWord);
    }
}
```

- Added optional Description (AltText) to images
- Added ability to wrap text around image

```cs
internal static void Example_AddingImagesSample4(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with some Images and Samples");
    var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithImagesSample4.docx");
    var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

    using var document = WordDocument.Create(filePath);

    var paragraph1 = document.AddParagraph("This paragraph starts with some text");
    paragraph1.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200);
    paragraph1.Image.Shape = ShapeTypeValues.Cube;

    var paragraph2 = document.AddParagraph("Image will be placed behind text");
    paragraph2.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.BehindText, "Przemek and Kulek on an image");


    var paragraph3 = document.AddParagraph("Image will be in front of text");
    paragraph3.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.InFrontText, "Przemek and Kulek on an image");


    var paragraph5 = document.AddParagraph("Image will be Square");
    paragraph5.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.Square, "Przemek and Kulek on an image");


    var paragraph6 = document.AddParagraph("Image will be Through");
    paragraph6.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.Through, "Przemek and Kulek on an image");


    var paragraph7 = document.AddParagraph("Image will be Tight");
    paragraph7.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.Tight, "Przemek and Kulek on an image");


    var paragraph8 = document.AddParagraph("Image will be Top And Bottom");
    paragraph8.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200, WrapImageText.TopAndBottom, "Przemek and Kulek on an image");
    paragraph8.Image.Shape = ShapeTypeValues.Can;

    document.Save(openWord);
}
```

* Fixes adding a bookmark on a document by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/68
* Detect read-only before saving by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/67
* Skip creation of run, which isn't needed by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/69
* Add net 7.0, bump dependencies for testing by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/72

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.3.1...v0.4.0

## 0.3.1 - 2022.10.18
- Adds BMP, GIF, PNG, TIFF image formats support [#42](https://github.com/EvotecIT/OfficeIMO/pull/42) by rstm-sf
- Fixes issue with read only documents crashing
- Fixes watermark throwing when Headers aren't added first [#46](https://github.com/EvotecIT/OfficeIMO/pull/46)
- Fixes Adding Watermark to section that has not initialized headers will throw an error [#27](https://github.com/EvotecIT/OfficeIMO/issues/27)
- New fields added to WordSettings allowing setting of default values for the whole document for FontSize, FontSizeComplexScript, FontFamily and Language

```csharp
public static void Example_BasicWordWithDefaultStyleChange(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with different default style");
    string filePath = System.IO.Path.Combine(folderPath, "BasicWordWithDefaultStyleChange.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        document.Settings.FontSize = 30;
        document.Settings.FontFamily = "Calibri Light";
        document.Settings.Language = "pl-PL";

        var paragraph1 = document.AddParagraph("To jest po polsku");

        var paragraph2 = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER");
        paragraph2.FontSize = 15;
        paragraph2.FontFamily = "Courier New";

        document.Save(openWord);
    }
}

```
- Add `IsLastRun` and `IsFirstRun` properties to WordParagraph

## 0.3.0 - 2022.10.11
- Update DocumentFormat.OpenXml from 2.16.0 to 2.18.0
- Update SixLabors.ImageSharp to 2.1.3

- Adds ability to add nested table into existing table.

```csharp
internal static void Example_NestedTables(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with nested tables");
    string filePath = System.IO.Path.Combine(folderPath, "Document with Nested Tables.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        var paragraph = document.AddParagraph("Lets add table ");
        paragraph.ParagraphAlignment = JustificationValues.Center;
        paragraph.Bold = true;
        paragraph.Underline = UnderlineValues.DotDash;

        WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
        wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
        wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
        wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
        wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

        wordTable.Rows[0].Cells[0].AddTable(3, 2, WordTableStyle.GridTable2Accent2);

        wordTable.Rows[0].Cells[1].AddTable(3, 2, WordTableStyle.GridTable2Accent5, true);

        document.Save(openWord);
    }
}
```

- Adds `NestedTables` property for WordTable to get all nested tables for given table
- Adds `HasNestedTables` property for WordTable to know if table has nested tables
- Adds `IsNestedTable` property for WordTable to know if table is nested table
- Adds `ParentTable` property for WordTable to find parent table if the table is nested
- Added some summaries to multiple table related methods/properties
- Adds `TablesIncludingNestedTables` property to Sections and Document to make it easy to find all tables within document and manipulate them

- Solves an issue with different word break required https://github.com/EvotecIT/OfficeIMO/issues/37

![image](https://user-images.githubusercontent.com/15063294/194375619-a43d0d20-50b9-445a-a8c7-d778a4c3fac2.png)

```csharp
public static void Example_BasicWordWithBreaks(string folderPath, bool openWord) {
    Console.WriteLine("[*] Creating standard document with paragraph & breaks");
    string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithParagraphsAndBreaks.docx");
    using (WordDocument document = WordDocument.Create(filePath)) {
        var paragraph1 = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER");


        var paragraph2 = document.AddParagraph("Adding paragraph2 with some text and pressing SHIFT+ENTER");
        paragraph2.AddBreak();
        paragraph2.AddText("Continue1");
        paragraph2.AddBreak();
        paragraph2.AddText("Continue2");

        var paragraph3 = document.AddParagraph("Adding paragraph3 with some text and pressing ENTER");


        document.Save(openWord);
    }
}
```

Additionally:
- Renames WordPageBreak to WordBreak to accommodate all Breaks, and not only PageBreak
- *BREAKING CHANGE* Removing WordBreak (or WordPageBreak) no longer by default removes paragraph, but instead requires bool set to true

```csharp
document.Breaks[0].Remove();
document.Breaks[0].Remove(includingParagraph: true);
```

- Add new `IsBreak` property for WordParagraph
- Add `Breaks` property for WordDocument

- Implement Save and Load to/from Stream [#43](https://github.com/EvotecIT/OfficeIMO/pull/43) tnx hisuwh

## - 0.2.1 - 2022.07.31
- Added basic support for Charts [#14](https://github.com/EvotecIT/OfficeIMO/pull/14)

## - 0.2.0 - 2022.07.31

- Added MIT License [#28](https://github.com/EvotecIT/OfficeIMO/pull/28)

- Adds new properties for Tables by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ Alignment
  - ☑️ WidthType
  - ☑️ Width
  - ☑️ ShadingFillColor
  - ☑️ ShadingFillColorHex
  - ☑️ ShadingFillPatern
  - ☑️ Title
  - ☑️ Description
  - ☑️ AllowOverlap
  - ☑️ AllowTextWrap
  - ☑️ ColumnWidth
  - ☑️ RowHeight

- Add table positioning along with properties by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ LeftFromText
  - ☑️ RightFromText
  - ☑️ BottomFromText
  - ☑️ TopFromText
  - ☑️ HorizontalAnchor
  - ☑️ TablePositionY
  - ☑️ TablePositionX
  - ☑️ TablePositionYAlignment
  - ☑️ TablePositionXAlignment
  - ☑️ TableOverlap

- Adds new properties for TableRow by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ FirstCell
  - ☑️ LastCell

- Renames some properties to better name them by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - FirstRow -> ConditionalFormattingFirstRow
  - LastRow -> ConditionalFormattingLastRow
  - FirstColumn -> ConditionalFormattingFirstColumn
  - LastColumn -> ConditionalFormattingLastColumn
  - NoHorizontalBand -> ConditionalFormattingNoHorizontalBand
  - NoVerticalBand -> ConditionalFormattingNoVerticalBand

- Adds new properties for Table by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ FirstRow
  - ☑️ LastRow

- Adds new methods for Table by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ AddComment(author, initials,comment)

- Adds new properties for TableCell by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - ☑️ TextDirection

## - 0.1.7 - 2022.06.12

### What's Changed
* Fixes **PageOrientation** of page/section if set before page size is applied. In case that happens it always reverted back to `Portrait` mode which is default for newly set **PageSizes**.
* Fixes **PageSize** detection when in `PageOrientationValues.Landscape` mode.

## - 0.1.6 - 2022.06.11

### What's Changed
* Rename `Color` to `ColorHex` property for Paragraphs *BREAKING CHANGE*
* Add `Color` property for Paragraphs as `SixLabors.ImageSharp.Color` *BREAKING CHANGE*

For example:

```csharp
var paragraph = document.AddParagraph("Basic paragraph");
paragraph.ParagraphAlignment = JustificationValues.Center;
paragraph.Color = SixLabors.ImageSharp.Color.Red;
```

```csharp
var paragraph = document.AddParagraph("Basic paragraph");
paragraph.ParagraphAlignment = JustificationValues.Center;
paragraph.ColorHex = "#FFFF00";
```

## - 0.1.5 - 2022.06.04

### What's Changed
* Fixes `TableOfContent.Update()`
* Fixed **SaveAs** functionality by @jordan-hemming in https://github.com/EvotecIT/OfficeIMO/pull/16
* Fixes **Azure Devops** tests to work properly on Linux and MacOs by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/18
* Tables styles are not being applied for loaded documents by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/20
* Add basic support for table cell borders by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/21
  - ☑️ LeftBorder
  - ☑️ RightBorder
  - ☑️ TopBorder
  - ☑️ BottomBorder
  - ☑️ EndBorder
  - ☑️ StartBorder
  - ☑️ InsideHorizontalBorder
  - ☑️ InsideVerticalBorder
  - ☑️ TopLeftToBottomRightBorder
  - ☑️ TopRightToBottomLeftCell
* Add additional tests for table cell borders by @PrzemyslawKlys in https://github.com/EvotecIT/OfficeIMO/pull/22

### New Contributors
* @jordan-hemming made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/16

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.1.4...v0.1.5

## - 0.1.4 - 2022.04.03
- First official release
