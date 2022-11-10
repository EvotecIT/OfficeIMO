#### 0.3.1 - 2022.10.18
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

#### 0.3.0 - 2022.10.11
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

#### - 0.2.1 - 2022.07.31
- Added basic support for Charts [#14](https://github.com/EvotecIT/OfficeIMO/pull/14)

#### - 0.2.0 - 2022.07.31

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

#### - 0.1.7 - 2022.06.12

##### What's Changed
* Fixes **PageOrientation** of page/section if set before page size is applied. In case that happens it always reverted back to `Portrait` mode which is default for newly set **PageSizes**.
* Fixes **PageSize** detection when in `PageOrientationValues.Landscape` mode.

#### - 0.1.6 - 2022.06.11

##### What's Changed
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

#### - 0.1.5 - 2022.06.04

##### What's Changed
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

##### New Contributors
* @jordan-hemming made their first contribution in https://github.com/EvotecIT/OfficeIMO/pull/16

**Full Changelog**: https://github.com/EvotecIT/OfficeIMO/compare/v0.1.4...v0.1.5

#### - 0.1.4 - 2022.04.03
- First official release
