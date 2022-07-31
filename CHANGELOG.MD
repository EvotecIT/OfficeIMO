#### - 0.2.0 - 2022.07.31

- Added MIT License [#28](https://github.com/EvotecIT/OfficeIMO/pull/28)

- Adds new properties for Tables by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - [x] Alignment
  - [x] WidthType
  - [x] Width
  - [x] ShadingFillColor
  - [x] ShadingFillColorHex
  - [x] ShadingFillPatern
  - [x] Title
  - [x] Description
  - [x] AllowOverlap
  - [x] AllowTextWrap
  - [x] ColumnWidth
  - [x] RowHeight

- Add table positioning along with properties by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - [x] LeftFromText
  - [x] RightFromText
  - [x] BottomFromText
  - [x] TopFromText
  - [x] HorizontalAnchor
  - [x] TablePositionY
  - [x] TablePositionX
  - [x] TablePositionYAlignment
  - [x] TablePositionXAlignment
  - [x] TableOverlap

- Adds new properties for TableRow by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - [x] FirstCell
  - [x] LastCell

- Renames some properties to better name them by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - FirstRow -> ConditionalFormattingFirstRow
  - LastRow -> ConditionalFormattingLastRow
  - FirstColumn -> ConditionalFormattingFirstColumn
  - LastColumn -> ConditionalFormattingLastColumn
  - NoHorizontalBand -> ConditionalFormattingNoHorizontalBand
  - NoVerticalBand -> ConditionalFormattingNoVerticalBand

- Adds new properties for Table by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - [x] FirstRow
  - [x] LastRow

- Adds new methods for Table by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - [x] AddComment(author, initials,comment)

- Adds new properties for TableCell by @PrzemyslawKlys in [#30](https://github.com/EvotecIT/OfficeIMO/pull/30)
  - [x] TextDirection

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