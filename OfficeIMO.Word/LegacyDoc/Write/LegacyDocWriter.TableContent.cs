using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendTable(StringBuilder text, List<LegacyDocWritableRun> runs, List<LegacyDocWritableParagraph> paragraphFormats, LegacyDocWritableBookmarksBuilder bookmarks, Table table, MainDocumentPart mainPart, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> styleIndexes, IReadOnlyDictionary<string, Style> tableStyleDefinitions, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, int tableDepth = 1, OpenXmlPart? relationshipOwner = null) {
            relationshipOwner ??= mainPart;
            if (tableDepth <= 0 || tableDepth > 2) {
                throw new NotSupportedException("Native DOC saving supports nested tables only to depth 2.");
            }

            // An ordinary paragraph mark separates adjacent top-level tables
            // in every story; otherwise DOC treats them as one table.
            if (tableDepth == 1 && text.Length > 0 && text[text.Length - 1] == '\a') text.Append('\r');

            ThrowIfUnsupportedTableShape(table, tableStyleDefinitions);

            TableRow[] rows = table.Elements<TableRow>().ToArray();
            if (rows.Length == 0) {
                throw new NotSupportedException("Native DOC saving supports simple tables only when at least one row is present.");
            }

            AppendLeadingTableBoundaryBookmarks(table, bookmarks, text.Length);
            TableProperties? tableProperties = ResolveSupportedEffectiveTableProperties(
                table.GetFirstChild<TableProperties>(), tableStyleDefinitions);
            LegacyDocTableAlignment? tableAlignment = ReadSupportedTableAlignment(tableProperties);
            tableAlignment ??= ReadSupportedTableStyleAlignment(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            int? tableLeftIndentTwips = ReadSupportedTableIndentation(tableProperties);
            tableLeftIndentTwips ??= ReadSupportedTableStyleIndentation(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocTablePreferredWidth? tablePreferredWidth = ReadSupportedTablePreferredWidth(tableProperties);
            tablePreferredWidth ??= ReadSupportedTableStylePreferredWidth(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            bool? tableAutofit = ReadSupportedTableAutofit(tableProperties);
            tableAutofit ??= ReadSupportedTableStyleAutofit(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocTableCellMargins? defaultCellMargins = ReadSupportedTableDefaultCellMargins(tableProperties);
            // Materialize the resolved DOCX inset for stable native DOC layout;
            // Word positions these cells differently when the padding is implicit.
            defaultCellMargins = new LegacyDocTableCellMargins(0, 108, 0, 108)
                .Merge(ReadSupportedTableStyleDefaultCellMargins(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions) ?? default)
                .Merge(defaultCellMargins ?? default);
            int? defaultCellSpacingTwips = ReadSupportedTableDefaultCellSpacing(tableProperties);
            defaultCellSpacingTwips ??= ReadSupportedTableStyleDefaultCellSpacing(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocTableBorders tableBorders = ReadSupportedTableBorders(tableProperties, tableStyleDefinitions);
            LegacyDocTableCellShading tableShading = ReadSupportedTableShading(tableProperties, tableStyleDefinitions);
            LegacyDocWritableParagraphFormatting tableStyleParagraphFormatting = ReadSupportedTableStyleParagraphFormatting(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocWritableFormatting tableStyleRunFormatting = ReadSupportedTableStyleRunFormatting(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocTableConditionalStyleSet conditionalStyles = ReadSupportedTableConditionalStyles(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocTableLook tableLook = ReadSupportedTableLook(tableProperties?.GetFirstChild<TableLook>());
            IReadOnlyList<int> gridColumnWidthsTwips = ReadSupportedTableGridWidths(table.GetFirstChild<TableGrid>());
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
                TableRow row = rows[rowIndex];
                LegacyDocWritableTableRowFormatting rowFormatting = ReadSupportedTableRowFormatting(row, out TableCell[] cells);
                rowFormatting = ApplySupportedTableConditionalRowFormatting(rowFormatting, conditionalStyles, tableLook, rowIndex, rows.Length);
                if (cells.Length == 0) {
                    throw new NotSupportedException("Native DOC saving supports simple tables only when every row contains at least one cell.");
                }

                IReadOnlyList<LegacyDocWritableTableCell> writableCells = ExpandSupportedTableCells(cells, gridColumnWidthsTwips, tableBorders, tableShading, conditionalStyles, tableLook, rowIndex, rows.Length);
                IReadOnlyList<int> cellWidthsTwips = ReadSupportedTableCellWidths(writableCells);
                IReadOnlyList<LegacyDocTableCellHorizontalMerge> cellHorizontalMerges = ReadSupportedTableCellHorizontalMerges(writableCells);
                IReadOnlyList<LegacyDocTableCellVerticalMerge> cellVerticalMerges = ReadSupportedTableCellVerticalMerges(writableCells);
                IReadOnlyList<LegacyDocTableCellVerticalAlignment> cellVerticalAlignments = ReadSupportedTableCellVerticalAlignments(writableCells);
                IReadOnlyList<LegacyDocTableCellTextDirection> cellTextDirections = ReadSupportedTableCellTextDirections(writableCells);
                IReadOnlyList<bool> cellFitTexts = ReadSupportedTableCellFitTexts(writableCells);
                IReadOnlyList<bool> cellNoWraps = ReadSupportedTableCellNoWraps(writableCells);
                IReadOnlyList<bool> cellHideMarks = ReadSupportedTableCellHideMarks(writableCells);
                IReadOnlyList<LegacyDocTableCellMargins> cellMargins = ReadSupportedTableCellMargins(writableCells);
                IReadOnlyList<LegacyDocTableCellShading> cellShadings = ReadSupportedTableCellShadings(writableCells);
                IReadOnlyList<LegacyDocTableCellBorders> cellBorders = ReadSupportedTableCellBorders(writableCells);
                foreach (LegacyDocWritableTableCell writableCell in writableCells) {
                    LegacyDocWritableParagraphFormatting cellParagraphFormatting = writableCell.ParagraphFormatting.WithInheritedParagraphFormatting(tableStyleParagraphFormatting);
                    LegacyDocWritableFormatting cellRunFormatting = writableCell.RunFormatting.WithInheritedFormatting(tableStyleRunFormatting);
                    LegacyDocWritableParagraphFormatting paragraphFormatting = AppendTableCell(text, runs, paragraphFormats, bookmarks, writableCell.SourceCell, mainPart, relationshipOwner, pictures, styleIndexes, tableStyleDefinitions, cellParagraphFormatting, cellRunFormatting, footnotes, endnotes, tableDepth, out int finalParagraphStart, out LegacyDocWritableFormatting finalParagraphMarkFormatting);
                    paragraphFormatting = tableDepth == 1
                        ? paragraphFormatting.WithTableMarkers(isTableTerminatingParagraph: false)
                        : paragraphFormatting.WithNestedTableMarkers(tableDepth);
                    text.Append(tableDepth == 1 ? '\a' : '\r');
                    AddParagraphMarkRunFormatting(runs, text.Length - 1, finalParagraphMarkFormatting);
                    paragraphFormats.Add(new LegacyDocWritableParagraph(finalParagraphStart, text.Length - finalParagraphStart, paragraphFormatting));
                }

                int rowTerminatorStart = text.Length;
                text.Append(tableDepth == 1 ? '\a' : '\r');
                LegacyDocWritableParagraphFormatting rowTerminatorFormatting = tableDepth == 1
                    ? LegacyDocWritableParagraphFormatting.Plain.WithTableMarkers(
                        isTableTerminatingParagraph: true,
                        tableCellWidthsTwips: cellWidthsTwips,
                        tableRowHeightTwips: rowFormatting.RowHeightTwips,
                        tableRowHeightIsExact: rowFormatting.RowHeightIsExact,
                        tableRowCantSplit: rowFormatting.RowCantSplit,
                        tableRowIsHeader: rowFormatting.RowIsHeader,
                        tableAlignment: tableAlignment,
                        tableLeftIndentTwips: tableLeftIndentTwips,
                        tableCellHorizontalMerges: cellHorizontalMerges,
                        tableCellVerticalMerges: cellVerticalMerges,
                        tableCellVerticalAlignments: cellVerticalAlignments,
                        tableCellTextDirections: cellTextDirections,
                        tableCellFitTexts: cellFitTexts,
                        tableCellNoWraps: cellNoWraps,
                        tableCellHideMarks: cellHideMarks,
                        tableCellMargins: cellMargins,
                        tableCellShadings: cellShadings,
                        tableCellBorders: cellBorders,
                        defaultTableCellMargins: defaultCellMargins,
                        defaultTableCellSpacingTwips: defaultCellSpacingTwips,
                        tablePreferredWidth: tablePreferredWidth,
                        tableAutofit: tableAutofit)
                    : LegacyDocWritableParagraphFormatting.Plain.WithNestedTableMarkers(tableDepth, isInnerTableTerminatingParagraph: true);
                paragraphFormats.Add(new LegacyDocWritableParagraph(rowTerminatorStart, 1, rowTerminatorFormatting));
                if (rowIndex + 1 < rows.Length) {
                    AppendTableRowBoundaryBookmarks(table, row, bookmarks, text.Length);
                }
            }

            if (tableDepth == 1) {
                AppendTrailingTableBoundaryBookmarks(table, bookmarks, text.Length);
            }
        }

        private static LegacyDocWritableParagraphFormatting AppendTableCell(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            List<LegacyDocWritableParagraph> paragraphFormats,
            LegacyDocWritableBookmarksBuilder bookmarks,
            TableCell? cell,
            MainDocumentPart mainPart,
            OpenXmlPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            IReadOnlyDictionary<string, Style> tableStyleDefinitions,
            LegacyDocWritableParagraphFormatting tableStyleParagraphFormatting,
            LegacyDocWritableFormatting tableStyleRunFormatting,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            int tableDepth,
            out int finalParagraphStart,
            out LegacyDocWritableFormatting finalParagraphMarkFormatting) {
            finalParagraphStart = text.Length;
            finalParagraphMarkFormatting = LegacyDocWritableFormatting.Plain;
            if (cell == null) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            var content = new List<OpenXmlElement>();
            foreach (OpenXmlElement child in cell.ChildElements) {
                switch (child) {
                    case TableCellProperties cellProperties:
                        ThrowIfUnsupportedTableCellProperties(cellProperties);
                        break;
                    case Paragraph cellParagraph:
                        content.Add(cellParagraph);
                        break;
                    case Table nestedTable:
                        if (tableDepth >= 2) {
                            throw new NotSupportedException("Native DOC saving supports nested tables only to depth 2.");
                        }

                        content.Add(nestedTable);
                        break;
                    case SdtBlock sdtBlock:
                        AddSupportedTableCellContentControlChildren(sdtBlock, content);
                        break;
                    case BookmarkStart:
                    case BookmarkEnd:
                        content.Add(child);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table cell element: {child.LocalName}.");
                }
            }

            if (content.Count == 0) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            bool hasFinalParagraph = false;
            LegacyDocWritableParagraphFormatting finalParagraphFormatting = LegacyDocWritableParagraphFormatting.Plain;
            for (int contentIndex = 0; contentIndex < content.Count; contentIndex++) {
                OpenXmlElement child = content[contentIndex];
                switch (child) {
                    case Paragraph paragraph:
                        int paragraphStart = text.Length;
                        LegacyDocWritableParagraphFormatting paragraphFormatting = AppendTableCellParagraph(text, runs, bookmarks, paragraph, mainPart, relationshipOwner, pictures, styleIndexes, tableStyleParagraphFormatting, tableStyleRunFormatting, footnotes, endnotes, out LegacyDocWritableFormatting paragraphMarkFormatting);
                        bool isFinalParagraph = !HasLaterParagraphOrTable(content, contentIndex);
                        if (isFinalParagraph) {
                            finalParagraphStart = paragraphStart;
                            finalParagraphFormatting = paragraphFormatting;
                            finalParagraphMarkFormatting = paragraphMarkFormatting;
                            hasFinalParagraph = true;
                        } else {
                            paragraphFormatting = tableDepth == 1
                                ? paragraphFormatting.WithTableMarkers(isTableTerminatingParagraph: false)
                                : paragraphFormatting.WithNestedTableMarkers(tableDepth, isCellTerminator: false);
                            text.Append('\r');
                            AddParagraphMarkRunFormatting(runs, text.Length - 1, paragraphMarkFormatting);
                            paragraphFormats.Add(new LegacyDocWritableParagraph(paragraphStart, text.Length - paragraphStart, paragraphFormatting));
                        }

                        break;
                    case Table nestedTable:
                        AppendTable(text, runs, paragraphFormats, bookmarks, nestedTable, mainPart, pictures, styleIndexes, tableStyleDefinitions, footnotes, endnotes, tableDepth + 1, relationshipOwner);
                        finalParagraphStart = text.Length;
                        finalParagraphMarkFormatting = LegacyDocWritableFormatting.Plain;
                        finalParagraphFormatting = LegacyDocWritableParagraphFormatting.Plain;
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, text.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, text.Length);
                        break;
                }
            }

            if (!hasFinalParagraph) {
                finalParagraphStart = text.Length;
                finalParagraphMarkFormatting = LegacyDocWritableFormatting.Plain;
                finalParagraphFormatting = LegacyDocWritableParagraphFormatting.Plain;
            }

            return finalParagraphFormatting;
        }

        private static bool HasLaterParagraphOrTable(IReadOnlyList<OpenXmlElement> content, int currentIndex) {
            for (int index = currentIndex + 1; index < content.Count; index++) {
                if (content[index] is Paragraph || content[index] is Table) {
                    return true;
                }
            }

            return false;
        }

        private static void AddSupportedTableCellContentControlChildren(SdtBlock sdtBlock, List<OpenXmlElement> content) {
            SdtContentBlock? contentBlock = sdtBlock.SdtContentBlock;
            if (contentBlock == null) {
                throw new NotSupportedException("Native DOC saving supports table cell content controls only when they contain simple paragraphs and bookmarks.");
            }

            foreach (OpenXmlElement child in contentBlock.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        if (paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>() != null) {
                            throw new NotSupportedException("Native DOC saving keeps section breaks scoped to supported body paragraph boundaries. Table cell content controls cannot contain section properties.");
                        }

                        content.Add(paragraph);
                        break;
                    case SdtBlock childContentControl:
                        AddSupportedTableCellContentControlChildren(childContentControl, content);
                        break;
                    case BookmarkStart:
                    case BookmarkEnd:
                        content.Add(child);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table cell content controls only when they contain simple paragraphs and bookmarks. Unsupported table cell content control element: {child.LocalName}.");
                }
            }
        }

        private static LegacyDocWritableParagraphFormatting AppendTableCellParagraph(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableBookmarksBuilder bookmarks, Paragraph paragraph, MainDocumentPart mainPart, OpenXmlPart relationshipOwner, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> styleIndexes, LegacyDocWritableParagraphFormatting tableStyleParagraphFormatting, LegacyDocWritableFormatting tableStyleRunFormatting, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, out LegacyDocWritableFormatting paragraphMarkFormatting) {
            paragraphMarkFormatting = ReadSupportedParagraphMarkRunFormatting(paragraph.ParagraphProperties);
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedParagraphFormatting(paragraph.ParagraphProperties, styleIndexes)
                .WithInheritedParagraphFormatting(ReadSupportedCellParagraphStyleFormatting(paragraph, mainPart))
                .WithInheritedParagraphFormatting(tableStyleParagraphFormatting);
            tableStyleRunFormatting = ReadSupportedCellParagraphStyleRunFormatting(paragraph, mainPart)
                .WithInheritedFormatting(tableStyleRunFormatting);

            OpenXmlElement[] children = paragraph.ChildElements.ToArray();
            for (int index = 0; index < children.Length; index++) {
                OpenXmlElement child = children[index];
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        if (IsComplexFieldBeginRun(run)) {
                            AppendSupportedComplexPageNumberField(children, ref index, text, runs, bookmarks, tableStyleRunFormatting);
                        } else {
                            AppendSupportedRunText(
                                text,
                                runs,
                                run,
                                footnotes,
                                endnotes,
                                tableStyleRunFormatting,
                                allowHyperlinkRunStyle: false,
                                pictures,
                                relationshipOwner);
                        }

                        break;
                    case InsertedRun insertedRun:
                        AppendSupportedRevisionText(text, runs, insertedRun, LegacyDocRevisionKind.Inserted, footnotes, endnotes, tableStyleRunFormatting, pictures, relationshipOwner);
                        break;
                    case DeletedRun deletedRun:
                        AppendSupportedRevisionText(text, runs, deletedRun, LegacyDocRevisionKind.Deleted, footnotes, endnotes, tableStyleRunFormatting, pictures, relationshipOwner);
                        break;
                    case Hyperlink hyperlink:
                        AppendSupportedHyperlinkText(text, runs, bookmarks, hyperlink, relationshipOwner, footnotes, endnotes, tableStyleRunFormatting);
                        break;
                    case SimpleField simpleField:
                        AppendSupportedPageNumberFieldFromSimpleField(text, runs, bookmarks, simpleField, tableStyleRunFormatting);
                        break;
                    case DocumentFormat.OpenXml.Math.OfficeMath officeMath:
                        AppendMathEquationField(text, runs, officeMath, tableStyleRunFormatting);
                        break;
                    case DocumentFormat.OpenXml.Math.Paragraph mathParagraph:
                        AppendMathEquationField(text, runs, mathParagraph, tableStyleRunFormatting);
                        break;
                    case SdtRun sdtRun:
                        AppendSupportedInlineContentControlText(text, runs, bookmarks, sdtRun, relationshipOwner, pictures, footnotes, endnotes, tableStyleRunFormatting, "table cell inline content control");
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, text.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, text.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports simple table cell paragraphs only with text runs, {SupportedFieldNames} simple fields, bookmarks, inline content controls, and simple hyperlinks. Unsupported paragraph element: {child.LocalName}.");
                }
            }

            return paragraphFormatting;
        }

    }
}
