using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendTable(StringBuilder text, List<LegacyDocWritableRun> runs, List<LegacyDocWritableParagraph> paragraphFormats, LegacyDocWritableBookmarksBuilder bookmarks, LegacyDocElement table, LegacyDocSourcePackage source, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> styleIndexes, LegacyDocTableStyleDefinitions tableStyleDefinitions, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, int tableDepth = 1, LegacyDocPart? relationshipOwner = null) {
            LegacyDocPart owner = relationshipOwner ?? source.Main;
            if (tableDepth <= 0 || tableDepth > 2) {
                throw new NotSupportedException("Native DOC saving supports nested tables only to depth 2.");
            }

            // An ordinary paragraph mark separates adjacent top-level tables
            // in every story; otherwise DOC treats them as one table.
            if (tableDepth == 1 && text.Length > 0 && text[text.Length - 1] == '\a') text.Append('\r');

            ThrowIfUnsupportedTableShape(table, tableStyleDefinitions);

            LegacyDocElement[] rows = table.Elements("tr").ToArray();
            if (rows.Length == 0) {
                throw new NotSupportedException("Native DOC saving supports simple tables only when at least one row is present.");
            }

            AppendLeadingTableBoundaryBookmarks(table, bookmarks, text.Length);
            // The table's own properties; a document default table style stands in for a missing w:tblStyle.
            LegacyDocElement tableProperties = table.Child("tblPr");
            string? tableStyleId = ResolveSupportedEffectiveTableStyleId(tableProperties, tableStyleDefinitions);
            LegacyDocTableAlignment? tableAlignment = ReadSupportedTableAlignmentFromProperties(tableProperties);
            tableAlignment ??= ReadSupportedTableStyleAlignment(tableStyleId, tableStyleDefinitions);
            int? tableLeftIndentTwips = ReadSupportedTableIndentationFromProperties(tableProperties);
            tableLeftIndentTwips ??= ReadSupportedTableStyleIndentation(tableStyleId, tableStyleDefinitions);
            LegacyDocTablePreferredWidth? tablePreferredWidth = ReadSupportedTablePreferredWidthFromProperties(tableProperties);
            tablePreferredWidth ??= ReadSupportedTableStylePreferredWidth(tableStyleId, tableStyleDefinitions);
            bool? tableAutofit = ReadSupportedTableAutofitFromProperties(tableProperties);
            tableAutofit ??= ReadSupportedTableStyleAutofit(tableStyleId, tableStyleDefinitions);
            // DOCX defaults to AutoFit, while binary DOC defaults to fixed columns.
            // Materialize the effective default after direct and inherited settings.
            tableAutofit ??= true;
            LegacyDocTableCellMargins? defaultCellMargins = ReadSupportedTableDefaultCellMarginsFromProperties(tableProperties);
            // Materialize the resolved DOCX inset for stable native DOC layout;
            // Word positions these cells differently when the padding is implicit.
            defaultCellMargins = new LegacyDocTableCellMargins(0, 108, 0, 108)
                .Merge(ReadSupportedTableStyleDefaultCellMargins(tableStyleId, tableStyleDefinitions) ?? default)
                .Merge(defaultCellMargins ?? default);
            int? defaultCellSpacingTwips = ReadSupportedTableDefaultCellSpacingFromProperties(tableProperties);
            defaultCellSpacingTwips ??= ReadSupportedTableStyleDefaultCellSpacing(tableStyleId, tableStyleDefinitions);
            LegacyDocTableBorders tableBorders = ReadSupportedTableBorders(tableProperties, tableStyleId, tableStyleDefinitions);
            LegacyDocTableCellShading tableShading = ReadSupportedTableShading(tableStyleId, tableStyleDefinitions);
            LegacyDocWritableParagraphFormatting tableStyleParagraphFormatting = ReadSupportedTableStyleParagraphFormatting(tableStyleId, tableStyleDefinitions);
            LegacyDocWritableFormatting tableStyleRunFormatting = ReadSupportedTableStyleRunFormatting(tableStyleId, tableStyleDefinitions);
            LegacyDocTableConditionalStyleSet conditionalStyles = ReadSupportedTableConditionalStyles(tableStyleId, tableStyleDefinitions);
            LegacyDocTableLook tableLook = ReadSupportedTableLook(tableProperties.Child("tblLook"));
            IReadOnlyList<int> gridColumnWidthsTwips = ReadSupportedTableGridWidths(table.Child("tblGrid"));
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
                LegacyDocElement row = rows[rowIndex];
                LegacyDocWritableTableRowFormatting rowFormatting = ReadSupportedTableRowFormatting(row, out LegacyDocElement[] cells);
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
                    LegacyDocWritableParagraphFormatting paragraphFormatting = AppendTableCell(text, runs, paragraphFormats, bookmarks, writableCell.SourceCell, source, owner, pictures, styleIndexes, tableStyleDefinitions, cellParagraphFormatting, cellRunFormatting, footnotes, endnotes, tableDepth, out int finalParagraphStart, out LegacyDocWritableFormatting finalParagraphMarkFormatting);
                    paragraphFormatting = tableDepth == 1
                        ? paragraphFormatting.WithTableMarkers(isTableTerminatingParagraph: false)
                        : paragraphFormatting.WithNestedTableMarkers(tableDepth);
                    text.Append(tableDepth == 1 ? '\a' : '\r');
                    AddParagraphMarkRunFormatting(runs, text.Length - 1, finalParagraphMarkFormatting);
                    paragraphFormats.Add(new LegacyDocWritableParagraph(finalParagraphStart, text.Length - finalParagraphStart, paragraphFormatting));
                }

                int rowTerminatorStart = text.Length;
                text.Append(tableDepth == 1 ? '\a' : '\r');
                // Both ordinary and nested row marks must carry the complete cell definitions.
                LegacyDocWritableParagraphFormatting rowTerminatorFormatting =
                    LegacyDocWritableParagraphFormatting.Plain.WithTableMarkers(
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
                        tableAutofit: tableAutofit);
                if (tableDepth > 1) {
                    rowTerminatorFormatting = rowTerminatorFormatting.WithNestedTableMarkers(
                        tableDepth, isInnerTableTerminatingParagraph: true);
                }
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
            LegacyDocElement cell,
            LegacyDocSourcePackage source,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            LegacyDocTableStyleDefinitions tableStyleDefinitions,
            LegacyDocWritableParagraphFormatting tableStyleParagraphFormatting,
            LegacyDocWritableFormatting tableStyleRunFormatting,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            int tableDepth,
            out int finalParagraphStart,
            out LegacyDocWritableFormatting finalParagraphMarkFormatting) {
            finalParagraphStart = text.Length;
            finalParagraphMarkFormatting = LegacyDocWritableFormatting.Plain;
            if (cell.IsNone) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            var content = new List<LegacyDocElement>();
            foreach (LegacyDocElement child in cell.Elements()) {
                switch (child.WName) {
                    case "tcPr":
                        ThrowIfUnsupportedTableCellProperties(child);
                        break;
                    case "p":
                        content.Add(child);
                        break;
                    case "tbl":
                        if (tableDepth >= 2) {
                            throw new NotSupportedException("Native DOC saving supports nested tables only to depth 2.");
                        }

                        content.Add(child);
                        break;
                    case "sdt":
                        AddSupportedTableCellContentControlChildren(child, content);
                        break;
                    case "bookmarkStart":
                    case "bookmarkEnd":
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
                LegacyDocElement child = content[contentIndex];
                switch (child.WName) {
                    case "p":
                        int paragraphStart = text.Length;
                        LegacyDocWritableParagraphFormatting paragraphFormatting = AppendTableCellParagraph(text, runs, bookmarks, child, source, relationshipOwner, pictures, styleIndexes, tableStyleParagraphFormatting, tableStyleRunFormatting, footnotes, endnotes, out LegacyDocWritableFormatting paragraphMarkFormatting);
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
                    case "tbl":
                        AppendTable(text, runs, paragraphFormats, bookmarks, child, source, pictures, styleIndexes, tableStyleDefinitions, footnotes, endnotes, tableDepth + 1, relationshipOwner);
                        finalParagraphStart = text.Length;
                        finalParagraphMarkFormatting = LegacyDocWritableFormatting.Plain;
                        finalParagraphFormatting = LegacyDocWritableParagraphFormatting.Plain;
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, text.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, text.Length);
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

        private static bool HasLaterParagraphOrTable(IReadOnlyList<LegacyDocElement> content, int currentIndex) {
            for (int index = currentIndex + 1; index < content.Count; index++) {
                if (content[index].Is("p") || content[index].Is("tbl")) {
                    return true;
                }
            }

            return false;
        }

        private static void AddSupportedTableCellContentControlChildren(LegacyDocElement sdtBlock, List<LegacyDocElement> content) {
            LegacyDocElement contentBlock = sdtBlock.Child("sdtContent");
            if (contentBlock.IsNone) {
                throw new NotSupportedException("Native DOC saving supports table cell content controls only when they contain simple paragraphs and bookmarks.");
            }

            foreach (LegacyDocElement child in contentBlock.Elements()) {
                switch (child.WName) {
                    case "p":
                        if (child.Child("pPr").Child("sectPr").Exists) {
                            throw new NotSupportedException("Native DOC saving keeps section breaks scoped to supported body paragraph boundaries. Table cell content controls cannot contain section properties.");
                        }

                        content.Add(child);
                        break;
                    case "sdt":
                        AddSupportedTableCellContentControlChildren(child, content);
                        break;
                    case "bookmarkStart":
                    case "bookmarkEnd":
                        content.Add(child);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports table cell content controls only when they contain simple paragraphs and bookmarks. Unsupported table cell content control element: {child.LocalName}.");
                }
            }
        }

        private static LegacyDocWritableParagraphFormatting AppendTableCellParagraph(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableBookmarksBuilder bookmarks, LegacyDocElement paragraph, LegacyDocSourcePackage source, LegacyDocPart relationshipOwner, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> styleIndexes, LegacyDocWritableParagraphFormatting tableStyleParagraphFormatting, LegacyDocWritableFormatting tableStyleRunFormatting, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, out LegacyDocWritableFormatting paragraphMarkFormatting) {
            LegacyDocElement paragraphProperties = paragraph.Child("pPr");
            paragraphMarkFormatting = ReadSupportedParagraphMarkRunFormatting(paragraphProperties);
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedParagraphFormatting(paragraphProperties, styleIndexes)
                .WithInheritedParagraphFormatting(ReadSupportedCellParagraphStyleFormatting(paragraph, source))
                .WithInheritedParagraphFormatting(tableStyleParagraphFormatting);
            tableStyleRunFormatting = ReadSupportedCellParagraphStyleRunFormatting(paragraph, source)
                .WithInheritedFormatting(tableStyleRunFormatting);

            LegacyDocElement[] children = paragraph.Elements().ToArray();
            for (int index = 0; index < children.Length; index++) {
                LegacyDocElement child = children[index];
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendMathEquationField(text, runs, child, tableStyleRunFormatting);
                    continue;
                }

                switch (child.WName) {
                    case "pPr":
                        break;
                    case "r":
                        if (IsComplexFieldBeginRun(child)) {
                            AppendSupportedComplexPageNumberField(children, ref index, text, runs, bookmarks, tableStyleRunFormatting);
                        } else {
                            AppendSupportedRunText(
                                text,
                                runs,
                                child,
                                footnotes,
                                endnotes,
                                tableStyleRunFormatting,
                                allowHyperlinkRunStyle: false,
                                pictures,
                                relationshipOwner);
                        }

                        break;
                    case "ins":
                        AppendSupportedRevisionText(text, runs, child, LegacyDocRevisionKind.Inserted, footnotes, endnotes, tableStyleRunFormatting, pictures, relationshipOwner);
                        break;
                    case "del":
                        AppendSupportedRevisionText(text, runs, child, LegacyDocRevisionKind.Deleted, footnotes, endnotes, tableStyleRunFormatting, pictures, relationshipOwner);
                        break;
                    case "hyperlink":
                        AppendSupportedHyperlinkText(text, runs, bookmarks, child, relationshipOwner, footnotes, endnotes, tableStyleRunFormatting);
                        break;
                    case "fldSimple":
                        AppendSupportedPageNumberFieldFromSimpleField(text, runs, bookmarks, child, tableStyleRunFormatting);
                        break;
                    case "sdt":
                        AppendSupportedInlineContentControlText(text, runs, bookmarks, child, relationshipOwner, pictures, footnotes, endnotes, tableStyleRunFormatting, "table cell inline content control");
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, text.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, text.Length);
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
