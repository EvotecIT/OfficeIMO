using OfficeIMO.Word.LegacyDoc.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static bool IsLegacyDocNestedTableParagraph(LegacyDocTableCellParagraph paragraph) =>
            paragraph.Format.MaximumTableDepth > 1
            || paragraph.Format.HasInnerTableCellMarker
            || paragraph.Format.HasInnerTableTerminatingParagraphMarker;

        private static int AddLegacyDocNestedTable(
            WordTableCell hostCell,
            IReadOnlyList<LegacyDocTableCellParagraph> paragraphs,
            int startIndex,
            LegacyDocStyleSheet styleSheet,
            LegacyDocNoteProjection notes,
            IReadOnlyList<LegacyDocBookmark>? pendingBookmarks = null) {
            var rows = new List<List<List<LegacyDocTableCellParagraph>>>();
            var rowFormats = new List<LegacyDocParagraphFormat>();
            var currentRow = new List<List<LegacyDocTableCellParagraph>>();
            var currentCell = new List<LegacyDocTableCellParagraph>();
            LegacyDocParagraphFormat lastRowFormat = LegacyDocParagraphFormat.Default;
            int index = startIndex;

            for (; index < paragraphs.Count; index++) {
                LegacyDocTableCellParagraph paragraph = paragraphs[index];
                if (!IsLegacyDocNestedTableParagraph(paragraph)) {
                    break;
                }
                lastRowFormat = paragraph.Format;

                if (!paragraph.Format.HasInnerTableTerminatingParagraphMarker || paragraph.Runs.Count > 0 || paragraph.Bookmarks.Count > 0) {
                    currentCell.Add(paragraph);
                }

                if (paragraph.Format.HasInnerTableCellMarker) {
                    CloseNestedCell();
                }

                if (paragraph.Format.HasInnerTableTerminatingParagraphMarker) {
                    if (currentCell.Count > 0 || currentRow.Count == 0) {
                        CloseNestedCell();
                    }

                    if (currentRow.Count > 0) {
                        rows.Add(currentRow);
                        rowFormats.Add(paragraph.Format);
                        currentRow = new List<List<LegacyDocTableCellParagraph>>();
                    }
                }
            }

            if (currentCell.Count > 0) {
                CloseNestedCell();
            }

            if (currentRow.Count > 0) {
                rows.Add(currentRow);
                rowFormats.Add(lastRowFormat);
            }

            if (rows.Count == 0) {
                return startIndex;
            }

            // Use the ordinary table projection so nested widths, row constraints,
            // cell settings and borders follow the same native DOC contract.
            var sourceRows = new List<LegacyDocTableRow>();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                LegacyDocParagraphFormat format = rowFormats[rowIndex];
                LegacyDocTableCell[] cells = rows[rowIndex].Select(cell => new LegacyDocTableCell(cell)).ToArray();
                sourceRows.Add(new LegacyDocTableRow(
                    cells, format.TableCellWidthsTwips, format.TableLeftIndentTwips,
                    format.TableRowHeightTwips, format.TableRowHeightIsExact,
                    format.TableRowCantSplit, format.TableRowIsHeader, format.TableAlignment,
                    format.TableCellHorizontalMerges, format.TableCellVerticalMerges,
                    format.TableCellVerticalAlignments, format.TableCellTextDirections,
                    format.TableCellFitTexts, format.TableCellNoWraps, format.TableCellHideMarks,
                    format.GetTableCellMarginsForCellCount(cells.Length),
                    format.GetTableCellShadingsForCellCount(cells.Length),
                    format.GetTableCellBordersForCellCount(cells.Length),
                    format.DefaultTableCellSpacingTwips, format.TablePreferredWidth, format.TableAutofit,
                    tableStyleIndex: format.TableStyleIndex, tableBorders: format.TableBorders));
            }
            WordTable? nestedTable = null;
            AddLegacyDocTableCore(
                (rowCount, columnCount) => nestedTable = hostCell.AddTable(rowCount, columnCount, WordTableStyle.TableNormal),
                new LegacyDocTableBlock(sourceRows, 0, 0), styleSheet, notes, projectNestedTables: false);
            if (nestedTable != null) AddPendingBookmarksAroundNestedTable(nestedTable, pendingBookmarks);

            return index - 1;

            void CloseNestedCell() {
                currentRow.Add(currentCell);
                currentCell = new List<LegacyDocTableCellParagraph>();
            }
        }

        private static void AddPendingBookmarksAroundNestedTable(WordTable nestedTable, IReadOnlyList<LegacyDocBookmark>? pendingBookmarks) {
            if (pendingBookmarks == null || pendingBookmarks.Count == 0 || nestedTable._table.Parent is not OpenXmlCompositeElement parent) {
                return;
            }

            foreach (LegacyDocBookmark bookmark in pendingBookmarks
                .OrderByDescending(bookmark => bookmark.EndCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                parent.InsertBefore(new BookmarkStart { Id = bookmark.ProjectionId, Name = bookmark.Name }, nestedTable._table);
            }

            OpenXmlElement afterAnchor = nestedTable._table;
            foreach (LegacyDocBookmark bookmark in pendingBookmarks
                .OrderBy(bookmark => bookmark.StartCharacter)
                .ThenByDescending(bookmark => bookmark.EndCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                afterAnchor = parent.InsertAfter(new BookmarkEnd { Id = bookmark.ProjectionId }, afterAnchor)!;
            }
        }
    }
}
