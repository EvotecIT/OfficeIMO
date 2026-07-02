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
            var currentRow = new List<List<LegacyDocTableCellParagraph>>();
            var currentCell = new List<LegacyDocTableCellParagraph>();
            int index = startIndex;

            for (; index < paragraphs.Count; index++) {
                LegacyDocTableCellParagraph paragraph = paragraphs[index];
                if (!IsLegacyDocNestedTableParagraph(paragraph)) {
                    break;
                }

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
                        currentRow = new List<List<LegacyDocTableCellParagraph>>();
                    }
                }
            }

            if (currentCell.Count > 0) {
                CloseNestedCell();
            }

            if (currentRow.Count > 0) {
                rows.Add(currentRow);
            }

            if (rows.Count == 0) {
                return startIndex;
            }

            int columnCount = rows.Max(row => row.Count);
            WordTable nestedTable = hostCell.AddTable(rows.Count, columnCount, WordTableStyle.TableNormal);
            AddPendingBookmarksAroundNestedTable(nestedTable, pendingBookmarks);
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                List<List<LegacyDocTableCellParagraph>> row = rows[rowIndex];
                for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                    AddLegacyDocTableCell(
                        nestedTable.Rows[rowIndex].Cells[columnIndex],
                        new LegacyDocTableCell(row[columnIndex]),
                        styleSheet,
                        notes,
                        projectNestedTables: false);
                }
            }

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
