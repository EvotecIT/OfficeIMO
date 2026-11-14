using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Word refuses a table without rows, a row without cells and a cell that does not end in a paragraph.
        /// Row, cell, paragraph and nested-table removal can each leave one behind, and an empty table may still be
        /// filled before saving, so these are settled when the document is written. A part whose XML was never
        /// parsed has not been edited and is written unchanged, so only loaded parts are repaired.
        /// </summary>
        private void RepairEmptyTableStructure() {
            MainDocumentPart? main = _wordprocessingDocument.MainDocumentPart;
            if (main == null) return;
            RepairEmptyTableStructure(main);
            RepairEmptyTableStructure(main.FootnotesPart);
            RepairEmptyTableStructure(main.EndnotesPart);
            foreach (HeaderPart part in main.HeaderParts) RepairEmptyTableStructure(part);
            foreach (FooterPart part in main.FooterParts) RepairEmptyTableStructure(part);
        }

        private static void RepairEmptyTableStructure(OpenXmlPart? part) {
            if (part is { IsRootElementLoaded: true, RootElement: { } root }) RepairEmptyTableStructure(root);
        }

        private static void RepairEmptyTableStructure(OpenXmlElement root) {
            var tables = new List<Table>();
            var cells = new List<TableCell>();
            CollectTableStructure(root, tables, cells);
            if (tables.Count == 0 && cells.Count == 0) return;

            bool removed = false;
            for (int index = tables.Count - 1; index >= 0; index--) {
                Table table = tables[index];
                foreach (TableRow row in WordTableContent.Rows(table).ToList()) {
                    if (!WordTableContent.Cells(row).Any()) {
                        row.Remove();
                        removed = true;
                    }
                }
                if (!WordTableContent.Rows(table).Any()) {
                    table.Remove();
                    removed = true;
                }
            }

            if (removed) {
                cells.Clear();
                tables.Clear();
                CollectTableStructure(root, tables, cells);
            }

            foreach (TableCell cell in cells) {
                OpenXmlElement? last = cell.ChildElements.LastOrDefault(child =>
                    child is not (TableCellProperties or BookmarkStart or BookmarkEnd or CommentRangeStart or CommentRangeEnd));
                if (last == null) cell.Append(new Paragraph());
            }
        }

        /// <summary>Collects tables and cells in document order in one walk, skipping run and paragraph formatting.</summary>
        private static void CollectTableStructure(OpenXmlElement root, List<Table> tables, List<TableCell> cells) {
            for (OpenXmlElement? element = root.FirstChild; element != null;
                 element = WordElementWalk.Next(element, root, descend: !WordElementWalk.IsFormattingOrText(element))) {
                if (element is Table table) tables.Add(table);
                else if (element is TableCell cell) cells.Add(cell);
            }
        }
    }
}
