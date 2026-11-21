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
        /// <returns>The remaining tables of each repaired part root, in document order.</returns>
        private Dictionary<OpenXmlElement, List<Table>> RepairEmptyTableStructure() {
            var tablesByRoot = new Dictionary<OpenXmlElement, List<Table>>();
            MainDocumentPart? main = _wordprocessingDocument.MainDocumentPart;
            if (main == null) return tablesByRoot;
            RepairEmptyTableStructure(main, tablesByRoot);
            RepairEmptyTableStructure(main.FootnotesPart, tablesByRoot);
            RepairEmptyTableStructure(main.EndnotesPart, tablesByRoot);
            foreach (HeaderPart part in main.HeaderParts) RepairEmptyTableStructure(part, tablesByRoot);
            foreach (FooterPart part in main.FooterParts) RepairEmptyTableStructure(part, tablesByRoot);
            return tablesByRoot;
        }

        private static void RepairEmptyTableStructure(OpenXmlPart? part, Dictionary<OpenXmlElement, List<Table>> tablesByRoot) {
            if (part is { IsRootElementLoaded: true, RootElement: { } root }) tablesByRoot[root] = RepairEmptyTableStructure(root);
        }

        private static List<Table> RepairEmptyTableStructure(OpenXmlElement root) {
            var tables = new List<Table>();
            var cells = new List<TableCell>();
            CollectTableStructure(root, tables, cells);
            if (tables.Count == 0 && cells.Count == 0) return tables;

            bool removed = false;
            for (int index = tables.Count - 1; index >= 0; index--) {
                Table table = tables[index];
                if (HasOnlyRowsWithCells(table)) continue;
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
                if (!HasCellContent(cell)) cell.Append(new Paragraph());
            }
            return tables;
        }

        /// <summary>
        /// Whether the table's children are only its properties, grid and at least one row, each row holding a cell
        /// directly: the common shape, which needs no repair, recognized without the content-control-aware row and
        /// cell enumeration.
        /// </summary>
        private static bool HasOnlyRowsWithCells(Table table) {
            bool hasRow = false;
            for (OpenXmlElement? child = table.FirstChild; child != null; child = child.NextSibling()) {
                Type type = child.GetType();
                if (type == typeof(TableRow)) {
                    if (!HasDirectCell(child)) return false;
                    hasRow = true;
                } else if (type != typeof(TableProperties) && type != typeof(TableGrid)) {
                    return false;
                }
            }
            return hasRow;
        }

        private static bool HasDirectCell(OpenXmlElement row) {
            for (OpenXmlElement? child = row.FirstChild; child != null; child = child.NextSibling()) {
                if (child.GetType() == typeof(TableCell)) return true;
            }
            return false;
        }

        private static bool HasCellContent(TableCell cell) {
            for (OpenXmlElement? child = cell.FirstChild; child != null; child = child.NextSibling()) {
                if (child is not (TableCellProperties or BookmarkStart or BookmarkEnd or CommentRangeStart or CommentRangeEnd)) return true;
            }
            return false;
        }

        /// <summary>Collects tables and cells in document order in one walk, skipping run and paragraph formatting.</summary>
        private static void CollectTableStructure(OpenXmlElement root, List<Table> tables, List<TableCell> cells) {
            for (OpenXmlElement? element = root.FirstChild; element != null;
                 element = WordElementWalk.Next(element, root, descend: !WordElementWalk.IsFormattingOrText(element))) {
                if (WordElementWalk.IsCommonContent(element)) continue;
                if (element is Table table) tables.Add(table);
                else if (element is TableCell cell) cells.Add(cell);
            }
        }
    }
}
