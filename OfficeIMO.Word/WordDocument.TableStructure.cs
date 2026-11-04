using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Word refuses a table without rows, a row without cells and a cell that does not end in a paragraph.
        /// Row, cell, paragraph and nested-table removal can each leave one behind, and an empty table may still be
        /// filled before saving, so these are settled when the document is written.
        /// </summary>
        private void RepairEmptyTableStructure() {
            MainDocumentPart? main = _wordprocessingDocument.MainDocumentPart;
            if (main == null) return;
            var roots = new List<OpenXmlElement?> { main.Document, main.FootnotesPart?.Footnotes, main.EndnotesPart?.Endnotes };
            roots.AddRange(main.HeaderParts.Select(part => (OpenXmlElement?)part.Header));
            roots.AddRange(main.FooterParts.Select(part => (OpenXmlElement?)part.Footer));
            foreach (OpenXmlElement? root in roots) {
                if (root != null) RepairEmptyTableStructure(root);
            }
        }

        private static void RepairEmptyTableStructure(OpenXmlElement root) {
            foreach (Table table in root.Descendants<Table>().Reverse().ToList()) {
                foreach (TableRow row in WordTableContent.Rows(table).ToList()) {
                    if (!WordTableContent.Cells(row).Any()) row.Remove();
                }
                if (!WordTableContent.Rows(table).Any()) table.Remove();
            }

            foreach (TableCell cell in root.Descendants<TableCell>().ToList()) {
                OpenXmlElement? last = cell.ChildElements.LastOrDefault(child =>
                    child is not (TableCellProperties or BookmarkStart or BookmarkEnd or CommentRangeStart or CommentRangeEnd));
                if (last == null) cell.Append(new Paragraph());
            }
        }
    }
}
