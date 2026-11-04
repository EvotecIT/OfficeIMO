using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace OfficeIMO.Word {
    /// <summary>Enumerates logical table rows and cells through their inert content-control and custom-XML containers.</summary>
    internal static class WordTableContent {
        private const int MaximumContainerDepth = 128;

        internal static IEnumerable<TableRow> Rows(Table table) => Rows(table, 0);

        internal static IEnumerable<TableCell> Cells(TableRow row) => Cells(row, 0);

        internal static IEnumerable<TableCell> CellsAfter(TableCell anchor, int count) {
            TableRow? row = anchor.Ancestors<TableRow>().FirstOrDefault();
            if (row == null) return Enumerable.Empty<TableCell>();
            return Cells(row).SkipWhile(cell => !ReferenceEquals(cell, anchor)).Skip(1).Take(count);
        }

        internal static IEnumerable<TableCell> CellsBelow(TableCell anchor, int count) {
            TableRow? row = anchor.Ancestors<TableRow>().FirstOrDefault();
            Table? table = row?.Ancestors<Table>().FirstOrDefault();
            if (row == null || table == null) yield break;
            int index = Cells(row).ToList().IndexOf(anchor);
            if (index < 0) yield break;
            foreach (TableRow next in Rows(table).SkipWhile(candidate => !ReferenceEquals(candidate, row)).Skip(1).Take(count)) {
                TableCell? cell = Cells(next).ElementAtOrDefault(index);
                if (cell != null) yield return cell;
            }
        }

        /// <summary>Grid columns a row occupies, counting spanned cells and the cells it skips.</summary>
        internal static int GridColumnCount(TableRow row) =>
            (row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0)
            + Cells(row).Sum(GridSpanOf)
            + (row.TableRowProperties?.GetFirstChild<GridAfter>()?.Val?.Value ?? 0);

        internal static int GridSpanOf(TableCell cell) => Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);

        /// <summary>The cell covering a grid column with the column it starts at, or null past the row's last cell.</summary>
        internal static (TableCell Cell, int Column)? CellAtColumn(TableRow row, int column) {
            int offset = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
            foreach (TableCell cell in Cells(row)) {
                int span = GridSpanOf(cell);
                if (column < offset + span) return (cell, offset);
                offset += span;
            }
            return null;
        }

        internal static bool ContinuesVerticalMerge(TableCell? cell) {
            VerticalMerge? merge = cell?.TableCellProperties?.VerticalMerge;
            return merge != null && (merge.Val == null || merge.Val.Value == MergedCellValues.Continue);
        }

        internal static bool ContinuesHorizontalMerge(TableCell? cell) {
            HorizontalMerge? merge = cell?.TableCellProperties?.HorizontalMerge;
            return merge != null && (merge.Val == null || merge.Val.Value == MergedCellValues.Continue);
        }

        internal static void InsertRowBefore(Table table, TableRow reference, TableRow row) =>
            RowInsertionAnchor(table, reference, before: true).InsertBeforeSelf(row);

        internal static void InsertRowAfter(Table table, TableRow reference, TableRow row) =>
            RowInsertionAnchor(table, reference, before: false).InsertAfterSelf(row);

        private static OpenXmlElement RowInsertionAnchor(Table table, TableRow row, bool before) {
            OpenXmlElement anchor = row;
            while (anchor.Parent is { } parent && !ReferenceEquals(parent, table)) {
                OpenXmlElement content = parent is SdtRow control ? control.SdtContentRow! : parent;
                if (content == null || parent is not (SdtContentRow or SdtRow or CustomXmlRow)) break;
                IEnumerable<TableRow> siblings = Rows(content, 0);
                // Lift a boundary insertion outside its enclosing control. An
                // insertion between rows in a shared container remains inside it.
                TableRow? boundary = before ? siblings.FirstOrDefault() : siblings.LastOrDefault();
                if (!ReferenceEquals(boundary, row)) break;
                anchor = parent;
            }
            return anchor;
        }

        private static IEnumerable<TableRow> Rows(OpenXmlElement parent, int depth) {
            foreach (OpenXmlElement child in parent.ChildElements) {
                if (child is TableRow row) {
                    yield return row;
                } else {
                    OpenXmlElement? content = child switch {
                        SdtRow control => control.SdtContentRow,
                        CustomXmlRow customXml => customXml,
                        _ => null
                    };
                    if (content == null) continue;
                    EnsureContainerDepth(depth);
                    foreach (TableRow contained in Rows(content, depth + 1)) yield return contained;
                }
            }
        }

        private static IEnumerable<TableCell> Cells(OpenXmlElement parent, int depth) {
            foreach (OpenXmlElement child in parent.ChildElements) {
                if (child is TableCell cell) {
                    // A cell is a boundary: its nested tables belong to that cell,
                    // rather than adding columns to the containing row.
                    yield return cell;
                } else {
                    OpenXmlElement? content = child switch {
                        SdtCell control => control.SdtContentCell,
                        CustomXmlCell customXml => customXml,
                        _ => null
                    };
                    if (content == null) continue;
                    EnsureContainerDepth(depth);
                    foreach (TableCell contained in Cells(content, depth + 1)) yield return contained;
                }
            }
        }

        private static void EnsureContainerDepth(int depth) {
            if (depth >= MaximumContainerDepth)
                throw new InvalidDataException($"Table row or cell containers exceed the supported limit of {MaximumContainerDepth} levels.");
        }
    }
}
