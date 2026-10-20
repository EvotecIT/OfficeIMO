using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace OfficeIMO.Word {
    /// <summary>Enumerates logical table rows and cells through their inert content-control and custom-XML containers.</summary>
    internal static class WordTableContent {
        private const int MaximumContainerDepth = 128;

        internal static IEnumerable<TableRow> Rows(Table table) => Rows(table, 0);

        internal static IEnumerable<TableCell> Cells(TableRow row) => Cells(row, 0);

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
