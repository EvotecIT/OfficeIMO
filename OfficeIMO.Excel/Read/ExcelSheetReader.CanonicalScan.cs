using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        /// <summary>
        /// Returns the grid a worksheet spans, or null when a cell is one the tabular fast paths
        /// resolve differently from the SDK route: a missing or out-of-grid reference, an unknown
        /// type, or a whitespace-only value.
        /// </summary>
        internal static WorksheetExtent? MeasureCanonicalWorksheet(XmlReader reader, CancellationToken cancellationToken) {
            int lastRow = 0;
            int lastColumn = 0;
            int currentRow = 0;
            bool afterTextElement = false;
            int nodes = 0;
            while (reader.Read()) {
                if ((++nodes & 16383) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType == XmlNodeType.Whitespace) {
                    if (afterTextElement) return null;
                    continue;
                }

                afterTextElement = false;
                if (reader.NodeType != XmlNodeType.Element) continue;
                string name = reader.LocalName;
                if (name == "v" || name == "t") {
                    afterTextElement = !reader.IsEmptyElement;
                } else if (name == "row") {
                    currentRow = ParsePositiveIntAttribute(ReadXmlAttributeText(reader, "r", out bool present));
                    if (!present) continue;
                    if (currentRow <= 0 || currentRow > ExcelA1.MaxRows) return null;
                    if (currentRow > lastRow) lastRow = currentRow;
                } else if (name == "c") {
                    if (!IsKnownXmlCellType(ReadXmlCellTypeAttribute(reader))) return null;
                    ReadOnlySpan<char> reference = ReadXmlAttributeText(reader, "r", out bool present);
                    if (!present
                        || !ExcelA1.TryParseCellReferenceFast(reference, out int row, out int column)
                        || row != currentRow
                        || column > ExcelA1.MaxColumns) {
                        return null;
                    }
                    if (column > lastColumn) lastColumn = column;
                }
            }

            return new WorksheetExtent(lastRow, lastColumn);
        }
    }
}
