using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Streaming APIs for large ranges.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Represents a rectangular block of rows produced during streaming.
        /// </summary>
        public sealed class RangeChunk {
            /// <summary>First row index (1-based) covered by this chunk.</summary>
            public int StartRow { get; }
            /// <summary>Number of rows in this chunk.</summary>
            public int RowCount { get; }
            /// <summary>First column index (1-based) covered by this chunk.</summary>
            public int StartCol { get; }
            /// <summary>Number of columns in this chunk.</summary>
            public int ColCount { get; }
            /// <summary>Row-major values array. Size is <see cref="RowCount"/> x <see cref="ColCount"/>.</summary>
            public object?[][] Rows { get; }

            internal RangeChunk(int startRow, int rowCount, int startCol, int colCount, object?[][] rows) {
                StartRow = startRow; RowCount = rowCount; StartCol = startCol; ColCount = colCount; Rows = rows;
            }
        }
    }
}