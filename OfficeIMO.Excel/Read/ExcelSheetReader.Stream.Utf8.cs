using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// UTF-8 streaming range helpers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private bool TryCreateRangeStreamUtf8(
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            CancellationToken ct,
            out ExcelUtf8RangeRowSource? source) {
            source = null;
            int width = lastColumn - firstColumn + 1;
            int rowCount = lastRow - firstRow + 1;
            if (width <= 0 || rowCount <= 0) {
                return false;
            }

            return ExcelUtf8RangeRowSource.TryCreate(this, firstRow, lastRow, firstColumn, width, ct, out source);
        }

        private static IEnumerable<RangeChunk> ReadRangeStreamUtf8(
            ExcelUtf8RangeRowSource source,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            int chunkRows,
            CancellationToken ct) {
            int width = lastColumn - firstColumn + 1;
            for (int startRow = firstRow; startRow <= lastRow; startRow += chunkRows) {
                ct.ThrowIfCancellationRequested();
                int currentRowCount = System.Math.Min(chunkRows, lastRow - startRow + 1);
                var rows = new object?[currentRowCount][];
                for (int rowOffset = 0; rowOffset < currentRowCount; rowOffset++) {
                    var rowValues = new object?[width];
                    int rowIndex = startRow + rowOffset;
                    if (source.SelectRow(rowIndex)) {
                        for (int columnOffset = 0; columnOffset < width; columnOffset++) {
                            source.ReadValue(
                                columnOffset,
                                XmlDataReaderTargetKind.None,
                                out _,
                                out _,
                                out _,
                                out _,
                                out rowValues[columnOffset]);
                        }
                    }

                    rows[rowOffset] = rowValues;
                }

                yield return new RangeChunk(startRow, currentRowCount, firstColumn, width, rows);
            }
        }
    }
}
