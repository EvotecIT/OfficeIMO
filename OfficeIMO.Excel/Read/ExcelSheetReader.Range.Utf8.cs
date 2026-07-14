using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// UTF-8 range materialization helpers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private static bool ShouldAttemptUtf8Range(int firstRow, int lastRow) {
            return firstRow > 1 || lastRow - firstRow + 1 > BufferedRangeStreamRowLimit;
        }

        private bool RangeReachesDeclaredWorksheetEnd(int lastRow) {
            return CanStreamWorksheetPart()
                && TryGetWorksheetDimensionReferenceFromXml(out string dimensionReference)
                && A1.TryParseRange(dimensionReference, out _, out _, out int dimensionLastRow, out _)
                && lastRow >= dimensionLastRow;
        }

        private bool TryFillRangeUtf8Fast(
            object?[,] result,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            CancellationToken ct,
            bool requireAllWorksheetCellsWithinRange = false) {
            int width = lastColumn - firstColumn + 1;
            if (width <= 0) {
                return false;
            }

            if (!ExcelUtf8RangeRowSource.TryCreate(this, firstRow, lastRow, firstColumn, width, ct, out var source)) {
                return false;
            }

            using (source) {
                if (requireAllWorksheetCellsWithinRange && !source!.CellsFitWithinRange) {
                    return false;
                }

                for (int row = firstRow; row <= lastRow; row++) {
                    if (!source!.SelectRow(row)) {
                        continue;
                    }

                    int rowOffset = row - firstRow;
                    for (int columnOffset = 0; columnOffset < width; columnOffset++) {
                        source.ReadValue(
                            columnOffset,
                            XmlDataReaderTargetKind.None,
                            out _,
                            out _,
                            out _,
                            out _,
                            out _,
                            out result[rowOffset, columnOffset]);
                    }
                }

                return true;
            }
        }
    }
}
