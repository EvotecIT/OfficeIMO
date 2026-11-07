namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private static void ValidateDataReaderRangeOptions(ExcelReaderOptions options, int chunkRows, int schemaSampleRows) {
            if (chunkRows <= 0 || chunkRows > options.MaxDataReaderChunkRows) {
                throw new ArgumentOutOfRangeException(nameof(chunkRows),
                    $"Chunk row count must be between 1 and {options.MaxDataReaderChunkRows}.");
            }
            if (schemaSampleRows < 0) {
                throw new ArgumentOutOfRangeException(nameof(schemaSampleRows), "Schema sample row count cannot be negative.");
            }
            if (schemaSampleRows > options.MaxDataReaderSchemaSampleRows) {
                throw new ArgumentOutOfRangeException(nameof(schemaSampleRows),
                    $"Schema sample row count exceeds {options.MaxDataReaderSchemaSampleRows}.");
            }
            if (options.MaxDataReaderColumns <= 0 || options.MaxDataReaderChunkRows <= 0
                || options.MaxDataReaderSchemaSampleRows < 0 || options.MaxDataReaderBufferedCells <= 0L) {
                throw new InvalidOperationException("Excel data-reader safety limits must be positive.");
            }
        }

        private static void ValidateDataReaderRangeShape(ExcelReaderOptions options, int rows, int columns,
            bool headersInFirstRow, int schemaSampleRows) {
            if (columns > options.MaxDataReaderColumns) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxDataReaderColumns), options.MaxDataReaderColumns,
                    $"Range data-reader column count exceeds {nameof(ExcelReaderOptions.MaxDataReaderColumns)}.");
            }
            long sampleCells = (long)Math.Min(Math.Max(0, rows - (headersInFirstRow ? 1 : 0)), schemaSampleRows) * columns;
            if (sampleCells > options.MaxDataReaderBufferedCells) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxDataReaderBufferedCells), options.MaxDataReaderBufferedCells,
                    $"Range data-reader buffering exceeds {nameof(ExcelReaderOptions.MaxDataReaderBufferedCells)}.");
            }
        }

        private static void ValidateDataReaderBufferedRows(ExcelReaderOptions options, int rows, int columns) {
            if ((long)rows * columns > options.MaxDataReaderBufferedCells) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxDataReaderBufferedCells), options.MaxDataReaderBufferedCells,
                    $"Range data-reader buffering exceeds {nameof(ExcelReaderOptions.MaxDataReaderBufferedCells)}.");
            }
        }
    }
}
