using System.Data.Common;
using System.Threading;
using OfficeIMO.Excel.Legacy;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        /// <summary>Reads parsed early cells through the existing sparse scalar range reader.</summary>
        internal static DbDataReader ReadEarlyBiffRows(LegacySpreadsheetSheet sheet, ExcelDateSystem dateSystem,
            ExcelReaderOptions options, CancellationToken cancellationToken) {
            int schemaSampleRows = options.InferSchema ? options.SchemaSampleRows : 0;
            if (sheet.Cells.Count == 0) {
                return new ExcelRangeDataReader(Array.Empty<RangeChunk>(), 1, 0, 0,
                    options.HasHeaderRow, schemaSampleRows, options, cancellationToken);
            }
            int firstRow = sheet.Cells[0].Row;
            int lastRow = sheet.Cells[sheet.Cells.Count - 1].Row;
            int firstColumn = int.MaxValue;
            int lastColumn = 0;
            foreach (LegacySpreadsheetCell cell in sheet.Cells) {
                cancellationToken.ThrowIfCancellationRequested();
                firstColumn = Math.Min(firstColumn, cell.Column);
                lastColumn = Math.Max(lastColumn, cell.Column);
            }
            int columns = lastColumn - firstColumn + 1;
            ValidateDataReaderRangeOptions(options, chunkRows: 1, schemaSampleRows);
            ValidateDataReaderRangeShape(options, lastRow - firstRow + 1, columns, options.HasHeaderRow, schemaSampleRows);
            ValidateDataReaderBufferedRows(options, 1, columns);
            return new ExcelRangeDataReader(ReadEarlyBiffChunks(sheet, firstColumn, columns, dateSystem, options, cancellationToken),
                firstRow, lastRow, columns, options.HasHeaderRow, schemaSampleRows, options, cancellationToken);
        }

        private static IEnumerable<RangeChunk> ReadEarlyBiffChunks(LegacySpreadsheetSheet sheet, int firstColumn,
            int columns, ExcelDateSystem dateSystem, ExcelReaderOptions options, CancellationToken cancellationToken) {
            var dateStyles = new Dictionary<ushort, ExcelSerialDateStyle>();
            int index = 0;
            while (index < sheet.Cells.Count) {
                cancellationToken.ThrowIfCancellationRequested();
                int row = sheet.Cells[index].Row;
                var values = new object?[columns];
                do {
                    cancellationToken.ThrowIfCancellationRequested();
                    LegacySpreadsheetCell cell = sheet.Cells[index++];
                    object? value = cell.Value;
                    if (value is double number) {
                        string text = InvariantNumberText.Get(number);
                        if (TryParseExcelNumberAsDouble(text, options.Culture, out double parsed)) {
                            ExcelSerialDateStyle style = options.TreatDatesUsingNumberFormat
                                ? ReadEarlyBiffDateStyle(sheet, cell, dateStyles) : ExcelSerialDateStyle.None;
                            value = style == ExcelSerialDateStyle.None ? (object)parsed
                                : new ExcelDataReaderDateSerial(parsed, dateSystem, style == ExcelSerialDateStyle.Calendar);
                        } else value = text;
                    }
                    values[cell.Column - firstColumn] = value;
                } while (index < sheet.Cells.Count && sheet.Cells[index].Row == row);
                yield return new RangeChunk(row, 1, firstColumn, columns, new[] { values });
            }
        }

        private static ExcelSerialDateStyle ReadEarlyBiffDateStyle(LegacySpreadsheetSheet sheet, LegacySpreadsheetCell cell,
            Dictionary<ushort, ExcelSerialDateStyle> styles) {
            if (!string.IsNullOrWhiteSpace(cell.NumberFormat))
                return ExcelNumberFormatClassifier.ClassifySerialDateStyle(0, cell.NumberFormat);
            if (!cell.StyleIndex.HasValue || sheet.StyleWorkbook == null) return ExcelSerialDateStyle.None;
            ushort index = cell.StyleIndex.Value;
            if (styles.TryGetValue(index, out ExcelSerialDateStyle cached)) return cached;
            LegacyXlsCellFormat? format = sheet.StyleWorkbook.GetEffectiveCellFormat(index);
            // Early format IDs are local slots; only built-in formats use OOXML's IDs.
            ExcelSerialDateStyle style = format == null || format.NumberFormatId == 0 || format.NumberFormatCode == null
                ? ExcelSerialDateStyle.None
                : ExcelNumberFormatClassifier.ClassifySerialDateStyle(format.IsBuiltInNumberFormat ? format.NumberFormatId : 0U,
                    format.NumberFormatCode);
            styles.Add(index, style);
            return style;
        }
    }
}
