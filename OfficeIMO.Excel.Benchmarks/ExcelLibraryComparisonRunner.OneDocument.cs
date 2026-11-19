#if NET10_0_OR_GREATER
using System.Data;
using System.Globalization;
using RecursiveAI.OneDocument.Spreadsheets.Streaming;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class ExcelLibraryComparisonRunner {
    private static int OneDocumentWriteDataReaderPlain(DataTable table)
        => ByteCount(OneDocumentWriteDataReaderPlainBytes(table));

    /// <summary>
    /// Streams the DataTable-backed reader through OneDocument's forward-only XLSX writer
    /// with its default options (shared strings, default deflate level).
    /// </summary>
    internal static byte[] OneDocumentWriteDataReaderPlainBytes(DataTable table) {
        using var stream = new MemoryStream();
        using (XlsxStreamWriter workbook = XlsxStreamWriter.Create(stream)) {
            int dateTimeStyle = workbook.AddFormat(StreamCellFormat.DateTime);
            SheetStreamWriter sheet = workbook.AddSheet("Data");

            using var reader = table.CreateDataReader();
            sheet.BeginRow();
            for (int column = 0; column < reader.FieldCount; column++) {
                sheet.WriteText(reader.GetName(column));
            }

            while (reader.Read()) {
                sheet.BeginRow();
                for (int column = 0; column < reader.FieldCount; column++) {
                    WriteOneDocumentValue(sheet, reader.GetValue(column), dateTimeStyle);
                }
            }

            workbook.Finish();
        }

        return stream.ToArray();
    }

    private static void WriteOneDocumentValue(SheetStreamWriter sheet, object? value, int dateTimeStyle) {
        switch (value) {
            case null or DBNull:
                sheet.SkipColumns(1);
                break;
            case string text:
                sheet.WriteText(text);
                break;
            case int number:
                sheet.WriteNumber(number);
                break;
            case double number:
                sheet.WriteNumber(number);
                break;
            case DateTime dateTime:
                sheet.WriteDate(dateTime, dateTimeStyle);
                break;
            case bool flag:
                sheet.WriteBoolean(flag);
                break;
            default:
                sheet.WriteText(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                break;
        }
    }
}
#endif
