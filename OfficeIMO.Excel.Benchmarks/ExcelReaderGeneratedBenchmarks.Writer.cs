#if NET10_0_OR_GREATER
using ExcelReader.Core.Writer;

namespace OfficeIMO.Excel.Benchmarks;

public partial class ExcelReaderGeneratedBenchmarks {
    internal static byte[] WriteOfficeIMO(ExcelFileFormat format, int rowCount, bool escapedNames = false) {
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(format, output,
            new ExcelWriterOptions { IncludeCellReferences = false, UseSharedStrings = false }, leaveOpen: true)) {
            writer.WriteTable("Data", Generate(rowCount, escapedNames), Headers,
                static (row, value) => row.Write(value.Name).Write(value.Id).Write(value.CreatedOn).Write(value.Amount));
        }
        return output.ToArray();
    }

    internal static byte[] WriteOfficeIMOBuffered(ExcelFileFormat format, int rowCount) {
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            var sheet = document.AddSheet("Data");
            sheet.WriteData(Generate(rowCount), ("Name", static row => row.Name), ("Id", static row => row.Id),
                ("CreatedOn", static row => row.CreatedOn), ("Amount", static row => row.Amount));
            document.Save(output, new ExcelSaveOptions { Format = format });
        }
        return output.ToArray();
    }

    internal static byte[] WriteExcelReader(ExcelFileFormat format, int rowCount) {
        using var output = new MemoryStream();
        switch (format) {
            case ExcelFileFormat.Xlsx:
                using (var workbook = XlsxWorkbookWriter.Create(output, leaveOpen: true)) {
                    using (var sheet = workbook.AddSheet("Data")) {
                        using (var row = sheet.StartRow()) WriteHeaders(row);
                        foreach (var value in Generate(rowCount)) { using var row = sheet.StartRow(); WriteValues(row, value); }
                        sheet.End();
                    }
                    workbook.End();
                }
                break;
            case ExcelFileFormat.Xlsb:
                using (var workbook = XlsbWorkbookWriter.Create(output, leaveOpen: true)) {
                    using (var sheet = workbook.AddSheet("Data")) {
                        using (var row = sheet.StartRow()) WriteHeaders(row);
                        foreach (var value in Generate(rowCount)) { using var row = sheet.StartRow(); WriteValues(row, value); }
                        sheet.End();
                    }
                    workbook.End();
                }
                break;
            case ExcelFileFormat.Xls:
                using (var workbook = XlsWorkbookWriter.Create(output, leaveOpen: true)) {
                    using (var sheet = workbook.AddSheet("Data")) {
                        using (var row = sheet.StartRow()) WriteHeaders(row);
                        foreach (var value in Generate(rowCount)) { using var row = sheet.StartRow(); WriteValues(row, value); }
                        sheet.End();
                    }
                    workbook.End();
                }
                break;
            default: throw new ArgumentOutOfRangeException(nameof(format));
        }
        return output.ToArray();
    }

    private static void WriteHeaders<TRow>(TRow row) where TRow : IRowWriter {
        foreach (string header in Headers) row.Write(header);
    }

    private static void WriteValues<TRow>(TRow row, GeneratedRecord value) where TRow : IRowWriter {
        row.Write(value.Name); row.Write(value.Id); row.Write(value.CreatedOn); row.Write(value.Amount);
    }

    private static IEnumerable<GeneratedRecord> Generate(int count, bool escapedNames = false) {
        for (int index = 0; index < count; index++) yield return CreateRecord(index, escapedNames);
    }
}
#endif
