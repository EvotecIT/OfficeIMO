using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>The three no-header raw workbook formats in the upstream read suite.</summary>
public enum RawWorkbookFormat { Xlsx, Xlsb, Xls }

internal static class RawWorkbookFixture {
    internal static async Task<byte[]> CreateAsync(int rows, RawWorkbookFormat format) {
        byte[] bytes = await (format switch {
            RawWorkbookFormat.Xlsx => CreateAsync<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(rows,
                static stream => XlsxWorkbookWriter.Create(stream, leaveOpen: true)),
            RawWorkbookFormat.Xlsb => CreateAsync<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(rows,
                static stream => XlsbWorkbookWriter.Create(stream, leaveOpen: true)),
            RawWorkbookFormat.Xls when rows <= 65_536 => Task.FromResult(XlsBenchmarkWorkbookGenerator.Build(rows)),
            _ => throw new ArgumentOutOfRangeException(nameof(rows), "The upstream XLS fixture supports up to 65536 rows."),
        });
        BenchmarkInput.WriteWorkbookFixtureIdentity($"read/raw/{format}/dataRows={rows}", bytes,
            zipPackage: format != RawWorkbookFormat.Xls);
        return bytes;
    }

    private static async Task<byte[]> CreateAsync<TWorkbook, TSheet, TRow>(int rows, Func<MemoryStream, TWorkbook> create)
        where TWorkbook : IWorkbookWriter<TSheet>
        where TSheet : ISheetWriter<TRow>
        where TRow : IRowWriter {
        await using var stream = new MemoryStream();
        await using (TWorkbook workbook = create(stream)) {
            TSheet sheet = workbook.AddSheet("S1");
            for (int index = 1; index <= rows; index++) {
                TypedRecord record = TypedWorkbookFixture.ExpectedRecord(index);
                await using TRow row = await sheet.StartRowAsync();
                row.Write(record.Name);
                row.Write(record.Id);
                row.Write(record.Date);
                row.Write(record.Value);
            }
            await sheet.EndAsync();
            await workbook.EndAsync();
        }
        return stream.ToArray();
    }
}
