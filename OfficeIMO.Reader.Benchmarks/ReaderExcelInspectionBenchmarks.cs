using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using OfficeIMO.Excel;
using OfficeIMO.Reader.Excel;

namespace OfficeIMO.Reader.Benchmarks;

/// <summary>
/// Measures workbook inspection and Reader extraction of the same workbook, from a few rows, where the per-workbook
/// cost dominates, to a large sheet, where the per-cell cost does.
/// </summary>
[MemoryDiagnoser]
[ShortRunJob(RuntimeMoniker.Net10_0)]
public class ReaderExcelInspectionBenchmarks {
    private OfficeDocumentReader _reader = null!;
    private ReaderOptions _options = null!;
    private byte[] _workbook = Array.Empty<byte>();

    [Params(10, 2_000, 50_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup() {
        _reader = new OfficeDocumentReaderBuilder().AddExcelHandler().Build();
        _workbook = BuildWorkbook(Rows);
        _options = new ReaderOptions {
            ComputeHashes = false,
            MaxChars = 4_000,
            MaxTableRows = 5_000,
            OpenXmlMaxCharactersInPart = 256L * 1024L * 1024L
        };
    }

    [Benchmark]
    public ExcelWorkbookInfo Inspect() {
        using var stream = new MemoryStream(_workbook, writable: false);
        using ExcelDocument document = ExcelDocument.Load(stream, new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly,
            OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = 256L * 1024L * 1024L }
        });
        return document.Inspect();
    }

    [Benchmark]
    public OfficeDocumentReadResult ReadDocument() =>
        _reader.ReadDocument(_workbook, "inspection.xlsx", _options);

    private static byte[] BuildWorkbook(int rows) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        string[] headers = { "Id", "Name", "Share", "Amount", "Date", "Ratio", "Code" };
        for (int column = 0; column < headers.Length; column++) sheet.Cell(1, column + 1).SetValue(headers[column]);
        var start = new DateTime(2020, 1, 1);
        for (int row = 2; row <= rows + 1; row++) {
            int item = row - 1;
            sheet.Cell(row, 1).SetValue(item);
            sheet.Cell(row, 2).SetValue("Item " + item);
            sheet.Cell(row, 3).SetValue(item % 1000 / 1000d);
            sheet.Cell(row, 4).SetValue(item * 12.5m);
            sheet.Cell(row, 5).SetValue(start.AddDays(item % 3650));
            sheet.Cell(row, 6).SetValue(item * 1234.5678d);
            sheet.Cell(row, 7).SetValue("C" + item);
        }
        sheet.Range(2, 3, rows + 1, 3).SetNumberFormat("0.00%");
        sheet.Range(2, 5, rows + 1, 5).SetNumberFormat("yyyy-mm-dd");
        ExcelSheet notes = document.AddSheet("Notes");
        notes.Cell(1, 1).SetValue("Note");
        notes.Cell(2, 1).SetValue("Reviewed");
        return document.ToBytes();
    }
}
