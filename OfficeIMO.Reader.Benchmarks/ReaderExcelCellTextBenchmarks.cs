using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using OfficeIMO.Excel;
using OfficeIMO.Reader.Excel;

namespace OfficeIMO.Reader.Benchmarks;

/// <summary>
/// Measures spreadsheet extraction on a large worksheet whose columns carry common number formats.
/// </summary>
[MemoryDiagnoser]
[ShortRunJob(RuntimeMoniker.Net80)]
public class ReaderExcelCellTextBenchmarks {
    private const int DataRows = 50_000;
    private OfficeDocumentReader _reader = null!;
    private ReaderOptions _options = null!;
    private byte[] _workbook = Array.Empty<byte>();

    [Params(ReaderExcelCellText.Raw, ReaderExcelCellText.Formatted)]
    public ReaderExcelCellText CellText { get; set; }

    [GlobalSetup]
    public void Setup() {
        _reader = new OfficeDocumentReaderBuilder().AddExcelHandler(new ReaderExcelOptions { CellText = CellText }).Build();
        _workbook = BuildWorkbook();
        _options = new ReaderOptions {
            ComputeHashes = false,
            MaxChars = 4_000,
            MaxTableRows = 5_000,
            // The formatted worksheet part is larger than the default Open XML part budget.
            OpenXmlMaxCharactersInPart = 256L * 1024L * 1024L
        };
    }

    [Benchmark]
    public OfficeDocumentReadResult ReadFormattedWorkbook() =>
        _reader.ReadDocument(_workbook, "formatted.xlsx", _options);

    private static byte[] BuildWorkbook() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-reader-cell-text-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddSheet("Data");
                string[] headers = { "Id", "Name", "Share", "Amount", "Date", "Ratio", "Code" };
                for (int column = 0; column < headers.Length; column++) sheet.Cell(1, column + 1).SetValue(headers[column]);
                var start = new DateTime(2020, 1, 1);
                for (int row = 2; row <= DataRows + 1; row++) {
                    int item = row - 1;
                    sheet.Cell(row, 1).SetValue(item);
                    sheet.Cell(row, 2).SetValue("Item " + item);
                    sheet.Cell(row, 3).SetValue(item % 1000 / 1000d);
                    sheet.Cell(row, 4).SetValue(item * 12.5m);
                    sheet.Cell(row, 5).SetValue(start.AddDays(item % 3650));
                    sheet.Cell(row, 6).SetValue(item * 1234.5678d);
                    sheet.Cell(row, 7).SetValue("C" + item);
                }
                int last = DataRows + 1;
                sheet.Range(2, 3, last, 3).SetNumberFormat("0.00%");
                sheet.Range(2, 4, last, 4).SetNumberFormat("\"$\"#,##0.00");
                sheet.Range(2, 5, last, 5).SetNumberFormat("yyyy-mm-dd");
                sheet.Range(2, 6, last, 6).SetNumberFormat("0.00E+00");
                sheet.Range(2, 7, last, 7).SetNumberFormat("@");
                document.Save();
            }
            return File.ReadAllBytes(path);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
