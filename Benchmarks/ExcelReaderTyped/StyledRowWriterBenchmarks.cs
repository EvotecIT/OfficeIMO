#if OFFICEIMO_BENCHMARK_NEW_APIS
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Writes the original 100,000 numeric rows with a real bold row default.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("WriteStyledRows")]
public class StyledRowWriterBenchmarks {
    internal const int Rows = 100_000;
    private readonly MemoryStream _output = new(32 * 1024 * 1024);

    [GlobalSetup]
    public void Setup() {
        BenchmarkInput.WriteDescription();
        OfficeIMODefaultRowStyle();
        StyledWorkbookValidation.Validate(_output.ToArray(), "OfficeIMO", requireCoordinates: true);
        ExcelReaderOriginalStyledRows();
        StyledWorkbookValidation.Validate(_output.ToArray(), "ExcelReader", requireCoordinates: false);
    }

    [GlobalCleanup]
    public void Cleanup() => _output.Dispose();

    [Benchmark]
    public long OfficeIMODefaultRowStyle() {
        _output.SetLength(0);
        ExcelDocument.WriteRows(_output, Enumerable.Range(0, Rows), ["Value"], static (row, value) => row.Write(value),
            new ExcelTabularWriteOptions {
                SheetName = "S", IncludeHeaders = false, RequireStreaming = true,
                Styles = new Dictionary<string, ExcelStyleDefinition> { ["Bold"] = new() { Bold = true } },
                DefaultRowStyle = "Bold",
            });
        return _output.Length;
    }

    [Benchmark(Baseline = true)]
    public long ExcelReaderOriginalStyledRows() {
        _output.SetLength(0);
        using (var workbook = XlsxWorkbookWriter.Create(_output, leaveOpen: true)) {
            int style = workbook.AddStyle(new CellStyle { Bold = true });
            using XlsxSheetWriter sheet = workbook.AddSheet("S");
            for (int index = 0; index < Rows; index++) {
                using XlsxRowWriter row = sheet.StartRow(style);
                row.Write(index);
            }
            sheet.End();
            workbook.End();
        }
        return _output.Length;
    }

    // Explicit qualification command only; BDN setup never creates artifact files.
    internal void SaveQualifiedArtifacts() {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_STYLED_WRITE_ARTIFACTS");
        if (directory == null) return;
        if (!Path.IsPathFullyQualified(directory) || !Directory.Exists(directory))
            throw new InvalidOperationException("OFFICEIMO_STYLED_WRITE_ARTIFACTS must name an existing absolute task directory.");
        foreach (string engine in new[] { "OfficeIMO", "ExcelReader" }) {
            if (engine == "OfficeIMO") OfficeIMODefaultRowStyle(); else ExcelReaderOriginalStyledRows();
            byte[] bytes = _output.ToArray();
            StyledWorkbookValidation.Validate(bytes, engine, requireCoordinates: engine == "OfficeIMO");
            string path = Path.Combine(directory, engine + "-StyledRows-100000.xlsx");
            using var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            output.Write(bytes);
            Console.WriteLine($"Preserved qualified styled output: {path}.");
        }
    }
}
#endif
