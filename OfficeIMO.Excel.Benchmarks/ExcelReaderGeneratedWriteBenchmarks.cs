#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Complete in-memory packages from the same generated typed rows, with independent output validation.</summary>
[MemoryDiagnoser]
public class ExcelReaderGeneratedWriteBenchmarks {
    [Params(ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb, ExcelFileFormat.Xls)]
    public ExcelFileFormat Format { get; set; }

    [Params(50_000)]
    public int RowCount { get; set; }

    [GlobalSetup(Target = nameof(OfficeIMO))]
    public void SetupOfficeIMO() {
        ApplyPriorityOverride();
        Record("OfficeIMO", "Default", () => Validate(OfficeIMO()));
    }

    [GlobalSetup(Target = nameof(ExcelReaderNet))]
    public void SetupExcelReader() {
        ApplyPriorityOverride();
        Record("ExcelReader.NET", "Default", () => Validate(ExcelReaderNet()));
    }

    [GlobalSetup(Target = nameof(OfficeIMO_WriteData))]
    public void SetupOfficeIMOBuffered() {
        ApplyPriorityOverride();
        Record("OfficeIMO", "WriteData", () => Validate(OfficeIMO_WriteData()));
    }

    [GlobalSetup(Target = nameof(OfficeIMO_MatchingWriteData))]
    public void SetupOfficeIMOMatching() {
        ApplyPriorityOverride();
        Record("OfficeIMO", "MatchingWriteData", () => Validate(OfficeIMO_MatchingWriteData()));
    }

    [Benchmark]
    [BenchmarkCategory("case:WriteData")]
    public byte[] OfficeIMO_WriteData() => ExcelReaderGeneratedBenchmarks.WriteOfficeIMOBuffered(Format, RowCount);

    /// <summary>The typed writer producing the same cells, references and string storage as <see cref="OfficeIMO_WriteData"/>.</summary>
    [Benchmark]
    [BenchmarkCategory("case:MatchingWriteData")]
    public byte[] OfficeIMO_MatchingWriteData() => ExcelReaderGeneratedBenchmarks.WriteOfficeIMOMatchingBuffered(Format, RowCount);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("case:Default")]
    public byte[] OfficeIMO() => ExcelReaderGeneratedBenchmarks.WriteOfficeIMO(Format, RowCount);

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public byte[] ExcelReaderNet() => ExcelReaderGeneratedBenchmarks.WriteExcelReader(Format, RowCount);

    /// <summary>Validates one lane's output and records the result; ExcelDataReader rereads every row of the package.</summary>
    internal void Record(string library, string caseName, Action validate) =>
        ExcelOutputValidation.Check(library, caseName,
            $"Format={Format}; RowCount={RowCount}; ExcelDataReader reread every row's name, id, date and amount", validate);

    // In-process execution may raise priority after the command-line override.
    private static void ApplyPriorityOverride() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority))
            Console.WriteLine($"Benchmark process priority: {BenchmarkProcessorAffinity.ApplyPriority(priority)}.");
    }

    internal void Validate(byte[] output) {
        if (Format == ExcelFileFormat.Xls) {
            var blocks = ExcelNativeBinaryWriteBenchmarks.InspectLegacyXlsDbCellBlocks(output, RowCount);
            if (blocks.Actual != blocks.Expected)
                throw new InvalidDataException($"XLS Index/DBCell structure missing or incomplete ({blocks.Actual} of {blocks.Expected} blocks). Timing withheld.");
        }
        using (var strict = ExcelDocument.OpenReader(output)) {
            while (strict.Read()) { }
        }
        // ExcelDataReader is independent of both writers and accepts all three formats.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using var stream = new MemoryStream(output);
        using var reader = global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read()) throw new InvalidDataException("Missing header.");
        ExcelReaderGeneratedBenchmarks.ValidateHeaders(reader);
        int count = 0;
        while (reader.Read()) {
            ExcelReaderGeneratedBenchmarks.Validate(count++, reader.GetString(0), checked((int)reader.GetDouble(1)), reader.GetDateTime(2), reader.GetDouble(3));
        }
        if (count != RowCount || reader.NextResult()) throw new InvalidDataException("Incorrect row or sheet count.");
        if (Format == ExcelFileFormat.Xlsb) {
            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(output), System.IO.Compression.ZipArchiveMode.Read);
            var worksheet = zip.GetEntry("xl/worksheets/sheet1.bin") ?? throw new InvalidDataException("Missing worksheet.");
            using var part = worksheet.Open();
            string hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(part));
            Console.WriteLine($"ValidatedXlsb: package={output.Length}; worksheet={worksheet.Length}; compressed={worksheet.CompressedLength}; sha256={hash}");
        }
    }
}
#endif
