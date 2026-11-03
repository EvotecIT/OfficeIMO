#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Complete in-memory packages from the same generated typed rows, with independent output validation.</summary>
[MemoryDiagnoser]
public class ExcelReaderGeneratedWriteBenchmarks {
    [Params(ExcelFileFormat.Xlsb, ExcelFileFormat.Xls)]
    public ExcelFileFormat Format { get; set; }

    [Params(50_000)]
    public int RowCount { get; set; }

    [GlobalSetup(Target = nameof(OfficeIMO))]
    public void SetupOfficeIMO() => Validate(OfficeIMO());

    [GlobalSetup(Target = nameof(ExcelReaderNet))]
    public void SetupExcelReader() => Validate(ExcelReaderNet());

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => ExcelReaderGeneratedBenchmarks.WriteOfficeIMO(Format, RowCount);

    [Benchmark]
    public byte[] ExcelReaderNet() => ExcelReaderGeneratedBenchmarks.WriteExcelReader(Format, RowCount);

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
