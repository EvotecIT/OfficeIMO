using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Measures range discovery alone, including sparse omitted row coordinates.</summary>
[MemoryDiagnoser]
public class ExcelUsedRangeDiscoveryBenchmarks {
    private string _path = string.Empty;
    private string _expected = string.Empty;

    [Params(2500, 25000)]
    public int RowCount { get; set; }

    [Params(false, true)]
    public bool Sparse { get; set; }

    [Params(false, true)]
    public bool OmitRowCoordinates { get; set; }

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA") ?? Path.GetTempPath();
        Directory.CreateDirectory(root);
        _path = Path.Combine(root, $"used-range-discovery-{Guid.NewGuid():N}.xlsx");
        _expected = $"A1:B{(Sparse ? 2 * RowCount : RowCount) + 1}";
        try {
            using (Stream output = File.Create(_path)) {
                ExcelDocument.WriteRows(output,
                    ExcelGeneratedRowStreamingBenchmarks.GenerateRows(RowCount), ["Id", "Amount"],
                    static (writer, row) => writer.Write(row.Id).Write(row.Amount),
                    new ExcelTabularWriteOptions { IncludeCellReferences = true, UseSharedStrings = false });
            }
            using (var package = ZipFile.Open(_path, ZipArchiveMode.Update)) {
                const string name = "xl/worksheets/sheet1.xml";
                var entry = package.GetEntry(name) ?? throw new InvalidDataException("Worksheet is missing.");
                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
                if (Regex.Matches(xml, "<row r=\"[0-9]+\"").Count != RowCount + 1
                    || Regex.Matches(xml, "<c r=\"[A-Z]+[0-9]+\"").Count != (RowCount + 1) * 2) {
                    throw new InvalidDataException("Generated worksheet coordinates differ.");
                }
                if (Sparse) {
                    xml = Regex.Replace(xml, "(<(?:row|c) r=\"[A-Z]*)([0-9]+)(\")", match =>
                        match.Groups[1].Value
                        + (2 * int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) - 1).ToString(CultureInfo.InvariantCulture)
                        + match.Groups[3].Value);
                }
                if (OmitRowCoordinates) xml = Regex.Replace(xml, "(<row) r=\"[0-9]+\"", "$1");
                // A dimension would let discovery avoid scanning the rows.
                xml = Regex.Replace(xml, "<dimension\\b[^>]*/>", string.Empty);
                entry.Delete();
                using var writer = new StreamWriter(package.CreateEntry(name, CompressionLevel.Fastest).Open(), Encoding.Unicode);
                writer.Write(xml.Replace("utf-8", "utf-16").Replace("UTF-8", "utf-16"));
            }
            DiscoverRange();
        } catch {
            Cleanup();
            throw;
        }
    }

    [Benchmark]
    public string DiscoverRange() {
        using var reader = ExcelDocumentReader.Open(_path);
        string range = reader.GetSheet("Data").GetUsedRangeA1();
        return range == _expected ? range : throw new InvalidDataException($"Unexpected used range: {range}.");
    }

    [GlobalCleanup]
    public void Cleanup() {
        if (_path.Length != 0) File.Delete(_path);
    }
}
