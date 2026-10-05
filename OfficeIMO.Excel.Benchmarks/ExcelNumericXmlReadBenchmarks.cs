using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Compares numeric XML decoding across public result shapes and number modes.
/// UTF-16 exercises the streaming fallback independently of worksheet size.
/// This is an OfficeIMO before/after lane, not a comparison between result shapes.
/// </summary>
[MemoryDiagnoser]
public class ExcelNumericXmlReadBenchmarks {
    private static readonly string[] Headers = ["Id", "Amount"];
    private string _path = string.Empty;
    private string _range = string.Empty;
    private string _tableRange = string.Empty;
    private long _expected;

    [Params(2500, 25000)]
    public int RowCount { get; set; }

    [Params(false, true)]
    public bool NumericAsDecimal { get; set; }

    [Params(false, true)]
    public bool InferDataTableColumnTypes { get; set; }

    [Params("DataReader", "TypedDataReader", "Range", "UsedRange", "DataTable")]
    public string Api { get; set; } = "DataReader";

    [Params("Explicit", "ImplicitRows", "ImplicitRowsAndCells")]
    public string Coordinates { get; set; } = "Explicit";

    [GlobalSetup]
    public void Setup() {
        if (Coordinates is not ("Explicit" or "ImplicitRows" or "ImplicitRowsAndCells"))
            throw new ArgumentOutOfRangeException(nameof(Coordinates));
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA") ?? Path.GetTempPath();
        Directory.CreateDirectory(root);
        _path = Path.Combine(root, $"numeric-xml-read-{Guid.NewGuid():N}.xlsx");
        _range = $"A2:B{RowCount + 1}";
        _tableRange = $"A1:B{RowCount + 1}";
        _expected = RowCount * (RowCount + 1L) / 2L + 125L * RowCount * (RowCount - 1L) / 2L;
        try {
            using (Stream output = File.Create(_path)) {
                ExcelDocument.WriteRows(output,
                    ExcelGeneratedRowStreamingBenchmarks.GenerateRows(RowCount), Headers,
                    static (writer, row) => writer.Write(row.Id).Write(row.Amount),
                    new ExcelTabularWriteOptions { IncludeCellReferences = Coordinates != "ImplicitRowsAndCells", UseSharedStrings = false });
            }
            using (var package = ZipFile.Open(_path, ZipArchiveMode.Update)) {
                const string name = "xl/worksheets/sheet1.xml";
                var entry = package.GetEntry(name) ?? throw new InvalidDataException("Worksheet is missing.");
                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
                if (Coordinates == "ImplicitRows") {
                    if (System.Text.RegularExpressions.Regex.Matches(xml, "(<row) r=\"[0-9]+\"").Count != RowCount + 1)
                        throw new InvalidDataException("Unexpected generated row references.");
                    xml = System.Text.RegularExpressions.Regex.Replace(xml, "(<row) r=\"[0-9]+\"", "$1");
                }
                entry.Delete();
                using var writer = new StreamWriter(package.CreateEntry(name, CompressionLevel.Fastest).Open(), Encoding.Unicode);
                writer.Write(xml.Replace("utf-8", "utf-16").Replace("UTF-8", "utf-16"));
            }
            Check(ReadCore(validate: true));
        } catch {
            Cleanup();
            throw;
        }
    }

    [Benchmark]
    public long Read() => Check(ReadCore(validate: false));

    [GlobalCleanup]
    public void Cleanup() {
        if (_path.Length != 0) File.Delete(_path);
    }

    private long ReadCore(bool validate) {
        var options = new ExcelReadOptions { NumericAsDecimal = NumericAsDecimal, InferDataTableColumnTypes = InferDataTableColumnTypes };
        long result = 0;
        int count = 0;
        if (Api == "DataReader" || Api == "TypedDataReader") {
            using var reader = ExcelDocument.OpenDataReader(_path, options);
            if (reader.FieldCount != 2 || reader.GetName(0) != Headers[0] || reader.GetName(1) != Headers[1]) {
                throw new InvalidDataException("Reader headers differ.");
            }
            while (reader.Read()) {
                if (Api == "TypedDataReader") {
                    int id = reader.GetInt32(0);
                    decimal amount = reader.GetDecimal(1);
                    if (validate) {
                        if (id != count + 1 || amount != count * 1.25m) throw new InvalidDataException("Typed numeric values differ.");
                        Add(ref result, reader.GetValue(0), reader.GetValue(1), count, validate: true);
                    } else {
                        result = unchecked(result + id + (long)(amount * 100));
                    }
                } else {
                    Add(ref result, reader.GetValue(0), reader.GetValue(1), count, validate);
                }
                count++;
            }
        } else {
            using var owner = ExcelDocumentReader.Open(_path, options);
            var sheet = owner.GetSheet("Data");
            if (Api == "Range" || Api == "UsedRange") {
                bool discoverRange = Api == "UsedRange";
                if (validate && !discoverRange) {
                    var headers = sheet.ReadRange("A1:B1", ExcelExecutionMode.Sequential);
                    if (!Equals(headers[0, 0], Headers[0]) || !Equals(headers[0, 1], Headers[1])) {
                        throw new InvalidDataException("Range headers differ.");
                    }
                }
                string range = discoverRange ? sheet.GetUsedRangeA1() : _range;
                if (discoverRange && range != _tableRange) throw new InvalidDataException("Used range differs.");
                var rows = sheet.ReadRange(range, ExcelExecutionMode.Sequential);
                if (rows.GetLength(1) != 2) throw new InvalidDataException("Range width differs.");
                if (discoverRange && (!Equals(rows[0, 0], Headers[0]) || !Equals(rows[0, 1], Headers[1]))) {
                    throw new InvalidDataException("Used range headers differ.");
                }
                for (int row = discoverRange ? 1 : 0; row < rows.GetLength(0); row++) Add(ref result, rows[row, 0], rows[row, 1], count++, validate);
            } else if (Api == "DataTable") {
                using DataTable table = sheet.ReadRangeAsDataTable(_tableRange, headersInFirstRow: true);
                if (table.Columns.Count != 2 || table.Columns[0].ColumnName != Headers[0] || table.Columns[1].ColumnName != Headers[1]) {
                    throw new InvalidDataException("DataTable headers differ.");
                }
                foreach (DataRow row in table.Rows) Add(ref result, row[0], row[1], count++, validate);
            } else {
                throw new InvalidOperationException($"Unknown API '{Api}'.");
            }
        }
        if (count != RowCount) throw new InvalidDataException("Row count differs.");
        return result;
    }

    private void Add(ref long result, object? id, object? amount, int row, bool validate) {
        decimal idValue = Convert.ToDecimal(id, CultureInfo.InvariantCulture);
        decimal amountValue = Convert.ToDecimal(amount, CultureInfo.InvariantCulture);
        if (validate) {
            Type expectedType = NumericAsDecimal ? typeof(decimal) : typeof(double);
            if (id?.GetType() != expectedType || amount?.GetType() != expectedType
                || idValue != row + 1 || amountValue != row * 1.25m) {
                throw new InvalidDataException($"Unexpected numeric value or type at row {row + 2}.");
            }
        }
        result = unchecked(result + (long)idValue + (long)(amountValue * 100));
    }

    private long Check(long result) => result == _expected
        ? result : throw new InvalidDataException("Numeric observation differs.");
}
