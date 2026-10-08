using System.Globalization;
using System.Text;
using System.Security.Cryptography;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

public sealed class CsvHeavyRecord {
    public string? Region { get; set; }
    public string? Country { get; set; }
    public DateTime OrderDate { get; set; }
    public decimal UnitPrice { get; set; }
    public decimal TotalRevenue { get; set; }
    public int Units { get; set; }
}

public sealed class CsvNarrowRecord {
    public int A { get; set; }
    public int B { get; set; }
    public int C { get; set; }
}

public readonly record struct CsvParallelInput(string Corpus, int Rows) {
    public override string ToString() => $"{Corpus}-{Rows}";
}

internal sealed class CsvParallelFixture : IDisposable {
    internal static readonly string[] Regions = ["Europe", "Asia", "North America", "Sub-Saharan Africa"];
    internal static readonly string[] Countries = ["Portugal", "Japan", "Canada", "Kenya", "Brazil", "Norway"];
    internal static readonly DateTime Start = new(2015, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
    private readonly string _directory;
    internal string Path { get; }
    internal CsvParallelInput Input { get; }
    internal bool Heavy => Input.Corpus.StartsWith("ConversionHeavy", StringComparison.Ordinal);
    internal long ExpectedSum { get; }
    internal long Bytes { get; }
    internal string Fingerprint { get; }

    internal CsvParallelFixture(CsvParallelInput input) {
        Input = input;
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_OUTPUT") ?? System.IO.Path.GetTempPath();
        if (!Directory.Exists(root)) throw new DirectoryNotFoundException("The configured benchmark output root must already exist.");
        _directory = System.IO.Path.Combine(root, $"officeimo-csv-parallel-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_directory);
        Path = System.IO.Path.Combine(_directory, "input.csv");
        try {
            using (var writer = new StreamWriter(Path, false, new UTF8Encoding(false))) {
                writer.WriteLine(Heavy ? "Region,Country,OrderDate,UnitPrice,TotalRevenue,Units" : "A,B,C");
                long sum = 0;
                for (int index = 0; index < input.Rows; index++) {
                    if (Heavy) {
                        decimal price = 10m + index % 9000 / 100m;
                        int units = index % 500 + 1;
                        writer.Write(Regions[index % 4]); writer.Write(',');
                        writer.Write(Countries[index % 6]); writer.Write(',');
                        writer.Write(Start.AddDays(index % 3650).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)); writer.Write(',');
                        writer.Write(price.ToString(CultureInfo.InvariantCulture)); writer.Write(',');
                        writer.Write((price * units).ToString(CultureInfo.InvariantCulture)); writer.Write(',');
                        writer.WriteLine(units.ToString(CultureInfo.InvariantCulture));
                        sum += units;
                    } else {
                        writer.Write(index.ToString(CultureInfo.InvariantCulture)); writer.Write(',');
                        writer.Write((index * 3).ToString(CultureInfo.InvariantCulture)); writer.Write(',');
                        writer.WriteLine((index * 7).ToString(CultureInfo.InvariantCulture));
                        sum += index;
                    }
                }
                ExpectedSum = sum;
            }
            Bytes = new FileInfo(Path).Length;
            using var inputStream = File.OpenRead(Path);
            Fingerprint = Convert.ToHexString(SHA256.HashData(inputStream));
        } catch { Dispose(); throw; }
    }

    internal static IEnumerable<CsvParallelInput> Cases() {
        string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_CSV_PARALLEL_BENCHMARK_ROWS");
        int? rows = configured is null ? null : int.Parse(configured, CultureInfo.InvariantCulture);
        if (rows is < 1 or > 10_000_000) throw new ArgumentOutOfRangeException(nameof(configured));
        yield return new("ConversionHeavy", rows ?? 4_300_000);
        yield return new("NarrowInt", rows ?? 8_000_000);
        yield return new("ConversionHeavy3M", rows ?? 3_000_000);
    }

    internal static IEnumerable<int> Degrees() {
        string configured = Environment.GetEnvironmentVariable("OFFICEIMO_CSV_PARALLEL_BENCHMARK_DOP") ?? "0,1,2,4,8,16";
        foreach (string part in configured.Split(',', StringSplitOptions.TrimEntries)) {
            int degree = int.Parse(part, CultureInfo.InvariantCulture);
            if (degree is < 0 or > 256) throw new ArgumentOutOfRangeException(nameof(configured));
            yield return degree;
        }
    }

    internal void Validate(CsvHeavyRecord row, int index) {
        decimal price = 10m + index % 9000 / 100m;
        int units = index % 500 + 1;
        if (row.Region != Regions[index % 4] || row.Country != Countries[index % 6]
            || row.OrderDate != Start.AddDays(index % 3650) || row.UnitPrice != price
            || row.TotalRevenue != price * units || row.Units != units)
            throw new InvalidDataException($"Conversion-heavy CSV fields differ at row {index}.");
    }

    internal void Validate(CsvNarrowRecord row, int index) {
        if (row.A != index || row.B != index * 3 || row.C != index * 7)
            throw new InvalidDataException($"Narrow CSV fields differ at row {index}.");
    }

    // Parallel accumulator callbacks do not have an ordered row index. The price
    // and date residues identify the fixture's repeating row pattern exactly.
    internal int ValidateHeavyParts(ReadOnlySpan<char> region, ReadOnlySpan<char> country,
        DateTime date, decimal price, decimal revenue, int units) {
        decimal priceResidue = (price - 10m) * 100m;
        int days = (int)(date - Start).TotalDays;
        if (priceResidue < 0 || priceResidue > 8999 || priceResidue != decimal.Truncate(priceResidue)
            || days < 0 || days >= 3650 || date != Start.AddDays(days))
            throw new InvalidDataException("Aggregate date or decimal precision differs.");
        int residue = (int)priceResidue;
        int difference = (days - residue % 3650 + 3650) % 3650;
        if (difference % 50 != 0) throw new InvalidDataException("Aggregate date/price pattern differs.");
        int representative = residue + 9000 * ((difference / 50 * 58) % 73);
        if (representative >= Input.Rows || !region.SequenceEqual(Regions[representative % 4])
            || !country.SequenceEqual(Countries[representative % 6]) || units != representative % 500 + 1
            || revenue != price * units)
            throw new InvalidDataException("Aggregate CSV field values differ.");
        return representative;
    }

    internal long Check(long sum, int count) => sum == ExpectedSum && count == Input.Rows
        ? sum : throw new InvalidDataException("Parallel CSV count or aggregate differs.");

    public void Dispose() {
        if (File.Exists(Path)) File.Delete(Path);
        if (Directory.Exists(_directory)) Directory.Delete(_directory);
    }
}
