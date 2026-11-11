using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Benchmarks;

/// <summary>One 65K sales row mapped into a mutable class.</summary>
public sealed class SalesRow {
    public string Region { get; set; } = string.Empty;
    public string Country { get; set; } = string.Empty;
    public string ItemType { get; set; } = string.Empty;
    public string SalesChannel { get; set; } = string.Empty;
    public string OrderPriority { get; set; } = string.Empty;
    public DateTime OrderDate { get; set; }
    public int OrderId { get; set; }
    public DateTime ShipDate { get; set; }
    public int UnitsSold { get; set; }
    public decimal UnitPrice { get; set; }
    public decimal UnitCost { get; set; }
    public decimal TotalRevenue { get; set; }
    public decimal TotalCost { get; set; }
    public decimal TotalProfit { get; set; }
}

/// <summary>One 65K sales row mapped through a positional record constructor.</summary>
public sealed record SalesRecord(
    string Region,
    string Country,
    string ItemType,
    string SalesChannel,
    string OrderPriority,
    DateTime OrderDate,
    int OrderId,
    DateTime ShipDate,
    int UnitsSold,
    decimal UnitPrice,
    decimal UnitCost,
    decimal TotalRevenue,
    decimal TotalCost,
    decimal TotalProfit);

/// <summary>
/// Supplies the hash-pinned 65K sales data with its header names written as property names,
/// so every library maps columns by name without library-specific attributes. Only the header
/// text changes: CSV rewrites its first line and XLSX rewrites the header entries of its shared
/// string table, keeping the worksheet part byte for byte.
/// </summary>
internal static class SalesMappingFixture {
    internal const int RowCount = 65_535;

    internal static readonly string[] Headers = [
        "Region", "Country", "ItemType", "SalesChannel", "OrderPriority", "OrderDate", "OrderId",
        "ShipDate", "UnitsSold", "UnitPrice", "UnitCost", "TotalRevenue", "TotalCost", "TotalProfit"
    ];

    private static readonly string[] OriginalHeaders = [
        "Region", "Country", "Item Type", "Sales Channel", "Order Priority", "Order Date", "Order ID",
        "Ship Date", "Units Sold", "Unit Price", "Unit Cost", "Total Revenue", "Total Cost", "Total Profit"
    ];

    /// <summary>Returns the CSV fixture with property-name headers.</summary>
    internal static byte[] CreateCsv() {
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.CsvFileName);
        byte[] source = File.ReadAllBytes(MarkPflug65KFixture.CsvPath);
        int lineEnd = Array.IndexOf(source, (byte)'\n');
        int bom = source.AsSpan().StartsWith(Encoding.UTF8.Preamble) ? Encoding.UTF8.Preamble.Length : 0;
        string header = Encoding.ASCII.GetString(source, bom, lineEnd - bom).TrimEnd('\r');
        if (header != string.Join(",", OriginalHeaders)) throw new InvalidDataException("Unexpected 65K CSV header.");
        // The byte order mark stays, so every reader still sees the original preamble.
        byte[] replacement = Encoding.UTF8.GetBytes((bom > 0 ? "﻿" : string.Empty) + string.Join(",", Headers));
        var output = new byte[replacement.Length + source.Length - lineEnd];
        replacement.CopyTo(output, 0);
        Array.Copy(source, lineEnd, output, replacement.Length, source.Length - lineEnd);
        return output;
    }

    /// <summary>Writes the XLSX fixture with property-name headers and returns its path.</summary>
    internal static string CreateXlsx(string directory) {
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.XlsxFileName);
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "65K_Records_Data.mapping.xlsx");
        File.Copy(MarkPflug65KFixture.XlsxPath, path, overwrite: true);
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry("xl/sharedStrings.xml") ?? throw new InvalidDataException("Missing shared strings.");
        string xml;
        using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
        for (int index = 0; index < Headers.Length; index++) {
            if (OriginalHeaders[index] == Headers[index]) continue;
            string original = "<t>" + OriginalHeaders[index] + "</t>";
            int at = xml.IndexOf(original, StringComparison.Ordinal);
            if (at < 0 || xml.IndexOf(original, at + 1, StringComparison.Ordinal) >= 0)
                throw new InvalidDataException($"Header '{OriginalHeaders[index]}' is not a single shared string.");
            xml = xml.Replace(original, "<t>" + Headers[index] + "</t>");
        }

        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry("xl/sharedStrings.xml", CompressionLevel.Optimal);
        using (var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false))) writer.Write(xml);
        return path;
    }

    /// <summary>Computes the expected observation by splitting the CSV text directly.</summary>
    internal static long ExpectedObservation(byte[] csv) {
        string[] lines = Encoding.UTF8.GetString(csv).Split('\n');
        var observation = new SalesObservation();
        for (int line = 1; line < lines.Length; line++) {
            if (lines[line].Length == 0) continue;
            string[] f = lines[line].TrimEnd('\r').Split(',');
            observation.Add(f[0], f[1], f[2], f[3], f[4],
                DateTime.ParseExact(f[5], "yyyy-MM-dd", CultureInfo.InvariantCulture),
                int.Parse(f[6], CultureInfo.InvariantCulture),
                DateTime.ParseExact(f[7], "yyyy-MM-dd", CultureInfo.InvariantCulture),
                int.Parse(f[8], CultureInfo.InvariantCulture),
                decimal.Parse(f[9], CultureInfo.InvariantCulture), decimal.Parse(f[10], CultureInfo.InvariantCulture),
                decimal.Parse(f[11], CultureInfo.InvariantCulture), decimal.Parse(f[12], CultureInfo.InvariantCulture),
                decimal.Parse(f[13], CultureInfo.InvariantCulture));
        }

        if (observation.Count != RowCount) throw new InvalidDataException("Unexpected 65K CSV row count.");
        return observation.Value;
    }
}

/// <summary>Order-sensitive digest of every mapped value; decimals compare to the cent.</summary>
internal struct SalesObservation {
    internal long Value;
    internal int Count;

    internal void Add(SalesRow row) => Add(row.Region, row.Country, row.ItemType, row.SalesChannel, row.OrderPriority,
        row.OrderDate, row.OrderId, row.ShipDate, row.UnitsSold, row.UnitPrice, row.UnitCost, row.TotalRevenue, row.TotalCost, row.TotalProfit);

    internal void Add(SalesRecord row) => Add(row.Region, row.Country, row.ItemType, row.SalesChannel, row.OrderPriority,
        row.OrderDate, row.OrderId, row.ShipDate, row.UnitsSold, row.UnitPrice, row.UnitCost, row.TotalRevenue, row.TotalCost, row.TotalProfit);

    internal void Add(string region, string country, string itemType, string salesChannel, string orderPriority,
        DateTime orderDate, int orderId, DateTime shipDate, int unitsSold,
        decimal unitPrice, decimal unitCost, decimal totalRevenue, decimal totalCost, decimal totalProfit) {
        unchecked {
            long value = Value;
            value = Mix(value, Text(region));
            value = Mix(value, Text(country));
            value = Mix(value, Text(itemType));
            value = Mix(value, Text(salesChannel));
            value = Mix(value, Text(orderPriority));
            value = Mix(value, orderDate.Ticks);
            value = Mix(value, orderId);
            value = Mix(value, shipDate.Ticks);
            value = Mix(value, unitsSold);
            value = Mix(value, Cents(unitPrice));
            value = Mix(value, Cents(unitCost));
            value = Mix(value, Cents(totalRevenue));
            value = Mix(value, Cents(totalCost));
            Value = Mix(value, Cents(totalProfit));
            Count++;
        }
    }

    internal readonly long Check(long expected, string lane) =>
        Count == SalesMappingFixture.RowCount && Value == expected
            ? Value
            : throw new InvalidDataException($"{lane} mapped {Count} rows with a different observation.");

    private static long Mix(long value, long item) => unchecked(value * 1_000_003 + item);

    private static long Cents(decimal value) => (long)decimal.Round(value * 100m, MidpointRounding.AwayFromZero);

    // The XLSX fixture keeps trailing spaces on some country names that its CSV twin drops,
    // so text compares without them; every other character must match.
    private static long Text(string? value) {
        if (value is null) return -1;
        unchecked {
            long hash = 1469598103934665603;
            int length = value.Length;
            while (length > 0 && value[length - 1] == ' ') length--;
            for (int index = 0; index < length; index++) hash = (hash ^ value[index]) * 1099511628211;
            return hash;
        }
    }
}
