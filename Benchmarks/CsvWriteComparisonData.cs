using System.Data;
using System.Globalization;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>Value shapes of the CSV write comparison.</summary>
public enum CsvWriteShape {
    /// <summary>Six text columns where most values need quoting: delimiters, quotes and line breaks.</summary>
    QuotedText,
    /// <summary>Integers, longs, doubles and decimals.</summary>
    Numbers,
    /// <summary>Date and time values with one explicit format, plus an id and a label.</summary>
    Dates,
    /// <summary>31 columns of short text, integers and doubles.</summary>
    Wide
}

/// <summary>Receives a record's fields in order with their own types, so writers that take typed fields avoid boxing.</summary>
public interface ICsvWriteFieldSink {
    void Int(int value);
    void Long(long value);
    void Double(double value);
    void Decimal(decimal value);
    void Date(DateTime value);
    void Text(string value);
}

internal enum CsvWriteFieldKind : byte { Int, Long, Double, Decimal, Date, Text }

/// <summary>A typed source record that writes its fields to a sink and exposes them as values.</summary>
internal interface ICsvWriteRecord {
    void Write<TSink>(ref TSink sink) where TSink : ICsvWriteFieldSink
#if NET9_0_OR_GREATER
        , allows ref struct
#endif
        ;
    object[] Values();
}

/// <summary>Deterministic source data of one shape in the three forms the writers accept.</summary>
internal sealed class CsvWriteSource {
    internal const string DateFormat = "yyyy-MM-ddTHH:mm:ss";

    private CsvWriteSource(string[] headers, CsvWriteFieldKind[] kinds, ICsvWriteRecord[] records) {
        Headers = headers;
        Kinds = kinds;
        Records = records;
        Types = kinds.Select(kind => kind switch {
            CsvWriteFieldKind.Int => typeof(int),
            CsvWriteFieldKind.Long => typeof(long),
            CsvWriteFieldKind.Double => typeof(double),
            CsvWriteFieldKind.Decimal => typeof(decimal),
            CsvWriteFieldKind.Date => typeof(DateTime),
            _ => typeof(string)
        }).ToArray();
        Rows = records.Select(record => (object?[])record.Values()).ToArray();
        Table = new DataTable("Data") { Locale = CultureInfo.InvariantCulture };
        for (int column = 0; column < headers.Length; column++) Table.Columns.Add(headers[column], Types[column]);
        Table.BeginLoadData();
        foreach (object?[] row in Rows) Table.Rows.Add(row);
        Table.EndLoadData();
    }

    internal string[] Headers { get; }
    internal CsvWriteFieldKind[] Kinds { get; }
    internal Type[] Types { get; }
    internal ICsvWriteRecord[] Records { get; }
    internal object?[][] Rows { get; }
    internal DataTable Table { get; }

    internal static CsvWriteSource Create(CsvWriteShape shape) {
        var random = new Random(20261111 + (int)shape);
        return shape switch {
            CsvWriteShape.QuotedText => Build(QuotedTextRecord.Headers, QuotedTextRecord.Kinds, 50_000, index => QuotedTextRecord.Create(index, random)),
            CsvWriteShape.Numbers => Build(NumberRecord.Headers, NumberRecord.Kinds, 50_000, index => NumberRecord.Create(index, random)),
            CsvWriteShape.Dates => Build(DateRecord.Headers, DateRecord.Kinds, 50_000, index => DateRecord.Create(index, random)),
            CsvWriteShape.Wide => Build(WideRecord.Headers, WideRecord.Kinds, 10_000, index => WideRecord.Create(index, random)),
            _ => throw new ArgumentOutOfRangeException(nameof(shape))
        };
    }

    private static CsvWriteSource Build(string[] headers, CsvWriteFieldKind[] kinds, int count, Func<int, ICsvWriteRecord> create) {
        var records = new ICsvWriteRecord[count];
        for (int index = 0; index < count; index++) records[index] = create(index);
        return new CsvWriteSource(headers, kinds, records);
    }

    internal static string Word(Random random, int minimum, int maximum) {
        const string letters = "abcdefghijklmnopqrstuvwxyz";
        var chars = new char[random.Next(minimum, maximum + 1)];
        for (int index = 0; index < chars.Length; index++) chars[index] = letters[random.Next(letters.Length)];
        return new string(chars);
    }
}

/// <summary>Text where nearly every value carries a delimiter, a quote or a line break.</summary>
public sealed class QuotedTextRecord : ICsvWriteRecord {
    internal static readonly string[] Headers = [nameof(Code), nameof(Name), nameof(Address), nameof(Quote), nameof(Notes), nameof(Tags)];
    internal static readonly CsvWriteFieldKind[] Kinds = [CsvWriteFieldKind.Text, CsvWriteFieldKind.Text, CsvWriteFieldKind.Text, CsvWriteFieldKind.Text, CsvWriteFieldKind.Text, CsvWriteFieldKind.Text];

    public string Code { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string Quote { get; set; } = string.Empty;
    public string Notes { get; set; } = string.Empty;
    public string Tags { get; set; } = string.Empty;

    internal static QuotedTextRecord Create(int index, Random random) => new() {
        Code = "C" + index.ToString(CultureInfo.InvariantCulture),
        Name = CsvWriteSource.Word(random, 4, 12) + ", " + CsvWriteSource.Word(random, 3, 10),
        Address = random.Next(1, 999).ToString(CultureInfo.InvariantCulture) + " " + CsvWriteSource.Word(random, 5, 12) + " St, Suite " + random.Next(1, 99).ToString(CultureInfo.InvariantCulture),
        Quote = "He said \"" + CsvWriteSource.Word(random, 3, 9) + "\" and \"" + CsvWriteSource.Word(random, 3, 9) + "\"",
        Notes = CsvWriteSource.Word(random, 8, 20) + "\n" + CsvWriteSource.Word(random, 8, 20) + (index % 3 == 0 ? "\r\n" + CsvWriteSource.Word(random, 4, 10) : string.Empty),
        Tags = index % 4 == 0 ? CsvWriteSource.Word(random, 5, 10) : CsvWriteSource.Word(random, 3, 6) + ";" + CsvWriteSource.Word(random, 3, 6) + "," + CsvWriteSource.Word(random, 3, 6)
    };

    public void Write<TSink>(ref TSink sink) where TSink : ICsvWriteFieldSink
#if NET9_0_OR_GREATER
        , allows ref struct
#endif
    {
        sink.Text(Code); sink.Text(Name); sink.Text(Address); sink.Text(Quote); sink.Text(Notes); sink.Text(Tags);
    }

    public object[] Values() => [Code, Name, Address, Quote, Notes, Tags];
}

/// <summary>Integers, longs, doubles and two-place decimals.</summary>
public sealed class NumberRecord : ICsvWriteRecord {
    internal static readonly string[] Headers = [nameof(Id), nameof(Count), nameof(Price), nameof(Ratio), nameof(Amount), nameof(Score), nameof(Units), nameof(Total)];
    internal static readonly CsvWriteFieldKind[] Kinds = [CsvWriteFieldKind.Int, CsvWriteFieldKind.Long, CsvWriteFieldKind.Decimal, CsvWriteFieldKind.Double, CsvWriteFieldKind.Decimal, CsvWriteFieldKind.Double, CsvWriteFieldKind.Int, CsvWriteFieldKind.Long];

    public int Id { get; set; }
    public long Count { get; set; }
    public decimal Price { get; set; }
    public double Ratio { get; set; }
    public decimal Amount { get; set; }
    public double Score { get; set; }
    public int Units { get; set; }
    public long Total { get; set; }

    internal static NumberRecord Create(int index, Random random) => new() {
        Id = index,
        Count = random.NextInt64(0, 10_000_000_000L),
        Price = random.Next(0, 10_000_000) / 100m,
        Ratio = random.NextDouble(),
        Amount = -50_000m + random.Next(0, 10_000_000) / 100m,
        Score = Math.Round(random.NextDouble() * 1000, 3),
        Units = random.Next(-1000, 100_000),
        Total = -random.NextInt64(0, 1_000_000_000_000L)
    };

    public void Write<TSink>(ref TSink sink) where TSink : ICsvWriteFieldSink
#if NET9_0_OR_GREATER
        , allows ref struct
#endif
    {
        sink.Int(Id); sink.Long(Count); sink.Decimal(Price); sink.Double(Ratio); sink.Decimal(Amount); sink.Double(Score); sink.Int(Units); sink.Long(Total);
    }

    public object[] Values() => [Id, Count, Price, Ratio, Amount, Score, Units, Total];
}

/// <summary>Date and time values written with one explicit format.</summary>
public sealed class DateRecord : ICsvWriteRecord {
    internal static readonly string[] Headers = [nameof(Id), nameof(Created), nameof(Modified), nameof(Due), nameof(Label)];
    internal static readonly CsvWriteFieldKind[] Kinds = [CsvWriteFieldKind.Int, CsvWriteFieldKind.Date, CsvWriteFieldKind.Date, CsvWriteFieldKind.Date, CsvWriteFieldKind.Text];
    private static readonly DateTime Origin = new(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

    public int Id { get; set; }
    public DateTime Created { get; set; }
    public DateTime Modified { get; set; }
    public DateTime Due { get; set; }
    public string Label { get; set; } = string.Empty;

    internal static DateRecord Create(int index, Random random) {
        DateTime created = Origin.AddSeconds(random.Next(0, 200_000_000));
        return new() {
            Id = index,
            Created = created,
            Modified = created.AddSeconds(random.Next(0, 5_000_000)),
            Due = created.Date.AddDays(random.Next(1, 400)),
            Label = CsvWriteSource.Word(random, 4, 10)
        };
    }

    public void Write<TSink>(ref TSink sink) where TSink : ICsvWriteFieldSink
#if NET9_0_OR_GREATER
        , allows ref struct
#endif
    {
        sink.Int(Id); sink.Date(Created); sink.Date(Modified); sink.Date(Due); sink.Text(Label);
    }

    public object[] Values() => [Id, Created, Modified, Due, Label];
}

/// <summary>An id followed by ten groups of a short label, an integer and a double.</summary>
public sealed class WideRecord : ICsvWriteRecord {
    internal static readonly string[] Headers = ["Id",
        "S0", "N0", "D0", "S1", "N1", "D1", "S2", "N2", "D2", "S3", "N3", "D3", "S4", "N4", "D4",
        "S5", "N5", "D5", "S6", "N6", "D6", "S7", "N7", "D7", "S8", "N8", "D8", "S9", "N9", "D9"];
    internal static readonly CsvWriteFieldKind[] Kinds = [CsvWriteFieldKind.Int, .. Enumerable.Repeat(new[] { CsvWriteFieldKind.Text, CsvWriteFieldKind.Int, CsvWriteFieldKind.Double }, 10).SelectMany(group => group)];

    public int Id { get; set; }
    public string S0 { get; set; } = string.Empty; public int N0 { get; set; } public double D0 { get; set; }
    public string S1 { get; set; } = string.Empty; public int N1 { get; set; } public double D1 { get; set; }
    public string S2 { get; set; } = string.Empty; public int N2 { get; set; } public double D2 { get; set; }
    public string S3 { get; set; } = string.Empty; public int N3 { get; set; } public double D3 { get; set; }
    public string S4 { get; set; } = string.Empty; public int N4 { get; set; } public double D4 { get; set; }
    public string S5 { get; set; } = string.Empty; public int N5 { get; set; } public double D5 { get; set; }
    public string S6 { get; set; } = string.Empty; public int N6 { get; set; } public double D6 { get; set; }
    public string S7 { get; set; } = string.Empty; public int N7 { get; set; } public double D7 { get; set; }
    public string S8 { get; set; } = string.Empty; public int N8 { get; set; } public double D8 { get; set; }
    public string S9 { get; set; } = string.Empty; public int N9 { get; set; } public double D9 { get; set; }

    internal static WideRecord Create(int index, Random random) {
        string S() => CsvWriteSource.Word(random, 2, 9);
        int N() => random.Next(-100_000, 1_000_000);
        double D() => Math.Round(random.NextDouble() * 10_000, 4);
        return new() {
            Id = index,
            S0 = S(), N0 = N(), D0 = D(), S1 = S(), N1 = N(), D1 = D(), S2 = S(), N2 = N(), D2 = D(),
            S3 = S(), N3 = N(), D3 = D(), S4 = S(), N4 = N(), D4 = D(), S5 = S(), N5 = N(), D5 = D(),
            S6 = S(), N6 = N(), D6 = D(), S7 = S(), N7 = N(), D7 = D(), S8 = S(), N8 = N(), D8 = D(),
            S9 = S(), N9 = N(), D9 = D()
        };
    }

    public void Write<TSink>(ref TSink sink) where TSink : ICsvWriteFieldSink
#if NET9_0_OR_GREATER
        , allows ref struct
#endif
    {
        sink.Int(Id);
        sink.Text(S0); sink.Int(N0); sink.Double(D0); sink.Text(S1); sink.Int(N1); sink.Double(D1);
        sink.Text(S2); sink.Int(N2); sink.Double(D2); sink.Text(S3); sink.Int(N3); sink.Double(D3);
        sink.Text(S4); sink.Int(N4); sink.Double(D4); sink.Text(S5); sink.Int(N5); sink.Double(D5);
        sink.Text(S6); sink.Int(N6); sink.Double(D6); sink.Text(S7); sink.Int(N7); sink.Double(D7);
        sink.Text(S8); sink.Int(N8); sink.Double(D8); sink.Text(S9); sink.Int(N9); sink.Double(D9);
    }

    public object[] Values() => [Id, S0, N0, D0, S1, N1, D1, S2, N2, D2, S3, N3, D3, S4, N4, D4,
        S5, N5, D5, S6, N6, D6, S7, N7, D7, S8, N8, D8, S9, N9, D9];
}
