using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using BenchmarkDotNet.Attributes;
using nietras.SeparatedValues;
using OfficeIMO.Benchmarks;
using SepLib = nietras.SeparatedValues.Sep;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>Compares current reader APIs on the shared generated UTF-8 input.</summary>
[MemoryDiagnoser]
public class CsvGeneratedModernReadBenchmarks
{
    private const int RowCount = 50_000;
    private byte[] _utf8 = [];
    private string?[][] _expectedRows = [];
    private int _expectedChecksum;
    private bool _validate;
    private static readonly SepReaderOptions SepOptions =
        SepLib.New(',').Reader(options => options with { Unescape = true });

    private static readonly SepReaderOptions SepPooledOptions =
        SepLib.New(',').Reader(options => options with { Unescape = true, CreateToString = SepToString.PoolPerCol() });

    [GlobalSetup]
    public void Setup()
    {
        BenchmarkProcessPriority.ApplyConfiguredPriority();
        CsvBenchmarkRow[] rows = CsvBenchmarkData.Create(RowCount, CsvBenchmarkShape.Mixed);
        _utf8 = Encoding.UTF8.GetBytes(CsvBenchmarkData.CreateText(rows));
        _expectedRows = rows.Select(CsvBenchmarkData.ProjectRow)
            .Select(CsvBenchmarkData.ProjectTextRow).ToArray();
        _expectedChecksum = _expectedRows.Sum(row => row.Sum(value => 1 + (value?.Length ?? 0)));
        _validate = true;
        try
        {
            OfficeIMO();
            OfficeIMO_InternStrings();
            CsvHelper();
            CsvHelper_PoolStrings();
            Sep();
            Sep_PoolPerCol();
            FourLambda();
        }
        finally
        {
            _validate = false;
        }
        Console.WriteLine($"Generated CSV input: {RowCount} rows, {CsvBenchmarkData.Headers.Length} columns, "
            + $"UTF-8 SHA256 {Convert.ToHexString(SHA256.HashData(_utf8))}.");
    }

    [Benchmark]
    public int OfficeIMO()
    {
        using var input = new MemoryStream(_utf8, writable: false);
        using var reader = CsvDocument.OpenReader(input);
        if (_validate) ValidateHeaders(Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
        int checksum = 0, rows = 0;
        while (reader.Read())
        {
            if (_validate) ValidateWidth(reader.FieldCount);
            for (int column = 0; column < reader.FieldCount; column++)
            {
                string value = reader.GetString(column);
                checksum += 1 + value.Length;
                if (_validate) ValidateField(rows, column, value);
            }
            rows++;
        }
        return CompleteRead(rows, checksum);
    }

    [Benchmark]
    public int CsvHelper()
    {
        using var input = new MemoryStream(_utf8, writable: false);
        using var reader = global::CsvHelper.CsvReader.Create(input,
            new global::CsvHelper.CsvReaderOptions { FormatProvider = CultureInfo.InvariantCulture });
        if (_validate) ValidateHeaders(reader.Headers.ToArray());
        int checksum = 0, rows = 0;
        while (reader.Read())
        {
            if (_validate) ValidateWidth(reader.FieldCount);
            for (int column = 0; column < reader.FieldCount; column++)
            {
                string value = reader.GetString(column);
                checksum += 1 + value.Length;
                if (_validate) ValidateField(rows, column, value);
            }
            rows++;
        }
        return CompleteRead(rows, checksum);
    }

    [Benchmark]
    public int Sep()
    {
        using var input = new MemoryStream(_utf8, writable: false);
        using var reader = SepOptions.From(input);
        if (_validate) ValidateHeaders(reader.Header.ColNames);
        int checksum = 0, rows = 0;
        foreach (var row in reader)
        {
            if (_validate) ValidateWidth(row.ColCount);
            for (int column = 0; column < row.ColCount; column++)
            {
                string value = row[column].ToString();
                checksum += 1 + value.Length;
                if (_validate) ValidateField(rows, column, value);
            }
            rows++;
        }
        return CompleteRead(rows, checksum);
    }

    [Benchmark]
    public int FourLambda()
    {
        using var input = new MemoryStream(_utf8, writable: false);
        using var reader = global::FourLambda.Csv.CsvReader.Create(input, hasHeaders: true);
        if (_validate) ValidateHeaders(reader.Headers!.OrderBy(pair => pair.Value).Select(pair => pair.Key));
        int checksum = 0, rows = 0;
        while (reader.ReadNext())
        {
            if (_validate) ValidateWidth(reader.FieldCount);
            for (int column = 0; column < reader.FieldCount; column++)
            {
                string value = reader.GetString(column);
                checksum += 1 + value.Length;
                if (_validate) ValidateField(rows, column, value);
            }
            rows++;
        }
        return CompleteRead(rows, checksum);
    }

    [Benchmark]
    public int OfficeIMO_InternStrings()
    {
        using var input = new MemoryStream(_utf8, writable: false);
        using var reader = CsvDocument.OpenReader(input, options: new CsvReaderOptions { InternStrings = true });
        if (_validate) ValidateHeaders(Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
        int checksum = 0, rows = 0;
        while (reader.Read())
        {
            if (_validate) ValidateWidth(reader.FieldCount);
            for (int column = 0; column < reader.FieldCount; column++)
            {
                string value = reader.GetString(column);
                checksum += 1 + value.Length;
                if (_validate) ValidateField(rows, column, value);
            }
            rows++;
        }
        return CompleteRead(rows, checksum);
    }

    [Benchmark]
    public int CsvHelper_PoolStrings()
    {
        using var input = new MemoryStream(_utf8, writable: false);
        using var reader = global::CsvHelper.CsvReader.Create(input,
            new global::CsvHelper.CsvReaderOptions { FormatProvider = CultureInfo.InvariantCulture, PoolStrings = true });
        if (_validate) ValidateHeaders(reader.Headers.ToArray());
        int checksum = 0, rows = 0;
        while (reader.Read())
        {
            if (_validate) ValidateWidth(reader.FieldCount);
            for (int column = 0; column < reader.FieldCount; column++)
            {
                string value = reader.GetString(column);
                checksum += 1 + value.Length;
                if (_validate) ValidateField(rows, column, value);
            }
            rows++;
        }
        return CompleteRead(rows, checksum);
    }

    [Benchmark]
    public int Sep_PoolPerCol()
    {
        using var input = new MemoryStream(_utf8, writable: false);
        using var reader = SepPooledOptions.From(input);
        if (_validate) ValidateHeaders(reader.Header.ColNames);
        int checksum = 0, rows = 0;
        foreach (var row in reader)
        {
            if (_validate) ValidateWidth(row.ColCount);
            for (int column = 0; column < row.ColCount; column++)
            {
                string value = row[column].ToString();
                checksum += 1 + value.Length;
                if (_validate) ValidateField(rows, column, value);
            }
            rows++;
        }
        return CompleteRead(rows, checksum);
    }

    private static void ValidateHeaders(IEnumerable<string> actual)
    {
        if (!actual.SequenceEqual(CsvBenchmarkData.Headers, StringComparer.Ordinal))
            throw new InvalidDataException("The generated CSV headers changed.");
    }

    private static void ValidateWidth(int actual)
    {
        if (actual != CsvBenchmarkData.Headers.Length)
            throw new InvalidDataException($"The generated CSV row contains {actual} fields.");
    }

    private void ValidateField(int row, int column, string actual)
    {
        if ((uint)row >= (uint)_expectedRows.Length
            || !string.Equals(actual, _expectedRows[row][column], StringComparison.Ordinal))
            throw new InvalidDataException($"The generated CSV changed row {row + 1}, field {column + 1}.");
    }

    private int CompleteRead(int rows, int checksum) => rows == RowCount && checksum == _expectedChecksum
        ? checksum : throw new InvalidDataException($"Incomplete generated CSV read: {rows} rows, checksum {checksum}.");
}
