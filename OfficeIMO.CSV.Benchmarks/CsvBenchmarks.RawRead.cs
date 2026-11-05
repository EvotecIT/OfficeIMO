using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using nietras.SeparatedValues;
using CsvHelperReader = CsvHelper.CsvReader;
using SylvanCsvDataReader = Sylvan.Data.Csv.CsvDataReader;
using DataplatCsvDataReader = Dataplat.Dbatools.Csv.Reader.CsvDataReader;

namespace OfficeIMO.CSV.Benchmarks;

public partial class CsvBenchmarks {
    /// <summary>Decodes the same prepared UTF-8 bytes inside each text-reader operation.</summary>
    private StreamReader OpenUtf8TextReader() => new(
        new MemoryStream(_csvUtf8, writable: false), Encoding.UTF8,
        detectEncodingFromByteOrderMarks: false);

    /// <summary>Checks the generated cohort's UTF-8 entry points against every expected field and typed row.</summary>
    internal void ValidateGeneratedUtf8Reads() {
        Console.WriteLine($"Generated CSV input: {_rows.Length} rows, {Headers.Length} columns, "
            + $"UTF-8 SHA256 {Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(_csvUtf8))}.");
        ValidateRawReadOutput(nameof(CsvHelper_ReadUtf8Fields), CsvHelper_ReadUtf8Fields);
        ValidateRawReadOutput(nameof(Sylvan_ReadUtf8Fields), Sylvan_ReadUtf8Fields);
        ValidateRawReadOutput(nameof(Dataplat_ReadUtf8Fields), Dataplat_ReadUtf8Fields);
        ValidateRawReadOutput(nameof(Sep_ReadUtf8Fields), Sep_ReadUtf8Fields);
        ValidateTypedReadOutput(nameof(OfficeIMO_ReadTypedRowsStream), OfficeIMO_ReadTypedRowsStream);
        ValidateTypedReadOutput(nameof(CsvHelper_ReadTypedRecordsUtf8), CsvHelper_ReadTypedRecordsUtf8);
    }

    [Benchmark]
    public int CsvHelper_ReadFields()
    {
        using var reader = new StringReader(_csvText);
        using var csv = new CsvHelperReader(reader, CultureInfo.InvariantCulture);
        return ReadCsvHelperFields(csv);
    }

    internal int CsvHelper_ReadUtf8Fields() {
        using var reader = OpenUtf8TextReader();
        using var csv = new CsvHelperReader(reader, CultureInfo.InvariantCulture);
        return ReadCsvHelperFields(csv);
    }

    private int ReadCsvHelperFields(CsvHelperReader csv) {
        var fieldCount = 0;
        if (!csv.Read())
        {
            return fieldCount;
        }

        csv.ReadHeader();
        while (csv.Read())
        {
            for (var i = 0; i < Headers.Length; i++)
            {
                string? value = csv.GetField(i);
                fieldCount += 1 + (value?.Length ?? 0);
                if (_validateRawRead) ValidateRawField(i, value.AsSpan());
            }
            if (_validateRawRead) CompleteRawRow();
        }

        return fieldCount;
    }

    [Benchmark]
    public int Sylvan_ReadFields()
    {
        using var reader = new StringReader(_csvText);
        using var csv = SylvanCsvDataReader.Create(reader);
        return ReadSylvanFields(csv);
    }

    internal int Sylvan_ReadUtf8Fields() {
        using var reader = OpenUtf8TextReader();
        using var csv = SylvanCsvDataReader.Create(reader);
        return ReadSylvanFields(csv);
    }

    private int ReadSylvanFields(SylvanCsvDataReader csv) {
        var fieldCount = 0;
        while (csv.Read())
        {
            for (var i = 0; i < csv.FieldCount; i++)
            {
                string value = csv.GetString(i);
                fieldCount += 1 + value.Length;
                if (_validateRawRead) ValidateRawField(i, value);
            }
            if (_validateRawRead) CompleteRawRow();
        }

        return fieldCount;
    }

    [Benchmark]
    public int Dataplat_ReadFields()
    {
        using var reader = new StringReader(_csvText);
        using var csv = new DataplatCsvDataReader(reader, DataplatReaderOptions);
        return ReadDataplatFields(csv);
    }

    internal int Dataplat_ReadUtf8Fields() {
        using var reader = OpenUtf8TextReader();
        using var csv = new DataplatCsvDataReader(reader, DataplatReaderOptions);
        return ReadDataplatFields(csv);
    }

    private int ReadDataplatFields(DataplatCsvDataReader csv) {
        var fieldCount = 0;
        while (csv.Read())
        {
            for (var i = 0; i < csv.FieldCount; i++)
            {
                string value = csv.GetString(i);
                fieldCount += 1 + value.Length;
                if (_validateRawRead) ValidateRawField(i, value);
            }
            if (_validateRawRead) CompleteRawRow();
        }

        return fieldCount;
    }

    [Benchmark]
    public int Sep_ReadFields()
    {
        var options = SepReadOptions;
        using var csv = options.FromText(_csvText);
        return ReadSepFields(csv);
    }

    internal int Sep_ReadUtf8Fields() {
        using var stream = new MemoryStream(_csvUtf8, writable: false);
        using var csv = SepReadOptions.From(stream);
        return ReadSepFields(csv);
    }

    private int ReadSepFields(SepReader csv) {
        var fieldCount = 0;
        foreach (var row in csv)
        {
            for (var i = 0; i < row.ColCount; i++)
            {
                string value = row[i].ToString();
                fieldCount += 1 + value.Length;
                if (_validateRawRead) ValidateRawField(i, value);
            }
            if (_validateRawRead) CompleteRawRow();
        }

        return fieldCount;
    }
}
