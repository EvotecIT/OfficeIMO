using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using nietras.SeparatedValues;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Compares materialized text reads when an escaped quoted field appears near the start,
/// middle or end of a record. Setup checks every value; each read checks the full observation.
/// </summary>
[MemoryDiagnoser]
public class CsvQuotedFieldPositionBenchmarks {
    private const int Columns = 14;
    private string _text = string.Empty;
    private string[] _values = [];
    private CsvReadObservation _expected;
    private bool _validateValues;

    [Params(5_000, 50_000)]
    public int RowCount { get; set; }

    [Params(',', '|')]
    public char Delimiter { get; set; }

    [Params(1, 6, 13)]
    public int QuotedColumn { get; set; }

    [GlobalSetup]
    public void Setup() {
        _values = new string[RowCount * Columns];
        var text = new StringBuilder();
        var expected = new CsvObservationAccumulator();
        for (int column = 0; column < Columns; column++) {
            if (column != 0) text.Append(Delimiter);
            text.Append('C').Append(column);
        }
        text.Append('\n');
        for (int row = 0; row < RowCount; row++) {
            for (int column = 0; column < Columns; column++) {
                if (column != 0) text.Append(Delimiter);
                string value = column == QuotedColumn ? $"name|, \"quoted\" {row}" : $"{row}-{column}";
                _values[row * Columns + column] = value;
                if (column == QuotedColumn) text.Append('"').Append(value.Replace("\"", "\"\"")).Append('"');
                else text.Append(value);
                expected.Add(value);
            }
            text.Append('\n');
        }
        _text = text.ToString();
        _expected = expected.ToObservation(RowCount);
        _validateValues = true;
        try {
            OfficeIMO();
            Sep();
            Sylvan();
            CsvHelper();
        } finally {
            _validateValues = false;
        }
    }

    [Benchmark]
    public CsvReadObservation OfficeIMO() {
        using var text = new StringReader(_text);
        using var reader = CsvDocument.OpenReader(text, new CsvReaderOptions { Delimiter = Delimiter, HasHeaderRow = true });
        if (reader.FieldCount != Columns) throw new InvalidDataException("Column count differs.");
        var observation = new CsvObservationAccumulator();
        int row = 0;
        while (reader.Read()) {
            for (int column = 0; column < Columns; column++) Visit(ref observation, reader.GetString(column), row, column);
            row++;
        }
        return Complete(observation, row);
    }

    [Benchmark]
    public CsvReadObservation Sep() {
        using var reader = nietras.SeparatedValues.Sep.New(Delimiter).Reader(options => options with { Unescape = true }).FromText(_text);
        var observation = new CsvObservationAccumulator();
        int row = 0;
        foreach (var record in reader) {
            if (record.ColCount != Columns) throw new InvalidDataException("Column count differs.");
            for (int column = 0; column < Columns; column++) Visit(ref observation, record[column].ToString(), row, column);
            row++;
        }
        return Complete(observation, row);
    }

    [Benchmark]
    public CsvReadObservation Sylvan() {
        using var text = new StringReader(_text);
        using var reader = global::Sylvan.Data.Csv.CsvDataReader.Create(text,
            new global::Sylvan.Data.Csv.CsvDataReaderOptions { Delimiter = Delimiter, HasHeaders = true });
        if (reader.FieldCount != Columns) throw new InvalidDataException("Column count differs.");
        var observation = new CsvObservationAccumulator();
        int row = 0;
        while (reader.Read()) {
            for (int column = 0; column < Columns; column++) Visit(ref observation, reader.GetString(column), row, column);
            row++;
        }
        return Complete(observation, row);
    }

    [Benchmark]
    public CsvReadObservation CsvHelper() {
        using var text = new StringReader(_text);
        using var reader = new global::CsvHelper.CsvReader(text,
            new global::CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture) {
                Delimiter = Delimiter.ToString(), HasHeaderRecord = true
            });
        if (!reader.Read() || !reader.ReadHeader() || reader.HeaderRecord?.Length != Columns)
            throw new InvalidDataException("Header differs.");
        var observation = new CsvObservationAccumulator();
        int row = 0;
        while (reader.Read()) {
            for (int column = 0; column < Columns; column++) Visit(ref observation, reader.GetField(column) ?? throw new InvalidDataException("Null field."), row, column);
            row++;
        }
        return Complete(observation, row);
    }

    private void Visit(ref CsvObservationAccumulator observation, string value, int row, int column) {
        if (_validateValues && (row >= RowCount || value != _values[row * Columns + column]))
            throw new InvalidDataException("Decoded value or order differs.");
        observation.Add(value);
    }

    private CsvReadObservation Complete(CsvObservationAccumulator observation, int rows) {
        CsvReadObservation actual = observation.ToObservation(rows);
        if (actual != _expected) throw new InvalidDataException($"Read observation differs. Expected {_expected}; actual {actual}.");
        return actual;
    }
}
