#nullable enable

namespace OfficeIMO.CSV.Benchmarks;

public partial class CsvBenchmarks
{
    private bool _validateRawRead;
    private string _rawReadMethod = string.Empty;
    private string[][] _expectedRawRows = [];
    private int _rawReadRow;
    private int _rawReadColumn;

    private void ValidateRawReadBenchmarkOutputs()
    {
        // The typed setup proves the input's semantic values. Check its raw
        // representation as well before comparing every peer against it.
        CsvBenchmarkOutputValidator.Validate(
            nameof(ValidateRawReadBenchmarkOutputs), _csvText, Headers, _rows.Length,
            expectedTextRows: null, expectedObjectRows: _projectedRows);

        var rows = new List<string[]>(_rows.Length);
        using (var input = new StringReader(_csvText))
        {
            CsvDocument.ReadRowsReusable(input, (headers, values) =>
            {
                if (!headers.SequenceEqual(Headers, StringComparer.Ordinal))
                    throw new InvalidOperationException("Raw CSV input has different headers.");
                rows.Add(values.ToArray());
            });
        }
        _expectedRawRows = rows.ToArray();
        if (_expectedRawRows.Length != _rows.Length)
            throw new InvalidOperationException("Raw CSV input has a different row count.");

        ValidateRawReadOutput(nameof(OfficeIMO_ReadDataReaderGetStrings), OfficeIMO_ReadDataReaderGetStrings);
        ValidateRawReadOutput(nameof(CsvHelper_ReadFields), CsvHelper_ReadFields);
        ValidateRawReadOutput(nameof(Sylvan_ReadFields), Sylvan_ReadFields);
        ValidateRawReadOutput(nameof(Sylvan_ReadFieldSpans), Sylvan_ReadFieldSpans);
        ValidateRawReadOutput(nameof(Dataplat_ReadFields), Dataplat_ReadFields);
        ValidateRawReadOutput(nameof(Sep_ReadFields), Sep_ReadFields);
        ValidateRawReadOutput(nameof(Sep_ReadFieldSpans), Sep_ReadFieldSpans);
    }

    private void ValidateRawReadOutput(string method, Func<int> read)
    {
        _rawReadMethod = method;
        _rawReadRow = 0;
        _rawReadColumn = 0;
        _validateRawRead = true;
        int actual;
        try { actual = read(); }
        finally { _validateRawRead = false; }
        if (_rawReadRow != _expectedRawRows.Length || _rawReadColumn != 0)
            throw new InvalidOperationException($"{method} returned {_rawReadRow} complete raw rows; expected {_expectedRawRows.Length}.");

        int expected = 0;
        foreach (string[] row in _expectedRawRows)
            foreach (string value in row)
                expected += 1 + value.Length;
        if (actual != expected)
            throw new InvalidOperationException($"{method} returned raw field count {actual}; expected {expected}.");
    }

    private void ValidateRawField(int column, ReadOnlySpan<char> actual)
    {
        if ((uint)_rawReadRow >= (uint)_expectedRawRows.Length
            || (uint)column >= (uint)Headers.Length
            || column != _rawReadColumn
            || _expectedRawRows[_rawReadRow].Length != Headers.Length
            || !actual.SequenceEqual(_expectedRawRows[_rawReadRow][column].AsSpan()))
            throw new InvalidOperationException($"{_rawReadMethod} changed raw row {_rawReadRow + 1}, field {column + 1}.");
        _rawReadColumn++;
    }

    private void CompleteRawRow()
    {
        if (_rawReadColumn != Headers.Length)
            throw new InvalidOperationException($"{_rawReadMethod} returned {_rawReadColumn} fields in raw row {_rawReadRow + 1}.");
        _rawReadRow++;
        _rawReadColumn = 0;
    }
}
