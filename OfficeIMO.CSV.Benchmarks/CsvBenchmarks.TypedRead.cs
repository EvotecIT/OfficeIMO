#nullable enable

using System.Data.Common;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using OfficeIMO.Data;
using CsvHelperReader = CsvHelper.CsvReader;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.CSV.Benchmarks;

public partial class CsvBenchmarks
{
    private int _expectedTypedReadChecksum;
#if NET10_0_OR_GREATER
    private readonly ExcelParser<CsvBenchmarkRow> _excelReaderParser = ExcelParser.FromAttributes<CsvBenchmarkRow>();
#else
    private readonly ExcelParser<CsvBenchmarkRow> _excelReaderParser = new();
#endif
    private bool _validateTypedRead;
    private int _typedReadValidationIndex;

    private void ValidateTypedReadBenchmarkOutputs()
    {
        ValidateTypedReadOutput(nameof(OfficeIMO_ReadTypedRowsForwardOnly), OfficeIMO_ReadTypedRowsForwardOnly);
        ValidateTypedReadOutput(nameof(OfficeIMO_ReadTypedRowsMaterialized), OfficeIMO_ReadTypedRowsMaterialized);
        ValidateTypedReadOutput(nameof(CsvHelper_ReadTypedRecords), CsvHelper_ReadTypedRecords);
        ValidateTypedReadOutput(nameof(ExcelReaderNet_ReadTypedRecords), ExcelReaderNet_ReadTypedRecords);
    }

    private void ValidateTypedReadOutput(string method, Func<int> read)
    {
        _typedReadValidationIndex = 0;
        _validateTypedRead = true;
        int actual;
        try
        {
            actual = read();
        }
        finally
        {
            _validateTypedRead = false;
        }
        if (_typedReadValidationIndex != _rows.Length)
        {
            throw new InvalidOperationException($"{method} returned {_typedReadValidationIndex} typed rows instead of {_rows.Length}.");
        }
        if (actual != _expectedTypedReadChecksum)
        {
            throw new InvalidOperationException(
                $"{method} returned typed-row checksum {actual} instead of {_expectedTypedReadChecksum}.");
        }
    }

    [Benchmark]
    public int OfficeIMO_ReadTypedRowsForwardOnly()
    {
        using DbDataReader reader = CsvDocument.OpenTextDataReader(_csvText);
        var checksum = 17;
        foreach (CsvBenchmarkRow row in reader.RowsAs<CsvBenchmarkRow>())
        {
            checksum = AddTypedRowChecksum(checksum, row);
        }

        return checksum;
    }

    [Benchmark]
    public int OfficeIMO_ReadTypedRowsMaterialized()
    {
        CsvDocument document = CsvDocument.Parse(_csvText);
        var checksum = 17;
        foreach (CsvBenchmarkRow row in document.RowsAs<CsvBenchmarkRow>())
        {
            checksum = AddTypedRowChecksum(checksum, row);
        }

        return checksum;
    }

    [Benchmark]
    public int CsvHelper_ReadTypedRecords()
    {
        using var reader = new StringReader(_csvText);
        using var csv = new CsvHelperReader(reader, CultureInfo.InvariantCulture);
        var checksum = 17;
        foreach (CsvBenchmarkRow row in csv.GetRecords<CsvBenchmarkRow>())
        {
            checksum = AddTypedRowChecksum(checksum, row);
        }

        return checksum;
    }

    [Benchmark]
    public int ExcelReaderNet_ReadTypedRecords()
    {
        using ExcelReaderNetCsvReader reader = ExcelReaderApi.FromCsv(_csvUtf8);
        var checksum = 17;
        foreach (CsvBenchmarkRow row in _excelReaderParser.Parse(reader))
        {
            checksum = AddTypedRowChecksum(checksum, row);
        }

        return checksum;
    }

    private int MeasureTypedRows(IEnumerable<CsvBenchmarkRow> rows)
    {
        var checksum = 17;
        foreach (CsvBenchmarkRow row in rows)
        {
            checksum = AddTypedRowChecksum(checksum, row);
        }

        return checksum;
    }

    private int AddTypedRowChecksum(int checksum, CsvBenchmarkRow row)
    {
        if (_validateTypedRead) ValidateTypedReadRow(row);
        unchecked
        {
            checksum = (checksum * 31) + row.Id;
            checksum = (checksum * 31) + StringComparer.Ordinal.GetHashCode(row.Name);
            checksum = (checksum * 31) + StringComparer.Ordinal.GetHashCode(row.Department);
            checksum = (checksum * 31) + StringComparer.Ordinal.GetHashCode(row.Region);
            checksum = (checksum * 31) + (row.IsEnabled ? 1 : 0);
            checksum = (checksum * 31) + row.Created.GetHashCode();
            checksum = (checksum * 31) + row.Score.GetHashCode();
            checksum = (checksum * 31) + StringComparer.Ordinal.GetHashCode(row.Owner);
            checksum = (checksum * 31) + row.TicketCount;
            checksum = (checksum * 31) + StringComparer.Ordinal.GetHashCode(row.Notes);
            return checksum;
        }
    }

    // Setup compares all decoded fields directly. Measurements retain the same
    // aggregate consumption, with validation disabled for every implementation.
    private void ValidateTypedReadRow(CsvBenchmarkRow actual)
    {
        int index = _typedReadValidationIndex;
        if ((uint)index >= (uint)_rows.Length)
            throw new InvalidOperationException("Typed CSV reader returned extra rows.");
        CsvBenchmarkRow expected = _rows[index];
        if (actual.Id != expected.Id
            || !string.Equals(actual.Name, expected.Name, StringComparison.Ordinal)
            || !string.Equals(actual.Department, expected.Department, StringComparison.Ordinal)
            || !string.Equals(actual.Region, expected.Region, StringComparison.Ordinal)
            || actual.IsEnabled != expected.IsEnabled
            || actual.Created != expected.Created
            || actual.Score != expected.Score
            || !string.Equals(actual.Owner, expected.Owner, StringComparison.Ordinal)
            || actual.TicketCount != expected.TicketCount
            || !string.Equals(actual.Notes, expected.Notes, StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"Typed CSV reader changed a field at data row {index + 1}.");
        }
        _typedReadValidationIndex++;
    }
}
