using BenchmarkDotNet.Attributes;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.CSV.Benchmarks;

public partial class CsvBenchmarks {
    /// <summary>Materializes every field, matching the string accessors used by the other raw read lanes.</summary>
    [Benchmark]
    public int ExcelReaderNet_ReadFields() {
        using ExcelReaderNetCsvReader reader = ExcelReaderApi.FromCsv(_csvUtf8);
        int fieldCount = 0;
        bool header = true;
#if NET10_0_OR_GREATER
        var rows = reader.FirstSheet;
#else
        var rows = reader;
#endif
        foreach (var row in rows) {
            if (header) {
                for (int column = 0; column < Headers.Length; column++)
                    if (row[column].GetString() != Headers[column]) throw new InvalidDataException("Incorrect CSV header.");
                header = false;
                continue;
            }
            if (row.ColumnCount != Headers.Length) throw new InvalidDataException("Incorrect CSV field count.");
            for (int column = 0; column < row.ColumnCount; column++) {
                string value = row[column].GetString();
                fieldCount += 1 + value.Length;
                if (_validateRawRead) ValidateRawField(column, value);
            }
            if (_validateRawRead) CompleteRawRow();
        }
        return fieldCount;
    }
}

/// <summary>A fixed generated-data comparison using the existing validated ten-column mixed CSV workload.</summary>
[MemoryDiagnoser]
public class CsvExcelReaderGeneratedBenchmarks {
    private readonly CsvBenchmarks _workload = new() { RowCount = 50_000, Shape = CsvBenchmarkShape.Mixed };
    [GlobalSetup]
    public void Setup() {
        OfficeIMO.Benchmarks.BenchmarkProcessPriority.ApplyConfiguredPriority();
        _workload.Setup();
        _workload.ValidateGeneratedUtf8Reads();
    }
    [Benchmark] public int OfficeIMO_Read() => _workload.OfficeIMO_ReadStreamGetStrings();
    [Benchmark] public int ExcelReaderNet_Read() => _workload.ExcelReaderNet_ReadFields();
    [Benchmark] public int Sylvan_Read() => _workload.Sylvan_ReadUtf8Fields();
    [Benchmark] public int Sep_Read() => _workload.Sep_ReadUtf8Fields();
    [Benchmark] public int CsvHelper_Read() => _workload.CsvHelper_ReadUtf8Fields();
    [Benchmark] public int Dataplat_Read() => _workload.Dataplat_ReadUtf8Fields();
    [Benchmark] public int OfficeIMO_Parse() => _workload.OfficeIMO_ReadTypedRowsStream();
    [Benchmark] public int ExcelReaderNet_Parse() => _workload.ExcelReaderNet_ReadTypedRecords();
    [Benchmark] public int CsvHelper_Parse() => _workload.CsvHelper_ReadTypedRecordsUtf8();
    [Benchmark] public int OfficeIMO_Write() => _workload.OfficeIMO_WriteProjectedRowsUtf8();
    [Benchmark] public int ExcelReaderNet_Write() => _workload.ExcelReaderNet_WriteProjectedRows();
}
