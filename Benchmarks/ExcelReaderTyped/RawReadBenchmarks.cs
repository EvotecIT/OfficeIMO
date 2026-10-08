using System.Data;
using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Reproduces the raw reader APIs, including their different string materialization work.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("RawOriginal")]
public partial class RawReadBenchmarks {
    private byte[] _workbook = [];
    private long _expected;

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;

    [ParamsAllValues]
    public RawWorkbookFormat Format { get; set; }

    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

    [GlobalSetup]
    public async Task SetupAsync() {
        BenchmarkInput.WriteDescription();
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _workbook = await RawWorkbookFixture.CreateAsync(RowCount, Format);
        _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
        ValidateExcelReader();
        ValidateOfficeIMO();
        ValidateSylvan();
        DescribeUpstreamSylvanDefaults();
        ExcelReaderOriginal();
        ExcelReaderStringsMaterialized();
        OfficeIMOOriginal();
        Sylvan();
        Console.WriteLine($"Validated raw {Format}: rows={RowCount}, columns=4, packageBytes={_workbook.Length}, checksum={_expected}.");
    }

    [Benchmark(Baseline = true)]
    public long ExcelReaderOriginal() => ReadExcelReader(materializeStrings: false);

    [Benchmark]
    public long OfficeIMOOriginal() {
        using var reader = ExcelDocument.OpenDataReader(_workbook, new ExcelReadOptions { HasHeaderRow = false });
        long sum = 0;
        int count = 0;
        while (reader.Read()) {
            for (int column = 0; column < reader.FieldCount; column++) {
                switch (reader.GetValue(column)) {
                    case string value: sum = unchecked(sum + value.Length); break;
                    case double value: sum = unchecked(sum + (long)value); break;
                    case DateTime value: sum = unchecked(sum + value.Ticks); break;
                }
            }
            count++;
        }
        return Check(sum, count);
    }

    [Benchmark]
    public long Sylvan() {
        using var stream = new MemoryStream(_workbook, writable: false);
        using var reader = OpenSylvan(stream);
        long sum = 0;
        int count = 0;
        do {
            while (reader.Read()) {
                for (int column = 0; column < reader.FieldCount; column++) {
                    if (reader.IsDBNull(column)) continue;
                    switch (reader.GetExcelDataType(column)) {
                        case ExcelDataType.String: sum = unchecked(sum + reader.GetString(column).Length); break;
                        case ExcelDataType.Numeric:
                            sum = unchecked(sum + (reader.GetFormat(column)?.Kind == FormatKind.Date
                                ? reader.GetDateTime(column).Ticks : (long)reader.GetDouble(column)));
                            break;
                        case ExcelDataType.DateTime: sum = unchecked(sum + reader.GetDateTime(column).Ticks); break;
                    }
                }
                count++;
            }
        } while (reader.NextResult());
        return Check(sum, count);
    }

    internal long ExcelReaderStringsMaterialized() => ReadExcelReader(materializeStrings: true);

    private long ReadExcelReader(bool materializeStrings) => Format switch {
        RawWorkbookFormat.Xlsx => ReadXlsx(materializeStrings),
        RawWorkbookFormat.Xlsb => ReadXlsb(materializeStrings),
        RawWorkbookFormat.Xls => ReadXls(materializeStrings),
        _ => throw new ArgumentOutOfRangeException(nameof(Format)),
    };

    // Keep each timed loop on the concrete workbook API used upstream, rather
    // than introducing IExcelWorkbook/IExcelSheet dispatch into the comparison.
    private long ReadXlsx(bool materializeStrings) {
        using var stream = new MemoryStream(_workbook, writable: false);
        using var workbook = ExcelReaderApi.FromXlsx(stream);
        long sum = 0;
        int count = 0;
        foreach (Row row in workbook.FirstSheet) {
            sum = unchecked(sum + AccumulateRow(row, materializeStrings));
            count++;
        }
        return Check(sum, count);
    }

    private long ReadXlsb(bool materializeStrings) {
        using var stream = new MemoryStream(_workbook, writable: false);
        using var workbook = ExcelReaderApi.FromXlsb(stream);
        long sum = 0;
        int count = 0;
        foreach (Row row in workbook.FirstSheet) {
            sum = unchecked(sum + AccumulateRow(row, materializeStrings));
            count++;
        }
        return Check(sum, count);
    }

    private long ReadXls(bool materializeStrings) {
        using var stream = new MemoryStream(_workbook, writable: false);
        using var workbook = ExcelReaderApi.FromXls(stream);
        long sum = 0;
        int count = 0;
        foreach (Row row in workbook.FirstSheet) {
            sum = unchecked(sum + AccumulateRow(row, materializeStrings));
            count++;
        }
        return Check(sum, count);
    }

    private static long AccumulateRow(Row row, bool materializeStrings) {
        long sum = 0;
        foreach (RowCell rowCell in row.Cells) {
            Cell cell = rowCell.Value;
            switch (cell.Type) {
                case CellType.ExcelString:
                    sum = unchecked(sum + (materializeStrings ? cell.GetString().Length : cell.Value.Length));
                    break;
                case CellType.Number:
                    if (cell.TryParse(null, out double number)) sum = unchecked(sum + (long)number);
                    break;
                case CellType.Date:
                    if (cell.TryGetDateTime(out DateTime date)) sum = unchecked(sum + date.Ticks);
                    break;
            }
        }
        return sum;
    }

    private IExcelWorkbook OpenExcelReader(Stream stream) => Format switch {
        RawWorkbookFormat.Xlsx => ExcelReaderApi.FromXlsx(stream),
        RawWorkbookFormat.Xlsb => ExcelReaderApi.FromXlsb(stream),
        RawWorkbookFormat.Xls => ExcelReaderApi.FromXls(stream),
        _ => throw new ArgumentOutOfRangeException(nameof(Format)),
    };

    private global::Sylvan.Data.Excel.ExcelDataReader OpenSylvan(Stream stream, bool hasNoHeaders = true) =>
        global::Sylvan.Data.Excel.ExcelDataReader.Create(stream, Format switch {
            RawWorkbookFormat.Xlsx => ExcelWorkbookType.ExcelXml,
            RawWorkbookFormat.Xlsb => ExcelWorkbookType.ExcelBinary,
            RawWorkbookFormat.Xls => ExcelWorkbookType.Excel,
            _ => throw new ArgumentOutOfRangeException(nameof(Format)),
        }, hasNoHeaders ? new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders } : new ExcelDataReaderOptions());

    private void ValidateExcelReader() {
        using var stream = new MemoryStream(_workbook, writable: false);
        using IExcelWorkbook workbook = OpenExcelReader(stream);
        if (workbook.SheetCount != 1) throw new InvalidDataException("Incorrect sheet count.");
        int count = 0;
        foreach (Row row in workbook.FirstSheet) {
            ValidateExcelReaderRow(row, ++count);
        }
        if (count != RowCount) throw new InvalidDataException("Incorrect raw row count.");
    }

    private static void ValidateExcelReaderRow(Row row, int index) {
        if (row.ColumnCount != 4 || row[0].Type != CellType.ExcelString || row[1].Type != CellType.Number
            || row[2].Type != CellType.Date || row[3].Type != CellType.Number
            || !row[1].TryParse(null, out double id) || !row[2].TryGetDateTime(out DateTime date)
            || !row[3].TryParse(null, out double value))
            throw new InvalidDataException("Incorrect raw field type or shape.");
        TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = row[0].GetString(), Id = (int)id,
            Date = date, Value = value }, index);
        if (id != index) throw new InvalidDataException("Incorrect numeric ID.");
    }

    private void ValidateOfficeIMO() {
        using var reader = ExcelDocument.OpenDataReader(_workbook, new ExcelReadOptions { HasHeaderRow = false });
        int count = 0;
        while (reader.Read()) {
            if (reader.FieldCount != 4 || reader.GetValue(0) is not string name || reader.GetValue(1) is not double id
                || reader.GetValue(2) is not DateTime date || reader.GetValue(3) is not double value)
                throw new InvalidDataException("OfficeIMO returned incorrect raw types or shape.");
            TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = name, Id = (int)id, Date = date, Value = value }, ++count);
            if (id != count) throw new InvalidDataException("Incorrect numeric ID.");
        }
        if (count != RowCount || reader.NextResult()) throw new InvalidDataException("Incorrect raw row or sheet count.");
    }

    private void ValidateSylvan() {
        using var stream = new MemoryStream(_workbook, writable: false);
        using var reader = OpenSylvan(stream);
        int count = 0;
        while (reader.Read()) {
            if (reader.FieldCount != 4 || reader.GetExcelDataType(0) != ExcelDataType.String
                || reader.GetExcelDataType(1) != ExcelDataType.Numeric || reader.GetExcelDataType(2) != ExcelDataType.Numeric
                || reader.GetFormat(2)?.Kind != FormatKind.Date
                || reader.GetExcelDataType(3) != ExcelDataType.Numeric)
                throw new InvalidDataException($"Sylvan returned raw shape {reader.FieldCount} and types "
                    + string.Join(',', Enumerable.Range(0, reader.FieldCount).Select(reader.GetExcelDataType)));
            double id = reader.GetDouble(1);
            TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = reader.GetString(0), Id = (int)id,
                Date = reader.GetDateTime(2), Value = reader.GetDouble(3) }, ++count);
            if (id != count) throw new InvalidDataException("Incorrect numeric ID.");
        }
        if (count != RowCount || reader.NextResult()) throw new InvalidDataException("Incorrect raw row or sheet count.");
    }

    private void DescribeUpstreamSylvanDefaults() {
        using var stream = new MemoryStream(_workbook, writable: false);
        using var reader = OpenSylvan(stream, hasNoHeaders: false);
        int count = 0;
        double? firstId = null;
        while (reader.Read()) {
            firstId ??= reader.GetDouble(1);
            count++;
        }
        Console.WriteLine($"Upstream Sylvan default-header diagnostic for {Format}: returnedRows={count}, firstId={firstId}.");
    }

    private long Check(long sum, int count) => sum == _expected && count == RowCount
        ? sum : throw new InvalidDataException("Raw scan returned an incorrect count or checksum.");
}

/// <summary>Separately compares raw APIs after ExcelReader also materializes text.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("RawStringsMaterialized")]
public class MaterializedRawReadBenchmarks {
    private readonly RawReadBenchmarks _workload = new();

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;

    [ParamsAllValues]
    public RawWorkbookFormat Format { get; set; }

    public IEnumerable<int> RowCounts() => _workload.RowCounts();

    [GlobalSetup]
    public Task SetupAsync() {
        _workload.RowCount = RowCount;
        _workload.Format = Format;
        return _workload.SetupAsync();
    }

    [Benchmark(Baseline = true)]
    public long ExcelReaderStringsMaterialized() => _workload.ExcelReaderStringsMaterialized();
    [Benchmark]
    public long OfficeIMOOriginal() => _workload.OfficeIMOOriginal();
    [Benchmark]
    public long Sylvan() => _workload.Sylvan();
}
