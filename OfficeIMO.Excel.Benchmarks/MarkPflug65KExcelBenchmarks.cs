using System.Data.Common;
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using OfficeIMO.Benchmarks;
using OfficeIMO.Excel.Xlsb.Read;
using OfficeOpenXml;
using MiniExcelApi = MiniExcelLibs.MiniExcel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;
using Sylvan.Data.Excel;
using SylvanExcelDataReader = Sylvan.Data.Excel.ExcelDataReader;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Neutral typed scan of the hash-pinned 65K sales XLSX fixture. Every compatible
/// library reads the same fourteen columns and must produce the same observation.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class MarkPflug65KXlsxBenchmarks {
    private static readonly string[] Headers = [
        "Region",
        "Country",
        "Item Type",
        "Sales Channel",
        "Order Priority",
        "Order Date",
        "Order ID",
        "Ship Date",
        "Units Sold",
        "Unit Price",
        "Unit Cost",
        "Total Revenue",
        "Total Cost",
        "Total Profit"
    ];

    private ExcelReadObservation _expected;

    [GlobalSetup]
    public void Setup() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        if (ExcelPackage.License.LicenseType != EPPlusLicenseType.NonCommercialPersonal
            && ExcelPackage.License.LicenseType != EPPlusLicenseType.NonCommercialOrganization
            && ExcelPackage.License.LicenseType != EPPlusLicenseType.Commercial) {
            ExcelPackage.License.SetNonCommercialOrganization("OfficeIMO local benchmarks");
        }

        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.XlsxFileName);
        _expected = ExpectedObservation();
    }

    [Benchmark]
    public ExcelReadObservation OfficeIMO() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            MarkPflug65KFixture.XlsxPath,
            new ExcelReadOptions {
                NumericAsDecimal = true,
                SheetName = null
            });
        return Validate(nameof(OfficeIMO), Observe(reader));
    }

    [Benchmark]
    public ExcelReadObservation OfficeIMO_Prefetch() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            MarkPflug65KFixture.XlsxPath,
            new ExcelReadOptions {
                NumericAsDecimal = true,
                SheetName = null,
                EnableWorksheetPrefetch = true
            });
        return Validate(nameof(OfficeIMO_Prefetch), Observe(reader));
    }

    [Benchmark]
    public ExcelReadObservation ExcelReaderNet() {
        using XlsxReader reader = ExcelReaderApi.FromFile(MarkPflug65KFixture.XlsxPath);
        return Validate(nameof(ExcelReaderNet), ObserveExcelReader(reader));
    }

    [Benchmark]
    public ExcelReadObservation ExcelReaderNet_Prefetch() {
        ExcelReaderOptions options = ExcelReaderOptions.Default with { PrefetchDecompression = true };
        using XlsxReader reader = ExcelReaderApi.FromFile(MarkPflug65KFixture.XlsxPath, options);
        return Validate(nameof(ExcelReaderNet_Prefetch), ObserveExcelReader(reader));
    }

    [Benchmark]
    public ExcelReadObservation Sylvan() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsxPath);
        using SylvanExcelDataReader reader = SylvanExcelDataReader.Create(
            stream,
            ExcelWorkbookType.ExcelXml,
            new ExcelDataReaderOptions { Schema = ExcelSchema.Default });
        return Validate(nameof(Sylvan), Observe(reader));
    }

    [Benchmark]
    public ExcelReadObservation ExcelDataReader() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsxPath);
        using global::ExcelDataReader.IExcelDataReader reader =
            global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read()) {
            return Validate(nameof(ExcelDataReader), default);
        }

        var observation = new ExcelObservationAccumulator();
        while (reader.Read()) {
            AddExcelDataReaderRow(ref observation, reader);
        }

        return Validate(nameof(ExcelDataReader), observation.Build());
    }

    [Benchmark]
    public ExcelReadObservation ClosedXML() {
        using var workbook = new XLWorkbook(MarkPflug65KFixture.XlsxPath);
        IXLWorksheet worksheet = workbook.Worksheet(1);
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var observation = new ExcelObservationAccumulator();
        for (int row = 2; row <= lastRow; row++) {
            AddClosedXmlRow(ref observation, worksheet, row);
        }

        return Validate(nameof(ClosedXML), observation.Build());
    }

    [Benchmark]
    public ExcelReadObservation EPPlus() {
        using var package = new ExcelPackage(new FileInfo(MarkPflug65KFixture.XlsxPath));
        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
        int lastRow = worksheet.Dimension?.End.Row ?? 0;
        var observation = new ExcelObservationAccumulator();
        for (int row = 2; row <= lastRow; row++) {
            AddEpplusRow(ref observation, worksheet, row);
        }

        return Validate(nameof(EPPlus), observation.Build());
    }

    [Benchmark]
    public ExcelReadObservation MiniExcel() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsxPath);
        IEnumerable<object> rows = MiniExcelApi.Query(
            stream,
            useHeaderRow: true,
            excelType: MiniExcelLibs.ExcelType.XLSX);
        var observation = new ExcelObservationAccumulator();
        foreach (object item in rows) {
            var row = (IDictionary<string, object?>)item;
            AddMiniExcelRow(ref observation, row);
        }

        return Validate(nameof(MiniExcel), observation.Build());
    }

    internal static ExcelReadObservation Observe(DbDataReader reader) {
        var observation = new ExcelObservationAccumulator();
        while (reader.Read()) {
            AddRow(ref observation, reader);
        }

        return observation.Build();
    }

    internal static ExcelReadObservation ObserveExcelReader(XlsxReader reader) {
        var observation = new ExcelObservationAccumulator();
        bool header = true;
        bool isDate1904 = reader.IsDate1904;
        foreach (Row row in reader) {
            if (header) {
                header = false;
                continue;
            }

            AddRow(ref observation, row, isDate1904);
        }

        return observation.Build();
    }

    internal static ExcelReadObservation ObserveExcelReader(XlsbReader reader) {
        var observation = new ExcelObservationAccumulator();
        bool header = true;
        bool isDate1904 = reader.IsDate1904;
        foreach (Row row in reader) {
            if (header) {
                header = false;
                continue;
            }

            AddRow(ref observation, row, isDate1904);
        }

        return observation.Build();
    }

    internal static ExcelReadObservation ObserveExcelReader(XlsReader reader) {
        var observation = new ExcelObservationAccumulator();
        bool header = true;
        bool isDate1904 = reader.IsDate1904;
        foreach (Row row in reader) {
            if (header) {
                header = false;
                continue;
            }

            AddRow(ref observation, row, isDate1904);
        }

        return observation.Build();
    }

    private static void AddRow(
        ref ExcelObservationAccumulator observation,
        Row row,
        bool isDate1904) {
        observation.BeginRow();
        for (int ordinal = 0; ordinal <= 4; ordinal++) {
            observation.Add(row[ordinal].GetString());
        }

        observation.Add(ReadExcelReaderDate(row[5], isDate1904, 5));
        observation.Add(ReadExcelReaderInteger(row[6], 6));
        observation.Add(ReadExcelReaderDate(row[7], isDate1904, 7));
        observation.Add(ReadExcelReaderInteger(row[8], 8));
        for (int ordinal = 9; ordinal <= 13; ordinal++) {
            observation.Add(ReadExcelReaderDecimal(row[ordinal], ordinal));
        }
    }

    private static DateTime ReadExcelReaderDate(Cell cell, bool isDate1904, int ordinal) {
        if (cell.TryGetDateTime(isDate1904, out DateTime value)) {
            return value;
        }

        throw new InvalidDataException($"ExcelReader.NET did not produce a date at column {ordinal}.");
    }

    private static int ReadExcelReaderInteger(Cell cell, int ordinal) {
        if (cell.TryParse(CultureInfo.InvariantCulture, out int value)) {
            return value;
        }

        throw new InvalidDataException($"ExcelReader.NET did not produce an integer at column {ordinal}.");
    }

    private static decimal ReadExcelReaderDecimal(Cell cell, int ordinal) {
        // The source workbook stores IEEE-754 numbers. Read that value first and
        // convert it to decimal so the benchmark preserves the intended monetary
        // precision instead of parsing the binary rendering's trailing digits.
        if (cell.TryGetDouble(out double value)) {
            return (decimal)value;
        }

        throw new InvalidDataException($"ExcelReader.NET did not produce a number at column {ordinal}.");
    }

    private static void AddRow(
        ref ExcelObservationAccumulator observation,
        DbDataReader reader) {
        observation.BeginRow();
        for (int ordinal = 0; ordinal <= 4; ordinal++) {
            observation.Add(reader.GetString(ordinal));
        }
        observation.Add(reader.GetDateTime(5));
        observation.Add(reader.GetInt32(6));
        observation.Add(reader.GetDateTime(7));
        observation.Add(reader.GetInt32(8));
        for (int ordinal = 9; ordinal <= 13; ordinal++) {
            observation.Add(reader.GetDecimal(ordinal));
        }
    }

    internal static void AddExcelDataReaderRow(
        ref ExcelObservationAccumulator observation,
        global::ExcelDataReader.IExcelDataReader reader) {
        observation.BeginRow();
        for (int ordinal = 0; ordinal <= 4; ordinal++) {
            observation.Add(Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty);
        }
        observation.Add(ReadDate(reader.GetValue(5)));
        observation.Add(Convert.ToInt32(reader.GetValue(6), CultureInfo.InvariantCulture));
        observation.Add(ReadDate(reader.GetValue(7)));
        observation.Add(Convert.ToInt32(reader.GetValue(8), CultureInfo.InvariantCulture));
        for (int ordinal = 9; ordinal <= 13; ordinal++) {
            observation.Add(Convert.ToDecimal(reader.GetValue(ordinal), CultureInfo.InvariantCulture));
        }
    }

    private static void AddClosedXmlRow(
        ref ExcelObservationAccumulator observation,
        IXLWorksheet worksheet,
        int row) {
        observation.BeginRow();
        for (int ordinal = 0; ordinal <= 4; ordinal++) {
            observation.Add(worksheet.Cell(row, ordinal + 1).GetString());
        }
        observation.Add(worksheet.Cell(row, 6).GetDateTime());
        observation.Add(worksheet.Cell(row, 7).GetValue<int>());
        observation.Add(worksheet.Cell(row, 8).GetDateTime());
        observation.Add(worksheet.Cell(row, 9).GetValue<int>());
        for (int ordinal = 9; ordinal <= 13; ordinal++) {
            observation.Add(worksheet.Cell(row, ordinal + 1).GetValue<decimal>());
        }
    }

    private static void AddEpplusRow(
        ref ExcelObservationAccumulator observation,
        ExcelWorksheet worksheet,
        int row) {
        observation.BeginRow();
        for (int ordinal = 0; ordinal <= 4; ordinal++) {
            observation.Add(Convert.ToString(worksheet.Cells[row, ordinal + 1].Value, CultureInfo.InvariantCulture) ?? string.Empty);
        }
        observation.Add(ReadDate(worksheet.Cells[row, 6].Value));
        observation.Add(Convert.ToInt32(worksheet.Cells[row, 7].Value, CultureInfo.InvariantCulture));
        observation.Add(ReadDate(worksheet.Cells[row, 8].Value));
        observation.Add(Convert.ToInt32(worksheet.Cells[row, 9].Value, CultureInfo.InvariantCulture));
        for (int ordinal = 9; ordinal <= 13; ordinal++) {
            observation.Add(Convert.ToDecimal(worksheet.Cells[row, ordinal + 1].Value, CultureInfo.InvariantCulture));
        }
    }

    private static void AddMiniExcelRow(
        ref ExcelObservationAccumulator observation,
        IDictionary<string, object?> row) {
        observation.BeginRow();
        for (int ordinal = 0; ordinal <= 4; ordinal++) {
            observation.Add(Convert.ToString(row[Headers[ordinal]], CultureInfo.InvariantCulture) ?? string.Empty);
        }
        observation.Add(ReadDate(row[Headers[5]]));
        observation.Add(Convert.ToInt32(row[Headers[6]], CultureInfo.InvariantCulture));
        observation.Add(ReadDate(row[Headers[7]]));
        observation.Add(Convert.ToInt32(row[Headers[8]], CultureInfo.InvariantCulture));
        for (int ordinal = 9; ordinal <= 13; ordinal++) {
            observation.Add(Convert.ToDecimal(row[Headers[ordinal]], CultureInfo.InvariantCulture));
        }
    }

    internal static DateTime ReadDate(object? value) =>
        value switch {
            DateTime date => date,
            double serial => DateTime.FromOADate(serial),
            _ => Convert.ToDateTime(value, CultureInfo.InvariantCulture)
        };

    internal static ExcelReadObservation ExpectedObservation() => new(
        MarkPflug65KFixture.ExpectedRows,
        MarkPflug65KFixture.ExpectedRows * MarkPflug65KFixture.ExpectedColumns,
        MarkPflug65KFixture.ExpectedExcelChecksum);

    private ExcelReadObservation Validate(string library, ExcelReadObservation actual) {
        if (actual != _expected
            || actual.Rows != MarkPflug65KFixture.ExpectedRows
            || actual.Cells != MarkPflug65KFixture.ExpectedRows * MarkPflug65KFixture.ExpectedColumns) {
            throw new InvalidDataException(
                $"{library} did not perform the same XLSX workload. Expected {_expected}; actual {actual}.");
        }

        return actual;
    }
}

/// <summary>
/// Neutral typed scan of the hash-pinned 65K sales XLSB fixture. Only libraries
/// with a compatible XLSB read path participate in this workload.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class MarkPflug65KXlsbBenchmarks {
    private ExcelReadObservation _expected;

    [GlobalSetup]
    public void Setup() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.XlsbFileName);
        _expected = MarkPflug65KXlsxBenchmarks.ExpectedObservation();
    }

    [Benchmark]
    public ExcelReadObservation OfficeIMO() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            MarkPflug65KFixture.XlsbPath,
            new ExcelReadOptions { NumericAsDecimal = true });
        return Validate(nameof(OfficeIMO), MarkPflug65KXlsxBenchmarks.Observe(reader));
    }

    [Benchmark]
    public ExcelReadObservation ExcelReaderNet() {
        using XlsbReader reader = ExcelReaderApi.FromXlsbFile(MarkPflug65KFixture.XlsbPath);
        return Validate(nameof(ExcelReaderNet), MarkPflug65KXlsxBenchmarks.ObserveExcelReader(reader));
    }

    [Benchmark]
    public ExcelReadObservation Sylvan() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsbPath);
        using SylvanExcelDataReader reader = SylvanExcelDataReader.Create(
            stream,
            ExcelWorkbookType.ExcelBinary,
            new ExcelDataReaderOptions { Schema = ExcelSchema.Default });
        return Validate(nameof(Sylvan), MarkPflug65KXlsxBenchmarks.Observe(reader));
    }

    [Benchmark]
    public ExcelReadObservation ExcelDataReader() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsbPath);
        using global::ExcelDataReader.IExcelDataReader reader =
            global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read()) {
            return Validate(nameof(ExcelDataReader), default);
        }

        var observation = new ExcelObservationAccumulator();
        while (reader.Read()) {
            MarkPflug65KXlsxBenchmarks.AddExcelDataReaderRow(ref observation, reader);
        }

        return Validate(nameof(ExcelDataReader), observation.Build());
    }

    private ExcelReadObservation Validate(string library, ExcelReadObservation actual) {
        if (actual != _expected
            || actual.Rows != MarkPflug65KFixture.ExpectedRows
            || actual.Cells != MarkPflug65KFixture.ExpectedRows * MarkPflug65KFixture.ExpectedColumns) {
            throw new InvalidDataException(
                $"{library} did not perform the same XLSB workload. Expected {_expected}; actual {actual}.");
        }

        return actual;
    }
}

/// <summary>
/// OfficeIMO-only diagnostic that isolates the public multi-sheet wrapper from
/// the package-owned XLSB worksheet reader. It is excluded from the public
/// multi-library comparison filter.
/// </summary>
[MemoryDiagnoser]
public class XlsbOfficeIMOPipelineBenchmarks {
    private ExcelReadObservation _expected;

    [GlobalSetup]
    public void Setup() {
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.XlsbFileName);
        _expected = PublicApi();
        ExcelReadObservation direct = DirectWorksheetReader();
        if (direct != _expected) {
            throw new InvalidDataException(
                $"The direct XLSB worksheet reader produced {direct}; the public API produced {_expected}.");
        }
    }

    [Benchmark(Baseline = true)]
    public ExcelReadObservation PublicApi() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            MarkPflug65KFixture.XlsbPath,
            new ExcelReadOptions { NumericAsDecimal = true });
        return MarkPflug65KXlsxBenchmarks.Observe(reader);
    }

    [Benchmark]
    public ExcelReadObservation DirectWorksheetReader() {
        var options = new ExcelReadOptions { NumericAsDecimal = true };
        using XlsbTabularWorkbook workbook = XlsbTabularWorkbook.Open(
            MarkPflug65KFixture.XlsbPath,
            options);
        using DbDataReader reader = workbook.OpenTable(
            workbook.TableNames[0],
            options.HasHeaderRow,
            options);
        return MarkPflug65KXlsxBenchmarks.Observe(reader);
    }
}

public readonly record struct ExcelReadObservation(int Rows, int Cells, ulong Checksum);

internal struct ExcelObservationAccumulator {
    private const ulong OffsetBasis = 14695981039346656037UL;
    private const ulong Prime = 1099511628211UL;
    private int _rows;
    private int _cells;
    private ulong _checksum;

    internal void BeginRow() {
        if (_checksum == 0) {
            _checksum = OffsetBasis;
        }
        _rows++;
    }

    internal void Add(string value) {
        AddTag(1);
        foreach (char character in value) {
            AddUInt64(character);
        }
        AddUInt64((ulong)value.Length);
        _cells++;
    }

    internal void Add(DateTime value) {
        AddTag(2);
        AddUInt64(unchecked((ulong)value.Ticks));
        _cells++;
    }

    internal void Add(int value) {
        AddTag(3);
        AddUInt64(unchecked((uint)value));
        _cells++;
    }

    internal void Add(decimal value) {
        AddTag(4);
        Span<int> parts = stackalloc int[4];
        decimal.GetBits(value, parts);
        for (int index = 0; index < parts.Length; index++) {
            AddUInt64(unchecked((uint)parts[index]));
        }
        _cells++;
    }

    internal readonly ExcelReadObservation Build() => new(_rows, _cells, _checksum);

    private void AddTag(byte value) {
        _checksum ^= value;
        _checksum *= Prime;
    }

    private void AddUInt64(ulong value) {
        // One deterministic FNV-style word mix keeps every observed value in the
        // timed contract without making validation dominate the reader itself.
        _checksum ^= value;
        _checksum *= Prime;
    }
}
