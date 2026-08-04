using System.Data.Common;
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using OfficeIMO.Benchmarks;
using OfficeIMO.Excel.Xlsb.Read;
using OfficeOpenXml;
using MiniExcelApi = MiniExcelLibs.MiniExcel;
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
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(Sylvan), Sylvan());
        Validate(nameof(ExcelDataReader), ExcelDataReader());
        Validate(nameof(ClosedXML), ClosedXML());
        Validate(nameof(EPPlus), EPPlus());
        Validate(nameof(MiniExcel), MiniExcel());
    }

    [Benchmark]
    public ExcelReadObservation OfficeIMO() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            MarkPflug65KFixture.XlsxPath,
            new ExcelReadOptions {
                NumericAsDecimal = true,
                SheetName = null
            });
        return Observe(reader);
    }

    [Benchmark]
    public ExcelReadObservation Sylvan() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsxPath);
        using SylvanExcelDataReader reader = SylvanExcelDataReader.Create(
            stream,
            ExcelWorkbookType.ExcelXml,
            new ExcelDataReaderOptions { Schema = ExcelSchema.Default });
        return Observe(reader);
    }

    [Benchmark]
    public ExcelReadObservation ExcelDataReader() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsxPath);
        using global::ExcelDataReader.IExcelDataReader reader =
            global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read()) {
            return default;
        }

        var observation = new ExcelObservationAccumulator();
        while (reader.Read()) {
            AddRow(
                ref observation,
                ordinal => Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty,
                ordinal => ReadDate(reader.GetValue(ordinal)),
                ordinal => Convert.ToInt32(reader.GetValue(ordinal), CultureInfo.InvariantCulture),
                ordinal => Convert.ToDecimal(reader.GetValue(ordinal), CultureInfo.InvariantCulture));
        }

        return observation.Build();
    }

    [Benchmark]
    public ExcelReadObservation ClosedXML() {
        using var workbook = new XLWorkbook(MarkPflug65KFixture.XlsxPath);
        IXLWorksheet worksheet = workbook.Worksheet(1);
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var observation = new ExcelObservationAccumulator();
        for (int row = 2; row <= lastRow; row++) {
            AddRow(
                ref observation,
                ordinal => worksheet.Cell(row, ordinal + 1).GetString(),
                ordinal => worksheet.Cell(row, ordinal + 1).GetDateTime(),
                ordinal => worksheet.Cell(row, ordinal + 1).GetValue<int>(),
                ordinal => worksheet.Cell(row, ordinal + 1).GetValue<decimal>());
        }

        return observation.Build();
    }

    [Benchmark]
    public ExcelReadObservation EPPlus() {
        using var package = new ExcelPackage(new FileInfo(MarkPflug65KFixture.XlsxPath));
        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
        int lastRow = worksheet.Dimension?.End.Row ?? 0;
        var observation = new ExcelObservationAccumulator();
        for (int row = 2; row <= lastRow; row++) {
            AddRow(
                ref observation,
                ordinal => Convert.ToString(worksheet.Cells[row, ordinal + 1].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                ordinal => ReadDate(worksheet.Cells[row, ordinal + 1].Value),
                ordinal => Convert.ToInt32(worksheet.Cells[row, ordinal + 1].Value, CultureInfo.InvariantCulture),
                ordinal => Convert.ToDecimal(worksheet.Cells[row, ordinal + 1].Value, CultureInfo.InvariantCulture));
        }

        return observation.Build();
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
            AddRow(
                ref observation,
                ordinal => Convert.ToString(row[Headers[ordinal]], CultureInfo.InvariantCulture) ?? string.Empty,
                ordinal => ReadDate(row[Headers[ordinal]]),
                ordinal => Convert.ToInt32(row[Headers[ordinal]], CultureInfo.InvariantCulture),
                ordinal => Convert.ToDecimal(row[Headers[ordinal]], CultureInfo.InvariantCulture));
        }

        return observation.Build();
    }

    internal static ExcelReadObservation Observe(DbDataReader reader) {
        var observation = new ExcelObservationAccumulator();
        while (reader.Read()) {
            AddRow(
                ref observation,
                reader.GetString,
                reader.GetDateTime,
                reader.GetInt32,
                reader.GetDecimal);
        }

        return observation.Build();
    }

    internal static void AddRow(
        ref ExcelObservationAccumulator observation,
        Func<int, string> text,
        Func<int, DateTime> date,
        Func<int, int> integer,
        Func<int, decimal> number) {
        observation.BeginRow();
        for (int ordinal = 0; ordinal <= 4; ordinal++) {
            observation.Add(text(ordinal));
        }
        observation.Add(date(5));
        observation.Add(integer(6));
        observation.Add(date(7));
        observation.Add(integer(8));
        for (int ordinal = 9; ordinal <= 13; ordinal++) {
            observation.Add(number(ordinal));
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

    private void Validate(string library, ExcelReadObservation actual) {
        if (actual != _expected
            || actual.Rows != MarkPflug65KFixture.ExpectedRows
            || actual.Cells != MarkPflug65KFixture.ExpectedRows * MarkPflug65KFixture.ExpectedColumns) {
            throw new InvalidDataException(
                $"{library} did not perform the same XLSX workload. Expected {_expected}; actual {actual}.");
        }
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
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(Sylvan), Sylvan());
        Validate(nameof(ExcelDataReader), ExcelDataReader());
    }

    [Benchmark]
    public ExcelReadObservation OfficeIMO() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            MarkPflug65KFixture.XlsbPath,
            new ExcelReadOptions { NumericAsDecimal = true });
        return MarkPflug65KXlsxBenchmarks.Observe(reader);
    }

    [Benchmark]
    public ExcelReadObservation Sylvan() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsbPath);
        using SylvanExcelDataReader reader = SylvanExcelDataReader.Create(
            stream,
            ExcelWorkbookType.ExcelBinary,
            new ExcelDataReaderOptions { Schema = ExcelSchema.Default });
        return MarkPflug65KXlsxBenchmarks.Observe(reader);
    }

    [Benchmark]
    public ExcelReadObservation ExcelDataReader() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsbPath);
        using global::ExcelDataReader.IExcelDataReader reader =
            global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read()) {
            return default;
        }

        var observation = new ExcelObservationAccumulator();
        while (reader.Read()) {
            MarkPflug65KXlsxBenchmarks.AddRow(
                ref observation,
                ordinal => Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty,
                ordinal => MarkPflug65KXlsxBenchmarks.ReadDate(reader.GetValue(ordinal)),
                ordinal => Convert.ToInt32(reader.GetValue(ordinal), CultureInfo.InvariantCulture),
                ordinal => Convert.ToDecimal(reader.GetValue(ordinal), CultureInfo.InvariantCulture));
        }

        return observation.Build();
    }

    private void Validate(string library, ExcelReadObservation actual) {
        if (actual != _expected
            || actual.Rows != MarkPflug65KFixture.ExpectedRows
            || actual.Cells != MarkPflug65KFixture.ExpectedRows * MarkPflug65KFixture.ExpectedColumns) {
            throw new InvalidDataException(
                $"{library} did not perform the same XLSB workload. Expected {_expected}; actual {actual}.");
        }
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
        foreach (int part in decimal.GetBits(value)) {
            AddUInt64(unchecked((uint)part));
        }
        _cells++;
    }

    internal readonly ExcelReadObservation Build() => new(_rows, _cells, _checksum);

    private void AddTag(byte value) {
        _checksum ^= value;
        _checksum *= Prime;
    }

    private void AddUInt64(ulong value) {
        for (int shift = 0; shift < 64; shift += 8) {
            _checksum ^= (byte)(value >> shift);
            _checksum *= Prime;
        }
    }
}
