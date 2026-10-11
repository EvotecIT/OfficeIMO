using System.ComponentModel;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed class ExcelAutomaticNativeBinaryMappingTests {
    [Theory]
    [InlineData(ExcelFileFormat.Xls, false, ExcelDateSystem.NineteenHundred)]
    [InlineData(ExcelFileFormat.Xls, true, ExcelDateSystem.NineteenHundred)]
    [InlineData(ExcelFileFormat.Xlsb, false, ExcelDateSystem.NineteenHundred)]
    [InlineData(ExcelFileFormat.Xlsb, true, ExcelDateSystem.NineteenHundred)]
    [InlineData(ExcelFileFormat.Xls, false, ExcelDateSystem.NineteenFour)]
    [InlineData(ExcelFileFormat.Xls, true, ExcelDateSystem.NineteenFour)]
    [InlineData(ExcelFileFormat.Xlsb, false, ExcelDateSystem.NineteenFour)]
    [InlineData(ExcelFileFormat.Xlsb, true, ExcelDateSystem.NineteenFour)]
    public void AllNativeScalarFieldsMatchOrdinaryValueConversion(
        ExcelFileFormat format, bool valueModel, ExcelDateSystem dateSystem) {
        byte[] workbook = CreateWorkbook(format,
            ["Name", "Flag", "Byte", "Short", "Order Id", "Long", "Single", "Double", "Amount", "Date", "OtherText", "OtherFlag", "OtherId", "OtherValue"],
            [
                ["Żółw 🐢\r\n", true, 126.5, -3.5, 7.5, 1234567890123.5, 1.0000000596046448, -0d, 12.125, new DateTime(2026, 1, 2, 3, 4, 5), "second", false, -1.5, 1.2345678901234567],
                ["漢字", false, 127.5, -2.5, 8.5, -1234567890123.5, -2.25, 42.5, -0.125, new DateTime(2026, 2, 3, 4, 5, 6), "third", true, 2.5, -42.125]
            ], dateSystem);
        using var oracle = ExcelDocument.OpenDataReader(workbook);
        using var reader = ExcelDocument.OpenDataReader(workbook);
        if (valueModel) {
            foreach (FullValueRow row in reader.RowsAs<FullValueRow>()) AssertNativeRow(oracle, row);
        } else {
            foreach (FullClassRow row in reader.RowsAs<FullClassRow>()) AssertNativeRow(oracle, row);
        }
        Assert.False(oracle.Read());
        Assert.False(reader.IsClosed);
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xls, false, false)]
    [InlineData(ExcelFileFormat.Xls, false, true)]
    [InlineData(ExcelFileFormat.Xls, true, false)]
    [InlineData(ExcelFileFormat.Xls, true, true)]
    [InlineData(ExcelFileFormat.Xlsb, false, false)]
    [InlineData(ExcelFileFormat.Xlsb, false, true)]
    [InlineData(ExcelFileFormat.Xlsb, true, false)]
    [InlineData(ExcelFileFormat.Xlsb, true, true)]
    public void HeterogeneousRowsKeepDecimalCultureTextDateAndMissingConversions(ExcelFileFormat format, bool numericAsDecimal, bool inferSchema) {
        byte[] workbook = CreateWorkbook(format,
            ["Text", "Single", "Id", "Flag", "Date", "Serial", "Name", "Optional"],
            [
                [1.2345678901234567, 1.0000000596046448, "1,0", "1", "2024-01-02T03:04:05.1234567Z", 1.25, null, null],
                [42.5, -2.25, 7.5, false, new DateTime(2026, 1, 2), 2.5, "following", 3.5]
            ], ExcelDateSystem.NineteenFour, styledSerialColumn: 6);
        CultureInfo culture = CultureInfo.GetCultureInfo("pl-PL");
        ExcelReadOptions Options() => new() { NumericAsDecimal = numericAsDecimal, Culture = culture, InferSchema = inferSchema, SchemaSampleRows = 1 };
        using var oracle = ExcelDocument.OpenDataReader(workbook, Options());
        using var reader = ExcelDocument.OpenDataReader(workbook, Options());
        FallbackRow[] rows = reader.RowsAs<FallbackRow>().ToArray();
        Assert.Equal(2, rows.Length);
        for (int index = 0; index < rows.Length; index++) {
            Assert.True(oracle.Read());
            Assert.Equal(Convert.ToString(oracle.GetValue(0), culture), rows[index].Text);
            Assert.Equal(Convert.ToSingle(oracle.GetValue(1), culture), rows[index].Single);
            Assert.Equal(index == 0 ? 1 : 8, rows[index].Id);
            Assert.Equal(index == 0, rows[index].Flag);
            DateTime expectedDate = index == 0
                ? DateTime.Parse("2024-01-02T03:04:05.1234567Z", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind)
                : Assert.IsType<DateTime>(oracle.GetValue(4));
            Assert.Equal(expectedDate, rows[index].Date);
            Assert.Equal(expectedDate.Kind, rows[index].Date.Kind);
            Assert.Equal(oracle.GetDouble(5), rows[index].Serial);
            Assert.Equal(index == 0 ? null : "following", rows[index].Name);
            Assert.Equal(index == 0 ? null : (int?)4, rows[index].Optional);
        }
        Assert.False(oracle.Read());
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xls, false)]
    [InlineData(ExcelFileFormat.Xls, true)]
    [InlineData(ExcelFileFormat.Xlsb, false)]
    [InlineData(ExcelFileFormat.Xlsb, true)]
    public void CustomConvertersReceiveTheOriginalValues(ExcelFileFormat format, bool cellConverter) {
        byte[] workbook = CreateWorkbook(format, ["Id", "Serial"], [[7.5, 1.25]],
            ExcelDateSystem.NineteenFour, styledSerialColumn: 2);
        int calls = 0;
        var options = new ExcelReadOptions {
            Culture = CultureInfo.GetCultureInfo("pl-PL"),
            CellValueConverter = cellConverter
                ? context => context.RawText == "7.5" ? new ExcelCellValue(9d) : ExcelCellValue.NotHandled
                : null,
            TypeConverter = (raw, type, culture) => {
                Assert.Equal("pl-PL", culture.Name);
                calls++;
                if (type == typeof(int)) {
                    Assert.Equal(cellConverter ? 9d : 7.5, Assert.IsType<double>(raw));
                    return (true, 42);
                }
                Assert.Equal(typeof(double), type);
                Assert.Equal(new DateTime(1904, 1, 2, 6, 0, 0), Assert.IsType<DateTime>(raw));
                return (false, null);
            }
        };
        using var reader = ExcelDocument.OpenDataReader(workbook, options);
        ConverterRow row = Assert.Single(reader.RowsAs<ConverterRow>());
        Assert.Equal(42, row.Id);
        Assert.Equal(1.25, row.Serial);
        Assert.Equal(2, calls);
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xls, 2147483648d)]
    [InlineData(ExcelFileFormat.Xls, 1e30)]
    [InlineData(ExcelFileFormat.Xls, null)]
    [InlineData(ExcelFileFormat.Xlsb, 2147483648d)]
    [InlineData(ExcelFileFormat.Xlsb, 1e30)]
    [InlineData(ExcelFileFormat.Xlsb, null)]
    public void NumericOverflowAndBlankFieldsKeepModelConstructionPropertyErrorsAndRedaction(ExcelFileFormat format, double? value) {
        byte[] workbook = CreateWorkbook(format, ["Id", "Name"], [[value, "present"]]);
        using var reader = ExcelDocument.OpenDataReader(workbook,
            new ExcelReadOptions { MappingErrorValuePolicy = DataMappingErrorValuePolicy.Redact });
        ConstructorRow.Calls = 0;
        DataMappingException error = Assert.Throws<DataMappingException>(() => reader.RowsAs<ConstructorRow>().ToArray());
        Assert.Contains("ConstructorRow.Id", error.Message);
        if (value.HasValue) Assert.DoesNotContain(value.Value.ToString("R", CultureInfo.InvariantCulture), error.Message);
        Assert.Equal(1, ConstructorRow.Calls);
        Assert.False(reader.IsClosed);
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xls)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void NativeMappingDoesNotRetryOrWrapAPropertySetterFailure(ExcelFileFormat format) {
        byte[] workbook = CreateWorkbook(format, ["Id"], [[42]]);
        using var reader = ExcelDocument.OpenDataReader(workbook);
        ThrowingSetterRow.Calls = 0;
        Assert.Same(ThrowingSetterRow.Failure,
            Assert.Throws<FormatException>(() => reader.RowsAs<ThrowingSetterRow>().ToArray()));
        Assert.Equal(1, ThrowingSetterRow.Calls);
        Assert.False(reader.IsClosed);
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xls)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void ErrorCellsRetainTheirTextDuringAutomaticMapping(ExcelFileFormat format) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Name");
        sheet.CellValue(2, 1, "error");
        Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference == "A2");
        cell.DataType = CellValues.Error;
        cell.CellValue = new CellValue("#VALUE!");
        byte[] workbook = document.ToBytes(format);
        using var reader = ExcelDocument.OpenDataReader(workbook);
        Assert.Equal("#VALUE!", Assert.Single(reader.RowsAs<TextRow>()).Name);
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xls)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void InvalidDateSerialKeepsGenericConstructionAndFormatSpecificFailure(ExcelFileFormat format) {
        byte[] workbook = CreateWorkbook(format, ["Date"], [[3000000d]], styledSerialColumn: 1);
        using var reader = ExcelDocument.OpenDataReader(workbook);
        DateConstructorRow.Calls = 0;
        if (format == ExcelFileFormat.Xlsb) {
            Assert.Throws<InvalidCastException>(() => reader.RowsAs<DateConstructorRow>().ToArray());
        } else {
            // XLS treats an unrepresentable styled serial as a number before mapping.
            Assert.Throws<DataMappingException>(() => reader.RowsAs<DateConstructorRow>().ToArray());
        }
        Assert.Equal(1, DateConstructorRow.Calls);
        Assert.False(reader.IsClosed);
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xls)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void ElapsedTimeStyleKeepsItsOaEpochInA1904Workbook(ExcelFileFormat format) {
        byte[] workbook = CreateWorkbook(format, ["Date", "Serial"], [[1.25, 1.25]],
            ExcelDateSystem.NineteenFour, styledSerialColumn: 1, serialFormat: "[h]:mm:ss");
        using var reader = ExcelDocument.OpenDataReader(workbook);
        DateAndSerialRow row = Assert.Single(reader.RowsAs<DateAndSerialRow>());
        Assert.Equal(DateTime.FromOADate(1.25), row.Date);
        Assert.Equal(DateTimeKind.Unspecified, row.Date.Kind);
        Assert.Equal(1.25, row.Serial);
    }

    private static byte[] CreateWorkbook(ExcelFileFormat format, string[] headers, object?[][] rows,
        ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred, int styledSerialColumn = 0,
        string serialFormat = "yyyy-mm-dd hh:mm:ss") {
        using ExcelDocument document = ExcelDocument.Create();
        document.DateSystem = dateSystem;
        ExcelSheet sheet = document.AddWorksheet("Data");
        for (int column = 0; column < headers.Length; column++) sheet.CellValue(1, column + 1, headers[column]);
        for (int row = 0; row < rows.Length; row++) {
            for (int column = 0; column < rows[row].Length; column++) {
                if (rows[row][column] != null) sheet.CellValue(row + 2, column + 1, rows[row][column]);
            }
            if (styledSerialColumn > 0) sheet.CellAt(row + 2, styledSerialColumn).SetNumberFormat(serialFormat);
        }
        return document.ToBytes(format);
    }

    private static void AssertNativeRow(ExcelWorkbookDataReader oracle, IFullRow row) {
        Assert.True(oracle.Read());
        object[] values = new object[oracle.FieldCount];
        Assert.Equal(14, oracle.GetValues(values));
        Assert.Equal(Convert.ToString(values[0], CultureInfo.InvariantCulture), row.Name);
        Assert.Equal(Convert.ToBoolean(values[1], CultureInfo.InvariantCulture), row.Flag);
        Assert.Equal(Convert.ToByte(values[2], CultureInfo.InvariantCulture), row.Byte);
        Assert.Equal(Convert.ToInt16(values[3], CultureInfo.InvariantCulture), row.Short);
        Assert.Equal(Convert.ToInt32(values[4], CultureInfo.InvariantCulture), row.Id);
        Assert.Equal(Convert.ToInt64(values[5], CultureInfo.InvariantCulture), row.Long);
        Assert.Equal(Convert.ToSingle(values[6], CultureInfo.InvariantCulture), row.Single);
        Assert.Equal(BitConverter.DoubleToInt64Bits(Convert.ToDouble(values[7], CultureInfo.InvariantCulture)), BitConverter.DoubleToInt64Bits(row.Double));
        Assert.Equal(Convert.ToDecimal(values[8], CultureInfo.InvariantCulture), row.Amount);
        DateTime date = Assert.IsType<DateTime>(values[9]);
        Assert.Equal(date, row.Date);
        Assert.Equal(date.Kind, row.Date.Kind);
        Assert.Equal(Convert.ToString(values[10], CultureInfo.InvariantCulture), row.OtherText);
        Assert.Equal(Convert.ToBoolean(values[11], CultureInfo.InvariantCulture), row.OtherFlag);
        Assert.Equal(Convert.ToInt32(values[12], CultureInfo.InvariantCulture), row.OtherId);
        Assert.Equal(Convert.ToDouble(values[13], CultureInfo.InvariantCulture), row.OtherValue);
    }

    public interface IFullRow {
        string? Name { get; }
        bool Flag { get; }
        byte Byte { get; }
        short Short { get; }
        int Id { get; }
        long Long { get; }
        float Single { get; }
        double Double { get; }
        decimal Amount { get; }
        DateTime Date { get; }
        string? OtherText { get; }
        bool OtherFlag { get; }
        int OtherId { get; }
        double OtherValue { get; }
    }
    public sealed class FullClassRow : IFullRow {
        public string? Name { get; set; }
        public bool Flag { get; set; }
        public byte Byte { get; set; }
        public short Short { get; set; }
        [DisplayName("Order Id")] public int Id { get; set; }
        public long Long { get; set; }
        public float Single { get; set; }
        public double Double { get; set; }
        public decimal Amount { get; set; }
        public DateTime Date { get; set; }
        public string? OtherText { get; set; }
        public bool OtherFlag { get; set; }
        public int OtherId { get; set; }
        public double OtherValue { get; set; }
    }
    public struct FullValueRow : IFullRow {
        public string? Name { get; set; }
        public bool Flag { get; set; }
        public byte Byte { get; set; }
        public short Short { get; set; }
        [DisplayName("Order Id")] public int Id { get; set; }
        public long Long { get; set; }
        public float Single { get; set; }
        public double Double { get; set; }
        public decimal Amount { get; set; }
        public DateTime Date { get; set; }
        public string? OtherText { get; set; }
        public bool OtherFlag { get; set; }
        public int OtherId { get; set; }
        public double OtherValue { get; set; }
    }
    public sealed class FallbackRow {
        public string? Text { get; set; }
        public float Single { get; set; }
        public int Id { get; set; }
        public bool Flag { get; set; }
        public DateTime Date { get; set; }
        public double Serial { get; set; }
        public string? Name { get; set; } = "initialized";
        public int? Optional { get; set; } = 99;
    }
    public sealed class ConverterRow { public int Id { get; set; } public double Serial { get; set; } }
    public sealed class DateAndSerialRow { public DateTime Date { get; set; } public double Serial { get; set; } }
    public sealed class ConstructorRow {
        internal static int Calls;
        public ConstructorRow() { Calls++; }
        public int Id { get; set; }
    }
    public sealed class DateConstructorRow {
        internal static int Calls;
        public DateConstructorRow() { Calls++; }
        public DateTime Date { get; set; }
    }
    public sealed class ThrowingSetterRow {
        internal static readonly FormatException Failure = new("property setter failure");
        internal static int Calls;
        public int Id { get => 0; set { Calls++; throw Failure; } }
    }
    public sealed class TextRow { public string? Name { get; set; } }
}
