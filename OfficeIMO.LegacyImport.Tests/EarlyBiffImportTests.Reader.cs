using OfficeIMO.Excel;
using System.Globalization;

namespace OfficeIMO.LegacyImport.Tests;

public sealed partial class EarlyBiffImportTests {
    [Fact]
    public void ForwardReaderEnforcesColumnBufferAndSchemaSampleAdmission() {
        byte[] bytes = BuildReaderScalarWorksheet(4, "bounded");
        using var columnSource = new MemoryStream(bytes);
        using ExcelWorkbookDataReader columns = ExcelDocument.OpenReader(columnSource,
            new ExcelReaderOptions { MaxDataReaderColumns = 5 });
        Assert.Throws<OfficeInspectionResourceLimitException>(() => columns.FieldCount);

        using var bufferSource = new MemoryStream(bytes);
        using ExcelWorkbookDataReader buffer = ExcelDocument.OpenReader(bufferSource,
            new ExcelReaderOptions { MaxDataReaderBufferedCells = 5 });
        Assert.Throws<OfficeInspectionResourceLimitException>(() => buffer.FieldCount);

        using var schemaSource = new MemoryStream(bytes);
        using ExcelWorkbookDataReader schema = ExcelDocument.OpenReader(schemaSource,
            new ExcelReaderOptions { InferSchema = true, SchemaSampleRows = 2, MaxDataReaderSchemaSampleRows = 1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => schema.FieldCount);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ForwardReaderUsesStoredCellBoundsRatherThanRowAndColumnGeometry(bool hasCell) {
        using var bytes = new MemoryStream();
        using (var output = new BinaryWriter(bytes, Encoding.UTF8, leaveOpen: true)) {
            Record(output, BofCode(4), record => WriteBof(record, 4, 0x10));
            Record(output, 0x0208, record => {
                record.Write((ushort)300);
                record.Write((ushort)0);
                record.Write((ushort)0);
                record.Write((ushort)400);
                record.Write(0U);
                record.Write(32U);
            });
            Record(output, 0x007D, record => {
                record.Write((ushort)200);
                record.Write((ushort)255);
                record.Write((ushort)2560);
                record.Write((ushort)0);
                record.Write((ushort)1);
            });
            if (hasCell) Record(output, 0x0203, record => {
                WriteReaderCoordinate(record, 4, 4, 2);
                record.Write(3d);
            });
            Record(output, 0x000A, _ => { });
        }
        bytes.Position = 0;
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(bytes, new ExcelReaderOptions { HasHeaderRow = false });
        Assert.Equal(hasCell ? 1 : 0, reader.FieldCount);
        Assert.Equal(hasCell, reader.HasRows);
        if (hasCell) {
            Assert.True(reader.Read());
            Assert.Equal(3d, reader.GetDouble(0));
        }
        Assert.False(reader.Read());
        Assert.False(reader.NextResult());
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void ForwardReaderPreservesHeadersSparseRowsAndOriginalCachedScalarTypes(int generation) {
        using var source = new MemoryStream(BuildReaderScalarWorksheet(generation, "Café"));
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(source);
        Assert.Equal(new[] { "Sheet1" }, reader.SheetNames);
        Assert.Equal(6, reader.FieldCount);
        Assert.Equal(new[] { "Name", "Amount", "Flag", "Error", "Formula", "Blank" },
            Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
        Assert.Equal(1, reader.GetOrdinal("Amount"));
        Assert.All(Enumerable.Range(0, reader.FieldCount), column => Assert.Equal(typeof(object), reader.GetFieldType(column)));
        var schema = reader.GetSchemaTable()!;
        Assert.Equal(6, schema.Rows.Count);
        Assert.Equal(typeof(object), schema.Rows[1]["DataType"]);
        Assert.True(reader.HasRows);

        Assert.True(reader.Read());
        Assert.All(Enumerable.Range(0, reader.FieldCount), column => Assert.True(reader.IsDBNull(column)));
        Assert.True(reader.Read());
        Assert.Equal("Café", reader.GetString(0));
        Assert.IsType<double>(reader.GetValue(1));
        Assert.Equal(42.25d, reader.GetDouble(1));
        Assert.IsType<bool>(reader.GetValue(2));
        Assert.True(reader.GetBoolean(2));
        Assert.Equal("#DIV/0!", reader.GetString(3));
        Assert.IsType<double>(reader.GetValue(4));
        Assert.Equal(77d, reader.GetDouble(4));
        Assert.True(reader.IsDBNull(5));
        Assert.True(reader.Read());
        Assert.All(Enumerable.Range(0, reader.FieldCount), column => Assert.True(reader.IsDBNull(column)));
        Assert.True(reader.Read());
        Assert.All(Enumerable.Range(0, reader.FieldCount), column => Assert.True(reader.IsDBNull(column)));
        Assert.False(reader.Read());
        Assert.False(reader.NextResult());
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(true, false)]
    public void ForwardReaderUsesInheritedDateFormatsWithoutTreatingLocalNumericFormatSlotsAsBuiltIns(bool use1904, bool parseDates) {
        CultureInfo culture = CultureInfo.GetCultureInfo("pl-PL");
        using var source = new MemoryStream(BuildReaderDateWorksheet(use1904));
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(source, new ExcelReaderOptions {
            HasHeaderRow = false, InferSchema = true, Culture = culture, TreatDatesUsingNumberFormat = parseDates
        });
        Assert.Equal(3, reader.FieldCount);
        Assert.Equal(parseDates ? typeof(DateTime) : typeof(double), reader.GetFieldType(0));
        Assert.Equal(typeof(double), reader.GetFieldType(1));
        Assert.Equal(parseDates ? typeof(DateTime) : typeof(double), reader.GetFieldType(2));
        Assert.True(reader.Read());
        Assert.Equal(60d, reader.GetDouble(0));
        Assert.Equal(60m, reader.GetDecimal(0));
        Assert.Equal(12.5d, reader.GetValue(1));
        Assert.Equal("12,5", reader.GetString(1));
        Assert.Equal(1.5d, reader.GetDouble(2));
        if (parseDates) {
            DateTime expectedCalendar = use1904 ? new DateTime(1904, 3, 1) : new DateTime(1900, 2, 28);
            var expectedElapsedCarrier = new DateTime(1899, 12, 31, 12, 0, 0);
            Assert.Equal(expectedCalendar, reader.GetValue(0));
            Assert.Equal(expectedCalendar, reader.GetDateTime(0));
            Assert.Equal(expectedCalendar.ToString(culture), reader.GetString(0));
            Assert.Equal(expectedElapsedCarrier, reader.GetDateTime(2));
        } else {
            Assert.Equal(60d, reader.GetValue(0));
            Assert.Equal(1.5d, reader.GetValue(2));
        }
        Assert.False(reader.Read());
    }

    [Fact]
    public void ForwardReaderDefersInvalidDateMaterializationUntilADateValueIsRequested() {
        using var source = new MemoryStream(BuildReaderDateWorksheet(use1904: false, includeInvalidDate: true));
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(source, new ExcelReaderOptions { HasHeaderRow = false });
        Assert.True(reader.Read());
        Assert.Equal(1_000_000_000d, reader.GetDouble(3));
        Assert.Throws<ArgumentException>(() => reader.GetValue(3));
        Assert.Equal(1_000_000_000d, reader.GetDouble(3));
    }

    [Fact]
    public void ForwardReaderRetainsBiff4WorkbookResultOrderAndSelectedWorkbookIndex() {
        byte[] bytes = BuildReaderWorkbook();
        using var source = new MemoryStream(bytes);
        using (ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(source)) {
            Assert.Equal(new[] { "First", "Second" }, reader.SheetNames);
            Assert.Equal("First", reader.CurrentSheetName);
            Assert.Equal(0, reader.CurrentSheetIndex);
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.True(reader.Read());
            Assert.Equal("First value", reader.GetString(0));
            Assert.True(reader.NextResult());
            Assert.Equal("Second", reader.CurrentSheetName);
            Assert.Equal(1, reader.CurrentSheetIndex);
            Assert.Equal(1, reader.CurrentResultIndex);
            Assert.Equal("Name", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.True(reader.Read());
            Assert.Equal("Second value", reader.GetString(0));
            Assert.False(reader.NextResult());
        }
        using var selectedSource = new MemoryStream(bytes);
        using ExcelWorkbookDataReader selected = ExcelDocument.OpenReader(selectedSource, new ExcelReaderOptions { SheetName = "second" });
        Assert.Equal(new[] { "Second" }, selected.SheetNames);
        Assert.Equal(1, selected.CurrentSheetIndex);
        Assert.Equal(0, selected.CurrentResultIndex);
        Assert.True(selected.Read());
        Assert.True(selected.Read());
        Assert.Equal("Second value", selected.GetString(0));
        Assert.False(selected.NextResult());
    }

    private static byte[] BuildReaderScalarWorksheet(int generation, string text) {
        using var bytes = new MemoryStream();
        using var output = new BinaryWriter(bytes);
        Record(output, BofCode(generation), record => WriteBof(record, generation, 0x10));
        Record(output, 0x0042, record => record.Write((ushort)1252));
        string[] headers = { " Name ", "Amount", "Flag", "Error", "Formula", "Blank" };
        for (ushort column = 0; column < headers.Length; column++) {
            Record(output, generation == 2 ? (ushort)4 : (ushort)0x204, record => {
                WriteReaderCoordinate(record, generation, 0, column);
                WriteString(record, Encoding.ASCII.GetBytes(headers[column]), shortLength: generation == 2);
            });
        }
        Record(output, generation == 2 ? (ushort)4 : (ushort)0x204, record => {
            WriteReaderCoordinate(record, generation, 2, 0);
            WriteString(record, Encoding.Latin1.GetBytes(text), shortLength: generation == 2);
        });
        Record(output, generation == 2 ? (ushort)3 : (ushort)0x203, record => {
            WriteReaderCoordinate(record, generation, 2, 1);
            record.Write(42.25d);
        });
        for (ushort column = 2; column <= 3; column++) {
            Record(output, generation == 2 ? (ushort)5 : (ushort)0x205, record => {
                WriteReaderCoordinate(record, generation, 2, column);
                record.Write((byte)(column == 2 ? 1 : 7));
                record.Write((byte)(column == 2 ? 0 : 1));
            });
        }
        Record(output, generation switch { 2 => (ushort)6, 3 => (ushort)0x206, _ => (ushort)0x406 }, record => {
            WriteReaderCoordinate(record, generation, 2, 4);
            record.Write(77d);
            if (generation == 2) { record.Write((byte)0); record.Write((byte)3); }
            else { record.Write((ushort)0); record.Write((ushort)3); }
            record.Write(new byte[] { 0x1E, 1, 0 });
        });
        Record(output, generation == 2 ? (ushort)1 : (ushort)0x201, record => WriteReaderCoordinate(record, generation, 4, 5));
        Record(output, 0x000A, _ => { });
        return bytes.ToArray();
    }

    private static byte[] BuildReaderDateWorksheet(bool use1904, bool includeInvalidDate = false) {
        using var bytes = new MemoryStream();
        using var output = new BinaryWriter(bytes);
        Record(output, BofCode(4), record => WriteBof(record, 4, 0x10));
        Record(output, 0x0022, record => record.Write((ushort)(use1904 ? 1 : 0)));
        for (int index = 0; index <= 14; index++) {
            string format = index switch { 1 => "yyyy-mm-dd", 2 => "[h]:mm", 14 => "0.00", _ => "General" };
            Record(output, 0x041E, record => {
                record.Write((ushort)0);
                WriteString(record, Encoding.ASCII.GetBytes(format), shortLength: true);
            });
        }
        WriteFormat(number: 1, type: 0xFFF4, used: 0);
        WriteFormat(number: 0, type: 0, used: 0);
        WriteFormat(number: 14, type: 0xFFF1, used: 0xFC);
        WriteFormat(number: 2, type: 0xFFF1, used: 0xFC);
        double[] values = includeInvalidDate ? new[] { 60d, 12.5d, 1.5d, 1_000_000_000d } : new[] { 60d, 12.5d, 1.5d };
        for (ushort column = 0; column < values.Length; column++) {
            Record(output, 0x0203, record => {
                WriteReaderCoordinate(record, 4, 0, column, column == 3 ? (ushort)1 : (ushort)(column + 1));
                record.Write(values[column]);
            });
        }
        Record(output, 0x000A, _ => { });
        return bytes.ToArray();

        void WriteFormat(byte number, ushort type, byte used) => Record(output, 0x0443, record => {
            record.Write((byte)0);
            record.Write(number);
            record.Write(type);
            record.Write((byte)0);
            record.Write(used);
            record.Write((ushort)0);
            record.Write(0U);
        });
    }

    private static byte[] BuildReaderWorkbook() {
        using var bytes = new MemoryStream();
        using var output = new BinaryWriter(bytes);
        Record(output, BofCode(4), record => WriteBof(record, 4, 0x0100));
        Record(output, 0x008E, record => record.Write(0));
        string[] names = { "First", "Second" };
        foreach (string name in names) Record(output, 0x0085, record => WriteString(record, Encoding.ASCII.GetBytes(name), shortLength: true));
        long firstSheet = bytes.Position;
        bytes.Position = 14;
        output.Write((int)firstSheet);
        bytes.Position = firstSheet;
        foreach (string name in names) {
            byte[] sheet = BuildReaderScalarWorksheet(4, name + " value");
            Record(output, 0x008F, record => {
                record.Write(sheet.Length);
                WriteString(record, Encoding.ASCII.GetBytes(name), shortLength: true);
            });
            output.Write(sheet);
        }
        Record(output, 0x000A, _ => { });
        return bytes.ToArray();
    }

    private static void WriteReaderCoordinate(BinaryWriter output, int generation, ushort row, ushort column, ushort style = 0) {
        output.Write(row);
        output.Write(column);
        output.Write(style);
        if (generation == 2) output.Write((byte)0);
    }
}
