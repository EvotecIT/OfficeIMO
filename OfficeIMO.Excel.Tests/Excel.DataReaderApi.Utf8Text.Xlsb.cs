#if NET8_0_OR_GREATER
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Data;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed class XlsbUtf8TextTests {
    [Theory]
    [InlineData("Schema", true)]
    [InlineData("Schema", false)]
    [InlineData("Converter", true)]
    [InlineData("Converter", false)]
    public void BorrowedXlsbText_LiveProjectionChangesKeepCreationTimeEligibilityAndCopyFallback(string projection, bool initiallyProjected) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Text");
        sheet.CellValue(2, 1, "first 🐢");
        sheet.CellValue(3, 1, "following 🐢");
        var options = new ExcelReadOptions { SchemaSampleRows = 1 };
        SetProjection(initiallyProjected);
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(Save(document, true), options);
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        Assert.True(reader.Read());
        AssertLiveProjection(reader, "first 🐢", !initiallyProjected);

        SetProjection(!initiallyProjected);
        AssertLiveProjection(reader, "first 🐢", false);
        Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(-1, out _));
        Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(1, out _));

        SetProjection(initiallyProjected);
        AssertLiveProjection(reader, "first 🐢", !initiallyProjected);
        Assert.True(reader.Read());
        AssertLiveProjection(reader, "following 🐢", !initiallyProjected);
        SetProjection(false);
        AssertLiveProjection(reader, "following 🐢", !initiallyProjected);
        Assert.False(reader.Read());
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));

        void SetProjection(bool enabled) {
            options.InferSchema = projection == "Schema" && enabled;
            options.CellValueConverter = projection == "Converter" && enabled ? static _ => default : null;
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BorrowedXlsbText_PreservesCanonicalStringsAndOrdinaryTypedValues(bool sharedStrings) {
        string[] values = ["Żółw 🐢", "", " A & <β>\r\n B ", "123", "Alpha", "Alpha", "alpha"];
        using var data = new DataSet();
        var table = data.Tables.Add("Data");
        table.Columns.Add("Text", typeof(string));
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Enabled", typeof(bool));
        table.Columns.Add("Date", typeof(DateTime));
        table.Columns.Add("Missing", typeof(string));
        DateTime date = new(2026, 10, 8);
        for (int index = 0; index < values.Length; index++) {
            table.Rows.Add(values[index], index + 1, index % 2 == 0, date, DBNull.Value);
        }
        using ExcelDocument document = ExcelDocument.Create();
        document.InsertDataSet(data, createTables: false, includeAutoFilter: false);
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(Save(document, sharedStrings));
        foreach (string expected in values) {
            Assert.True(reader.Read());
            AssertBorrowed(reader, 0, expected);
            Assert.False(reader.IsDBNull(0));
            Assert.True(reader.GetInt32(1) > 0);
            Assert.IsType<bool>(reader.GetValue(2));
            Assert.Equal(date, reader.GetDateTime(3));
            Assert.True(reader.IsDBNull(4));
            for (int ordinal = 1; ordinal < reader.FieldCount; ordinal++) {
                Assert.False(reader.TryGetUtf8Text(ordinal, out ReadOnlySpan<byte> declined));
                Assert.True(declined.IsEmpty);
            }
            Assert.Throws<InvalidCastException>(() => reader.GetBytes(0, 0, null, 0, 0));
        }
        Assert.False(reader.Read());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BorrowedXlsbText_SameRowCollisionsAndLargeValuesKeepEarlierSpansValid(bool sharedStrings) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        for (int column = 1; column <= 257; column++) {
            sheet.CellValue(1, column, "Column" + column);
            sheet.CellValue(2, column, "value-" + column);
        }
        sheet.CellValue(1, 258, "Empty");
        sheet.CellValue(2, 258, string.Empty);
        string large = new string('λ', 3000) + " 🐢";
        sheet.CellValue(1, 259, "Large");
        sheet.CellValue(2, 259, large);
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(Save(document, sharedStrings));
        Assert.Equal(259, reader.FieldCount);
        Assert.True(reader.Read());
        Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> first));
        Assert.True(reader.TryGetUtf8Text(256, out ReadOnlySpan<byte> collision));
        Assert.True(reader.TryGetUtf8Text(257, out ReadOnlySpan<byte> empty));
        Assert.True(empty.IsEmpty);
        Assert.True(reader.TryGetUtf8Text(258, out ReadOnlySpan<byte> oversized));
        AssertBorrowed(reader, 258, large);
        Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> restored));
        Assert.Equal("value-1", Encoding.UTF8.GetString(first));
        Assert.Equal("value-257", Encoding.UTF8.GetString(collision));
        Assert.Equal(large, Encoding.UTF8.GetString(oversized));
        Assert.Equal("value-1", Encoding.UTF8.GetString(restored));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BorrowedXlsbText_ClearsEligibilityAcrossFormulaMissingAndTextRows(bool sharedStrings) {
        using var data = new DataSet();
        var table = data.Tables.Add("Data");
        table.Columns.Add("Text", typeof(string));
        table.Columns.Add("Id", typeof(int));
        table.Rows.Add("first", 1);
        table.Rows.Add("cached", 2);
        table.Rows.Add(DBNull.Value, 3);
        table.Rows.Add("last", 4);
        using ExcelDocument document = ExcelDocument.Create();
        document.InsertDataSet(data, createTables: false, includeAutoFilter: false);
        ExcelSheet sheet = document.Sheets[0];
        sheet.CellFormula(3, 1, "\"cached\"");
        Cell formula = sheet.WorksheetPart!.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A3");
        formula.DataType = CellValues.String;
        formula.CellValue = new CellValue("cached");
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(Save(document, sharedStrings));
        Assert.True(reader.Read());
        AssertBorrowed(reader, 0, "first");
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out _));
        Assert.Equal("cached", reader.GetString(0));
        AssertCopy(reader, 0, "cached");
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out _));
        Assert.True(reader.IsDBNull(0));
        Assert.Throws<InvalidCastException>(() => reader.GetUtf8Bytes(0, 0, null, 0, 0));
        Assert.True(reader.Read());
        AssertBorrowed(reader, 0, "last");
        Assert.False(reader.Read());
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
    }

    [Theory]
    [InlineData("Schema")]
    [InlineData("Converter")]
    [InlineData("UnhandledConverter")]
    public void BorrowedXlsbText_DeclinesSchemaReplayAndConverters(string projection) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Text");
        sheet.CellValue(2, 1, "sample");
        sheet.CellValue(3, 1, "following");
        var options = new ExcelReadOptions {
            InferSchema = projection == "Schema",
            SchemaSampleRows = 1,
            CellValueConverter = projection switch {
                "Converter" => static _ => new ExcelCellValue("converted"),
                "UnhandledConverter" => static _ => default,
                _ => null
            }
        };
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(Save(document, true), options);
        foreach (string original in new[] { "sample", "following" }) {
            Assert.True(reader.Read());
            Assert.False(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
            Assert.True(text.IsEmpty);
            string expected = projection == "Converter" ? "converted" : original;
            Assert.Equal(expected, reader.GetString(0));
            AssertCopy(reader, 0, expected);
        }
    }

    [Fact]
    public void BorrowedXlsbText_LeavesConvertedBinaryGetBytesAvailable() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Text");
        sheet.CellValue(2, 1, "original");
        byte[] binary = [0, 255, 42];
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(Save(document, true),
            new ExcelReadOptions { CellValueConverter = _ => new ExcelCellValue(binary) });
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out _));
        byte[] destination = new byte[binary.Length];
        Assert.Equal(binary.Length, reader.GetBytes(0, 0, destination, 0, destination.Length));
        Assert.Equal(binary, destination);
    }

    [Fact]
    public async Task BorrowedXlsbText_FollowsAsyncCursorResultAndDisposalContracts() {
        using ExcelDocument document = ExcelDocument.Create();
        foreach (string name in new[] { "First", "Second" }) {
            ExcelSheet sheet = document.AddWorksheet(name);
            sheet.CellValue(1, 1, "Text");
            sheet.CellValue(2, 1, name);
        }
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(Save(document, true));
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        Assert.True(await reader.ReadAsync());
        AssertBorrowed(reader, 0, "First");
        Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(-1, out _));
        Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(1, out _));
        Assert.True(await reader.NextResultAsync());
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        Assert.True(await reader.ReadAsync());
        AssertBorrowed(reader, 0, "Second");
        Assert.False(await reader.ReadAsync());
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        reader.Dispose();
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
    }

    [Fact]
    public void BorrowedXlsbText_UsesIndependentExcelProducedSharedStrings() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "XlsbCorpus", "excel-generated", "basic-values-formula.xlsb");
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path,
            new ExcelReadOptions { HasHeaderRow = false });
        Assert.True(reader.Read());
        AssertBorrowed(reader, 0, "Name");
        AssertBorrowed(reader, 1, "Amount");
        Assert.True(reader.Read());
        AssertBorrowed(reader, 0, "Alpha");
        Assert.False(reader.TryGetUtf8Text(1, out _));
        Assert.Equal(42, reader.GetInt32(1));
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(1, out _));
        Assert.Equal(50, reader.GetInt32(1));
    }

    [Fact]
    public void BorrowedXlsbText_UsesCanonicalRichTextCellValue() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Text");
        sheet.CellValue(2, 1, " rich Żółw 🐢 \r\n");
        using var package = new MemoryStream();
        byte[] original = Save(document, false);
        package.Write(original);
        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry part = archive.GetEntry("xl/worksheets/sheet1.bin")!;
            IReadOnlyList<XlsbRecord> records;
            using (Stream input = part.Open()) records = XlsbRecordReader.ReadAll(input);
            part.Delete();
            using Stream output = archive.CreateEntry("xl/worksheets/sheet1.bin").Open();
            foreach (XlsbRecord record in records) {
                if (record.Type != 6) {
                    XlsbRecordWriter.Write(output, record.Type, record.Data);
                    continue;
                }
                byte[] rich = new byte[record.Size + 1];
                Array.Copy(record.Data, 0, rich, 0, 8);
                Array.Copy(record.Data, 8, rich, 9, record.Size - 8);
                XlsbRecordWriter.Write(output, 62, rich);
            }
        }
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package.ToArray());
        Assert.True(reader.Read());
        AssertBorrowed(reader, 0, " rich Żółw 🐢 \r\n");
    }

    private static byte[] Save(ExcelDocument document, bool sharedStrings) =>
        document.ToBytes(ExcelFileFormat.Xlsb, new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings });

    private static void AssertLiveProjection(ExcelWorkbookDataReader reader, string expected, bool canBorrow) {
        Assert.Equal(expected, reader.GetString(0));
        Assert.False(reader.IsDBNull(0));
        bool borrowed = false;
        byte[]? borrowedBytes = null;
        Exception? borrowError = Record.Exception(() => {
            borrowed = reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text);
            if (borrowed) borrowedBytes = text.ToArray();
        });
        Exception? copyError = Record.Exception(() => AssertCopy(reader, 0, expected));
        Assert.True(borrowError == null && copyError == null,
            $"Borrow: {borrowError?.GetType().Name ?? "success"}; UTF8 copy: {copyError?.GetType().Name ?? "success"}");
        Assert.Equal(canBorrow, borrowed);
        if (canBorrow) Assert.Equal(Encoding.UTF8.GetBytes(expected), borrowedBytes);
    }

    private static void AssertBorrowed(ExcelWorkbookDataReader reader, int ordinal, string expected) {
        Assert.True(reader.TryGetUtf8Text(ordinal, out ReadOnlySpan<byte> text));
        Assert.Equal(expected, Encoding.UTF8.GetString(text));
        Assert.Equal(expected, reader.GetString(ordinal));
        Assert.True(((IDataRecord)reader).TryGetUtf8Text(ordinal, out text));
        Assert.Equal(expected, Encoding.UTF8.GetString(text));
        AssertCopy(reader, ordinal, expected);
    }

    private static void AssertCopy(ExcelWorkbookDataReader reader, int ordinal, string expected) {
        byte[] utf8 = Encoding.UTF8.GetBytes(expected);
        Assert.Equal(utf8.Length, reader.GetUtf8Bytes(ordinal, 0, null, 0, 0));
        byte[] destination = new byte[utf8.Length];
        Assert.Equal(utf8.Length, reader.GetUtf8Bytes(ordinal, 0, destination, 0, destination.Length));
        Assert.Equal(utf8, destination);
    }
}
#endif
