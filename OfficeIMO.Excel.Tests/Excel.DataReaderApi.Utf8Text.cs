#if NET8_0_OR_GREATER
using System.Data;
using System.Data.Common;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed partial class ExcelUtf8TextTests {
    [Fact]
    public void Utf8Copy_UsesOrdinaryExcelStringConversionsAndCursorValidation() {
        string path = CreateWorkbook(
            "<c r=\"A2\" t=\"inlineStr\"><is><t>Żółw 🐢</t></is></c>" +
            "<c r=\"B2\"><v>123</v></c>", columns: 2);
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Throws<InvalidOperationException>(() => reader.GetUtf8Bytes(0, 0, null, 0, 0));
            Assert.True(reader.Read());
            byte[] bytes = new byte[32];
            int count = (int)reader.GetUtf8Bytes(0, 0, bytes, 0, bytes.Length);
            Assert.Equal("Żółw 🐢", Encoding.UTF8.GetString(bytes, 0, count));
            count = (int)reader.GetUtf8Bytes(1, 0, bytes, 0, bytes.Length);
            Assert.Equal("123", Encoding.UTF8.GetString(bytes, 0, count));
            Assert.False(reader.Read());
            Assert.Throws<InvalidOperationException>(() => reader.GetUtf8Bytes(0, 0, null, 0, 0));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void BorrowedText_PreservesUnicodeEmptyTextAndOrdinaryGetters() {
        string path = CreateWorkbook(
            "<c r=\"A2\" t=\"inlineStr\"><is><t>Żółw 🐢</t></is></c>" +
            "<c r=\"B2\" t=\"str\"><v/></c>", columns: 2);
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
            Assert.Equal("Żółw 🐢", Decode(text));
            Assert.Equal("Żółw 🐢", reader.GetString(0));
            Assert.True(((IDataRecord)reader).TryGetUtf8Text(0, out text));
            Assert.Equal("Żółw 🐢", Decode(text));
            Assert.True(reader.TryGetUtf8Text(1, out text));
            Assert.True(text.IsEmpty);
            Assert.Equal(string.Empty, reader.GetString(1));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void BorrowedText_UsesTheUsedRangeOrdinalAfterADeclaredBlankColumn() {
        string path = CreateWorkbook(
            "<c r=\"B2\" t=\"str\"><v>selected</v></c>", firstColumn: 2);
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(1, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
            Assert.Equal("selected", Decode(text));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<c r=\"A2\" t=\"str\"><v>A&amp;B</v></c>", "A&B")]
    [InlineData("<c r=\"A2\" t=\"inlineStr\"><is><t>A\r\nB</t></is></c>", "A\nB")]
    [InlineData("<c r=\"A2\" t=\"e\"><v>#DIV/0!</v></c>", "#DIV/0!")]
    [InlineData("<c r=\"A2\" t=\"str\"><f>\"cached\"</f><v>cached</v></c>", "cached")]
    [InlineData("<c r=\"A2\" t=\"inlineStr\"><is><r><t>rich</t></r><r><t> text</t></r></is></c>", "rich text")]
    public void BorrowedText_DeclinesFieldsRequiringDecodingOrProjection(string cell, string expected) {
        string path = CreateWorkbook(cell);
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.False(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
            Assert.True(text.IsEmpty);
            Assert.Equal(expected, reader.GetString(0));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void BorrowedText_DeclinesNumericBooleanAndMissingFields() {
        string path = CreateWorkbook(
            "<c r=\"A2\"><v>42</v></c><c r=\"B2\" t=\"b\"><v>1</v></c>", columns: 3);
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            for (int ordinal = 0; ordinal < reader.FieldCount; ordinal++) {
                Assert.False(reader.TryGetUtf8Text(ordinal, out ReadOnlySpan<byte> text));
                Assert.True(text.IsEmpty);
            }
            Assert.Equal(42, reader.GetInt32(0));
            Assert.True(reader.GetBoolean(1));
            Assert.True(reader.IsDBNull(2));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void BorrowedText_DeclinesCellConvertersAndUtf16Input() {
        foreach (bool utf16 in new[] { false, true }) {
            string path = CreateWorkbook("<c r=\"A2\" t=\"str\"><v>original</v></c>", utf16: utf16);
            try {
                var options = new ExcelReadOptions {
                    CellValueConverter = utf16 ? null : static _ => new ExcelCellValue("converted")
                };
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path, options);
                Assert.True(reader.Read());
                Assert.False(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
                Assert.True(text.IsEmpty);
                Assert.Equal(utf16 ? "original" : "converted", reader.GetString(0));
            } finally {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void BorrowedText_ValidatesCursorAndOrdinalsAndFollowsEnumeratorReads() {
        string path = CreateWorkbook("<c r=\"A2\" t=\"str\"><v>row</v></c>");
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
            System.Collections.IEnumerator rows = reader.GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
            Assert.Equal("row", Decode(text));
            Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(-1, out _));
            Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(reader.FieldCount, out _));
            Assert.False(rows.MoveNext());
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
            reader.Close();
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task BorrowedText_FollowsAsyncReadsAndResultTransitions() {
        string path = CreateWorkbook("<c r=\"A2\" t=\"str\"><v>first</v></c>", secondSheet: true);
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(await reader.ReadAsync());
            AssertBorrowedText(reader, "first");
            Assert.True(await reader.NextResultAsync());
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
            Assert.True(await reader.ReadAsync());
            AssertBorrowedText(reader, "second");
            Assert.False(await reader.ReadAsync());
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
            reader.Dispose();
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task BorrowedText_SchemaReplayNeverBorrowsTheInnerReadersDifferentRow() {
        string path = CreateWorkbook("<c r=\"A2\" t=\"str\"><v>sampled</v></c>",
            extraRows: "<row r=\"3\"><c r=\"A3\" t=\"str\"><v>streamed</v></c></row>");
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path,
                new ExcelReadOptions { InferSchema = true, SchemaSampleRows = 1 });
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
            Assert.True(await reader.ReadAsync());
            Assert.False(reader.TryGetUtf8Text(0, out _));
            Assert.Equal("sampled", reader.GetString(0));
            Assert.True(await reader.ReadAsync());
            Assert.False(reader.TryGetUtf8Text(0, out _));
            Assert.Equal("streamed", reader.GetString(0));
            Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(1, out _));
            Assert.False(await reader.ReadAsync());
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("SpreadsheetDateSerialCorpus", "", "serials-1900.xls")]
    public void BorrowedText_LegacyXlsValidatesStateAndReturnsFalse(string corpus, string producer, string file) {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", corpus, producer, file);
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path,
            new ExcelReadOptions { HasHeaderRow = false });
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        Assert.True(reader.Read());
        Assert.False(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
        Assert.True(text.IsEmpty);
        Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(reader.FieldCount, out _));
        while (reader.Read()) { }
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        reader.Dispose();
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
    }

    [Fact]
    public void BorrowedText_UnsupportedDataRecordReturnsFalseWithoutConverting() {
        using var table = new DataTable();
        table.Columns.Add("Value", typeof(object));
        table.Rows.Add(new NoTextConversion());
        using DbDataReader reader = table.CreateDataReader();
        Assert.True(reader.Read());
        Assert.False(((IDataRecord)reader).TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
        Assert.True(text.IsEmpty);
        Assert.Throws<IndexOutOfRangeException>(() => ((IDataRecord)reader).TryGetUtf8Text(-1, out _));
        Assert.Throws<ArgumentNullException>(() => DataReaderUtf8TextExtensions.TryGetUtf8Text(null!, 0, out _));
    }

    private static string Decode(ReadOnlySpan<byte> text) => Encoding.UTF8.GetString(text.ToArray());

    private static void AssertBorrowedText(ExcelWorkbookDataReader reader, string expected) {
        Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
        Assert.Equal(expected, Decode(text));
    }

    private static string CreateWorkbook(string cells, int columns = 1, bool utf16 = false, string extraRows = "", int firstColumn = 1, bool secondSheet = false, string? sharedStringItems = null) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Utf8Text.{Guid.NewGuid():N}.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            if (secondSheet) {
                document.AddWorksheet("Second").CellValue(1, 1, "Value");
            }
            document.Save();
        }

        string headers = string.Concat(Enumerable.Range(firstColumn, columns).Select(ordinal =>
            $"<c r=\"{(char)('A' + ordinal - 1)}1\" t=\"str\"><v>Column{ordinal}</v></c>"));
        string xml = $"<?xml version=\"1.0\" encoding=\"{(utf16 ? "utf-16" : "utf-8")}\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            $"<dimension ref=\"A1:{(char)('A' + firstColumn + columns - 2)}{(extraRows.Length == 0 ? 2 : 3)}\"/>" +
            $"<sheetData><row r=\"1\">{headers}</row><row r=\"2\">{cells}</row>{extraRows}</sheetData></worksheet>";
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ReplaceWorksheet(archive, "xl/worksheets/sheet1.xml", xml, utf16);
        if (secondSheet) {
            ReplaceWorksheet(archive, "xl/worksheets/sheet2.xml", xml.Replace("first", "second"), utf16);
        }
        if (sharedStringItems != null) {
            ReplaceWorksheet(archive, "xl/sharedStrings.xml",
                "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + sharedStringItems + "</sst>", utf16: false);
        }
        return path;
    }

    private static void ReplaceWorksheet(ZipArchive archive, string entryName, string xml, bool utf16) {
        archive.GetEntry(entryName)!.Delete();
        using Stream output = archive.CreateEntry(entryName, CompressionLevel.Optimal).Open();
        byte[] bytes = utf16 ? Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray()
            : Encoding.UTF8.GetBytes(xml);
        output.Write(bytes, 0, bytes.Length);
    }

    private sealed class NoTextConversion {
        public override string ToString() => throw new InvalidOperationException("A capability probe must not convert this field.");
    }
}
#endif
