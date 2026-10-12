#if NET8_0_OR_GREATER
using System.Data;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed partial class ExcelUtf8TextTests {
    [Theory]
    [InlineData("<si><t>Żółw 🐢</t></si>", "Żółw 🐢")]
    [InlineData("<si><t/></si>", "")]
    [InlineData("<si><t>A&amp;B &#xD;tail</t></si>", "A&B \rtail")]
    [InlineData("<si><t>A\r\nB</t></si>", "A\nB")]
    [InlineData("<si><r><t>A&amp;B</t></r><r><t xml:space=\"preserve\"> text</t></r></si>", "A&B text")]
    [InlineData("<si><t xml:space=\"preserve\"> \t spaced </t></si>", " \t spaced ")]
    [InlineData("<si><t>45200</t></si>", "45200")]
    public void BorrowedSharedText_UsesCanonicalTextForSpansStringsAndCopies(string item, string expected) {
        string path = CreateWorkbook("<c r=\"A2\" t=\"s\"><v>0</v></c>", sharedStringItems: item);
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
            Assert.Equal(expected, Decode(text));
            Assert.Equal(expected, reader.GetString(0));
            Assert.False(reader.IsDBNull(0));
            Assert.True(((IDataRecord)reader).TryGetUtf8Text(0, out text));
            Assert.Equal(expected, Decode(text));

            byte[] bytes = new byte[Encoding.UTF8.GetByteCount(expected)];
            Assert.Equal(bytes.Length, reader.GetUtf8Bytes(0, 0, null, 0, 0));
            Assert.Equal(bytes.Length, reader.GetUtf8Bytes(0, 0, bytes, 0, bytes.Length));
            Assert.Equal(Encoding.UTF8.GetBytes(expected), bytes);
            Assert.Throws<NotSupportedException>(() => reader.GetBytes(0, 0, null, 0, 0));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void BorrowedSharedText_RemainsValidWhenAnotherFieldUpdatesTheCache() {
        string items = "<si><t>first 🐢</t></si>" +
            string.Concat(Enumerable.Range(1, 255).Select(index => "<si><t>Item" + index + "</t></si>")) +
            "<si><t>last Ż</t></si>";
        string path = CreateWorkbook("<c r=\"A2\" t=\"s\"><v>0</v></c><c r=\"B2\" t=\"s\"><v>256</v></c>",
            columns: 2, sharedStringItems: items);
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> first));
            Assert.True(reader.TryGetUtf8Text(1, out ReadOnlySpan<byte> last));
            Assert.Equal("last Ż", Decode(last));
            Assert.Equal("first 🐢", Decode(first));
            Assert.Equal("first 🐢", reader.GetString(0));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void BorrowedSharedText_LargeItemsPreserveTheFullUtf8Value() {
        string expected = new string('Ż', 3000) + "🐢";
        string path = CreateWorkbook("<c r=\"A2\" t=\"s\"><v>0</v></c>",
            sharedStringItems: "<si><t>" + expected + "</t></si>");
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            AssertBorrowedText(reader, expected);
            Assert.Equal(expected, reader.GetString(0));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("Converter")]
    [InlineData("Schema")]
    [InlineData("Utf16")]
    [InlineData("Formula")]
    public void BorrowedSharedText_PreservesUnsupportedProjectionFallbacks(string projection) {
        string cell = projection == "Formula"
            ? "<c r=\"A2\" t=\"s\"><f>\"original\"</f><v>0</v></c>"
            : "<c r=\"A2\" t=\"s\"><v>0</v></c>";
        string path = CreateWorkbook(cell, utf16: projection == "Utf16",
            sharedStringItems: "<si><t>original</t></si>");
        try {
            var options = new ExcelReadOptions {
                InferSchema = projection == "Schema",
                SchemaSampleRows = 1,
                CellValueConverter = projection == "Converter" ? static _ => new ExcelCellValue("converted") : null
            };
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path, options);
            Assert.True(reader.Read());
            Assert.False(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
            Assert.True(text.IsEmpty);
            Assert.Equal(projection == "Converter" ? "converted" : "original", reader.GetString(0));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task BorrowedSharedText_FollowsRowResultAndDisposalBoundaries() {
        string path = CreateWorkbook("<c r=\"A2\" t=\"s\"><v>0</v></c>",
            extraRows: "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>1</v></c></row>", secondSheet: true,
            sharedStringItems: "<si><t>first Ż</t></si><si><t>second 🐢</t></si>");
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
            Assert.True(await reader.ReadAsync());
            AssertBorrowedText(reader, "first Ż");
            Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(-1, out _));
            Assert.Throws<IndexOutOfRangeException>(() => reader.TryGetUtf8Text(1, out _));
            Assert.True(await reader.ReadAsync());
            AssertBorrowedText(reader, "second 🐢");
            Assert.False(await reader.ReadAsync());
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
            Assert.True(await reader.NextResultAsync());
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
            Assert.True(await reader.ReadAsync());
            AssertBorrowedText(reader, "first Ż");
            reader.Dispose();
            Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
        } finally {
            File.Delete(path);
        }
    }
}
#endif
