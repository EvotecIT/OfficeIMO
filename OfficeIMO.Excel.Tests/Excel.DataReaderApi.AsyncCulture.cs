using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task AsyncOpeningDefaultCultureRetainsBorrowedTextAndTypedValues(bool seekable) {
        byte[] workbook = CreateAsyncCultureWorkbook(AsyncCultureRows);
        using var source = new AsyncOpeningReadStream(workbook, seekable);
        using (ExcelWorkbookDataReader reader = await ExcelDocument.OpenDataReaderAsync(source)) {
            Assert.Equal(new[] { "Name", "Id", "Date", "Value" }, Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
            Assert.True(await reader.ReadAsync());
            AssertAsyncCultureRow(reader, "Alpha", 42, 45292.25, 1.25);
            Assert.True(await reader.ReadAsync());
            AssertAsyncCultureRow(reader, "Zażółć 😀", 43, 45293, 2.5);
            Assert.False(await reader.ReadAsync());
            Assert.True(source.AsyncReads > 0);
        }
        Assert.True(source.CanRead);
        if (seekable) Assert.Equal(0, source.Position);
    }

    [Fact]
    public async Task AsyncOpeningSnapshotsCustomInvariantNamedCultureBeforeAwaitingSource() {
        var culture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        culture.NumberFormat.NumberDecimalSeparator = "|";
        var options = new ExcelReadOptions { Culture = culture };
        var gate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var source = new AsyncOpeningReadStream(
            CreateAsyncCultureWorkbook(AsyncCultureRows.Replace("Alpha", "12|5")), seekable: true) {
            ReadGate = gate.Task
        };
        Task<ExcelWorkbookDataReader> opening = ExcelDocument.OpenDataReaderAsync(source, options);
        Assert.False(opening.IsCompleted);
        culture.NumberFormat.NumberDecimalSeparator = ":";
        gate.SetResult(true);
        using ExcelWorkbookDataReader reader = await opening;
        Assert.True(await reader.ReadAsync());
        Assert.Equal("12|5", reader.GetString(0));
        Assert.Equal(12.5, reader.GetDouble(0));
        Assert.Equal(42, reader.GetInt32(1));
        Assert.Equal(":", options.Culture.NumberFormat.NumberDecimalSeparator);
        Assert.True(source.CanRead);
        Assert.Equal(0, source.Position);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public async Task AsyncOpeningDefaultCultureValidatesLateXmlAndStylesBeforeDelivery(bool seekable, bool invalidStyle) {
        string rows = invalidStyle
            ? AsyncCultureRows.Replace("<c><v>2.5</v></c>", "<c s=\"99999\"><v>2.5</v></c>")
            : AsyncCultureRows;
        byte[] workbook = CreateAsyncCultureWorkbook(rows, invalidStyle ? string.Empty : "<trailing>");
        using var source = new AsyncOpeningReadStream(workbook, seekable);
        if (invalidStyle) {
            await Assert.ThrowsAsync<InvalidDataException>(() => ExcelDocument.OpenDataReaderAsync(source));
        } else {
            await Assert.ThrowsAnyAsync<XmlException>(() => ExcelDocument.OpenDataReaderAsync(source));
        }
        Assert.True(source.CanRead);
        if (seekable) Assert.Equal(0, source.Position);
    }

    private static void AssertAsyncCultureRow(ExcelWorkbookDataReader reader, string text, int id, double date, double number) {
#if NET8_0_OR_GREATER
        Assert.True(reader.TryGetUtf8Text(0, out var borrowed));
        Assert.Equal(text, Encoding.UTF8.GetString(borrowed));
#endif
        Assert.Equal(text, reader.GetString(0));
        Assert.Equal(id, reader.GetInt32(1));
        Assert.Equal(DateTime.FromOADate(date), reader.GetDateTime(2));
        Assert.Equal(number, reader.GetDouble(3));
    }

    private const string AsyncCultureRows = "<row><c t=\"inlineStr\"><is><t>Alpha</t></is></c><c><v>42</v></c><c s=\"1\"><v>45292.25</v></c><c><v>1.25</v></c></row>"
        + "<row><c t=\"inlineStr\"><is><t>Zażółć 😀</t></is></c><c><v>43</v></c><c s=\"1\"><v>45293</v></c><c><v>2.5</v></c></row>";

    private static byte[] CreateAsyncCultureWorkbook(string rows, string trailingXml = "") {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddWorksheet("Data").CellValue(1, 1, "Name");
        byte[] original = document.ToBytes();
        using var buffer = new MemoryStream();
        buffer.Write(original, 0, original.Length);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true)) {
            string worksheet = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:D3\"/><sheetData>"
                + "<row><c t=\"inlineStr\"><is><t>Name</t></is></c><c t=\"inlineStr\"><is><t>Id</t></is></c><c t=\"inlineStr\"><is><t>Date</t></is></c><c t=\"inlineStr\"><is><t>Value</t></is></c></row>"
                + rows + "</sheetData></worksheet>" + trailingXml;
            ReplaceAsyncCultureEntry(archive, "xl/worksheets/sheet1.xml", worksheet);
            ReplaceAsyncCultureEntry(archive, "xl/styles.xml", "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs></styleSheet>");
        }
        return buffer.ToArray();
    }

    private static void ReplaceAsyncCultureEntry(ZipArchive archive, string name, string xml) {
        archive.GetEntry(name)!.Delete();
        using Stream destination = archive.CreateEntry(name).Open();
        byte[] bytes = Encoding.UTF8.GetBytes(xml);
        destination.Write(bytes, 0, bytes.Length);
    }
}
