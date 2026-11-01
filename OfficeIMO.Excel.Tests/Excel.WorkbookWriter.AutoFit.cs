using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class StreamingAutoFitTests {
    [Fact]
    public void Writer_AutoFitsSinglePassRowsAndWritesColumnsBeforeData() {
        using var output = new MemoryStream();
        var rows = new OnePassRows();
        using (var writer = ExcelDocument.CreateWriter(output, new ExcelWriterOptions {
            AutoFit = true, RequireStreaming = true, UseSharedStrings = false, CreateTable = true
        }, leaveOpen: true)) {
            var result = writer.WriteTable("Data", rows, new[] { "Label", "Value" },
                (row, value) => row.Write(value).Write(12345.67m));
            Assert.Equal(3, result.RowCount);
            Assert.Equal("A1:B4", result.Range);
        }
        Assert.Equal(1, rows.Enumerations);
        using var package = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var sheet = Assert.Single(package.WorkbookPart!.WorksheetParts).Worksheet;
        var children = sheet!.ChildElements.ToList();
        Assert.True(children.FindIndex(child => child is Columns) < children.FindIndex(child => child is SheetData));
        var columns = sheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
        Assert.Equal(2, columns.Length);
        Assert.All(columns, column => Assert.True(column.Width!.Value > 1));
        Assert.True(columns[0].Width!.Value > columns[1].Width!.Value);
        using var reader = ExcelDocument.OpenReader(new MemoryStream(output.ToArray()));
        foreach (string value in new[] { "short", "a much longer streamed value", "last" }) {
            Assert.True(reader.Read());
            Assert.Equal(value, reader.GetString(0));
            Assert.Equal(12345.67, reader.GetDouble(1), 8);
        }
        Assert.False(reader.Read());
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred, true)]
    [InlineData(ExcelDateSystem.NineteenFour, false)]
    public void Writer_AutoFitUsesCanonicalFormatsFontsAndTableHeaders(ExcelDateSystem dateSystem, bool filter) {
        using var output = new MemoryStream();
        var options = new ExcelWriterOptions {
            AutoFit = true, UseSharedStrings = false, CreateTable = true, DateSystem = dateSystem, IncludeAutoFilter = filter,
            DefaultRowStyle = "NormalBody",
            Styles = new Dictionary<string, ExcelStyleDefinition> {
                ["NormalBody"] = new(),
                ["Money"] = new() { NumberFormat = "\"USD \"#,##0.0000\" dollars\"" },
                ["Percent"] = new() { NumberFormat = "0.0000%" },
                ["Large"] = new() { FontSize = OfficeIMO.OfficeLength.Points(28), Bold = true }
            },
            ColumnStyles = new Dictionary<int, string> { [2] = "Money" }
        };
        using (var writer = ExcelDocument.CreateWriter(output, options, leaveOpen: true)) {
            writer.WriteTable("Styled", new[] { 1, 2 }, new[] { "Percent", "Money", "Date", "Duration", "Small", "Large", "Lines", "Boolean", "Blank" }, (row, index) => {
                row.SetRowStyle(null); // Exercise column inheritance as well as explicit cell and row styles.
                row.SetNextCellStyle("Percent").Write(0.123456);
                row.Write(12.5m);
                row.SetNextCellStyle("Large").Write(new DateTime(2026, 11, index, 13, 45, 0));
                row.Write(TimeSpan.FromHours(125));
                row.Write("matched text");
                row.SetNextCellStyle("Large").Write("matched text");
#if NET8_0_OR_GREATER
                row.WriteUtf8(System.Text.Encoding.UTF8.GetBytes("one & <two>\r\nlonger Unicode 日本語"));
#else
                row.Write("one & <two>\r\nlonger Unicode 日本語");
#endif
                row.Write(false).WriteBlank();
            });
            // A different schema also exercises the workbook's deferred multi-sheet metadata.
            writer.WriteTable("RowStyle", new[] { 1 }, new[] { "Text", "Other" }, (row, _) => row.SetRowStyle("Large").Write("row font").WriteBlank());
        }
        byte[] bytes = output.ToArray();
        using (var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false)) {
            Assert.Empty(new OpenXmlValidator().Validate(package));
            var first = package.WorkbookPart!.WorksheetParts.OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal).First();
            var widths = first.Worksheet!.GetFirstChild<Columns>()!.Elements<Column>().Select(column => column.Width!.Value).ToArray();
            Assert.True(widths[1] > widths[0]); // Literal currency suffix must affect the displayed width.
            Assert.True(widths[5] > widths[4]); // Same text, larger and bold font.
            Assert.True(widths[2] > 20); // The formatted date uses its resolved temporal variant and font.
            Assert.Equal(dateSystem == ExcelDateSystem.NineteenFour, package.WorkbookPart!.Workbook!.WorkbookProperties?.Date1904?.Value ?? false);
        }
        // The ordinary native worksheet AutoFit is the canonical contract, including table header adornments.
        using var editable = ExcelDocument.Load(new MemoryStream(bytes));
        foreach (var sheet in editable.Sheets) {
            double[] streamed = sheet.GetColumnDefinitions().Select(column => column.Width!.Value.CharacterCount!.Value).ToArray();
            sheet.AutoFitColumns();
            double[] ordinary = sheet.GetColumnDefinitions().Select(column => column.Width!.Value.CharacterCount!.Value).ToArray();
            Assert.Equal(streamed.Length, ordinary.Length);
            for (int column = 0; column < ordinary.Length; column++) Assert.InRange(Math.Abs(streamed[column] - ordinary[column]), 0, 0.001);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Writer_AutoFitsAsyncReusedRowsIntoNonSeekableOutput(bool headers) {
        using var output = new NonSeekableReadWriteBuffer(Array.Empty<byte>());
        int enumerations = 0;
        bool disposed = false;
        using (var writer = ExcelDocument.CreateWriter(output, new ExcelWriterOptions {
            AutoFit = true, UseSharedStrings = false, IncludeHeaders = headers, CreateTable = headers,
            IncludeCellReferences = false, RequireStreaming = true
        }, leaveOpen: true)) {
            var result = await writer.WriteTable("Async", Rows(), new[] { "Number", "Text" }, (row, value) => row.Write(value[0]).Write(value[1]));
            Assert.Equal(100, result.RowCount);
            Assert.Equal(headers ? "A1:B101" : "A1:B100", result.Range);
        }
        Assert.True(disposed);
        Assert.Equal(1, enumerations);
        byte[] bytes = output.ToArray();
        output.WriteByte(0); // A real write proves leaveOpen after finalizing the non-seekable destination.
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var sheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
        Assert.NotNull(sheet!.GetFirstChild<Columns>());
        var values = sheet.GetFirstChild<SheetData>()!.Elements<Row>().Skip(headers ? 1 : 0).ToArray();
        Assert.Equal(100, values.Length);
        for (int index = 0; index < values.Length; index++) {
            var cells = values[index].Elements<Cell>().ToArray();
            Assert.Null(cells[0].CellReference);
            Assert.Equal(index.ToString(System.Globalization.CultureInfo.InvariantCulture), cells[0].CellValue!.Text);
            Assert.Equal("row " + index, cells[1].InlineString!.InnerText);
        }

        async IAsyncEnumerable<object?[]> Rows() {
            if (++enumerations != 1) throw new InvalidOperationException("Source was enumerated twice.");
            var buffer = new object?[2];
            try {
                for (int index = 0; index < 100; index++) {
                    buffer[0] = index;
                    buffer[1] = "row " + index;
                    yield return buffer;
                    await Task.Yield();
                }
            } finally { disposed = true; }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Writer_AutoFitFailureOrCancellationRemovesSpoolAndRetainsDestination(bool cancel) {
        string directory = Path.Combine(Path.GetTempPath(), "AutoFit-Atomic-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string destination = Path.Combine(directory, "book.xlsx");
        File.WriteAllText(destination, "original destination");
        var before = new HashSet<string>(Directory.GetFiles(Path.GetTempPath(), "OfficeIMO-Excel-AutoFit-*.xml"));
        string? spool = null;
        bool disposed = false;
        using var cancellation = new CancellationTokenSource();
        try {
            using (var writer = ExcelDocument.CreateWriter(destination, new ExcelWriterOptions { AutoFit = true, UseSharedStrings = false })) {
                void Write() => writer.WriteTable("Atomic", Rows(), new[] { "Text" }, (row, value) => {
                    spool ??= Assert.Single(Directory.GetFiles(Path.GetTempPath(), "OfficeIMO-Excel-AutoFit-*.xml"), path => !before.Contains(path));
                    row.Write(value);
                    if (cancel) cancellation.Cancel();
                    else throw new InvalidDataException("row callback failed");
                }, cancellation.Token);
                if (cancel) Assert.ThrowsAny<OperationCanceledException>(Write);
                else Assert.Throws<InvalidDataException>(Write);
                Assert.NotNull(spool);
                Assert.False(File.Exists(spool));
            }
            Assert.True(disposed);
            Assert.Equal("original destination", File.ReadAllText(destination));
            Assert.Equal(new[] { destination }, Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, true); }

        IEnumerable<string> Rows() {
            try { yield return "already serialized"; yield return "later"; }
            finally { disposed = true; }
        }
    }

    [Fact]
    public void Writer_AutoFitPublishesOnlyAfterDisposeAndCleansSuccessfulSpool() {
        string directory = Path.Combine(Path.GetTempPath(), "AutoFit-Success-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string destination = Path.Combine(directory, "book.xlsx");
        File.WriteAllText(destination, "original destination");
        var before = new HashSet<string>(Directory.GetFiles(Path.GetTempPath(), "OfficeIMO-Excel-AutoFit-*.xml"));
        string? spool = null;
        try {
            using (var writer = ExcelDocument.CreateWriter(destination, new ExcelWriterOptions { AutoFit = true, UseSharedStrings = false })) {
                writer.WriteTable("Data", new[] { "fits" }, new[] { "Text" }, (row, value) => {
                    spool = Assert.Single(Directory.GetFiles(Path.GetTempPath(), "OfficeIMO-Excel-AutoFit-*.xml"), path => !before.Contains(path));
                    row.Write(value);
                });
                Assert.Equal("original destination", File.ReadAllText(destination));
                Assert.False(File.Exists(spool));
            }
            using var package = SpreadsheetDocument.Open(destination, false);
            Assert.Empty(new OpenXmlValidator().Validate(package));
            Assert.Single(Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void Writer_AutoFitSpoolsDuringEnumerationAndClampsStyledWidths() {
        using var output = new MemoryStream();
        var before = new HashSet<string>(Directory.GetFiles(Path.GetTempPath(), "OfficeIMO-Excel-AutoFit-*.xml"));
        bool observedSpoolData = false;
        using (var writer = ExcelDocument.CreateWriter(output, new ExcelWriterOptions {
            AutoFit = true, UseSharedStrings = false, CreateTable = false,
            Styles = new Dictionary<string, ExcelStyleDefinition> { ["Large"] = new() { FontSize = OfficeIMO.OfficeLength.Points(28) } },
            DefaultRowStyle = "Large"
        }, leaveOpen: true)) {
            writer.WriteTable("Long", Enumerable.Range(0, 128), new[] { "Text" }, (row, index) => {
                if (index == 127) {
                    string spool = Assert.Single(Directory.GetFiles(Path.GetTempPath(), "OfficeIMO-Excel-AutoFit-*.xml"), path => !before.Contains(path));
                    observedSpoolData = new FileInfo(spool).Length > 0;
                }
                row.Write(new string('W', 1024));
            });
        }
        Assert.True(observedSpoolData);
        using var package = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.Equal(255D, package.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<Columns>()!.Elements<Column>().Single().Width!.Value);
    }

    [Fact]
    public void Writer_AutoFitOutputFailureRemovesSpool() {
        using var output = new OfficeIMO.Tests.TestStreams.ThrowAfterBytesWriteStream(0);
        var before = new HashSet<string>(Directory.GetFiles(Path.GetTempPath(), "OfficeIMO-Excel-AutoFit-*.xml"));
        string? spool = null;
        var writer = ExcelDocument.CreateWriter(output, new ExcelWriterOptions { AutoFit = true, UseSharedStrings = false }, leaveOpen: true);
        try {
            Assert.ThrowsAny<IOException>(() => writer.WriteTable("Fail", new[] { "value" }, new[] { "Text" }, (row, value) => {
                spool = Assert.Single(Directory.GetFiles(Path.GetTempPath(), "OfficeIMO-Excel-AutoFit-*.xml"), path => !before.Contains(path));
                row.Write(value);
            }));
            Assert.NotNull(spool);
            Assert.False(File.Exists(spool));
        } finally {
            try { writer.Dispose(); } catch (IOException) { /* The failed destination can also reject the ZIP trailer. */ }
        }
    }

    private sealed class OnePassRows : IEnumerable<string> {
        internal int Enumerations { get; private set; }
        public IEnumerator<string> GetEnumerator() {
            if (++Enumerations != 1) throw new InvalidOperationException("Source was enumerated twice.");
            return ((IEnumerable<string>)new[] { "short", "a much longer streamed value", "last" }).GetEnumerator();
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
