using System.Collections;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.SharedSource.IO;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public async Task WorkbookWriter_StreamsMultipleTypedTablesAndPreservesTheme() {
        using var output = new MemoryStream();
        var first = new SinglePassRows<WriterEvent>(new[] { new WriterEvent { Id = 7, Name = "<first>", Amount = 12.5m } });
        using (var writer = ExcelDocument.CreateWriter(output, theme: OfficeTheme.Dark, leaveOpen: true)) {
            ExcelDataSetImportResult firstResult = writer.WriteTable("Events", first);
            ExcelDataSetImportResult laterResult = await writer.WriteTable("Later", LaterRows());
            Assert.Equal(("Events", "Events", "A1:C2", 1, 3),
                (firstResult.SheetName, firstResult.TableName, firstResult.Range, firstResult.RowCount, firstResult.ColumnCount));
            Assert.Equal(("Later", "Later", "A1:C3", 2, 3),
                (laterResult.SheetName, laterResult.TableName, laterResult.Range, laterResult.RowCount, laterResult.ColumnCount));
        }
        Assert.Equal(1, first.Enumerations);
        using var document = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var workbook = document.WorkbookPart!;
        Assert.Equal(new[] { "Events", "Later" }, workbook.Workbook.Sheets!.Elements<Sheet>().Select(sheet => sheet.Name!.Value));
        Assert.Equal(new[] { "A1:C2", "A1:C3" }, workbook.WorksheetParts.OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal)
            .Select(part => part.TableDefinitionParts.Single().Table.Reference!.Value));
        Assert.Contains("<first>", workbook.WorksheetParts.First().Worksheet.InnerText);
        Assert.Equal(OfficeTheme.Dark.Name, workbook.ThemePart!.Theme.Name!.Value);
        Assert.Contains(OfficeTheme.Dark.Fonts.Body.Family, workbook.WorkbookStylesPart!.Stylesheet.InnerXml);

        static async IAsyncEnumerable<WriterEvent> LaterRows() {
            yield return new WriterEvent { Id = 8, Name = "next", Amount = 13m };
            await Task.Yield();
            yield return new WriterEvent { Id = 9, Name = "last", Amount = 14m };
        }
    }

    [Fact]
    public void WorkbookWriter_WritesOutputWhileTheSinglePassSourceIsStillActive() {
        using var output = new MemoryStream();
        bool observedOutput = false;
        using (var writer = ExcelDocument.CreateWriter(output, leaveOpen: true)) writer.WriteTable("Events", Rows());
        Assert.True(observedOutput);
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        Assert.Null(archive.GetEntry("xl/sharedStrings.xml"));

        IEnumerable<WriterEvent> Rows() {
            for (int index = 0; index < 64; index++) {
                if (index == 32) observedOutput = output.Length > 0;
                yield return new WriterEvent { Id = index, Name = new string((char)('A' + index % 26), 4096) };
            }
        }
    }

    [Fact]
    public void WorkbookWriter_DataReaderWritesToNonSeekableOutputAndLeavesCallerResourcesOpen() {
        var data = new System.Data.DataTable();
        data.Columns.Add("Id", typeof(int));
        data.Columns.Add("Name", typeof(string));
        data.Rows.Add(42, "streamed");
        using var reader = data.CreateDataReader();
        using var output = new NonSeekableReadWriteBuffer(Array.Empty<byte>());
        using (var writer = ExcelDocument.CreateWriter(output, leaveOpen: true)) {
            ExcelDataSetImportResult result = writer.WriteTable("Events", reader);
            Assert.Equal(("Events", "Events", "A1:B2", 1, 2),
                (result.SheetName, result.TableName, result.Range, result.RowCount, result.ColumnCount));
        }
        Assert.False(reader.IsClosed);
        Assert.True(output.CanWrite);
        using var document = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var table = document.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table;
        Assert.Equal("Events", table.Name!.Value);
        Assert.Equal("A1:B2", table.Reference!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WorkbookWriter_FailureOrCancellationDiscardsTheStagedFile(bool cancel) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-Writer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "events.xlsx");
        File.WriteAllText(path, "original");
        using var cancellation = new CancellationTokenSource();
        bool disposed = false;
        try {
            using (var writer = ExcelDocument.CreateWriter(path)) {
                if (cancel) await Assert.ThrowsAnyAsync<OperationCanceledException>(() => writer.WriteTable("Events", Rows(), cancellation.Token));
                else await Assert.ThrowsAsync<InvalidDataException>(() => writer.WriteTable("Events", Rows(), cancellation.Token));
            }
            Assert.True(disposed);
            Assert.Equal("original", File.ReadAllText(path));
            Assert.Single(Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, true); }

        async IAsyncEnumerable<WriterEvent> Rows([EnumeratorCancellation] CancellationToken token = default) {
            try {
                yield return new WriterEvent { Id = 1 };
                await Task.Yield();
                if (cancel) { cancellation.Cancel(); token.ThrowIfCancellationRequested(); }
                throw new InvalidDataException("Source failed.");
            } finally { disposed = true; }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WorkbookWriter_DataTableResultsMatchTheWrittenRowsWithoutATable(bool headers) {
        var data = new System.Data.DataTable();
        data.Columns.Add("Id", typeof(int));
        data.Columns.Add("Name", typeof(string));
        data.Rows.Add(1, "First");
        data.Rows.Add(2, "Second");
        using var output = new MemoryStream();
        ExcelDataSetImportResult result;
        using (var writer = ExcelDocument.CreateWriter(output, new ExcelWriterOptions {
            IncludeHeaders = headers, CreateTable = false, UseSharedStrings = false
        }, leaveOpen: true)) {
            result = writer.WriteTable("Written", data);
        }
        Assert.Null(result.TableName);
        Assert.Equal("Written", result.SheetName);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(headers ? "A1:B3" : "A1:B2", result.Range);
        using var document = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
        WorksheetPart part = Assert.Single(document.WorkbookPart!.WorksheetParts);
        Assert.Equal(headers ? 3 : 2, part.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Count());
        Assert.Empty(part.TableDefinitionParts);
    }

    [Fact]
    public void WorkbookWriter_DataReaderUsesTheExistingBlankAndDuplicateHeaderRules() {
        var data = new System.Data.DataTable();
        data.Columns.Add(" ", typeof(int));
        data.Columns.Add("Column1", typeof(string));
        data.Rows.Add(42, "value");
        using var reader = data.CreateDataReader();
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(output, leaveOpen: true)) {
            ExcelDataSetImportResult result = writer.WriteTable("Headers", reader);
            Assert.Equal("A1:B2", result.Range);
        }
        using var loaded = ExcelDocument.OpenReader(new MemoryStream(output.ToArray()));
        Assert.Equal("Column1", loaded.GetName(0));
        Assert.Equal("Column1 (2)", loaded.GetName(1));
        Assert.True(loaded.Read());
        Assert.Equal(42D, loaded.GetDouble(0));
        Assert.Equal("value", loaded.GetString(1));
    }

    public sealed class WriterEvent {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public decimal Amount { get; set; }
    }

    private sealed class SinglePassRows<T>(IEnumerable<T> rows) : IEnumerable<T> {
        internal int Enumerations { get; private set; }
        public IEnumerator<T> GetEnumerator() {
            if (++Enumerations != 1) throw new InvalidOperationException("A second source enumeration is not supported.");
            return rows.GetEnumerator();
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
