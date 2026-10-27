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
            writer.WriteTable("Events", first);
            await writer.WriteTable("Later", LaterRows());
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
        using (var writer = ExcelDocument.CreateWriter(output, leaveOpen: true)) writer.WriteTable("Events", reader);
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
