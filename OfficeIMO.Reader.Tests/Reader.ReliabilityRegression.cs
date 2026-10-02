using OfficeIMO.Excel;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using System.Globalization;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderReliabilityRegressionTests {
    [Theory]
    [InlineData("folder")]
    [InlineData("documents")]
    [InlineData("detailed")]
    [InlineData("path-file")]
    [InlineData("path-folder")]
    public void FolderAndDetailedRoutesApplyProcessorsExactlyOnce(string route) {
        string directory = Path.Combine(Path.GetTempPath(), "reader-parity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "original.txt");
            File.WriteAllText(path, "original");
            int calls = 0;
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
                .AddProcessor(new DelegateOfficeDocumentProcessor("replace", (document, _) => {
                    calls++;
                    foreach (var chunk in document.Chunks) chunk.Text = "processed";
                    return document;
                })).Build();
            IEnumerable<ReaderChunk> chunks = route switch {
                "folder" => reader.ReadFolder(directory),
                "documents" => reader.ReadFolderDocuments(directory).SelectMany(document => document.Chunks),
                "detailed" => reader.ReadFolderDetailed(directory).Chunks,
                "path-file" => reader.ReadPathDocumentsDetailed(path).Documents.SelectMany(document => document.Chunks),
                _ => reader.ReadPathDocumentsDetailed(directory).Documents.SelectMany(document => document.Chunks)
            };
            Assert.Equal("processed", Assert.Single(chunks).Text);
            Assert.Equal(1, calls);
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnicodeChunkBoundariesPreserveIndependentEncoding(bool rtf) {
        string text = new string('a', 255) + "\U0001F600b";
        var reader = rtf ? new OfficeDocumentReaderBuilder().AddRtfHandler().Build()
            : new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        byte[] bytes = rtf
            ? Encoding.ASCII.GetBytes("{\\rtf1\\ansi\\uc1 " + new string('a', 255) + "\\u-10179?\\u-8704?b}")
            : Encoding.UTF8.GetBytes(text);
        var chunks = reader.Read(bytes, rtf ? "unicode.rtf" : "unicode.txt", new ReaderOptions { MaxChars = 256 }).ToArray();
        var strict = new UTF8Encoding(false, true);
        Assert.Equal(text, string.Concat(chunks.Select(chunk => strict.GetString(strict.GetBytes(chunk.Text)))));
        if (rtf) Assert.All(chunks, chunk => strict.GetBytes(chunk.Markdown!));
    }

    [Fact]
    public void RtfSplitParagraphDoesNotDuplicatePlainText() {
        string text = new string('a', 1000);
        var reader = new OfficeDocumentReaderBuilder().AddRtfHandler().Build();
        var chunks = reader.Read(Encoding.ASCII.GetBytes("{\\rtf1\\ansi " + text + "}"), "long.rtf",
            new ReaderOptions { MaxChars = 256 }).ToArray();
        Assert.Equal(text, string.Concat(chunks.Select(chunk => chunk.Text)));
        Assert.All(chunks, chunk => Assert.InRange(chunk.Text.Length, 1, 256));
        Assert.All(chunks, chunk => Assert.Equal(chunk.Text, chunk.Markdown));
    }

    [Fact]
    public void WordTableMarkdownRetainsCompleteCellsBeyondChunkTarget() {
        using var stream = new MemoryStream();
        using (var document = WordDocument.Create(stream)) {
            var table = document.AddTable(2, 1);
            table.RepeatAsHeaderRowAtTheTopOfEachPage = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Header";
            table.Rows[1].Cells[0].Paragraphs[0].Text = new string('a', 1000) + "TAIL_SENTINEL";
            document.Save();
        }
        var result = new OfficeDocumentReaderBuilder().AddWordHandler().Build()
            .ReadDocument(stream.ToArray(), "table.docx", new ReaderOptions { MaxChars = 256 });
        Assert.Contains("TAIL_SENTINEL", result.Markdown, StringComparison.Ordinal);
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
            warning => warning.Contains("preserved", StringComparison.Ordinal));
    }

    [Fact]
    public void ChunkHashesDistinguishFieldBoundaries() {
        ReaderChunk Read(string text, string markdown) => Assert.Single(new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "framing", Kind = ReaderInputKind.Text, Extensions = new[] { ".framing" },
                ReadStream = (_, _, _, _) => new[] { new ReaderChunk { Kind = ReaderInputKind.Text, Text = text, Markdown = markdown } }
            }).Build().Read(new byte[] { 1 }, "same.framing"));
        Assert.NotEqual(Read("alpha|beta", "gamma").ChunkHash, Read("alpha", "beta|gamma").ChunkHash);
        Assert.Equal(Read("alpha", "beta").ChunkHash, Read("alpha", "beta").ChunkHash);
    }

    [Fact]
    public void ExcelRegistrationSnapshotsMutableReadOptions() {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorksheet("Data");
            sheet.Cell(1, 1, "Name");
            sheet.Cell(2, 1, "alpha");
            document.Save();
        }
        var policy = new ExcelReadOptions();
        var reader = new OfficeDocumentReaderBuilder().AddExcelHandler(new ReaderExcelOptions { ReadOptions = policy }).Build();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        policy.CancellationToken = cancellation.Token;
        Assert.NotEmpty(reader.ReadDocument(stream.ToArray(), "sample.xlsx").Chunks);
    }

    [Fact]
    public void ExcelOptionsCloneIsolatesCultureAndExecutionThresholds() {
        var options = new ExcelReadOptions { Culture = new CultureInfo("pl-PL") };
        var clone = options.Clone();
        options.Culture.NumberFormat.NumberDecimalSeparator = "!";
        options.Execution.OperationThresholds["ReadRange"] = 1;
        Assert.Equal(",", clone.Culture.NumberFormat.NumberDecimalSeparator);
        Assert.Equal(100_000, clone.Execution.OperationThresholds["ReadRange"]);
    }
}
