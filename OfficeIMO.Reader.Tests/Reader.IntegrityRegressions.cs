using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Yaml;
using OfficeIMO.Reader.Zip;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderIntegrityRegressionsTests {
    [Theory]
    [InlineData(256)]
    [InlineData(8000)]
    public void TextMarkdownPreservesContentAcrossChunkSizes(int maximum) {
        string text = new string('a', 300) + "\n\n" + new string('b', 300);
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(text), "sample.txt",
            new ReaderOptions { ComputeHashes = false, MaxChars = maximum });
        Assert.Equal(text, result.Markdown);
    }

    [Fact]
    public void PageMarkdownPreservesTableCellsInSourceOrderWithoutDuplicatePreview() {
        var table = new ReaderTable {
            Columns = new[] { "Name", "Count" },
            Rows = new[] { (IReadOnlyList<string>)new[] { "alpha", "2" } },
            Location = new ReaderLocation { BlockIndex = 1, BlockAnchor = "table-1" }
        };
        var page = new OfficeDocumentPage {
            Number = 1,
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "before", Text = "Before", Location = new ReaderLocation { BlockIndex = 0 } },
                new OfficeDocumentBlock { Id = "table-1", Kind = "table", Text = "alpha preview", Location = table.Location },
                new OfficeDocumentBlock { Id = "after", Text = "After", Location = new ReaderLocation { BlockIndex = 2 } }
            },
            Tables = new[] { table }
        };
        string markdown = new OfficeDocumentReadResult { Pages = new[] { page } }.GetPageMarkdown()[0].Markdown;
        Assert.Contains("| alpha | 2 |", markdown);
        Assert.DoesNotContain("alpha preview", markdown);
        Assert.True(markdown.IndexOf("Before", StringComparison.Ordinal) < markdown.IndexOf("alpha", StringComparison.Ordinal));
        Assert.True(markdown.IndexOf("alpha", StringComparison.Ordinal) < markdown.IndexOf("After", StringComparison.Ordinal));
        Assert.Equal("alpha", table.Rows[0][0]);
    }

    [Fact]
    public void ZipPreservesDuplicateAndNormalizedEntryPayloadsAndIdentities() {
        using var source = new MemoryStream();
        using (var zip = new ZipArchive(source, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (var item in new[] { ("same.txt", "FIRST"), ("same.txt", "SECOND"), ("./note.txt", "DOT"), ("folder\\note.txt", "BACKSLASH") }) {
                using var writer = new StreamWriter(zip.CreateEntry(item.Item1).Open());
                writer.Write(item.Item2);
            }
        }
        source.Position = 0;
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(source, "members.zip", new ReaderOptions { ComputeHashes = false });
        Assert.Equal(new[] { "BACKSLASH", "DOT", "FIRST", "SECOND" }, result.Chunks.Select(chunk => chunk.Text).OrderBy(value => value, StringComparer.Ordinal));
        Assert.Equal(4, result.Chunks.Select(chunk => chunk.SourceId).Distinct().Count());
        Assert.Equal(4, result.NestedDocuments.Select(item => item.Path).Distinct().Count());
        Assert.Empty(result.Diagnostics);
    }

    [Theory]
    [InlineData("plain")]
    [InlineData("quoted")]
    [InlineData("literal")]
    public void StructuredReadersPreserveLongScalarValues(string style) {
        string value = new string('x', 3000);
        string yaml = style == "quoted" ? "value: '" + value + "'"
            : style == "literal" ? "value: |-\n  " + value : "value: " + value;
        var reader = new OfficeDocumentReaderBuilder().AddXmlHandler().AddYamlHandler().Build();
        foreach (var input in new[] { ("scalar.yaml", yaml), ("scalar.xml", "<root value='" + value + "'>" + value + "</root>") }) {
            OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(input.Item2), input.Item1,
                new ReaderOptions { ComputeHashes = false });
            Assert.Contains(result.Tables.SelectMany(table => table.Rows).SelectMany(row => row), cell => cell == value);
            Assert.All(result.Tables, table => Assert.False(table.Truncated));
            Assert.Empty(result.Diagnostics);
        }
    }

    [Theory]
    [InlineData("depth")]
    [InlineData("nodes")]
    [InlineData("scalar")]
    public void XmlRejectsStructuralLimitsBeforeProjection(string limit) {
        string xml = limit == "depth" ? string.Concat(Enumerable.Repeat("<r>", 64)) + string.Concat(Enumerable.Repeat("</r>", 64))
            : limit == "nodes" ? "<r>" + string.Concat(Enumerable.Repeat("<x/>", 100)) + "</r>"
            : "<r>" + new string('x', 5000) + "</r>";
        var options = new XmlReadOptions {
            MaxDepth = limit == "depth" ? 8 : 128,
            MaxNodes = limit == "nodes" ? 4 : 200_000,
            MaxScalarLength = limit == "scalar" ? 8 : 1_048_576
        };
        var reader = new OfficeDocumentReaderBuilder().AddXmlHandler(options).Build();
        ReaderResourceLimitException error = Assert.Throws<ReaderResourceLimitException>(() =>
            reader.ReadDocument(Encoding.UTF8.GetBytes(xml), "bounded.xml", new ReaderOptions { ComputeHashes = false }));
        Assert.Equal(limit == "depth" ? "MaxDepth" : limit == "nodes" ? "MaxNodes" : "MaxScalarLength", error.LimitName);
    }

    [Fact]
    public void CsvCancellationInterruptsAQuotedRecordBeforeReadingItsRemainder() {
        using var cancellation = new CancellationTokenSource();
        using var source = new CancelOnReadStream(Encoding.UTF8.GetBytes("\"" + new string('x', 200_000) + "\"\n"), cancellation);
        Assert.Throws<OperationCanceledException>(() => CsvReaderAdapter.Read(source, "quoted.csv",
            new ReaderOptions { ComputeHashes = false }, cancellationToken: cancellation.Token).ToArray());
        Assert.True(source.Position < source.Length, "Cancellation must reach the parser before it consumes the complete quoted record.");
    }

    [Theory]
    [InlineData("document")]
    [InlineData("chunks")]
    [InlineData("incremental")]
    [InlineData("document-async")]
    [InlineData("chunks-async")]
    public async Task PathReadsRejectAChangedSourceInsteadOfMislabelingItsHash(string surface) {
        string path = Path.Combine(Path.GetTempPath(), "reader-race-" + Guid.NewGuid().ToString("N") + ".rdr");
        OfficeDocumentReadResult ReadAndChange(string input) {
            string text = File.ReadAllText(input);
            File.WriteAllText(input, "NEW");
            return new OfficeDocumentReadResult { Kind = ReaderInputKind.Text,
                Chunks = new[] { new ReaderChunk { Id = "body", Kind = ReaderInputKind.Text, Text = text } } };
        }
        var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "changing-source", Kind = ReaderInputKind.Text, Extensions = new[] { ".rdr" },
            SupportsIncrementalPath = true,
            ReadPath = (input, options, token) => ReadAndChange(input).Chunks,
            ReadDocumentPath = (input, options, token) => ReadAndChange(input),
            ReadDocumentPathAsync = (input, options, token) => Task.FromResult(ReadAndChange(input))
        }).Build();
        try {
            File.WriteAllText(path, "OLD");
            IOException error = await Assert.ThrowsAsync<IOException>(async () => {
                if (surface == "document-async") await reader.ReadDocumentAsync(path);
                else if (surface == "chunks-async") await reader.ReadAsync(path);
                else if (surface == "document") reader.ReadDocument(path);
                else if (surface == "chunks") reader.Read(path).ToArray();
                else reader.EnumerateChunks(path).ToArray();
            });
            Assert.Contains("Source changed", error.Message);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmailItemEnumerationEnforcesAggregateBudgetsAcrossItems(bool hashes) {
        string directory = Path.Combine(Path.GetTempPath(), "reader-mail-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            for (int i = 0; i < 3; i++) File.WriteAllText(Path.Combine(directory, i + ".eml"),
                "From: a@example.test\r\nTo: b@example.test\r\nSubject: Budget\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nBody\r\n");
            var reader = new OfficeDocumentReaderBuilder().AddEmailStoreHandler().Build();
            ReaderOptions Options(long budget) => new ReaderOptions { ComputeHashes = hashes,
                ResourceLimits = new ReaderResourceLimits { MaxChunks = budget } };
            ReaderResourceLimitException zero = Assert.Throws<ReaderResourceLimitException>(() =>
                reader.ReadEmailStoreItems(directory, Options(0)).ToArray());
            Assert.Equal("MaxChunks", zero.LimitName);
            ReaderEmailStoreItemResult first = reader.ReadEmailStoreItems(directory, new ReaderOptions { ComputeHashes = hashes }).First();
            using var iterator = reader.ReadEmailStoreItems(directory, Options(first.Chunks.Count)).GetEnumerator();
            Assert.True(iterator.MoveNext());
            Assert.Throws<ReaderResourceLimitException>(() => iterator.MoveNext());
            Assert.Throws<ReaderResourceLimitException>(() => reader.ReadEmailStoreItem(directory, first.Reference.Id, Options(0)));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    private sealed class CancelOnReadStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal CancelOnReadStream(byte[] bytes, CancellationTokenSource cancellation) : base(bytes, writable: false) { _cancellation = cancellation; }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            _cancellation.Cancel();
            return read;
        }
    }
}
