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
    [InlineData(256, "paragraphs")]
    [InlineData(8000, "paragraphs")]
    [InlineData(256, "leading-spaces")]
    [InlineData(8000, "leading-spaces")]
    [InlineData(256, "leading-newlines")]
    [InlineData(8000, "leading-newlines")]
    [InlineData(256, "whitespace")]
    [InlineData(8000, "whitespace")]
    public void TextMarkdownPreservesContentAcrossChunkSizes(int maximum, string shape) {
        string text = shape == "leading-spaces" ? new string(' ', 300) + "body"
            : shape == "leading-newlines" ? new string('\n', 300) + "body"
            : shape == "whitespace" ? new string(' ', 300)
            : new string('a', 300) + "\n\n" + new string('b', 300);
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

    [Theory]
    [InlineData("path")]
    [InlineData("page")]
    [InlineData("slide")]
    [InlineData("sheet")]
    public void PageMarkdownKeepsTablePreviewsFromOtherSourceContainers(string container) {
        var previewLocation = new ReaderLocation { BlockAnchor = "table-1", Path = "a.doc", Page = 1, Slide = 1, Sheet = "A" };
        var tableLocation = new ReaderLocation { BlockAnchor = "table-1", Path = "a.doc", Page = 1, Slide = 1, Sheet = "A" };
        if (container == "path") tableLocation.Path = "b.doc";
        else if (container == "page") tableLocation.Page = 2;
        else if (container == "slide") tableLocation.Slide = 2;
        else tableLocation.Sheet = "B";
        var page = new OfficeDocumentPage {
            Blocks = new[] { new OfficeDocumentBlock { Kind = "table", Text = "Other source preview", Location = previewLocation } },
            Tables = new[] { new ReaderTable { Columns = new[] { "Value" },
                Rows = new[] { (IReadOnlyList<string>)new[] { "Structured cell" } }, Location = tableLocation } }
        };
        string markdown = new OfficeDocumentReadResult { Pages = new[] { page } }.GetPageMarkdown()[0].Markdown;
        Assert.Contains("Other source preview", markdown);
        Assert.Contains("| Structured cell |", markdown);
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

    [Theory]
    [InlineData("<root xmlns='urn:x' xmlns:p='urn:y' p:attr='v'><p:child/></root>",
        "{urn:x}root[1]|{urn:x}root[1]/@xmlns|{urn:x}root[1]/@xmlns:p|{urn:x}root[1]/@p:attr|{urn:x}root[1]/p:child[1]")]
    [InlineData("<root xmlns:p='urn:x' xmlns:q='urn:x'><q:child q:value='v'/></root>",
        "root[1]|root[1]/@xmlns:p|root[1]/@xmlns:q|root[1]/p:child[1]|root[1]/p:child[1]/@p:value")]
    [InlineData("<root xmlns='urn:x' xmlns:p='urn:x'><p:child/></root>",
        "p:root[1]|p:root[1]/@xmlns|p:root[1]/@xmlns:p|p:root[1]/p:child[1]")]
    [InlineData("<root xmlns:p='urn:x'><child xmlns:p='urn:y' xmlns:q='urn:x'><q:leaf/></child></root>",
        "root[1]|root[1]/@xmlns:p|root[1]/child[1]|root[1]/child[1]/@xmlns:p|root[1]/child[1]/@xmlns:q|root[1]/child[1]/q:leaf[1]")]
    public void XmlPathsPreserveNamespaceDeclarationAndAliasConventions(string xml, string expected) {
        var reader = new OfficeDocumentReaderBuilder().AddXmlHandler().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(xml), "namespaces.xml",
            new ReaderOptions { ComputeHashes = false });
        Assert.Equal(expected.Split('|'), result.Tables.SelectMany(table => table.Rows).Select(row => row[0]));
        Assert.Empty(result.Diagnostics);
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
