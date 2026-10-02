using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Reader.Email;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderResourceLimitTests {
    [Fact]
    public async Task NestedReadInsideProcessorCannotDowngradeAnOperationLimit() {
        var childReader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
            .AddProcessor(new DelegateOfficeDocumentProcessor("nested-read", (document, _) => {
                childReader.ReadDocument(Encoding.UTF8.GetBytes("child"), "child.txt");
                return document;
            })).Build();
        byte[] bytes = Encoding.UTF8.GetBytes("root");
        var options = Limited(new ReaderResourceLimits { MaxChunks = 1 });
        Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(bytes, "root.txt", options));
        await Assert.ThrowsAsync<ReaderResourceLimitException>(() => reader.ReadDocumentAsync(bytes, "root.txt", options));
    }

    [Theory]
    [InlineData(3, false)]
    [InlineData(4, true)]
    public void MaterializedAssetCountsItsBytesOnceAndChecksActualLength(int bytes, bool exceeds) {
        var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "asset", Extensions = new[] { ".asset" },
            ReadDocumentStream = (_, _, _, _) => new OfficeDocumentReadResult {
                Assets = new[] { new OfficeDocumentAsset { Id = "payload", LengthBytes = 3 } }
            }
        }).AddProcessor(new DelegateOfficeDocumentProcessor("materialize", (document, _) => {
            document.Assets[0].PayloadBytes = new byte[bytes];
            return document;
        })).Build();
        using var input = new MemoryStream(new byte[] { 1 });
        var options = Limited(new ReaderResourceLimits { MaxAssetBytes = 3 });
        if (exceeds) Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(input, "input.asset", options));
        else Assert.Equal(3, Assert.Single(reader.ReadDocument(input, "input.asset", options).Assets).PayloadBytes!.Length);
    }

    [Theory]
    [InlineData("ordered")]
    [InlineData("detailed")]
    [InlineData("completed")]
    public async Task BatchBudgetSpansWorkersAndCannotBecomeAFailedOutcome(string route) {
        string folder = Path.Combine(Path.GetTempPath(), "reader-batch-budget-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string[] paths = { Path.Combine(folder, "one.txt"), Path.Combine(folder, "two.txt") };
            foreach (string path in paths) File.WriteAllText(path, "content");
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
            await Assert.ThrowsAsync<ReaderResourceLimitException>(async () => {
                var options = Limited(new ReaderResourceLimits { MaxChunks = 1 });
                if (route == "ordered") await reader.ReadDocumentsAsync(paths, options);
                else if (route == "detailed") await reader.ReadDocumentsDetailedAsync(paths, options);
                else await reader.ReadDocumentsAsCompletedAsync(paths, _ => { }, options);
            });
        } finally { Directory.Delete(folder, true); }
    }

    [Theory]
    [InlineData("chunks")]
    [InlineData("document")]
    [InlineData("incremental")]
    public void OutputBudgetFailsAllPublicReadRoutes(string route) {
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        var options = Limited(new ReaderResourceLimits { MaxChunks = 1 });
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(new string('a', 600)));
        var error = Assert.Throws<ReaderResourceLimitException>(() => {
            if (route == "chunks") reader.Read(input, "source.txt", options).ToArray();
            else if (route == "document") reader.ReadDocument(input, "source.txt", options);
            else reader.EnumerateChunks(input, "source.txt", options).ToArray();
        });
        Assert.Equal(nameof(ReaderResourceLimits.MaxChunks), error.LimitName);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FolderBudgetSpansFilesAndRemainsTerminal(bool detailed) {
        string folder = Path.Combine(Path.GetTempPath(), "reader-budget-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            File.WriteAllText(Path.Combine(folder, "one.txt"), "one");
            File.WriteAllText(Path.Combine(folder, "two.txt"), "two");
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
            var options = Limited(new ReaderResourceLimits { MaxChunks = 1 });
            Assert.Throws<ReaderResourceLimitException>(() => {
                if (detailed) reader.ReadFolderDetailed(folder, options: options);
                else reader.ReadFolder(folder, options: options).ToArray();
            });
            Assert.Equal(2, reader.ReadFolder(folder, options: Limited(new ReaderResourceLimits { MaxChunks = 2 })).Count());
        } finally { Directory.Delete(folder, true); }
    }

    [Theory]
    [InlineData("chunks")]
    [InlineData("characters")]
    [InlineData("blocks")]
    [InlineData("assets")]
    [InlineData("bytes")]
    public async Task RichAndProcessorOutputCannotBypassBudgets(string limit) {
        var limits = new ReaderResourceLimits();
        switch (limit) {
            case "chunks": limits.MaxChunks = 0; break;
            case "characters": limits.MaxChunkCharacters = 2; break;
            case "blocks": limits.MaxBlocks = 0; break;
            case "assets": limits.MaxAssets = 0; break;
            case "bytes": limits.MaxAssetBytes = 1; break;
        }
        var reader = RichBuilder().Build();
        using var input = new MemoryStream(new byte[] { 1 });
        Assert.Throws<ReaderResourceLimitException>(() => reader.Read(input, "item.rich", Limited(limits)).ToArray());
        await Assert.ThrowsAsync<ReaderResourceLimitException>(() => reader.ReadDocumentAsync(input, "item.rich", Limited(limits)));

        var processed = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
            .AddProcessor(new DelegateOfficeDocumentProcessor("grow", (document, _) => {
                document.Chunks[0].Text = new string('x', 50);
                return document;
            })).Build();
        using var text = new MemoryStream(Encoding.UTF8.GetBytes("a"));
        await Assert.ThrowsAsync<ReaderResourceLimitException>(() => processed.ReadDocumentAsync(text, "source.txt",
            Limited(new ReaderResourceLimits { MaxChunkCharacters = 10 })));
    }

    [Theory]
    [InlineData("bytes")]
    [InlineData("documents")]
    [InlineData("depth")]
    public void NestedZipLimitsSpanSiblingArchives(string limit) {
        byte[] inner = Zip(("one.txt", Encoding.UTF8.GetBytes("abc")));
        byte[] outer = Zip(("a.zip", inner), ("b.zip", inner));
        var limits = new ReaderResourceLimits();
        if (limit == "bytes") limits.MaxNestedInputBytes = inner.Length + 3;
        if (limit == "documents") limits.MaxNestedDocuments = 2;
        if (limit == "depth") limits.MaxNestedDepth = 1;
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
        Assert.Throws<ReaderResourceLimitException>(() => reader.Read(outer, "outer.zip", Limited(limits)).ToArray());
        Assert.Equal(2, reader.Read(outer, "outer.zip", Limited(new ReaderResourceLimits {
            MaxNestedInputBytes = 2 * (inner.Length + 3), MaxNestedDocuments = 4, MaxNestedDepth = 2 })).Count());
    }

    [Fact]
    public void ZipRetainsRichChildAndRoundTripsTransport() {
        var reader = RichBuilder().AddZipHandler().Build();
        var result = reader.ReadDocument(Zip(("child.rich", new byte[] { 1 })), "bundle.zip", Limited(new ReaderResourceLimits { MaxAssets = 1 }));
        var nested = Assert.Single(result.NestedDocuments);
        Assert.Equal("bundle.zip::child.rich", nested.Path);
        Assert.Single(nested.Document.Links);
        Assert.Single(nested.Document.Forms);
        Assert.Single(nested.Document.Assets);
        Assert.Single(nested.Document.Metadata);
        Assert.Single(nested.Document.Blocks);
        Assert.Equal("child.rich", nested.Document.Source.Path);
        Assert.Equal(nested.Document.Source.SourceId, nested.Document.Chunks[0].SourceId);
        var restored = OfficeDocumentReadResultJson.Deserialize(result.ToJson());
        var child = Assert.Single(restored.NestedDocuments).Document;
        Assert.Equal("https://example.test", Assert.Single(child.Links).Uri);
        Assert.Equal(3, Assert.Single(child.Assets).LengthBytes);
        Assert.Null(Assert.Single(child.Assets).PayloadBytes);
    }

    [Fact]
    public void EmailBudgetsStayTerminalAndChildChunksKeepTheirIdentity() {
        const string email = "MIME-Version: 1.0\r\nFrom: a@example.test\r\nTo: b@example.test\r\nSubject: Test\r\nContent-Type: multipart/mixed; boundary=parts\r\n\r\n--parts\r\nContent-Type: text/plain\r\n\r\nBody\r\n--parts\r\nContent-Type: application/octet-stream\r\nContent-Disposition: attachment; filename=child.rich\r\nContent-Transfer-Encoding: base64\r\n\r\nAQ==\r\n--parts--\r\n";
        var reader = RichBuilder().AddEmailHandler().Build();
        byte[] bytes = Encoding.UTF8.GetBytes(email);
        Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(bytes, "mail.eml",
            Limited(new ReaderResourceLimits { MaxNestedDocuments = 0 })));
        var result = reader.ReadDocument(bytes, "mail.eml", Limited(new ReaderResourceLimits { MaxNestedDocuments = 1 }));
        var child = Assert.Single(result.NestedDocuments).Document;
        Assert.Equal("child", child.Chunks[0].Id);
        Assert.Contains(result.Chunks, c => c.Id.StartsWith("email:attachment-content:", StringComparison.Ordinal));
        Assert.NotNull(child.Chunks[0].SourceId);
    }

    [Fact]
    public void TransportRejectsInvalidNestedContractsAndKeepsOldSchemas() {
        var result = RichResult("test.rich");
        result.NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "child.rich", Document = RichResult("child.rich") } };
        var node = JsonNode.Parse(result.ToJson())!.AsObject();
        node["nestedDocuments"]![0]!["document"]!["schemaVersion"] = 999;
        Assert.Throws<OfficeDocumentReadResultSchemaException>(() => OfficeDocumentReadResultJson.Deserialize(node.ToJsonString()));
        result.SchemaVersion = 8;
        Assert.Throws<JsonException>(() => result.ToJson());
        result.NestedDocuments = Array.Empty<OfficeDocumentNestedResult>();
        using var old = JsonDocument.Parse(result.ToJson());
        Assert.False(old.RootElement.TryGetProperty("nestedDocuments", out _));
        Assert.Empty(OfficeDocumentReadResultJson.Deserialize(old.RootElement.GetRawText()).NestedDocuments);
        result.SchemaVersion = 9;
        result.NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "cycle", Document = result } };
        Assert.Throws<JsonException>(() => result.ToJson());
    }

    private static ReaderOptions Limited(ReaderResourceLimits limits) => new ReaderOptions {
        ComputeHashes = false, MaxChars = 256, ResourceLimits = limits
    };
    private static OfficeDocumentReaderBuilder RichBuilder() => new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
        Id = "rich", Kind = ReaderInputKind.Pdf, Extensions = new[] { ".rich" },
        ReadDocumentStream = (_, name, _, _) => RichResult(name!)
    });
    private static OfficeDocumentReadResult RichResult(string name) => new OfficeDocumentReadResult {
        Kind = ReaderInputKind.Pdf, Source = new OfficeDocumentSource { Path = name },
        Chunks = new[] { new ReaderChunk { Id = "child", Kind = ReaderInputKind.Pdf, Text = "hello", Location = new ReaderLocation { Path = name } } },
        Blocks = new[] { new OfficeDocumentBlock { Id = "block", Text = "hello" } },
        Assets = new[] { new OfficeDocumentAsset { Id = "asset", LengthBytes = 3, PayloadBytes = new byte[] { 1, 2, 3 } } },
        Links = new[] { new OfficeDocumentLink { Id = "link", Uri = "https://example.test" } },
        Forms = new[] { new OfficeDocumentFormField { Id = "form", Name = "answer", Value = "yes" } },
        Metadata = new[] { new OfficeDocumentMetadataEntry { Name = "author", Value = "writer" } }
    };
    private static byte[] Zip(params (string Name, byte[] Bytes)[] entries) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
            foreach (var entry in entries) { using var output = archive.CreateEntry(entry.Name).Open(); output.Write(entry.Bytes, 0, entry.Bytes.Length); }
        return stream.ToArray();
    }
}
