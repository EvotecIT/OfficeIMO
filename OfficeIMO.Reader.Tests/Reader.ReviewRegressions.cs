using OfficeIMO.Email;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderReviewRegressionTests {
    [Theory]
    [InlineData("path", "blocks")]
    [InlineData("stream", "assets")]
    [InlineData("bytes", "bytes")]
    public async Task NativeAsyncChunkReadsEnforceRichOutputBudgets(string route, string limit) {
        var limits = limit == "blocks" ? new ReaderResourceLimits { MaxBlocks = 0 }
            : limit == "assets" ? new ReaderResourceLimits { MaxAssets = 0 }
            : new ReaderResourceLimits { MaxAssetBytes = 0 };
        var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "async", Extensions = new[] { ".native" },
            ReadDocumentPathAsync = (_, _, _) => Task.FromResult(Rich()),
            ReadDocumentStreamAsync = (_, _, _, _) => Task.FromResult(Rich())
        }).Build();
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".native");
        File.WriteAllBytes(path, new byte[] { 1 });
        try {
            await Assert.ThrowsAsync<ReaderResourceLimitException>(async () => {
                var options = Options(limits);
                if (route == "path") await reader.ReadAsync(path, options);
                else if (route == "bytes") await reader.ReadAsync(new byte[] { 1 }, "input.native", options);
                else { using var stream = new MemoryStream(new byte[] { 1 }); await reader.ReadAsync(stream, "input.native", options); }
            });
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageAssetsAreCountedAndSharedRootRecordsAreNotCountedTwice(bool bytes) {
        var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "page-assets", Extensions = new[] { ".page" },
            ReadDocumentStream = (_, _, _, _) => {
                var result = Rich();
                result.Pages = new[] { new OfficeDocumentPage { Assets = result.Assets } };
                result.Assets = Array.Empty<OfficeDocumentAsset>();
                return result;
            }
        }).Build();
        Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(new byte[] { 1 }, "input.page",
            Options(bytes ? new ReaderResourceLimits { MaxAssetBytes = 2 } : new ReaderResourceLimits { MaxAssets = 0 })));
        Assert.Single(reader.ReadDocument(new byte[] { 1 }, "input.page", Options(new ReaderResourceLimits {
            MaxAssets = 1, MaxAssetBytes = 3 })).Pages[0].Assets);
        var shared = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "shared", Extensions = new[] { ".page" },
            ReadDocumentStream = (_, _, _, _) => { var result = Rich(); result.Pages = new[] { new OfficeDocumentPage { Assets = result.Assets } }; return result; }
        }).Build();
        Assert.Single(shared.ReadDocument(new byte[] { 1 }, "input.page", Options(new ReaderResourceLimits {
            MaxAssets = 1, MaxAssetBytes = 3 })).Assets);
    }

    [Fact]
    public void ClearingPayloadDoesNotIncreaseAccountedAssetBytes() {
        var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "asset", Extensions = new[] { ".asset" }, ReadDocumentStream = (_, _, _, _) => Rich()
        }).AddProcessor(new DelegateOfficeDocumentProcessor("clear", (document, _) => {
            document.Assets[0].PayloadBytes = null;
            return document;
        })).Build();
        var result = reader.ReadDocument(new byte[] { 1 }, "input.asset", Options(new ReaderResourceLimits { MaxAssetBytes = 3 }));
        Assert.Null(Assert.Single(result.Assets).PayloadBytes);
    }

    [Theory]
    [InlineData("default", false)]
    [InlineData("ceiling", false)]
    [InlineData("prefix", false)]
    [InlineData("prefix", true)]
    public void IncrementalStreamOnlyHandlerEnforcesItsAdmissionPolicy(string policy, bool seekable) {
        var handler = PolicyHandler(policy);
        handler.SupportsIncrementalStream = true;
        var reader = new OfficeDocumentReaderBuilder().AddHandler(handler).Build();
        using var forward = new CountingStream(new byte[10_000]);
        using var memory = new MemoryStream(new byte[10_000]);
        Stream input = seekable ? memory : forward;
        Assert.Throws<IOException>(() => reader.EnumerateChunks(input, "input.policy", new ReaderOptions {
            ComputeHashes = false, MaxInputBytes = policy == "ceiling" ? 1024 : null
        }).ToArray());
        if (!seekable) Assert.InRange(forward.BytesRead, 1, policy == "prefix" ? 9 : 129);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData("default")]
    [InlineData("ceiling")]
    [InlineData("prefix")]
    public void NestedForwardInputAppliesHandlerPolicyDuringBuffering(string policy) {
        using var input = new CountingStream(new byte[10_000]);
        var reader = new OfficeDocumentReaderBuilder().AddHandler(PolicyHandler(policy)).AddHandler(new ReaderHandlerRegistration {
            Id = "container", Extensions = new[] { ".parent" },
            ReadDocumentStream = (_, _, options, token) => ReaderNestedContent.ReadDocument(input, "child.policy", options, token)
        }).Build();
        Assert.Throws<IOException>(() => reader.ReadDocument(new byte[] { 1 }, "input.parent", new ReaderOptions {
            ComputeHashes = false, MaxInputBytes = policy == "ceiling" ? 1024 : null,
            ResourceLimits = new ReaderResourceLimits { MaxNestedInputBytes = 100_000 }
        }));
        Assert.InRange(input.BytesRead, 1, policy == "prefix" ? 9 : 129);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData("depth")]
    [InlineData("documents")]
    [InlineData("bytes")]
    public void NativeEmbeddedEmailMessagesConsumeNestedBudgets(string limit) {
        var child = new EmailDocument { Subject = "Child" }; child.Body.Text = "nested body";
        var parent = new EmailDocument { Subject = "Parent" }; parent.Body.Text = "root body";
        parent.Attachments.Add(new EmailAttachment { FileName = "child.eml", ContentType = "message/rfc822", EmbeddedDocument = child });
        byte[] bytes = new EmailDocumentWriter().ToBytes(parent, EmailFileFormat.Eml);
        var limits = limit == "depth" ? new ReaderResourceLimits { MaxNestedDepth = 0 }
            : limit == "documents" ? new ReaderResourceLimits { MaxNestedDocuments = 0 }
            : new ReaderResourceLimits { MaxNestedInputBytes = 0 };
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandler().Build();
        Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(bytes, "parent.eml", Options(limits)));
    }

    [Theory]
    [InlineData(EmailFileFormat.Eml)]
    [InlineData(EmailFileFormat.OutlookMsg)]
    public void NativeEmbeddedEmailRetainsRichChildAndRespectsDepthAcrossGrandchildren(EmailFileFormat format) {
        var child = new EmailDocument { Subject = "Child" }; child.Body.Text = "nested body";
        var parent = new EmailDocument { Subject = "Parent" }; parent.Body.Text = "root body";
        parent.Attachments.Add(new EmailAttachment { FileName = "child.eml", ContentType = "message/rfc822", EmbeddedDocument = child });
        var writer = new EmailDocumentWriter();
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandler().Build();
        string name = format == EmailFileFormat.OutlookMsg ? "parent.msg" : "parent.eml";
        var result = reader.ReadDocument(writer.ToBytes(parent, format), name, Options(new ReaderResourceLimits {
            MaxNestedDocuments = 1, MaxNestedDepth = 1, MaxNestedInputBytes = 100_000
        }));
        var nested = Assert.Single(result.NestedDocuments);
        Assert.Equal("Child", nested.Document.Source.Subject);
        Assert.Contains(nested.Document.Chunks, chunk => chunk.Text == "nested body");
        Assert.Contains(nested.Document.Metadata, entry => entry.Name == "Subject" && entry.Value == "Child");
        Assert.All(nested.Document.Chunks, chunk => Assert.Equal(nested.Document.Source.SourceId, chunk.SourceId));
        Assert.All(result.Chunks, chunk => Assert.Equal(result.Source.SourceId, chunk.SourceId));
        child.Attachments.Add(new EmailAttachment { FileName = "grandchild.eml", ContentType = "message/rfc822", EmbeddedDocument = new EmailDocument { Subject = "Grandchild" } });
        Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(writer.ToBytes(parent, format), name,
            Options(new ReaderResourceLimits { MaxNestedDepth = 1 })));
    }

    [Fact]
    public void EmbeddedEmailAttachmentsKeepDistinctAssetIdsAndPayloads() {
        var child = new EmailDocument { Subject = "Child" }; child.Body.Text = "child body";
        child.Attachments.Add(new EmailAttachment { FileName = "notes.txt", ContentType = "text/plain", Content = Encoding.UTF8.GetBytes("notes") });
        var parent = new EmailDocument { Subject = "Parent" }; parent.Body.Text = "root body";
        parent.Attachments.Add(new EmailAttachment { FileName = "child.eml", ContentType = "message/rfc822", EmbeddedDocument = child });
        var result = new OfficeDocumentReaderBuilder().AddEmailHandler().Build()
            .ReadDocument(new EmailDocumentWriter().ToBytes(parent, EmailFileFormat.Eml), "parent.eml");
        Assert.Equal(2, result.Assets.Count);
        Assert.Equal(2, result.BuildAssetDataUriMap().Count);
        Assert.Equal(result.Assets.Count, result.Assets.Select(asset => asset.Id).Distinct(StringComparer.Ordinal).Count());
        var childAsset = Assert.Single(Assert.Single(result.NestedDocuments).Document.Assets);
        Assert.Contains(result.Assets, asset => asset.Id == childAsset.Id && ReferenceEquals(asset.PayloadBytes, childAsset.PayloadBytes));
    }

    [Fact]
    public void TokenizerCancellationDuringNoFitSearchRemainsCancellation() {
        using var cancellation = new System.Threading.CancellationTokenSource();
        var counter = new ReaderDelegateTokenCounter("no-fit", text => text.Length * 2,
            (_, source, _, length) => {
                if (length == source.Length) cancellation.Cancel();
                return length * 2;
            });
        Assert.Throws<OperationCanceledException>(() => ReaderHierarchicalChunker.Chunk(
            new[] { new ReaderChunk { Text = new string('x', 4096) } },
            new ReaderHierarchicalChunkingOptions { MaxTokens = 1, OverlapTokens = 0, IncludeContextInText = false, TokenCounter = counter },
            cancellation.Token));
    }

    [Fact]
    public void IncrementalPrefixIsReplayedWithoutBufferingTheWholeInput() {
        byte[] bytes = Encoding.UTF8.GetBytes(new string('a', 100_000));
        var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "prefix", Extensions = new[] { ".prefix" }, SupportsIncrementalStream = true,
            InputLimitProbeBytes = 8, ResolveMaxInputBytesFromPrefix = prefix => { Assert.Equal("aaaaaaaa", Encoding.UTF8.GetString(prefix.ToArray())); return 200_000; },
            ReadStream = (input, _, _, _) => ReadOne(input)
        }).Build();
        using var input = new CountingStream(bytes);
        using var iterator = reader.EnumerateChunks(input, "input.prefix", new ReaderOptions { ComputeHashes = false }).GetEnumerator();
        Assert.True(iterator.MoveNext());
        Assert.Equal("aaaaaaaaaa", iterator.Current.Text);
        Assert.Equal(10, input.BytesRead);
    }
    private static IEnumerable<ReaderChunk> ReadOne(Stream input) {
        var bytes = new byte[10]; int count = 0;
        while (count < bytes.Length) { int read = input.Read(bytes, count, bytes.Length - count); if (read == 0) break; count += read; }
        yield return new ReaderChunk { Text = Encoding.UTF8.GetString(bytes, 0, count) };
    }

    [Fact]
    public void TokenizerCanFitACompleteTokenWhoseFirstScalarExceedsTheBudget() {
        var counter = new ReaderDelegateTokenCounter("vocabulary", text => text == "你好" ? 1 : text.Length * 2);
        var result = ReaderHierarchicalChunker.Chunk(new[] { new ReaderChunk { Text = "你好" } },
            new ReaderHierarchicalChunkingOptions { MaxTokens = 1, OverlapTokens = 0, IncludeContextInText = false, TokenCounter = counter });
        Assert.Equal("你好", Assert.Single(result.Chunks).Text);
        Assert.Equal(1, result.Chunks[0].TokenEstimate);
    }

    [Fact]
    public void EmailHashesIncludeSourceIdentityAndUseTheCoreFramedContract() {
        var document = new EmailDocument { Subject = "Same" }; document.Body.Text = "Same body";
        byte[] bytes = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml);
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandler().Build();
        var first = reader.ReadDocument(bytes, "one.eml");
        var second = reader.ReadDocument(bytes, "two.eml");
        Assert.Equal(first.Chunks.Select(c => c.Text), second.Chunks.Select(c => c.Text));
        Assert.All(first.Chunks.Zip(second.Chunks, (left, right) => (First: left, Second: right)), pair => Assert.NotEqual(pair.First.ChunkHash, pair.Second.ChunkHash));
        var refreshed = new OfficeDocumentReaderBuilder().AddEmailHandler()
            .AddProcessor(new DelegateOfficeDocumentProcessor("identity", (result, _) => result)).Build()
            .ReadDocument(bytes, "one.eml");
        Assert.Equal(first.Chunks.Select(chunk => chunk.ChunkHash), refreshed.Chunks.Select(chunk => chunk.ChunkHash));
    }

    [Fact]
    public void ForwardOnlyIncrementalReadPreservesRequestedSourceHash() {
        byte[] bytes = Encoding.UTF8.GetBytes("hash me");
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        using var input = new CountingStream(bytes);
        var chunk = Assert.Single(reader.EnumerateChunks(input, "input.txt", new ReaderOptions { ComputeHashes = true }));
        Assert.Equal(Assert.Single(reader.Read(bytes, "input.txt")).SourceHash, chunk.SourceHash);
        Assert.NotNull(chunk.SourceHash);
        Assert.True(input.CanRead);
    }

    private static ReaderHandlerRegistration PolicyHandler(string policy) => new ReaderHandlerRegistration {
        Id = "policy", Extensions = new[] { ".policy" },
        DefaultMaxInputBytes = policy == "default" ? 128 : null,
        MaxInputBytesCeiling = policy == "ceiling" ? 128 : null,
        InputLimitProbeBytes = policy == "prefix" ? 8 : 0,
        ResolveMaxInputBytesFromPrefix = policy == "prefix" ? _ => 8 : null,
        ReadStream = (stream, _, _, _) => { stream.CopyTo(Stream.Null); return Array.Empty<ReaderChunk>(); }
    };
    private static ReaderOptions Options(ReaderResourceLimits limits) => new ReaderOptions { ComputeHashes = false, ResourceLimits = limits };
    private static OfficeDocumentReadResult Rich() => new OfficeDocumentReadResult {
        Blocks = new[] { new OfficeDocumentBlock { Text = "block" } },
        Assets = new[] { new OfficeDocumentAsset { LengthBytes = 3, PayloadBytes = new byte[3] } }
    };
    private sealed class CountingStream : Stream {
        private readonly MemoryStream _inner;
        internal CountingStream(byte[] bytes) { _inner = new MemoryStream(bytes); }
        internal long BytesRead { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) { int read = _inner.Read(buffer, offset, count); BytesRead += read; return read; }
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
