using System.Security.Cryptography;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;

namespace OfficeIMO.Email.Tests;

public sealed class EmailAsyncReaderProjectionTests {
    [Theory]
    [InlineData("message.eml", "Subject: Native async\r\nContent-Type: text/plain\r\n\r\nhello", true)]
    [InlineData("message.eml", "Subject: Native async\r\nContent-Type: text/plain\r\n\r\nhello", false)]
    [InlineData("mail.mbox", "From a@example.test Wed Sep 30 12:00:00 2026\nSubject: Native async\n\nhello\n", true)]
    [InlineData("calendar.ics", "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:one\r\nSUMMARY:hello\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n", true)]
    [InlineData("contact.vcf", "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:hello\r\nEND:VCARD\r\n", true)]
    public async Task NativeHandlersAndSourceHashingUseAsyncSourceIoAndPreservePosition(string name, string text, bool hashes) {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().Build();
        var options = new ReaderOptions { ComputeHashes = hashes, DetectionMode = ReaderDetectionMode.ExtensionOnly };
        using var source = new AsyncOnlySource(bytes);
        source.Position = 3;
        OfficeDocumentReadResult result = await reader.ReadDocumentAsync(source, name, options);
        Assert.NotEmpty(result.Chunks);
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("hello"));
        Assert.Equal(3, source.Position);
        Assert.True(source.CanRead);
        using var sha = SHA256.Create();
        string? expected = hashes ? string.Concat(sha.ComputeHash(bytes).Select(value => value.ToString("x2"))) : null;
        Assert.Equal(expected, result.Source.SourceHash);
        // Core snapshots a caller stream once; parsing and hashing use the owned snapshot.
        Assert.Equal(bytes.Length, source.BytesRead);
        Assert.All(result.Chunks, chunk => {
            Assert.Equal(expected, chunk.SourceHash);
            Assert.Equal(result.Source.SourceId, chunk.SourceId);
            Assert.Equal(hashes, !string.IsNullOrWhiteSpace(chunk.ChunkHash));
        });
        using var syncSource = new MemoryStream(bytes);
        OfficeDocumentReadResult syncResult = reader.ReadDocument(syncSource, name, options);
        Assert.Equal(syncResult.Source.SourceId, result.Source.SourceId);
        Assert.Equal(syncResult.Chunks.Select(chunk => chunk.Text), result.Chunks.Select(chunk => chunk.Text));
        Assert.Equal(syncResult.Chunks.Select(chunk => chunk.ChunkHash), result.Chunks.Select(chunk => chunk.ChunkHash));
    }

    [Fact]
    public void DefaultBodyProjectionProducesTextAndSemanticMarkdownAndHonorsCustomHtmlHandler() {
        byte[] bytes = Encoding.UTF8.GetBytes("Subject: HTML\r\nContent-Type: text/html\r\n\r\n<p>Hello <strong>world</strong></p>");
        var builder = new OfficeDocumentReaderBuilder().AddEmailHandlers();
        using var source = new MemoryStream(bytes);
        var result = builder.Build().ReadDocument(source, "message.eml");
        ReaderChunk body = Assert.Single(result.Chunks, chunk => chunk.Location.SourceBlockKind == "email-body-html");
        Assert.Contains("Hello", body.Text);
        Assert.Contains("**world**", body.Markdown);
        Assert.DoesNotContain("<strong>", body.Text);
        builder.AddHandler(new ReaderHandlerRegistration {
            Id = "custom-html", Kind = ReaderInputKind.Html, Extensions = new[] { ".html" },
            ReadStream = (input, name, options, token) => new[] { new ReaderChunk { Text = "custom", Markdown = "custom" } }
        });
        var customized = builder.Build().ReadDocument(source, "message.eml");
        Assert.Contains(customized.Chunks, chunk => chunk.Text == "custom");
    }

    [Fact]
    public async Task AsyncReadCancellationRemainsObservable() {
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().Build();
        using var stream = new AsyncOnlySource(Encoding.ASCII.GetBytes("Subject: cancel\r\n\r\nbody"));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.ReadDocumentAsync(stream, "message.eml",
            cancellationToken: new CancellationToken(true)));
        Assert.Equal(0, stream.BytesRead);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void DirectArtifactHandlersAdvertiseNativeAsyncEntryPoints() {
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandler().Build();
        Assert.All(reader.GetCapabilities(), capability => {
            Assert.True(capability.SupportsAsyncPath);
            Assert.True(capability.SupportsAsyncStream);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ContentDetectionKeepsTheSelectedAsyncContentLineOwner(bool vcard) {
        byte[] bytes = ContentLines(vcard);
        ReaderInputKind expected = vcard ? ReaderInputKind.VCard : ReaderInputKind.Calendar;
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().Build();
        var options = new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent };
        foreach (string? name in new string?[] { null, "renamed.bin", vcard ? "wrong.ics" : "wrong.vcf" }) {
            using var stream = new MemoryStream(bytes);
            var result = await reader.ReadDocumentAsync(stream, name, options);
            Assert.Equal(expected, result.Kind);
            Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("hello"));
        }
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + (vcard ? ".ics" : ".vcf"));
        try {
            File.WriteAllBytes(path, bytes);
            Assert.Equal(expected, (await reader.ReadDocumentAsync(path, options)).Kind);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task StableContentLineFileHeldByProducerRemainsReadableSynchronouslyAndAsynchronously(bool vcard) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + (vcard ? ".vcf" : ".ics"));
        try {
            File.WriteAllBytes(path, ContentLines(vcard));
            using var producer = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete);
            var reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().Build();
            var sync = reader.ReadDocument(path);
            var asyncResult = await reader.ReadDocumentAsync(path);
            Assert.Equal(sync.Kind, asyncResult.Kind);
            Assert.Equal(sync.Source.SourceHash, asyncResult.Source.SourceHash);
            Assert.Equal(sync.Chunks.Select(chunk => chunk.Text), asyncResult.Chunks.Select(chunk => chunk.Text));
        } finally { File.Delete(path); }
    }

    private static byte[] ContentLines(bool vcard) => Encoding.ASCII.GetBytes(vcard
        ? "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:hello\r\nEND:VCARD\r\n"
        : "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:one\r\nSUMMARY:hello\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

    private sealed class AsyncOnlySource : MemoryStream {
        internal AsyncOnlySource(byte[] bytes) : base(bytes, writable: false) { }
        internal int BytesRead { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Synchronous source I/O is forbidden.");
        public override int ReadByte() => throw new InvalidOperationException("Synchronous source I/O is forbidden.");
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = base.Read(buffer, offset, count);
            BytesRead += read;
            return Task.FromResult(read);
        }
    }
}
