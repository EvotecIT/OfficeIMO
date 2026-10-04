using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests.Emlx;

public sealed class EmlxStreamingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void SelectedReadsRejectSameLengthChangesUnlessSessionOwnsASnapshot(bool snapshot, bool duringRead) {
        byte[] bytes = Encoding.UTF8.GetBytes("Subject: Original\r\n\r\nBody\r\n");
        using var input = new MutatingInput();
        byte[] prefix = Encoding.ASCII.GetBytes(bytes.Length + "\n");
        input.Write(prefix, 0, prefix.Length);
        input.Write(bytes, 0, bytes.Length);
        input.Position = 2;
        using (EmailStoreSession session = snapshot ? EmailStoreSession.OpenSnapshot(input, "mutable.emlx") :
            EmailStoreSession.Open(input, "mutable.emlx")) {
            EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
            Assert.Equal("Original", session.ReadSummary(reference).Subject);
            Assert.Equal("Original", session.ReadItem(reference).Document.Subject);
            void Mutate() {
                long position = input.Position;
                input.Position = prefix.Length + "Subject: ".Length;
                input.Write(Encoding.ASCII.GetBytes("Modified"), 0, 8);
                input.Position = position;
            }
            if (duringRead) input.AfterEndOfRead = Mutate;
            else Mutate();
            if (snapshot) Assert.Equal("Original", session.ReadItem(reference).Document.Subject);
            else Assert.Throws<InvalidDataException>(() => session.ReadItem(reference));
        }
        Assert.True(input.CanRead);
        Assert.Equal(2, input.Position);
    }

    private sealed class MutatingInput : MemoryStream {
        internal Action? AfterEndOfRead { get; set; }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            if (read == 0 && AfterEndOfRead is { } change) { AfterEndOfRead = null; change(); }
            return read;
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StreamingAttachmentsSurviveTheReadAndExpireWithTheSession(bool directory) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-emlx-stream-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            byte[] expected = Enumerable.Range(0, 65536).Select(index => (byte)index).ToArray();
            string path = Path.Combine(root, "message.emlx");
            new EmailStoreEmlxWriter().Write(Document(expected), path).RequireNoLoss();
            IEmailContentSource source;
            Stream outstanding;
            using (EmailStoreSession session = EmailStoreSession.Open(directory ? root : path)) {
                EmailStoreItem item = session.ReadItem(Assert.Single(session.EnumerateItems()),
                    new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true));
                EmailAttachment attachment = Assert.Single(item.Document.Attachments);
                Assert.Null(attachment.Content);
                source = Assert.IsAssignableFrom<IEmailContentSource>(attachment.ContentSource);
                outstanding = source.OpenRead();
                using Stream input = source.OpenRead();
                using var content = new MemoryStream();
                input.CopyTo(content);
                Assert.Equal(expected, content.ToArray());
                using var rewritten = new MemoryStream();
                new EmailStoreEmlxWriter().Write(item.Document, rewritten).RequireNoLoss();
                Assert.Equal(expected, new EmailStoreReader().Read(rewritten, "rewritten.emlx")
                    .Store.Folders.Single().Items.Single().Document.Attachments.Single().Content);
            }
            Assert.Throws<ObjectDisposedException>(() => source.OpenRead());
            Assert.Throws<ObjectDisposedException>(() => outstanding.ReadByte());
            outstanding.Dispose();
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public void RepeatedStreamingReadsCannotAccumulateUnboundedTemporaryAttachments() {
        byte[] bytes = new EmailStoreEmlxWriter().ToBytes(Document(new byte[16]));
        using var input = new MemoryStream(bytes);
        using EmailStoreSession session = EmailStoreSession.Open(input, "bounded.emlx",
            new EmailStoreReaderOptions(maxAttachmentBytes: 16, maxTotalAttachmentBytes: 16));
        EmailStoreItemReference reference = Assert.Single(session.EnumerateItems());
        var options = new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true);
        IEmailContentSource source = session.ReadItem(reference, options).Document.Attachments.Single().ContentSource!;
        EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(() => session.ReadItem(reference, options));
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes), error.LimitName);
        using Stream content = source.OpenRead();
        Assert.Equal(16, content.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AsyncWritesUseAsyncSourceIoAndCancellationProtectsTheDestination(bool cancel) {
        using var cancellation = new CancellationTokenSource();
        var source = new AsyncSource(cancel ? cancellation : null);
        EmailDocument document = Document(null);
        document.Attachments.Single().ContentSource = source;
        string root = Path.Combine(Path.GetTempPath(), "officeimo-emlx-async-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "destination.emlx");
        File.WriteAllText(path, "Existing artifact");
        try {
            if (cancel) {
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                    new EmailStoreEmlxWriter().WriteAsync(document, path, cancellation.Token));
                Assert.Equal("Existing artifact", File.ReadAllText(path));
            } else {
                EmailWriteResult result = await new EmailStoreEmlxWriter().WriteAsync(document, path, cancellation.Token);
                result.RequireNoLoss();
                Assert.Equal(new byte[] { 1, 2, 3, 4 }, new EmailStoreReader().Read(path)
                    .Store.Folders.Single().Items.Single().Document.Attachments.Single().Content);
            }
            Assert.Equal(1, source.OpenCount);
            Assert.True(source.Closed);
        } finally { Directory.Delete(root, recursive: true); }
    }

    private static EmailDocument Document(byte[]? content) {
        var document = new EmailDocument { Subject = "Streaming EMLX" };
        document.Body.Text = "Body";
        document.Attachments.Add(new EmailAttachment {
            FileName = "payload.bin", ContentType = "application/octet-stream", Content = content,
            Length = content?.Length ?? 4
        });
        return document;
    }

    private sealed class AsyncSource : IEmailContentSource {
        private readonly CancellationTokenSource? _cancellation;
        internal AsyncSource(CancellationTokenSource? cancellation) => _cancellation = cancellation;
        public long? Length => 4;
        internal int OpenCount { get; private set; }
        internal bool Closed { get; private set; }
        public Stream OpenRead() => throw new InvalidOperationException("The EMLX async writer must use asynchronous source I/O.");
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OpenCount++;
            return Task.FromResult<Stream>(new AsyncInput(this));
        }
        private sealed class AsyncInput : MemoryStream {
            private readonly AsyncSource _owner;
            internal AsyncInput(AsyncSource owner) : base(new byte[] { 1, 2, 3, 4 }) => _owner = owner;
            public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Synchronous read is forbidden.");
            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
                _owner._cancellation?.Cancel();
                cancellationToken.ThrowIfCancellationRequested();
                return Task.FromResult(base.Read(buffer, offset, count));
            }
#if NET8_0_OR_GREATER
            public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) {
                _owner._cancellation?.Cancel();
                cancellationToken.ThrowIfCancellationRequested();
                return ValueTask.FromResult(base.Read(buffer.Span));
            }
#endif
            protected override void Dispose(bool disposing) {
                _owner.Closed = true;
                base.Dispose(disposing);
            }
        }
    }
}
