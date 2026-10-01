using System.Security.Cryptography;

namespace OfficeIMO.Email.Tests;

public sealed class EmailAttachmentExtractionTests {
    [Fact]
    public void EmbeddedExportsShareTheOperationWideAttachmentVisitBudget() {
        InDirectory(root => {
            var sources = new List<CountingSource>();
            var document = new EmailDocument();
            for (int message = 0; message < 8; message++) {
                var child = new EmailDocument();
                for (int index = 0; index < 2; index++) {
                    var source = new CountingSource(bytes: 0);
                    sources.Add(source);
                    child.Attachments.Add(new EmailAttachment { ContentSource = source });
                }
                document.Attachments.Add(new EmailAttachment { EmbeddedDocument = child });
            }
            var result = EmailAttachmentExtractor.Extract(document, root, new EmailAttachmentExtractionOptions(maxAttachments: 7));
            Assert.True(result.Truncated);
            Assert.True(result.Entries.Count <= 7);
            Assert.Equal(4, sources.Sum(source => source.OpenCount));
            Assert.All(sources.Skip(4), source => Assert.Equal(0, source.OpenCount));
            Assert.Equal(2, Directory.GetFiles(root).Length);
        });
    }

    [Fact]
    public void EmbeddedSerializationCannotResetTheDepthBudget() {
        InDirectory(root => {
            var source = new CountingSource(bytes: 0);
            var deep = new EmailDocument();
            deep.Attachments.Add(new EmailAttachment { ContentSource = source });
            var nested = new EmailDocument();
            nested.Attachments.Add(new EmailAttachment { EmbeddedDocument = deep });
            var document = new EmailDocument();
            document.Attachments.Add(new EmailAttachment { EmbeddedDocument = nested });
            var result = EmailAttachmentExtractor.Extract(document, root, new EmailAttachmentExtractionOptions(maxDepth: 1));
            Assert.True(result.Truncated);
            Assert.Null(Assert.Single(result.Entries).OutputPath);
            Assert.Equal(0, source.OpenCount);
            Assert.Empty(Directory.GetFiles(root));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SynchronousExtractionDoesNotExposeBlockedCallerContextToAsyncSources(bool embedded) {
        InDirectory(root => {
            var source = new ContextAwareSource();
            var payload = new EmailDocument();
            payload.Attachments.Add(new EmailAttachment { FileName = "content.bin", ContentSource = source });
            var document = payload;
            if (embedded) {
                document = new EmailDocument();
                document.Attachments.Add(new EmailAttachment { EmbeddedDocument = payload });
            }
            var previous = SynchronizationContext.Current;
            EmailAttachmentExtractionResult result;
            try {
                SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
                result = EmailAttachmentExtractor.Extract(document, root);
            } finally { SynchronizationContext.SetSynchronizationContext(previous); }
            Assert.False(source.SawCallerContext);
            Assert.NotNull(Assert.Single(result.Entries).OutputPath);
            Assert.Equal(1, source.Opens);
        });
    }

    [Fact]
    public async Task AsyncEmbeddedExtractionReadsSourceOnceAndIncludesInputAndGeneratedBytes() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO.Email.AsyncExtraction.Tests." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var source = new CountingSource(bytes: 80, seekable: false);
            var nested = new EmailDocument { Subject = "Source once" };
            nested.Attachments.Add(new EmailAttachment { FileName = "content.dat", ContentSource = source });
            var document = new EmailDocument();
            document.Attachments.Add(new EmailAttachment { EmbeddedDocument = nested });
            var result = await EmailAttachmentExtractor.ExtractAsync(document, root);
            var entry = Assert.Single(result.Entries);
            Assert.NotNull(entry.OutputPath);
            Assert.Equal(1, source.OpenCount);
            Assert.Equal(80, source.Stream!.BytesRead);
            Assert.Equal(80 + entry.BytesWritten, result.BytesRead);
            Assert.True(source.Stream.Disposed);
            using var read = new EmailDocumentReader().Read(entry.OutputPath!);
            Assert.Equal(new byte[80], Assert.Single(read.Document.Attachments).Content);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void BlockedEmbeddedSignaturesDoNotOpenPayloadSources() {
        InDirectory(root => {
            var source = new CountingSource();
            var nested = new EmailDocument();
            nested.Headers.Add(new EmailHeader("DKIM-Signature", "original"));
            nested.Attachments.Add(new EmailAttachment { ContentSource = source });
            var document = new EmailDocument();
            document.Attachments.Add(new EmailAttachment { EmbeddedDocument = nested });
            var result = EmailAttachmentExtractor.Extract(document, root);
            var entry = Assert.Single(result.Entries);
            Assert.Null(entry.OutputPath);
            Assert.Contains(entry.Diagnostics, d => d.Code == "EMAIL_TRANSPORT_SIGNATURE_INVALIDATED");
            Assert.Equal(0, source.OpenCount);
            Assert.Empty(Directory.GetFiles(root));
        });
    }
    [Fact]
    public void FailedEmbeddedPreparationAccountsSourceBytesAndStopsAggregateBudget() {
        InDirectory(root => {
            var first = new CountingSource(bytes: 20000, seekable: false);
            var second = new CountingSource(bytes: 20000, seekable: false);
            var document = new EmailDocument();
            foreach (var source in new[] { first, second }) {
                var nested = new EmailDocument();
                var payload = new EmailAttachment { ContentType = "multipart/mixed", ContentSource = source };
                payload.ContentTypeParameters["boundary"] = "retained";
                nested.Attachments.Add(payload);
                document.Attachments.Add(new EmailAttachment { EmbeddedDocument = nested });
            }
            var result = EmailAttachmentExtractor.Extract(document, root, new EmailAttachmentExtractionOptions(maxAttachmentBytes: 4096, maxTotalBytes: 4096));
            Assert.True(result.Truncated);
            Assert.Single(result.Entries);
            Assert.Equal(4097, first.Stream!.BytesRead);
            Assert.True(first.Stream.Disposed);
            Assert.Null(second.Stream);
            Assert.Equal(4097, result.BytesRead);
            Assert.Empty(Directory.GetFiles(root));
        });
    }

    [Fact]
    public void EmbeddedPreparationCancellationWinsOverLimitFailureAndClosesSource() {
        InDirectory(root => {
            using var cancellation = new CancellationTokenSource();
            var source = new CountingSource(cancellation, bytes: 20000, seekable: false);
            var nested = new EmailDocument();
            var payload = new EmailAttachment { ContentType = "multipart/mixed", ContentSource = source };
            payload.ContentTypeParameters["boundary"] = "retained";
            nested.Attachments.Add(payload);
            var document = new EmailDocument();
            document.Attachments.Add(new EmailAttachment { EmbeddedDocument = nested });
            Assert.Throws<OperationCanceledException>(() => EmailAttachmentExtractor.Extract(document, root,
                new EmailAttachmentExtractionOptions(maxAttachmentBytes: 4096), cancellation.Token));
            Assert.True(source.Stream!.Disposed);
            Assert.Equal(4097, source.Stream.BytesRead);
            Assert.Empty(Directory.GetFiles(root));
        });
    }
    [Fact]
    public void PortableNamesPreventTraversalCollisionsAndReservedNamesWithHashEvidence() {
        InDirectory(root => {
            var document = new EmailDocument();
            string[] names = { "../outside.txt", "CON", "same.txt", "same.txt", new string('\u4e2d', 250) + ".txt" };
            foreach (string name in names) document.Attachments.Add(new EmailAttachment { FileName = name, Content = new byte[] { 1, 2, 3 } });
            var result = EmailAttachmentExtractor.Extract(document, root);
            Assert.False(result.Truncated);
            Assert.Equal(15, result.BytesRead);
            Assert.Equal(names.Length, result.Entries.Select(item => item.OutputPath).Distinct().Count());
            foreach (var entry in result.Entries) {
                Assert.Equal(root, Path.GetDirectoryName(entry.OutputPath));
                Assert.True(Encoding.UTF8.GetByteCount(Path.GetFileName(entry.OutputPath)!) < 218);
                Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(entry.OutputPath!));
                using var hash = SHA256.Create();
                Assert.Equal(BitConverter.ToString(hash.ComputeHash(new byte[] { 1, 2, 3 })).Replace("-", "").ToLowerInvariant(), entry.Sha256);
                Assert.Empty(entry.Diagnostics);
            }
            var second = EmailAttachmentExtractor.Extract(document, root);
            Assert.All(second.Entries, item => { Assert.Null(item.OutputPath); Assert.Contains(item.Diagnostics, d => d.Code == "EMAIL_EXTRACTION_FAILED"); });
            Assert.Equal(names.Length, Directory.GetFiles(root).Length);
        });
    }

    [Fact]
    public void EmbeddedMessagesReopenAndNestedPathsAreRecordedWithoutMutatingSource() {
        InDirectory(root => {
            var nested = new EmailDocument { Subject = "Nested" };
            nested.Body.Text = "Body";
            nested.Attachments.Add(new EmailAttachment { FileName = "nested.txt", Content = Encoding.UTF8.GetBytes("payload") });
            var parent = new EmailDocument();
            parent.Attachments.Add(new EmailAttachment { FileName = "original.msg", EmbeddedDocument = nested });
            var result = EmailAttachmentExtractor.Extract(parent, root, new EmailAttachmentExtractionOptions(recurseEmbeddedMessages: true));
            Assert.Equal(new[] { "0", "0/0" }, result.Entries.Select(item => item.SourcePath));
            using var read = new EmailDocumentReader().Read(result.Entries[0].OutputPath!);
            Assert.Equal("Nested", read.Document.Subject);
            using var hash = SHA256.Create();
            Assert.Equal(BitConverter.ToString(hash.ComputeHash(File.ReadAllBytes(result.Entries[0].OutputPath!))).Replace("-", "").ToLowerInvariant(), result.Entries[0].Sha256);
            Assert.Equal("payload", File.ReadAllText(result.Entries[1].OutputPath!));
            Assert.Same(nested, parent.Attachments[0].EmbeddedDocument);
            Assert.Empty(Directory.GetFiles(root, ".officeimo-*"));
        });
    }

    [Fact]
    public void UnknownLengthSourceStopsAtBudgetAndDisposesWithoutCommittingPartialFile() {
        InDirectory(root => {
            var content = new CountingSource();
            var document = new EmailDocument();
            document.Attachments.Add(new EmailAttachment { FileName = "large.dat", ContentSource = content });
            document.Attachments.Add(new EmailAttachment { FileName = "later.dat", Content = new byte[] { 1 } });
            var result = EmailAttachmentExtractor.Extract(document, root, new EmailAttachmentExtractionOptions(maxAttachmentBytes: 20, maxTotalBytes: 20));
            Assert.True(result.Truncated);
            Assert.Null(Assert.Single(result.Entries).OutputPath);
            Assert.Equal(21, result.BytesRead);
            Assert.Equal(21, content.Stream!.BytesRead);
            Assert.True(content.Stream.Disposed);
            Assert.Empty(Directory.GetFiles(root));
        });
    }

    [Fact]
    public void MissingLinkedAndHiddenContentIsVisibleAndDepthAndCountLimitsStopTraversal() {
        InDirectory(root => {
            var nested = new EmailDocument();
            nested.Attachments.Add(new EmailAttachment { Content = new byte[] { 1 } });
            var document = new EmailDocument();
            document.Attachments.Add(new EmailAttachment { LinkedPath = "must-not-open" });
            document.Attachments.Add(new EmailAttachment { IsHidden = true, Content = new byte[] { 2 } });
            document.Attachments.Add(new EmailAttachment { EmbeddedDocument = nested });
            var result = EmailAttachmentExtractor.Extract(document, root, new EmailAttachmentExtractionOptions(maxDepth: 0, recurseEmbeddedMessages: true));
            Assert.Equal(3, result.Entries.Count);
            Assert.Contains(result.Entries[0].Diagnostics, d => d.Code == "EMAIL_EXTRACTION_CONTENT_UNAVAILABLE");
            Assert.Contains(result.Entries[1].Diagnostics, d => d.Code == "EMAIL_EXTRACTION_SKIPPED");
            Assert.True(result.Truncated);
            Assert.Contains(result.Diagnostics, d => d.Code == "EMAIL_EXTRACTION_DEPTH");
            var limited = EmailAttachmentExtractor.Extract(document, root, new EmailAttachmentExtractionOptions(maxAttachments: 1));
            Assert.Single(limited.Entries);
            Assert.True(limited.Truncated);
        });
    }

    [Fact]
    public void CancellationCleansPartialOutputAndClosesContentSource() {
        InDirectory(root => {
            using var cancellation = new CancellationTokenSource();
            var content = new CountingSource(cancellation);
            var document = new EmailDocument();
            document.Attachments.Add(new EmailAttachment { ContentSource = content });
            Assert.Throws<OperationCanceledException>(() => EmailAttachmentExtractor.Extract(document, root, cancellationToken: cancellation.Token));
            Assert.True(content.Stream!.Disposed);
            Assert.Empty(Directory.GetFiles(root));
        });
    }

    [Fact]
    public void HiddenEmbeddedMessageDoesNotLeakNestedAttachments() {
        InDirectory(root => {
            var hidden = new EmailDocument();
            hidden.Attachments.Add(new EmailAttachment { Content = new byte[] { 1 } });
            var document = new EmailDocument();
            document.Attachments.Add(new EmailAttachment { IsHidden = true, EmbeddedDocument = hidden });
            var result = EmailAttachmentExtractor.Extract(document, root, new EmailAttachmentExtractionOptions(recurseEmbeddedMessages: true));
            Assert.Null(Assert.Single(result.Entries).OutputPath);
            Assert.Empty(Directory.GetFiles(root));
        });
    }

    private static void InDirectory(Action<string> action) {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO.Email.Extraction.Tests." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try { action(root); } finally { Directory.Delete(root, true); }
    }

    private sealed class ContextAwareSource : IEmailContentSource {
        public long? Length => null;
        public bool SawCallerContext { get; private set; }
        public int Opens { get; private set; }
        public Stream OpenRead() => new MemoryStream(new byte[] { 1, 2, 3 });
        public async Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            SawCallerContext = SynchronizationContext.Current != null;
            // Fail deterministically instead of hanging the test on a blocked UI context.
            if (SawCallerContext) throw new InvalidOperationException("An awaited source would need the blocked caller context.");
            await Task.Yield();
            cancellationToken.ThrowIfCancellationRequested();
            Opens++;
            return OpenRead();
        }
    }

    private sealed class CountingSource : IEmailContentSource {
        private readonly CancellationTokenSource? _cancellation;
        private readonly int _bytes;
        private readonly bool _seekable;
        public CountingSource(CancellationTokenSource? cancellation = null, int bytes = 100, bool seekable = true) {
            _cancellation = cancellation; _bytes = bytes; _seekable = seekable;
        }
        public CountingStream? Stream { get; private set; }
        public int OpenCount { get; private set; }
        public long? Length => null;
        public Stream OpenRead() { OpenCount++; return Stream = new CountingStream(_cancellation, _bytes, _seekable); }
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) => Task.FromResult(OpenRead());
    }

    private sealed class CountingStream : MemoryStream {
        private readonly CancellationTokenSource? _cancellation;
        private readonly bool _seekable;
        public CountingStream(CancellationTokenSource? cancellation, int bytes, bool seekable) : base(new byte[bytes]) {
            _cancellation = cancellation; _seekable = seekable;
        }
        public override bool CanSeek => _seekable;
        public int BytesRead { get; private set; }
        public bool Disposed { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            BytesRead += read;
            _cancellation?.Cancel();
            return read;
        }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(Read(buffer, offset, count));
        }
        protected override void Dispose(bool disposing) { Disposed = true; base.Dispose(disposing); }
    }
}
