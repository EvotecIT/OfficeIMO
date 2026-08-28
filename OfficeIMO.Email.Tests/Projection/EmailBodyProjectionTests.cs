using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailBodyProjectionTests {
    [Fact]
    public void ImageDocument_DiscoversAndRewritesOnlyActualSrcAttributes() {
        const string html = "<p data-src='not-an-image'>before</p><img data-src='lazy.png' SRC='one.png'><p>one.png</p>";

        EmailHtmlImageDocument document = EmailHtmlImageDocument.Parse(html);

        EmailHtmlImageReference image = Assert.Single(document.Images);
        Assert.Equal("one.png", image.Source);
        document.SetImageSource(image.Index, "cid:one.png");
        string rendered = document.ToHtml();
        Assert.Contains("data-src=\"lazy.png\"", rendered, StringComparison.Ordinal);
        Assert.Contains("src=\"cid:one.png\"", rendered, StringComparison.Ordinal);
        Assert.Contains(">one.png</p>", rendered, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageDocument_RewritesDuplicateSourcesIndependently() {
        EmailHtmlImageDocument document = EmailHtmlImageDocument.Parse(
            "<img src='same.png'><img src='same.png'>");

        Assert.Equal(2, document.Images.Count);
        document.SetImageSource(0, "cid:first.png");
        document.SetImageSource(1, "cid:second.png");

        string rendered = document.ToHtml();
        Assert.Contains("src=\"cid:first.png\"", rendered, StringComparison.Ordinal);
        Assert.Contains("src=\"cid:second.png\"", rendered, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageDocument_PreservesDocumentEnvelopeWhenPresent() {
        EmailHtmlImageDocument document = EmailHtmlImageDocument.Parse(
            "<html><head><title>Mail</title></head><body><img src='one.png'></body></html>");

        string rendered = document.ToHtml();

        Assert.StartsWith("<html", rendered, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<head>", rendered, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<body>", rendered, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("<!-- <html><body>comment</body></html> --><img src='one.png'>")]
    [InlineData("<bodyguard><img src='one.png'></bodyguard>")]
    [InlineData("prefix &lt;html without a tag <img src='one.png'>")]
    [InlineData("<template><body>template text</body></template><img src='one.png'>")]
    public void ImageDocument_DoesNotTreatEnvelopeLikeTextAsDocument(string html) {
        EmailHtmlImageDocument document = EmailHtmlImageDocument.Parse(html);

        document.SetImageSource(0, "cid:one.png");
        string rendered = document.ToHtml();

        Assert.DoesNotContain("<html", rendered, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("src=\"cid:one.png\"", rendered, StringComparison.Ordinal);
    }

    [Fact]
    public void Sanitizes_html_blocks_remote_resources_and_resolves_embedded_content_once() {
        var document = new EmailDocument();
        document.Body.Html = "<html><body onload='alert(1)'><script>alert(2)</script>" +
            "<img src='https://tracking.example/pixel'><img src='cid:logo@example.test'>" +
            "<a href='javascript:alert(3)'>unsafe</a></body></html>";
        document.Attachments.Add(new EmailAttachment {
            FileName = "logo.png",
            ContentType = "image/png",
            ContentId = "<logo@example.test>",
            ContentLocation = "images/logo.png",
            IsInline = true,
            Content = new byte[] { 1, 2, 3 },
            Length = 3
        });

        EmailBodyProjectionResult result = EmailBodyProjection.Create(document);

        Assert.Equal(EmailBodySourceKind.Html, result.SourceKind);
        Assert.DoesNotContain("<script", result.Html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("onload", result.Html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("javascript:", result.Html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("tracking.example", result.Html, StringComparison.OrdinalIgnoreCase);
        string prepared = result.Document.CreateDocumentForConversion().DocumentElement.OuterHtml;
        Assert.DoesNotContain("<script", prepared, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("onload", prepared, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("tracking.example", prepared, StringComparison.OrdinalIgnoreCase);
        EmailBodyResource cid = Assert.IsType<EmailBodyResource>(
            result.ResolveResource("cid:logo@example.test"));
        Assert.Same(cid, result.ResolveResource("images/logo.png"));
        Assert.Equal(new byte[] { 1, 2, 3 }, cid.ReadAllBytes());
        Assert.Equal(new byte[] { 1, 2, 3 }, cid.ReadAllBytes());
    }

    [Fact]
    public void Applies_consumer_selection_without_duplicating_Rtf_fallback_logic() {
        var document = new EmailDocument();
        document.Body.Text = "plain choice";
        document.Body.Html = "<p>html choice</p>";
        document.Body.Rtf = @"{\rtf1\ansi rtf choice}";

        EmailBodyProjectionResult reader = EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions {
                SelectionPolicy = EmailBodySelectionPolicy.PlainTextFirst
            });
        EmailBodyProjectionResult renderer = EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions {
                SelectionPolicy = EmailBodySelectionPolicy.Richest
            });
        EmailBodyProjectionResult rtf = EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions {
                SelectionPolicy = EmailBodySelectionPolicy.RtfFirst
            });

        Assert.Equal(EmailBodySourceKind.PlainText, reader.SourceKind);
        Assert.Contains("plain choice", reader.Html, StringComparison.Ordinal);
        Assert.Equal(EmailBodySourceKind.Html, renderer.SourceKind);
        Assert.Contains("html choice", renderer.Html, StringComparison.Ordinal);
        Assert.Equal(EmailBodySourceKind.Rtf, rtf.SourceKind);
        Assert.Contains("rtf choice", rtf.Html, StringComparison.Ordinal);
        Assert.Contains(rtf.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_BODY_RTF_PROJECTED");
    }

    [Fact]
    public void Enforces_bounded_operation_scoped_attachment_reads() {
        var document = new EmailDocument { Body = { Html = "<img src='cid:large'>" } };
        document.Attachments.Add(new EmailAttachment {
            ContentId = "large",
            IsInline = true,
            Content = new byte[] { 1, 2, 3, 4 },
            Length = 4
        });
        EmailBodyProjectionResult result = EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions { MaxResourceBytes = 3 });

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            Assert.IsType<EmailBodyResource>(result.ResolveResource("cid:large")).ReadAllBytes());

        Assert.Equal("EmailBodyProjectionOptions.MaxResourceBytes", exception.LimitName);
    }

    [Fact]
    public void Rejects_resource_count_before_opening_attachment_content() {
        EmailDocument document = CreateInlineResourceDocument(
            new byte[] { 1 },
            new byte[] { 2 });

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            EmailBodyProjection.Create(document,
                new EmailBodyProjectionOptions { MaxResourceCount = 1 }));

        Assert.Equal("EmailBodyProjectionOptions.MaxResourceCount", exception.LimitName);
        Assert.Equal(2, exception.ActualValue);
    }

    [Fact]
    public void Rejects_declared_aggregate_resource_size() {
        EmailDocument document = CreateInlineResourceDocument(
            new byte[] { 1, 2, 3 },
            new byte[] { 4, 5, 6 });

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            EmailBodyProjection.Create(document,
                new EmailBodyProjectionOptions { MaxTotalResourceBytes = 5 }));

        Assert.Equal("EmailBodyProjectionOptions.MaxTotalResourceBytes", exception.LimitName);
        Assert.Equal(6, exception.ActualValue);
    }

    [Fact]
    public void Open_stream_enforces_actual_size_when_declared_length_is_unknown() {
        EmailDocument document = CreateInlineResourceDocument(new byte[] { 1, 2, 3, 4 });
        document.Attachments[0].Length = 0;
        EmailBodyResource resource = EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions { MaxResourceBytes = 3 }).Resources[0];

        using Stream source = resource.OpenReadStream();
        using var output = new MemoryStream();
        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            source.CopyTo(output));

        Assert.Equal("EmailBodyProjectionOptions.MaxResourceBytes", exception.LimitName);
        Assert.Equal(4, exception.ActualValue);
    }

    [Fact]
    public async Task Async_copies_share_one_projection_wide_resource_budget() {
        EmailDocument document = CreateInlineResourceDocument(
            new byte[] { 1, 2, 3 },
            new byte[] { 4, 5, 6 });
        document.Attachments[0].Length = 0;
        document.Attachments[1].Length = 0;
        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions { MaxTotalResourceBytes = 5 });

        using (var first = new MemoryStream()) {
            await projection.Resources[0].CopyToAsync(first);
            Assert.Equal(new byte[] { 1, 2, 3 }, first.ToArray());
        }
        using var second = new MemoryStream();
        EmailLimitExceededException exception = await Assert.ThrowsAsync<EmailLimitExceededException>(() =>
            projection.Resources[1].CopyToAsync(second));

        Assert.Equal("EmailBodyProjectionOptions.MaxTotalResourceBytes", exception.LimitName);
        Assert.Equal(6, exception.ActualValue);
    }

    [Fact]
    public async Task Open_stream_honors_operation_and_read_cancellation() {
        EmailDocument document = CreateInlineResourceDocument(new byte[] { 1, 2, 3 });
        EmailBodyResource resource = EmailBodyProjection.Create(document).Resources[0];
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => resource.OpenReadStream(cancellation.Token));
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            resource.OpenReadStreamAsync(cancellation.Token));
    }

    [Fact]
    public void Direct_reads_never_write_rejected_bytes_and_poison_an_oversized_resource() {
        var content = new TrackingContentSource(new byte[] { 1, 2, 3, 4 });
        EmailDocument document = CreateInlineResourceDocument(new byte[] { 0 });
        document.Attachments[0].Content = null;
        document.Attachments[0].ContentSource = content;
        document.Attachments[0].Length = 0;
        EmailBodyResource resource = EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions { MaxResourceBytes = 3 }).Resources[0];
        var buffer = new byte[] { 9, 9, 9, 9 };

        using Stream source = resource.OpenReadStream();
        Assert.Equal(3, source.Read(buffer, 0, buffer.Length));
        Assert.Equal(new byte[] { 1, 2, 3, 9 }, buffer);
        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            source.Read(buffer, 0, buffer.Length));
        int readsAfterFailure = content.ReadCount;

        Assert.Equal("EmailBodyProjectionOptions.MaxResourceBytes", exception.LimitName);
        Assert.Throws<EmailLimitExceededException>(() => source.Read(buffer, 0, buffer.Length));
        Assert.Throws<EmailLimitExceededException>(() => resource.OpenReadStream());
        Assert.Equal(readsAfterFailure, content.ReadCount);
        source.Dispose();
        Assert.Equal(1, content.DisposeCount);
    }

    [Fact]
    public void Aggregate_failure_poisoning_prevents_reopening_other_resources() {
        var firstContent = new TrackingContentSource(new byte[] { 1, 2, 3 });
        var secondContent = new TrackingContentSource(new byte[] { 4, 5, 6 });
        EmailDocument document = CreateInlineResourceDocument(new byte[] { 0 }, new byte[] { 0 });
        document.Attachments[0].Content = null;
        document.Attachments[0].ContentSource = firstContent;
        document.Attachments[0].Length = 0;
        document.Attachments[1].Content = null;
        document.Attachments[1].ContentSource = secondContent;
        document.Attachments[1].Length = 0;
        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions { MaxTotalResourceBytes = 3 });

        Assert.Equal(new byte[] { 1, 2, 3 }, projection.Resources[0].ReadAllBytes());
        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            projection.Resources[1].ReadAllBytes());
        int opensAfterFailure = secondContent.OpenCount;
        int readsAfterFailure = secondContent.ReadCount;

        Assert.Equal("EmailBodyProjectionOptions.MaxTotalResourceBytes", exception.LimitName);
        Assert.Throws<EmailLimitExceededException>(() => projection.Resources[1].OpenReadStream());
        Assert.Equal(opensAfterFailure, secondContent.OpenCount);
        Assert.Equal(readsAfterFailure, secondContent.ReadCount);
    }

    [Fact]
    public async Task Operation_cancellation_interrupts_an_active_async_source_read() {
        var content = new BlockingContentSource();
        EmailDocument document = CreateInlineResourceDocument(new byte[] { 0 });
        document.Attachments[0].Content = null;
        document.Attachments[0].ContentSource = content;
        document.Attachments[0].Length = 0;
        EmailBodyResource resource = EmailBodyProjection.Create(document).Resources[0];
        using var cancellation = new CancellationTokenSource();
        using Stream source = await resource.OpenReadStreamAsync(cancellation.Token);
        Task<int> read = source.ReadAsync(new byte[1], 0, 1, CancellationToken.None);
        await content.ReadStarted;

        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => read);
    }

    [Fact]
    public async Task Concurrent_reads_wait_for_reservations_before_deciding_aggregate_exhaustion() {
        var firstContent = new ControlledContentSource(new byte[] { 1 });
        var secondContent = new TrackingContentSource(new byte[] { 2 });
        EmailDocument document = CreateInlineResourceDocument(new byte[] { 0 }, new byte[] { 0 });
        document.Attachments[0].Content = null;
        document.Attachments[0].ContentSource = firstContent;
        document.Attachments[0].Length = 0;
        document.Attachments[1].Content = null;
        document.Attachments[1].ContentSource = secondContent;
        document.Attachments[1].Length = 0;
        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions { MaxTotalResourceBytes = 2 });
        using Stream first = await projection.Resources[0].OpenReadStreamAsync();
        using Stream second = await projection.Resources[1].OpenReadStreamAsync();
        var firstBuffer = new byte[2];
        var secondBuffer = new byte[1];

        Task<int> firstRead = first.ReadAsync(firstBuffer, 0, firstBuffer.Length);
        await firstContent.ReadStarted;
        Task<int> secondRead = second.ReadAsync(secondBuffer, 0, secondBuffer.Length);
        Assert.False(secondRead.IsCompleted);

        firstContent.ReleaseRead();

        Assert.Equal(1, await firstRead);
        Assert.Equal(1, await secondRead);
        Assert.Equal(1, firstBuffer[0]);
        Assert.Equal(2, secondBuffer[0]);
    }

    [Fact]
    public void Resource_indexing_can_be_disabled_for_body_only_consumers() {
        EmailDocument document = CreateInlineResourceDocument(
            new byte[] { 1, 2 },
            new byte[] { 3, 4 });

        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions {
                IncludeResources = false,
                MaxResourceCount = 1,
                MaxTotalResourceBytes = 1
            });

        Assert.Empty(projection.Resources);
    }

    private static EmailDocument CreateInlineResourceDocument(params byte[][] resources) {
        var document = new EmailDocument { Body = { Html = "<p>inline resources</p>" } };
        for (int index = 0; index < resources.Length; index++) {
            byte[] content = resources[index];
            document.Attachments.Add(new EmailAttachment {
                FileName = $"resource-{index}.bin",
                ContentId = $"resource-{index}",
                ContentType = "application/octet-stream",
                IsInline = true,
                Content = content,
                Length = content.LongLength
            });
        }
        return document;
    }

    private sealed class TrackingContentSource : IEmailContentSource {
        private readonly byte[] _content;

        internal TrackingContentSource(byte[] content) {
            _content = (byte[])content.Clone();
        }

        public long? Length => null;
        internal int OpenCount { get; private set; }
        internal int ReadCount { get; private set; }
        internal int DisposeCount { get; private set; }

        public Stream OpenRead() {
            OpenCount++;
            return new TrackingReadStream(
                _content,
                () => ReadCount++,
                () => DisposeCount++);
        }

        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(OpenRead());
        }
    }

    private sealed class TrackingReadStream : Stream {
        private readonly MemoryStream _inner;
        private readonly Action _onRead;
        private readonly Action _onDispose;
        private bool _disposed;

        internal TrackingReadStream(byte[] content, Action onRead, Action onDispose) {
            _inner = new MemoryStream(content, writable: false);
            _onRead = onRead;
            _onDispose = onDispose;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) {
            _onRead();
            return _inner.Read(buffer, offset, count);
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(Read(buffer, offset, count));
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing && !_disposed) {
                _disposed = true;
                _inner.Dispose();
                _onDispose();
            }
            base.Dispose(disposing);
        }
    }

    private sealed class BlockingContentSource : IEmailContentSource {
        private readonly TaskCompletionSource<bool> _readStarted =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        public long? Length => null;
        internal Task ReadStarted => _readStarted.Task;
        public Stream OpenRead() => new BlockingReadStream(_readStarted);
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(OpenRead());
        }
    }

    private sealed class ControlledContentSource : IEmailContentSource {
        private readonly byte[] _content;
        private readonly TaskCompletionSource<bool> _readStarted =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<bool> _releaseRead =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        internal ControlledContentSource(byte[] content) {
            _content = (byte[])content.Clone();
        }

        public long? Length => null;
        internal Task ReadStarted => _readStarted.Task;
        internal void ReleaseRead() => _releaseRead.TrySetResult(true);
        public Stream OpenRead() => new ControlledReadStream(
            _content,
            _readStarted,
            _releaseRead.Task);
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(OpenRead());
        }
    }

    private sealed class BlockingReadStream : Stream {
        private readonly TaskCompletionSource<bool> _readStarted;

        internal BlockingReadStream(TaskCompletionSource<bool> readStarted) {
            _readStarted = readStarted;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException();

        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) {
            _readStarted.TrySetResult(true);
            await Task.Delay(Timeout.Infinite, cancellationToken);
            return 0;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class ControlledReadStream : Stream {
        private readonly MemoryStream _inner;
        private readonly TaskCompletionSource<bool> _readStarted;
        private readonly Task _releaseRead;

        internal ControlledReadStream(
            byte[] content,
            TaskCompletionSource<bool> readStarted,
            Task releaseRead) {
            _inner = new MemoryStream(content, writable: false);
            _readStarted = readStarted;
            _releaseRead = releaseRead;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException();

        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) {
            _readStarted.TrySetResult(true);
            await WaitWithCancellationAsync(_releaseRead, cancellationToken);
            return _inner.Read(buffer, offset, count);
        }

        private static async Task WaitWithCancellationAsync(
            Task operation,
            CancellationToken cancellationToken) {
            if (!cancellationToken.CanBeCanceled) {
                await operation;
                return;
            }
            var canceled = new TaskCompletionSource<bool>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            using (cancellationToken.Register(() => canceled.TrySetCanceled())) {
                await await Task.WhenAny(operation, canceled.Task);
            }
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
