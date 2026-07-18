using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Email;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEmailImageExportTests {
    private static readonly byte[] PixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+X8m0WQAAAABJRU5ErkJggg==");

    [Fact]
    public void PlainTextEmailExportsThroughHtmlWithMessageChrome() {
        var email = new EmailDocument {
            Subject = "Quarterly update",
            From = new EmailAddress("sender@example.com", "Sender"),
            Date = new DateTimeOffset(2026, 7, 18, 12, 0, 0, TimeSpan.Zero)
        };
        email.Recipients.Add(new EmailRecipient(
            EmailRecipientKind.To,
            new EmailAddress("reader@example.com", "Reader")));
        email.Body.Text = "Hello,\nThis is a useful plain-text message.";

        OfficeImageExportResult result = email.ExportImage(
            OfficeImageExportFormat.Svg);

        Assert.Equal(OfficeImageExportFormat.Svg, result.Format);
        string svg = System.Text.Encoding.UTF8.GetString(result.Bytes);
        Assert.Contains("Quarterly update", svg);
        Assert.Contains("Hello,", svg);
        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_IMAGE_BODY_MISSING");
    }

    [Fact]
    public async Task HtmlEmailResolvesInlineContentIdImagesAsynchronously() {
        var email = new EmailDocument { Subject = "Inline image" };
        email.Body.Html = "<p>Logo</p><img src=\"cid:logo@example\" alt=\"Logo\">";
        email.Attachments.Add(new EmailAttachment {
            FileName = "logo.png",
            ContentType = "image/png",
            ContentId = "logo@example",
            IsInline = true,
            Content = PixelPng,
            Length = PixelPng.Length
        });

        OfficeImageExportResult result = await email.ExportImageAsync(
            OfficeImageExportFormat.Png);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalImagePending ||
                          diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable ||
                          diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
    }

    [Fact]
    public void HtmlEmailResolvesRetainedInlineContentIdImagesSynchronously() {
        var email = new EmailDocument { Subject = "Inline image" };
        email.Body.Html = "<p>Logo</p><img src=\"cid:logo@example\" alt=\"Logo\">";
        email.Attachments.Add(new EmailAttachment {
            FileName = "logo.png",
            ContentType = "image/png",
            ContentId = "logo@example",
            IsInline = true,
            Content = PixelPng,
            Length = PixelPng.Length
        });

        OfficeImageExportResult result = email.ExportImage(
            OfficeImageExportFormat.Png);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalImagePending ||
                          diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable ||
                          diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EmailRetainedResourceByteLimitHasPreciseDiagnostic(
        bool asynchronous) {
        var email = new EmailDocument { Subject = "Oversized inline image" };
        email.Body.Html = "<img src=\"cid:logo@example\" alt=\"Logo\">";
        email.Attachments.Add(new EmailAttachment {
            FileName = "logo.png",
            ContentType = "image/png",
            ContentId = "logo@example",
            IsInline = true,
            Content = PixelPng,
            Length = PixelPng.Length
        });
        var options = new EmailImageExportOptions {
            MaxResourceBytes = PixelPng.Length - 1L
        };

        OfficeImageExportResult result = asynchronous
            ? await email.ExportImageAsync(
                OfficeImageExportFormat.Png,
                options)
            : email.ExportImage(
                OfficeImageExportFormat.Png,
                options);

        Assert.Contains(
            result.Diagnostics,
            diagnostic => diagnostic.Code ==
                          HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded);
        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code ==
                          HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public void EmailSyncBatchCancellationReachesRetainedResourceRead() {
        using var cancellation = new CancellationTokenSource();
        var email = new EmailDocument { Subject = "Cancelable inline image" };
        email.Body.Html = "<img src=\"cid:logo@example\" alt=\"Logo\">";
        email.Attachments.Add(new EmailAttachment {
            FileName = "logo.png",
            ContentType = "image/png",
            ContentId = "logo@example",
            IsInline = true,
            ContentSource = new CancelOnReadContentSource(
                PixelPng,
                cancellation),
            Length = PixelPng.Length
        });

        Assert.ThrowsAny<OperationCanceledException>(() =>
            email.ExportImages(
                OfficeImageExportFormat.Png,
                _ => { },
                cancellationToken: cancellation.Token));
    }

    [Fact]
    public void EmailRtfProjectionParticipatesInNoLossPolicy() {
        var email = new EmailDocument();
        email.Body.Rtf = "{\\rtf1\\ansi Rendered RTF body}";
        var options = new EmailImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoLoss = true }
        };

        OfficeImageExportPolicyException exception =
            Assert.Throws<OfficeImageExportPolicyException>(() =>
                email.ExportImage(
                    OfficeImageExportFormat.Png,
                    options));

        Assert.Contains(
            exception.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_IMAGE_RTF_BODY_PROJECTED");
    }

    [Fact]
    public async Task EmailInlineResolverDoesNotWeakenFallbackUrlPolicy() {
        var email = new EmailDocument { Subject = "Policy boundary" };
        email.Body.Html = "<img src=\"file:///private/secret.png\" alt=\"blocked\">";
        int fallbackCalls = 0;
        var options = new EmailImageExportOptions {
            UrlPolicy = new HtmlUrlPolicy {
                RestrictUrlSchemes = false,
                DisallowFileUrls = true
            },
            ResourceResolver = (request, cancellationToken) => {
                cancellationToken.ThrowIfCancellationRequested();
                fallbackCalls++;
                return Task.FromResult<HtmlResolvedResource?>(
                    new HtmlResolvedResource(PixelPng, "image/png"));
            }
        };

        await email.ExportImageAsync(OfficeImageExportFormat.Png, options);

        Assert.Equal(0, fallbackCalls);
    }

    [Fact]
    public async Task FluentEmailBatchSaveStreamsPagesAndReturnsPayloadFreeMetadata() {
        var email = new EmailDocument { Subject = "Paged message" };
        email.Body.Html =
            "<h1>Message</h1><p>First page</p>" +
            "<section style=\"break-before:page\"><h2>Continued</h2>" +
            "<p>Second page</p></section>";
        string folder = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO.EmailImages",
            Guid.NewGuid().ToString("N"));
        try {
            OfficeImageExportBatchSaveResult saved = await email
                .ToImages(new EmailImageExportOptions {
                    PageSize = new OfficePageSize(
                        360D / HtmlRenderOptions.CssPixelsPerInch,
                        260D / HtmlRenderOptions.CssPixelsPerInch)
                })
                .Paged()
                .AsPng()
                .WithBatchLimits(20, 20_000_000, 20_000_000)
                .SaveFilesAsync(folder);

            Assert.True(saved.Files.Count > 1);
            Assert.All(saved.Files, file => {
                Assert.True(File.Exists(file.Path));
                Assert.Equal(OfficeImageExportFormat.Png, file.Format);
                Assert.True(file.EncodedLength > 0);
            });
        } finally {
            if (Directory.Exists(folder)) {
                Directory.Delete(folder, recursive: true);
            }
        }
    }

    private sealed class CancelOnReadContentSource : IEmailContentSource {
        private readonly byte[] _bytes;
        private readonly CancellationTokenSource _cancellation;

        internal CancelOnReadContentSource(
            byte[] bytes,
            CancellationTokenSource cancellation) {
            _bytes = (byte[])bytes.Clone();
            _cancellation = cancellation;
        }

        public long? Length => _bytes.LongLength;

        public Stream OpenRead() =>
            new CancelOnReadStream(_bytes, _cancellation);

        public Task<Stream> OpenReadAsync(
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(OpenRead());
        }
    }

    private sealed class CancelOnReadStream : Stream {
        private readonly MemoryStream _inner;
        private readonly CancellationTokenSource _cancellation;

        internal CancelOnReadStream(
            byte[] bytes,
            CancellationTokenSource cancellation) {
            _inner = new MemoryStream(bytes, writable: false);
            _cancellation = cancellation;
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
            int read = _inner.Read(buffer, offset, count);
            if (read > 0) _cancellation.Cancel();
            return read;
        }

        public override void Flush() {
        }

        public override long Seek(long offset, SeekOrigin origin) =>
            throw new NotSupportedException();

        public override void SetLength(long value) =>
            throw new NotSupportedException();

        public override void Write(
            byte[] buffer,
            int offset,
            int count) =>
            throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
