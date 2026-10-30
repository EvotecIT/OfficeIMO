using System;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Email;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEmailImageExportTests {
    private static readonly byte[] PixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg==");

    [Fact]
    public void EmailFitWithinBoundsHighRequestedScaleBeforeHtmlSurfaceValidation() {
        var email = new EmailDocument { Subject = "Bounded message" };
        email.Body.Html = "<h1>Bounded</h1><p>High requested scale.</p>";

        var imageOptions1 = new EmailImageExportOptions();
        imageOptions1.Scale = 100D;
        imageOptions1.TargetDpi = null;
        imageOptions1.MaximumOutputWidth = 360;
        imageOptions1.MaximumOutputHeight = 360;
        imageOptions1.Format = OfficeImageExportFormat.Png;
        OfficeImageExportResult result = email.ExportImage(OfficeImageExportFormat.Png, imageOptions1, pageIndex: imageOptions1.PageIndex);

        Assert.True(result.Width <= 360);
        Assert.True(result.Height <= 360);
    }

    [Fact]
    public void EmailRtfBatchDiagnosticsPreserveSequenceMetadata() {
        var email = new EmailDocument();
        email.Body.Rtf = "{\\rtf1\\ansi Rendered RTF body}";

        OfficeImageExportResult result = Assert.Single(email.ExportImages(OfficeImageExportFormat.Png));

        Assert.Equal(0, result.SequenceIndex);
        Assert.Equal(1, result.SequenceCount);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_IMAGE_RTF_BODY_PROJECTED");
    }

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
        Assert.Contains("Quarterly update", SvgPaintedText.Read(svg));
        Assert.Contains("Hello,", SvgPaintedText.Read(svg));
        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_IMAGE_BODY_MISSING");
    }

    [Fact]
    public void EmailMessageChromeUsesTheConfiguredCallerScopedDefaultFont() {
        var email = new EmailDocument { Subject = "A" };
        email.Body.Html = "<p style='font-family:Arial'>A</p>";
        var options = new EmailImageExportOptions {
            DefaultFontFamily = ManagedTextShapingTestAssets.FamilyName
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));

        OfficeImageExportResult result = email.ExportImage(
            OfficeImageExportFormat.Svg,
            options);

        XDocument svg = XDocument.Parse(System.Text.Encoding.UTF8.GetString(result.Bytes));
        XElement[] slices = svg.Descendants().Where(element => element.Name.LocalName == "text"
            || element.Name.LocalName == "g" && element.Attribute("role")?.Value == "img"
                && element.Attribute("aria-label") != null
                && element.Descendants().Any(child => child.Name.LocalName == "path")).ToArray();
        Assert.Equal(2, slices.Length);
        Assert.Equal("AA", SvgPaintedText.Read(svg.ToString()));
        XElement subject = slices[0], body = slices[1];
        Assert.Equal("A", subject.Attribute("aria-label")?.Value);
        OfficeDrawingShape subjectInk = ReadInk(subject);

        Assert.True(options.Fonts.TryResolveFaceForText("A", ManagedTextShapingTestAssets.FamilyName,
            OfficeFontFaceDescriptor.Regular, out OfficeFontFace? face));
        Assert.True(face!.Program.TryGetGlyphMetrics('A', out _, out int advance));
        Assert.Equal(500, advance);
        Assert.Equal(1000, face.Program.UnitsPerEm);
        const double subjectSizePoints = 24D * 72D / HtmlRenderOptions.CssPixelsPerInch;
        OfficePoint[] expected = face.Program.GetTextContours("A", 0D, 0D, subjectSizePoints)
            .SelectMany(contour => contour).ToArray();
        Assert.NotEmpty(expected);
        Assert.Equal(expected.Max(point => point.X) - expected.Min(point => point.X), subjectInk.Shape.Width, 3);
        Assert.Equal(expected.Max(point => point.Y) - expected.Min(point => point.Y), subjectInk.Shape.Height, 3);
        if (body.Name.LocalName == "text") {
            Assert.True(subjectInk.Y < double.Parse(body.Attribute("y")!.Value, CultureInfo.InvariantCulture));
            Assert.DoesNotContain(ManagedTextShapingTestAssets.FamilyName,
                OfficeFontFamilyParser.Parse(body.Attribute("font-family")?.Value));
        } else {
            OfficeDrawingShape bodyInk = ReadInk(body);
            Assert.True(subjectInk.Y < bodyInk.Y);
            // The explicit Arial body must not become the supplied rectangular glyph.
            Assert.True(bodyInk.Shape.PathCommands.Count != subjectInk.Shape.PathCommands.Count
                || Math.Abs(bodyInk.Shape.Width / 16D - subjectInk.Shape.Width / 24D) > 0.001D
                || Math.Abs(bodyInk.Shape.Height / 16D - subjectInk.Shape.Height / 24D) > 0.001D);
        }

        OfficeDrawingShape ReadInk(XElement projection) {
            XElement content = new XElement(projection);
            foreach (XElement ancestor in projection.Ancestors().TakeWhile(element => element != svg.Root))
                content = new XElement(ancestor.Name, ancestor.Attributes(), content);
            var isolated = new XElement(svg.Root!.Name,
                new XAttribute("width", "4096"), new XAttribute("height", "4096"),
                new XAttribute("viewBox", "0 0 4096 4096"),
                svg.Descendants().Where(element => element.Name.LocalName == "defs")
                    .Select(element => new XElement(element)), content);
            OfficeDrawing drawing = HtmlSvgPaintAssertions.Read(System.Text.Encoding.UTF8.GetBytes(isolated.ToString()));
            return Assert.Single(DrawingTestTraversal.Elements(drawing).OfType<OfficeDrawingShape>());
        }
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void StyledEmailExportsThroughEverySharedImageFormat(OfficeImageExportFormat format) {
        var email = new EmailDocument { Subject = "Typography" };
        email.Body.Html = "<p><span style='font-family:Aptos;font-size:18px;color:#336699;font-weight:700;font-style:italic;text-decoration-line:underline line-through;text-decoration-style:wavy;vertical-align:super'>Styled</span></p>";

        OfficeImageExportResult result = Assert.Single(email.ExportImages(format));

        Assert.Equal(format, result.Format);
        Assert.True(result.Bytes.Length > 32);
        if (format == OfficeImageExportFormat.Svg) {
            OfficeDrawing drawing = HtmlSvgPaintAssertions.Read(result.Bytes);
            Assert.Contains(DrawingTestTraversal.Elements(drawing).OfType<OfficeDrawingText>(), text =>
                text.Text == "Styled" && (text.Font.Style & OfficeFontStyle.Italic) != 0);
            HtmlSvgPaintAssertions.WavyStrokes(drawing, OfficeColor.FromRgb(0x33, 0x66, 0x99), 2);
        }
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
    public void EmailImageExportRejectsTooManyInlineResourcesBeforeRendering() {
        var email = new EmailDocument { Subject = "Too many inline images" };
        email.Body.Html = "<img src=\"cid:first@example\"><img src=\"cid:second@example\">";
        email.Attachments.Add(CreateInlineImage("first@example"));
        email.Attachments.Add(CreateInlineImage("second@example"));

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            email.ExportImage(
                OfficeImageExportFormat.Png,
                new EmailImageExportOptions { MaxInlineResourceCount = 1 }));

        Assert.Equal("EmailBodyProjectionOptions.MaxResourceCount", exception.LimitName);
    }

    [Fact]
    public async Task EmailImageExportRejectsAggregateInlineBytesBeforeRendering() {
        var email = new EmailDocument { Subject = "Oversized inline image set" };
        email.Body.Html = "<img src=\"cid:first@example\"><img src=\"cid:second@example\">";
        email.Attachments.Add(CreateInlineImage("first@example"));
        email.Attachments.Add(CreateInlineImage("second@example"));

        EmailLimitExceededException exception = await Assert.ThrowsAsync<EmailLimitExceededException>(() =>
            email.ExportImageAsync(
                OfficeImageExportFormat.Png,
                new EmailImageExportOptions {
                    MaxTotalInlineResourceBytes = PixelPng.LongLength * 2L - 1L
                }));

        Assert.Equal("EmailBodyProjectionOptions.MaxTotalResourceBytes", exception.LimitName);
    }

    [Fact]
    public void DisabledInlineResourcesBypassResourceInventoryLimits() {
        var email = new EmailDocument { Subject = "Body-only image" };
        email.Body.Html = "<p>No inline resources are needed.</p>";
        email.Attachments.Add(CreateInlineImage("first@example"));
        email.Attachments.Add(CreateInlineImage("second@example"));

        OfficeImageExportResult result = email.ExportImage(
            OfficeImageExportFormat.Svg,
            new EmailImageExportOptions {
                IncludeInlineResources = false,
                MaxInlineResourceCount = 1,
                MaxTotalInlineResourceBytes = 1
            });

        Assert.NotEmpty(result.Bytes);
    }

    [Fact]
    public async Task AggregateInlineReadFailureUsesTheOperationWideDiagnostic() {
        var email = new EmailDocument { Subject = "Aggregate inline limit" };
        email.Body.Html = "<img src=\"cid:logo@example\" alt=\"Logo\">";
        EmailAttachment attachment = CreateInlineImage("logo@example");
        attachment.Length = 0;
        email.Attachments.Add(attachment);

        OfficeImageExportPolicyException exception = await Assert.ThrowsAsync<OfficeImageExportPolicyException>(() =>
            email.ExportImageAsync(
                OfficeImageExportFormat.Png,
                new EmailImageExportOptions {
                    MaxTotalInlineResourceBytes = PixelPng.LongLength - 1L
                }));

        Assert.Contains(
            exception.Diagnostics,
            diagnostic => diagnostic.Code ==
                          HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded);
        Assert.DoesNotContain(
            exception.Diagnostics,
            diagnostic => diagnostic.Code ==
                          HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded);
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
    public async Task EmptyRetainedHttpsResourceDoesNotFallThroughToTheAsyncFallback() {
        var email = new EmailDocument { Subject = "Empty retained image" };
        email.Body.Html = "<img src=\"https://assets.example/empty.png\" alt=\"empty\">";
        email.Attachments.Add(new EmailAttachment {
            FileName = "empty.png",
            ContentType = "image/png",
            ContentLocation = "https://assets.example/empty.png",
            IsInline = true,
            Content = Array.Empty<byte>(),
            Length = 0
        });
        int fallbackCalls = 0;
        var options = new EmailImageExportOptions {
            RemoteResourcePolicy = EmailRemoteResourcePolicy.AllowByConsumerResolver,
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
    public async Task UnknownContentIdDoesNotFallThroughToAnExplicitCidFallback() {
        var email = new EmailDocument { Subject = "Unknown CID" };
        email.Body.Html = "<img src=\"cid:missing@example.test\" alt=\"missing\">";
        int fallbackCalls = 0;
        HtmlUrlPolicy fallbackPolicy = HtmlUrlPolicy.CreateWebResourceProfile();
        fallbackPolicy.AllowedUrlSchemes.Add("cid");
        var options = new EmailImageExportOptions {
            RemoteResourcePolicy = EmailRemoteResourcePolicy.AllowByConsumerResolver,
            ResourceUrlPolicy = fallbackPolicy,
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
            var imageOptions2 = new EmailImageExportOptions {
                    PageSize = new OfficePageSize(
                        global::OfficeIMO.OfficeLength.Inches(360D / HtmlRenderOptions.CssPixelsPerInch),
                        global::OfficeIMO.OfficeLength.Inches(260D / HtmlRenderOptions.CssPixelsPerInch))
                };
            imageOptions2.Mode = HtmlRenderMode.Paged;
            imageOptions2.Format = OfficeImageExportFormat.Png;
            imageOptions2.MaximumOutputCount = 20;
            imageOptions2.MaximumTotalRasterPixels = 20_000_000;
            imageOptions2.MaximumTotalEncodedBytes = 20_000_000;
            OfficeConversionReport saved = await email
                .SaveAsImagesAsync(folder, imageOptions2);

            Assert.True(saved.ItemsWritten > 1);
            string[] paths = Directory.GetFiles(folder, "*.png");
            Assert.Equal(saved.ItemsWritten, paths.Length);
            Assert.All(paths, path => {
                byte[] bytes = File.ReadAllBytes(path);
                Assert.NotEmpty(bytes);
                Assert.Equal("image/png", OfficeImageReader.Identify(bytes).MimeType);
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

    private static EmailAttachment CreateInlineImage(string contentId) => new EmailAttachment {
        FileName = contentId + ".png",
        ContentType = "image/png",
        ContentId = contentId,
        IsInline = true,
        Content = PixelPng,
        Length = PixelPng.Length
    };

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
