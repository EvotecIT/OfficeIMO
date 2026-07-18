using System;
using System.IO;
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
}
