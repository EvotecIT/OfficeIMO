using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailBodyProjectionTests {
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
}
