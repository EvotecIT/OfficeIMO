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
}
