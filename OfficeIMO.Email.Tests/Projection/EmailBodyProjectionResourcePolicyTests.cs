using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailBodyProjectionResourcePolicyTests {
    [Fact]
    public void Offline_projection_uses_the_document_base_for_relative_content_locations() {
        var document = new EmailDocument();
        document.Body.Html = "<base href='https://mail.example/message/'><img src='images/logo.png'>";
        document.Attachments.Add(new EmailAttachment {
            FileName = "logo.png",
            ContentType = "image/png",
            ContentLocation = "images/logo.png",
            IsInline = true,
            Content = new byte[] { 1 },
            Length = 1
        });

        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document);

        Assert.Contains("src=\"cid:officeimo-resource-", projection.Html,
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal(new Uri("https://mail.example/message/"), projection.Document.BaseUri);
        Assert.Same(projection.Resources[0], projection.ResolveResource(
            "https://mail.example/message/images/logo.png",
            new Uri("https://mail.example/message/images/logo.png")));
    }

    [Fact]
    public void Declared_aggregate_overflow_is_reported_as_a_limit_violation() {
        var document = new EmailDocument { Body = { Html = "<p>resources</p>" } };
        document.Attachments.Add(CreateInlineAttachment("first", long.MaxValue));
        document.Attachments.Add(CreateInlineAttachment("second", 1));

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            EmailBodyProjection.Create(document, new EmailBodyProjectionOptions {
                MaxResourceBytes = long.MaxValue,
                MaxTotalResourceBytes = long.MaxValue
            }));

        Assert.Equal("EmailBodyProjectionOptions.MaxTotalResourceBytes", exception.LimitName);
        Assert.Equal(long.MaxValue, exception.ActualValue);
        Assert.Equal(long.MaxValue, exception.MaximumValue);
    }

    [Fact]
    public void Duplicate_content_ids_are_rewritten_to_unique_aliases_and_remain_ambiguous_by_original_id() {
        var document = new EmailDocument();
        document.Body.HtmlContentLocation = "https://mail.example/message/";
        document.Body.Html = "<img src='images/second.png'>";
        document.Attachments.Add(CreateInlineAttachment(
            "duplicate@example.test", 1, "first.png", "images/first.png"));
        document.Attachments.Add(CreateInlineAttachment(
            "duplicate@example.test", 1, "second.png", "images/second.png"));

        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document);
        string alias = projection.Document.CreateDocumentForConversion()
            .QuerySelector("img")!.GetAttribute("src")!;

        Assert.StartsWith("cid:officeimo-resource-", alias, StringComparison.OrdinalIgnoreCase);
        Assert.Same(projection.Resources[1], projection.ResolveResource(alias));
        Assert.Null(projection.ResolveResource("cid:duplicate@example.test"));
    }

    [Fact]
    public void Content_location_matching_precedes_filename_fallback() {
        var document = new EmailDocument { Body = { Html = "<img src='logo.png'>" } };
        document.Attachments.Add(CreateInlineAttachment(
            "filename@example.test", 1, "logo.png", null));
        document.Attachments.Add(CreateInlineAttachment(
            "location@example.test", 1, "other.png", "logo.png"));

        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document);

        Assert.Contains("src=\"cid:location@example.test\"", projection.Html,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Filename_fallback_preserves_case_insensitive_compatibility() {
        var document = new EmailDocument { Body = { Html = "<img src='LOGO.PNG'>" } };
        document.Attachments.Add(CreateInlineAttachment(
            "filename@example.test", 1, "logo.png", null));

        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document);

        Assert.Contains("src=\"cid:filename@example.test\"", projection.Html,
            StringComparison.OrdinalIgnoreCase);
        Assert.Same(projection.Resources[0], projection.ResolveResource("LOGO.PNG"));
    }

    [Fact]
    public async Task Projection_snapshots_resource_identity_metadata_and_content_selection() {
        var attachment = CreateInlineAttachment(
            "old@example.test", 1, "old.png", "images/old.png");
        attachment.ContentType = "image/png";
        attachment.Content = new byte[] { 1 };
        var document = new EmailDocument { Body = { Html = "<img src='cid:old@example.test'>" } };
        document.Attachments.Add(attachment);

        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document);
        EmailBodyResource resource = Assert.Single(projection.Resources);

        attachment.ContentId = "new@example.test";
        attachment.ContentLocation = "images/new.png";
        attachment.FileName = "new.png";
        attachment.ContentType = "image/jpeg";
        attachment.Length = 2;
        attachment.Content = new byte[] { 2, 3 };

        Assert.Equal("old@example.test", resource.ContentId);
        Assert.Equal("images/old.png", resource.ContentLocation);
        Assert.Equal("old.png", resource.FileName);
        Assert.Equal("image/png", resource.ContentType);
        Assert.Equal(1, resource.Length);
        Assert.Equal(new byte[] { 1 }, resource.ReadAllBytes());
        Assert.Equal(new byte[] { 1 }, await resource.ReadAllBytesAsync());
        Assert.Same(resource, projection.ResolveResource("cid:old@example.test"));
        Assert.Null(projection.ResolveResource("cid:new@example.test"));
    }

    [Fact]
    public void Absolute_resource_matching_preserves_path_case() {
        var document = new EmailDocument();
        document.Body.HtmlContentLocation = "https://assets.example/";
        document.Body.Html = "<img src='images/logo.png'>";
        document.Attachments.Add(CreateInlineAttachment(
            "upper@example.test", 1, "upper.png", "Images/logo.png"));
        document.Attachments.Add(CreateInlineAttachment(
            "lower@example.test", 1, "lower.png", "images/logo.png"));

        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document);

        Assert.Contains("src=\"cid:lower@example.test\"", projection.Html,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Ambiguous_content_locations_are_not_rewritten_or_resolved() {
        var document = new EmailDocument { Body = { Html = "<img src='shared.png'>" } };
        document.Attachments.Add(CreateInlineAttachment(
            "first@example.test", 1, "first.png", "shared.png"));
        document.Attachments.Add(CreateInlineAttachment(
            "second@example.test", 1, "second.png", "shared.png"));

        EmailBodyProjectionResult projection = EmailBodyProjection.Create(document);
        AngleSharp.Dom.IElement image = projection.Document.CreateDocumentForConversion()
            .QuerySelector("img")!;

        Assert.False(image.HasAttribute("src"));
        Assert.Null(projection.ResolveResource("shared.png"));
    }

    private static EmailAttachment CreateInlineAttachment(
        string contentId,
        long length,
        string? fileName = null,
        string? contentLocation = null) => new EmailAttachment {
        FileName = fileName,
        ContentLocation = contentLocation,
        ContentId = contentId,
        ContentType = "application/octet-stream",
        IsInline = true,
        Content = new byte[] { 1 },
        Length = length
    };
}
