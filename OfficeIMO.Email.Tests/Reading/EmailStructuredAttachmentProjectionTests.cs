using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Rtf;

namespace OfficeIMO.Email.Tests;

public sealed class EmailStructuredAttachmentProjectionTests {
    [Fact]
    public void MimeCharsetRemainsAuthoritativeWhenHtmlAttachmentDeclaresLegacyMetaCharset() {
        byte[] content = Encoding.GetEncoding(28591).GetBytes("<meta charset='iso-8859-1'><p>café</p>");
        var attachment = new EmailAttachment { FileName = "body.html", ContentType = "text/html", Content = content };
        attachment.ContentTypeParameters["charset"] = "iso-8859-1";
        var document = new EmailDocument();
        document.Attachments.Add(attachment);
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().AddHtmlHandler().Build();
        using var stream = new MemoryStream(document.ToBytes());
        var result = reader.ReadDocument(stream, "message.eml");
        string projected = string.Concat(result.Chunks.Where(chunk => chunk.Id.StartsWith("email:attachment-content:")).Select(chunk => chunk.Text));
        Assert.Contains("café", projected);
        Assert.DoesNotContain("cafÃ©", projected);
    }

    [Fact]
    public void RtfAttachmentRetainsItsByteEncodingForTheRtfOwner() {
        byte[] content = Encoding.ASCII.GetBytes("{\\rtf1\\ansi\\ansicpg1252 caf")
            .Concat(new byte[] { 0xe9 }).Concat(Encoding.ASCII.GetBytes("}" )).ToArray();
        var document = new EmailDocument();
        document.Attachments.Add(new EmailAttachment { FileName = "body.rtf", ContentType = "text/rtf", Content = content });
        var reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().AddRtfHandler().Build();
        using var stream = new MemoryStream(document.ToBytes());
        var result = reader.ReadDocument(stream, "message.eml");
        string projected = string.Concat(result.Chunks.Where(chunk => chunk.Id.StartsWith("email:attachment-content:")).Select(chunk => chunk.Text));
        Assert.Contains("café", projected);
        Assert.DoesNotContain("cafÃ©", projected);
    }
}
