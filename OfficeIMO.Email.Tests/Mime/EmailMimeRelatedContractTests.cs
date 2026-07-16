using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMimeRelatedContractTests {
    [Fact]
    public void ReaderRetainsRelatedMembershipAndHtmlRootIdentity() {
        const string eml = "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/related; boundary=archive; type=\"text/html\"; start=\"<root>\"\r\n\r\n" +
            "--archive\r\nContent-Type: image/png\r\nContent-ID: <logo>\r\n\r\npng\r\n" +
            "--archive\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root>\r\n" +
            "Content-Location: https://example.test/page.html\r\n\r\n<img src=\"cid:logo\">\r\n" +
            "--archive--\r\n";

        EmailDocument document = new EmailDocumentReader().Read(Encoding.UTF8.GetBytes(eml)).Document;

        Assert.Equal("<img src=\"cid:logo\">", document.Body.Html);
        Assert.Equal("root", document.Body.HtmlContentId);
        Assert.Equal("https://example.test/page.html", document.Body.HtmlContentLocation);
        Assert.True(Assert.Single(document.Attachments).IsMimeRelated);
    }

    [Fact]
    public void WriterPreservesExplicitRelatedPartsAndHtmlRootIdentity() {
        var document = new EmailDocument();
        document.Body.Html = "<html><body>archive</body></html>";
        document.Body.HtmlContentId = "root";
        document.Body.HtmlContentLocation = "https://example.test/page.html";
        byte[] css = Encoding.UTF8.GetBytes("body { color: black; }");
        document.Attachments.Add(new EmailAttachment {
            ContentType = "text/css",
            ContentLocation = "styles/site.css",
            IsInline = true,
            IsMimeRelated = true,
            Content = css,
            Length = css.LongLength
        });

        byte[] bytes = new EmailDocumentWriter().ToBytes(document);
        string serialized = Encoding.ASCII.GetString(bytes);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes).Document;

        Assert.Contains("multipart/related", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("type=\"text/html\"", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("start=\"<root>\"", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Content-ID: <root>", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Content-Location: https://example.test/page.html", serialized,
            StringComparison.OrdinalIgnoreCase);
        Assert.True(Assert.Single(roundTrip.Attachments).IsMimeRelated);
        Assert.Equal("root", roundTrip.Body.HtmlContentId);
        Assert.Equal("https://example.test/page.html", roundTrip.Body.HtmlContentLocation);
    }
}
