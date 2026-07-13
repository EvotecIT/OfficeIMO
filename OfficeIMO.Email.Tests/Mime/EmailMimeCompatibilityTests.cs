using MimeKit;
using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMimeCompatibilityTests {
    [Fact]
    public void ReadsMimeKitMultipartMessageAsMailozaurrCompatibilityOracle() {
        var source = new MimeMessage();
        source.From.Add(new MailboxAddress("José Sender", "sender@example.com"));
        source.To.Add(new MailboxAddress("Receiver", "receiver@example.com"));
        source.Cc.Add(new MailboxAddress("Team", "team@example.com"));
        source.Subject = "MimeKit żółć";
        source.MessageId = "mimekit@example.com";
        source.Date = new DateTimeOffset(2026, 7, 10, 12, 30, 0, TimeSpan.FromHours(2));

        var body = new BodyBuilder {
            TextBody = "plain body",
            HtmlBody = "<p>html body</p>"
        };
        body.Attachments.Add("dane-ą.bin", new byte[] { 1, 2, 3, 4 },
            ContentType.Parse("application/octet-stream"));
        source.Body = body.ToMessageBody();

        using var stream = new MemoryStream();
        source.WriteTo(stream);

        EmailReadResult result = new EmailDocumentReader().Read(stream.ToArray());

        Assert.Equal("MimeKit żółć", result.Document.Subject);
        Assert.Equal("José Sender", result.Document.From!.DisplayName);
        Assert.Equal("sender@example.com", result.Document.From.Address);
        Assert.Equal(2, result.Document.Recipients.Count);
        Assert.Equal("plain body", result.Document.Body.Text!.Trim());
        Assert.Equal("<p>html body</p>", result.Document.Body.Html!.Trim());
        EmailAttachment attachment = Assert.Single(result.Document.Attachments);
        Assert.Equal("dane-ą.bin", attachment.FileName);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, attachment.Content);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void MimeKitReadsOfficeImoDeterministicOutput() {
        var source = new EmailDocument {
            Format = EmailFileFormat.Eml,
            From = new EmailAddress("sender@example.com", "José Sender"),
            Subject = "OfficeIMO żółć",
            MessageId = "officeimo@example.com"
        };
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("receiver@example.com", "Receiver")));
        source.Body.Text = "plain body";
        source.Body.Html = "<p>html body</p>";
        source.Attachments.Add(new EmailAttachment {
            FileName = "dane-ą.bin",
            ContentType = "application/octet-stream",
            Content = new byte[] { 4, 3, 2, 1 },
            Length = 4
        });

        byte[] bytes = new EmailDocumentWriter().ToBytes(source);
        using var stream = new MemoryStream(bytes, writable: false);
        MimeMessage parsed = MimeMessage.Load(stream);

        Assert.Equal("OfficeIMO żółć", parsed.Subject);
        Assert.Equal("José Sender", Assert.IsType<MailboxAddress>(Assert.Single(parsed.From)).Name);
        Assert.Equal("plain body", parsed.TextBody);
        Assert.Equal("<p>html body</p>", parsed.HtmlBody);
        MimePart attachment = Assert.IsType<MimePart>(Assert.Single(parsed.Attachments));
        Assert.Equal("dane-ą.bin", attachment.FileName);
        using var content = new MemoryStream();
        attachment.Content!.DecodeTo(content);
        Assert.Equal(new byte[] { 4, 3, 2, 1 }, content.ToArray());
    }
}
