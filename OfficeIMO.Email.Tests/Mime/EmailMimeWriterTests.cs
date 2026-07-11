using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMimeWriterTests {
    [Fact]
    public void DeterministicWriterRoundTripsEnvelopeBodiesAndAttachments() {
        EmailDocument document = new EmailDocument {
            Format = EmailFileFormat.Eml,
            Subject = "Zażółć",
            From = new EmailAddress("sender@example.com", "José"),
            MessageId = "stable@example.com",
            Date = new DateTimeOffset(2026, 7, 10, 15, 0, 0, TimeSpan.FromHours(2))
        };
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("to@example.com", "Receiver")));
        document.Body.Text = "plain";
        document.Body.Html = "<strong>html</strong>";
        document.Headers.Add(new EmailHeader("X-Correlation", "abc"));
        document.Attachments.Add(new EmailAttachment {
            FileName = "dane-ą.bin",
            ContentType = "application/octet-stream",
            Content = new byte[] { 9, 8, 7 },
            Length = 3
        });

        EmailDocumentWriter writer = new EmailDocumentWriter();
        byte[] first = writer.WriteToBytes(document);
        byte[] second = writer.WriteToBytes(document);
        EmailReadResult parsed = new EmailDocumentReader().Read(first);

        Assert.Equal(first, second);
        Assert.Equal(document.Subject, parsed.Document.Subject);
        Assert.Equal(document.Body.Text, parsed.Document.Body.Text);
        Assert.Equal(document.Body.Html, parsed.Document.Body.Html);
        Assert.Equal(document.Attachments[0].FileName, parsed.Document.Attachments[0].FileName);
        Assert.Equal(document.Attachments[0].Content, parsed.Document.Attachments[0].Content);
        Assert.Equal("abc", parsed.Document.Headers.Single(header => header.Name == "X-Correlation").Value);
    }

    [Fact]
    public void PreservedSourceWritingIsExplicitAndVerbatim() {
        byte[] source = Encoding.ASCII.GetBytes("Subject: raw\n\nbody\n");
        EmailReaderOptions readerOptions = new EmailReaderOptions(preserveRawSource: true);
        EmailDocument document = new EmailDocumentReader(readerOptions).Read(source).Document;
        EmailDocumentWriter writer = new EmailDocumentWriter(new EmailWriterOptions(usePreservedRawSource: true));

        byte[] result = writer.WriteToBytes(document);

        Assert.Equal(source, result);
    }

    [Fact]
    public void WritesAndReadsEmbeddedMessages() {
        EmailDocument child = new EmailDocument { Format = EmailFileFormat.Eml, Subject = "Child" };
        child.Body.Text = "inside";
        EmailDocument parent = new EmailDocument { Format = EmailFileFormat.Eml, Subject = "Parent" };
        parent.Body.Text = "outside";
        parent.Attachments.Add(new EmailAttachment {
            FileName = "child.eml",
            ContentType = "message/rfc822",
            EmbeddedDocument = child
        });

        byte[] bytes = new EmailDocumentWriter().WriteToBytes(parent);
        EmailDocument parsed = new EmailDocumentReader().Read(bytes).Document;

        Assert.Equal("Child", Assert.Single(parsed.Attachments).EmbeddedDocument!.Subject);
    }

    [Fact]
    public void AppliesConfiguredBase64LineLengthToTextBodies() {
        var document = new EmailDocument { Subject = "line length" };
        document.Body.Text = new string('x', 120);

        byte[] bytes = new EmailDocumentWriter(new EmailWriterOptions(base64LineLength: 20))
            .WriteToBytes(document);
        string[] lines = Encoding.ASCII.GetString(bytes).Split(new[] { "\r\n" }, StringSplitOptions.None);
        int bodyStart = Array.IndexOf(lines, string.Empty) + 1;
        string[] payloadLines = lines.Skip(bodyStart).Where(line => line.Length > 0).ToArray();

        Assert.True(payloadLines.Length > 1);
        Assert.All(payloadLines, line => Assert.InRange(line.Length, 1, 20));
    }
}
