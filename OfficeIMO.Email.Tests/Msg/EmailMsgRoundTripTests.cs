using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMsgRoundTripTests {
    [Fact]
    public void RoundTripsMessageRecipientsAttachmentsEmbeddedMessagesAndUnknownProperties() {
        EmailDocument child = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "Embedded",
            MessageClass = "IPM.Note"
        };
        child.Body.Text = "inside";
        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "MSG subject",
            MessageClass = "IPM.Note",
            From = new EmailAddress("from@example.com", "From Person"),
            Sender = new EmailAddress("sender@example.com", "Sender Person"),
            MessageId = "id@example.com",
            Date = new DateTimeOffset(2026, 7, 10, 12, 30, 0, TimeSpan.Zero),
            ReceivedDate = new DateTimeOffset(2026, 7, 10, 12, 31, 0, TimeSpan.Zero)
        };
        source.Body.Text = "plain body";
        source.Body.Html = "<p>html body</p>";
        source.Headers.Add(new EmailHeader("X-Test", "value"));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("to@example.com", "To")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Cc, new EmailAddress("cc@example.com", "Cc")));
        source.MapiProperties.Add(new MapiProperty(0x66AA, MapiPropertyType.MultipleUnicode,
            new object[] { "one", "dwa" }));
        source.MapiProperties.Add(new MapiProperty(0x8000, MapiPropertyType.Unicode, "named value",
            name: new MapiNamedProperty(MsgProjection.PsetidCommon, 0x85FF)));
        source.Attachments.Add(new EmailAttachment {
            FileName = "data.bin",
            ContentType = "application/octet-stream",
            Content = new byte[] { 1, 2, 3, 4 },
            Length = 4
        });
        source.Attachments.Add(new EmailAttachment {
            FileName = "child.msg",
            ContentType = "application/vnd.ms-outlook",
            EmbeddedDocument = child
        });
        var structured = new EmailAttachment { FileName = "object.ole", MapiAttachMethod = 6 };
        structured.StructuredStorageStreams["Contents"] = new byte[] { 9, 8, 7 };
        structured.StructuredStorageStreams["Nested/Metadata"] = Encoding.UTF8.GetBytes("meta");
        source.Attachments.Add(structured);

        byte[] first = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);
        byte[] second = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);
        EmailReadResult result = new EmailDocumentReader().Read(first);

        Assert.Equal(first, second);
        Assert.Equal(EmailFileFormat.OutlookMsg, result.Document.Format);
        Assert.Equal(source.Subject, result.Document.Subject);
        Assert.Equal(source.Body.Text, result.Document.Body.Text);
        Assert.Equal(source.Body.Html, result.Document.Body.Html);
        Assert.Equal("from@example.com", result.Document.From!.Address);
        Assert.Equal("sender@example.com", result.Document.Sender!.Address);
        Assert.Equal(2, result.Document.Recipients.Count);
        Assert.Equal(3, result.Document.Attachments.Count);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, result.Document.Attachments[0].Content);
        Assert.Equal("Embedded", result.Document.Attachments[1].EmbeddedDocument!.Subject);
        Assert.Equal("meta", Encoding.UTF8.GetString(result.Document.Attachments[2].StructuredStorageStreams["Nested/Metadata"]));
        Assert.Equal(new object[] { "one", "dwa" }, result.Document.MapiProperties.Single(property => property.PropertyId == 0x66AA).Value);
        Assert.Equal("named value", result.Document.MapiProperties.Single(property => property.Name?.LocalId == 0x85FF).Value);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void OutputIsReadableByMsgReaderOracle() {
        EmailDocument source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "Oracle subject",
            From = new EmailAddress("sender@example.com", "Sender")
        };
        source.Body.Text = "Oracle body";
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("receiver@example.com", "Receiver")));
        source.Attachments.Add(new EmailAttachment {
            FileName = "a.txt",
            ContentType = "text/plain",
            Content = Encoding.UTF8.GetBytes("attachment"),
            Length = 10
        });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        using MemoryStream stream = new MemoryStream(bytes);
        using var oracle = new global::MsgReader.Outlook.Storage.Message(stream, FileAccess.Read, true);

        Assert.Equal("Oracle subject", oracle.Subject);
        Assert.Equal("Oracle body", oracle.BodyText!.TrimEnd());
        Assert.Single(oracle.Recipients!);
        Assert.Single(oracle.Attachments!);
    }

    [Fact]
    public void ReaderCanSkipMsgAttachmentBytes() {
        EmailDocument source = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "skip" };
        source.Attachments.Add(new EmailAttachment { FileName = "a.bin", Content = new byte[] { 1, 2, 3 }, Length = 3 });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        EmailDocument parsed = new EmailDocumentReader(new EmailReaderOptions(includeAttachmentContent: false)).Read(bytes).Document;

        EmailAttachment attachment = Assert.Single(parsed.Attachments);
        Assert.Equal(3, attachment.Length);
        Assert.Null(attachment.Content);
        Assert.Null(attachment.MapiProperties.Single(property => property.PropertyId == 0x3701).RawData);
    }

    [Fact]
    public void ReadsMsgKitGeneratedMessageAsCompatibilityOracle() {
        string directory = Path.Combine(Path.GetTempPath(), string.Concat("OfficeIMO-Email-", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(directory);
        string emlPath = Path.Combine(directory, "source.eml");
        string msgPath = Path.Combine(directory, "source.msg");
        const string eml = "From: Sender <sender@example.com>\r\nTo: Receiver <receiver@example.com>\r\n" +
            "Subject: MsgKit source\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nbody from MsgKit\r\n" +
            "--x\r\nContent-Type: application/octet-stream; name=a.bin\r\n" +
            "Content-Disposition: attachment; filename=a.bin\r\nContent-Transfer-Encoding: base64\r\n\r\nAQID\r\n--x--\r\n";
        try {
            File.WriteAllText(emlPath, eml, new UTF8Encoding(false));
            MsgKit.Converter.ConvertEmlToMsg(emlPath, msgPath);

            EmailReadResult result = new EmailDocumentReader().Read(msgPath);

            Assert.Equal("MsgKit source", result.Document.Subject);
            Assert.Contains("body from MsgKit", result.Document.Body.Text);
            Assert.Equal("receiver@example.com", Assert.Single(result.Document.Recipients).Address.Address);
            Assert.Equal(new byte[] { 1, 2, 3 }, Assert.Single(result.Document.Attachments).Content);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, true);
        }
    }
}
