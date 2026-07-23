using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests.Emlx;

public sealed class EmlxStoreReaderTests {
    [Fact]
    public void ReadsExactMimeSegmentAndAppleMetadata() {
        byte[] message = CreateMultipartMessage();
        long flags = (1L << 0) | (1L << 6) | (1L << 25);
        string plist = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                       "<!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\" \"http://www.apple.com/DTDs/PropertyList-1.0.dtd\">\n" +
                       "<plist version=\"1.0\"><dict>" +
                       "<key>conversation-id</key><integer>123456</integer>" +
                       "<key>date-received</key><integer>1710000000</integer>" +
                       "<key>flags</key><integer>" + flags.ToString() + "</integer>" +
                       "<key>remote-id</key><string>remote-42</string>" +
                       "</dict></plist>";
        byte[] emlx = CreateEmlx(message, "\n", plist);

        EmailStoreReadResult result = Read(emlx, "42.emlx");

        Assert.Equal(EmailStoreFormat.Emlx, result.Store.Format);
        Assert.Equal("42", result.Store.DisplayName);
        EmailStoreFolder folder = Assert.Single(result.Store.Folders);
        Assert.Equal("Apple Mail", folder.Name);
        EmailDocument document = Assert.Single(folder.Items).Document;
        Assert.Equal(EmailFileFormat.Eml, document.Format);
        Assert.Equal("EMLX contract", document.Subject);
        Assert.Equal("sender@example.test", document.From?.Address);
        Assert.Equal("Plain body", document.Body.Text?.Trim());
        EmailAttachment attachment = Assert.Single(document.Attachments);
        Assert.Equal("payload.bin", attachment.FileName);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, attachment.Content);
        Assert.True(document.MessageMetadata.IsRead);
        Assert.True(document.MessageMetadata.IsDraft);
        Assert.Equal(123456L, document.Properties["Emlx:Metadata:conversation-id"]);
        Assert.Equal("remote-42", document.Properties["Emlx:Metadata:remote-id"]);
        Assert.Equal(DateTimeOffset.FromUnixTimeSeconds(1710000000), document.ReceivedDate);
        Assert.Equal(message.LongLength, document.Properties["Emlx:DeclaredMessageBytes"]);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void AcceptsCrLfPrefixAndRestoresStreamPosition() {
        byte[] message = Encoding.ASCII.GetBytes("Subject: CRLF prefix\r\nFrom: a@example.test\r\n\r\nBody\r\n");
        byte[] emlx = CreateEmlx(message, "\r\n", null);
        using (var stream = new MemoryStream(emlx, writable: false)) {
            stream.Position = 4;

            EmailStoreReadResult result = new EmailStoreReader().Read(stream, "message.emlx");

            Assert.Equal(4, stream.Position);
            Assert.Equal("CRLF prefix", Assert.Single(Assert.Single(result.Store.Folders).Items).Document.Subject);
        }
    }

    [Fact]
    public void HonorsAttachmentRetentionPolicyThroughSharedMimeReader() {
        byte[] emlx = CreateEmlx(CreateMultipartMessage(), "\n", null);
        var options = new EmailStoreReaderOptions(retainAttachmentContent: false);

        EmailStoreReadResult result = Read(emlx, "message.emlx", options);
        EmailAttachment attachment = Assert.Single(Assert.Single(Assert.Single(result.Store.Folders).Items)
            .Document.Attachments);

        Assert.Equal(4, attachment.Length);
        Assert.Null(attachment.Content);
    }

    [Fact]
    public void EnforcesAttachmentCountAfterMimeProjection() {
        byte[] emlx = CreateEmlx(CreateMultipartMessageWithTwoAttachments(), "\n", null);
        var options = new EmailStoreReaderOptions(maxAttachmentsPerItem: 1);

        EmailStoreLimitExceededException exception = Assert.Throws<EmailStoreLimitExceededException>(
            () => Read(emlx, "message.emlx", options));

        Assert.Equal(nameof(EmailStoreReaderOptions.MaxAttachmentsPerItem), exception.LimitName);
    }

    [Fact]
    public void ReportsInvalidMetadataWithoutDiscardingTheMessage() {
        byte[] message = Encoding.ASCII.GetBytes("Subject: Keep me\r\n\r\nBody\r\n");
        byte[] emlx = CreateEmlx(message, "\n", "<plist><dict><key>broken</key></dict>");

        EmailStoreReadResult result = Read(emlx, "message.emlx");

        Assert.Equal("Keep me", Assert.Single(Assert.Single(result.Store.Folders).Items).Document.Subject);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_EMLX_METADATA_INVALID" &&
            diagnostic.Severity == EmailStoreDiagnosticSeverity.Warning);
    }

    [Fact]
    public void KeepsTheMessageWhenMetadataUsesBinaryPlistEncoding() {
        byte[] message = Encoding.ASCII.GetBytes("Subject: Binary metadata\r\n\r\nBody\r\n");
        byte[] emlx = CreateEmlx(message, "\n", "bplist00opaque");

        EmailStoreReadResult result = Read(emlx, "binary.emlx");

        Assert.Equal("Binary metadata", Assert.Single(Assert.Single(result.Store.Folders).Items).Document.Subject);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_EMLX_METADATA_UNSUPPORTED");
    }

    [Fact]
    public void IdentifiesPartialMessagesWithoutInventingSiblingAttachments() {
        byte[] message = Encoding.ASCII.GetBytes("Subject: Headers only\r\n\r\n");
        byte[] emlx = CreateEmlx(message, "\n", null);

        EmailStoreReadResult result = Read(emlx, "314.partial.emlx");
        EmailStoreItem item = Assert.Single(Assert.Single(result.Store.Folders).Items);
        EmailDocument document = item.Document;

        Assert.Equal(true, document.Properties["Emlx:IsPartial"]);
        Assert.Null(item.ContentAvailability.IsHeaderOnly);
        Assert.True(item.ContentAvailability.IsPotentiallyPartial);
        Assert.True(item.ContentAvailability.IndeterminateParts.HasFlag(
            EmailStoreItemReadParts.Bodies));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_EMLX_PARTIAL_MESSAGE");
    }

    [Fact]
    public void RejectsInvalidOrTruncatedLengthPrefixes() {
        byte[][] invalid = {
            Encoding.ASCII.GetBytes("abc\nSubject: no\r\n\r\n"),
            Encoding.ASCII.GetBytes("-1\nSubject: no\r\n\r\n"),
            Encoding.ASCII.GetBytes("999\nSubject: short\r\n\r\n"),
            Encoding.ASCII.GetBytes("12345678901234567890123456789012345678901234567890123456789012345")
        };

        foreach (byte[] data in invalid) {
            Assert.Throws<InvalidDataException>(() => Read(data, "invalid.emlx"));
        }
    }

    [Fact]
    public void EnforcesMessageMetadataAndPropertyLimits() {
        byte[] message = Encoding.ASCII.GetBytes("Subject: limits\r\n\r\nBody\r\n");
        string plist = "<plist><dict><key>one</key><integer>1</integer><key>two</key><integer>2</integer></dict></plist>";
        byte[] emlx = CreateEmlx(message, "\n", plist);

        var messageOptions = new EmailStoreReaderOptions(maxMessageBytes: message.Length - 1L);
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxMessageBytes),
            Assert.Throws<EmailStoreLimitExceededException>(() => Read(emlx, "limits.emlx", messageOptions)).LimitName);

        var metadataOptions = new EmailStoreReaderOptions(maxDecodedPropertyBytesPerItem: 16);
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
            Assert.Throws<EmailStoreLimitExceededException>(() => Read(emlx, "limits.emlx", metadataOptions)).LimitName);

        var propertyOptions = new EmailStoreReaderOptions(maxPropertiesPerItem: 1);
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxPropertiesPerItem),
            Assert.Throws<EmailStoreLimitExceededException>(() => Read(emlx, "limits.emlx", propertyOptions)).LimitName);
    }

    [Fact]
    public void DetectsEmlxByItsExtensionWithoutConsumingTheStream() {
        byte[] message = Encoding.ASCII.GetBytes("Subject: detection\r\n\r\n");
        byte[] emlx = CreateEmlx(message, "\n", null);
        using (var stream = new MemoryStream(emlx, writable: false)) {
            stream.Position = 3;

            EmailStoreFormat format = EmailStoreReader.DetectFormat(stream, "7.partial.emlx");

            Assert.Equal(EmailStoreFormat.Emlx, format);
            Assert.Equal(3, stream.Position);
        }
    }

    private static EmailStoreReadResult Read(byte[] data, string sourceName,
        EmailStoreReaderOptions? options = null) {
        using (var stream = new MemoryStream(data, writable: false)) {
            return new EmailStoreReader(options).Read(stream, sourceName);
        }
    }

    private static byte[] CreateEmlx(byte[] message, string prefixNewline, string? plist) {
        byte[] prefix = Encoding.ASCII.GetBytes(message.Length.ToString() + prefixNewline);
        byte[] metadata = plist == null ? Array.Empty<byte>() : Encoding.UTF8.GetBytes(plist);
        var result = new byte[prefix.Length + message.Length + metadata.Length];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        Buffer.BlockCopy(message, 0, result, prefix.Length, message.Length);
        Buffer.BlockCopy(metadata, 0, result, prefix.Length + message.Length, metadata.Length);
        return result;
    }

    private static byte[] CreateMultipartMessage() {
        const string message =
            "From: Sender <sender@example.test>\r\n" +
            "To: Receiver <receiver@example.test>\r\n" +
            "Subject: EMLX contract\r\n" +
            "Message-ID: <emlx@example.test>\r\n" +
            "Date: Tue, 14 Jul 2026 08:30:00 +0000\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=\"emlx-boundary\"\r\n" +
            "\r\n" +
            "--emlx-boundary\r\n" +
            "Content-Type: text/plain; charset=utf-8\r\n" +
            "\r\n" +
            "Plain body\r\n" +
            "--emlx-boundary\r\n" +
            "Content-Type: application/octet-stream; name=\"payload.bin\"\r\n" +
            "Content-Disposition: attachment; filename=\"payload.bin\"\r\n" +
            "Content-Transfer-Encoding: base64\r\n" +
            "\r\n" +
            "AQIDBA==\r\n" +
            "--emlx-boundary--\r\n";
        return Encoding.ASCII.GetBytes(message);
    }

    private static byte[] CreateMultipartMessageWithTwoAttachments() {
        const string secondAttachment =
            "--emlx-boundary\r\n" +
            "Content-Type: application/octet-stream; name=\"second.bin\"\r\n" +
            "Content-Disposition: attachment; filename=\"second.bin\"\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\n" +
            "BQYHCA==\r\n";
        string message = Encoding.ASCII.GetString(CreateMultipartMessage());
        return Encoding.ASCII.GetBytes(message.Replace(
            "--emlx-boundary--\r\n",
            secondAttachment + "--emlx-boundary--\r\n"));
    }
}
