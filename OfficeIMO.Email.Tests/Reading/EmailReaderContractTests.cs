using OfficeIMO.Email;
using System.Threading;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailReaderContractTests {
    [Fact]
    public void EmailDocumentToStreamMatchesToBytesAndStartsAtBeginning() {
        var document = new EmailDocument { Subject = "Lifecycle" };

        using MemoryStream stream = document.ToStream();

        Assert.Equal(0, stream.Position);
        Assert.Equal(document.ToBytes(), stream.ToArray());
        Assert.True(stream.CanWrite);
    }

    [Fact]
    public void PublicEnumValuesRemainStable() {
        Assert.Equal(1, (int)EmailFileFormat.Eml);
        Assert.Equal(2, (int)EmailFileFormat.OutlookMsg);
        Assert.Equal(3, (int)EmailFileFormat.Tnef);
        Assert.Equal(4, (int)EmailFileFormat.Mbox);
        Assert.Equal(6, (int)OutlookItemKind.Note);
        Assert.Equal(6, (int)EmailRecipientKind.Room);
        Assert.Equal(2, (int)EmailDiagnosticSeverity.Error);
        Assert.Equal(2, (int)MboxVariant.Mboxrd);
    }

    [Fact]
    public void PublicMapiLookupSupportsStandardNumericAndStringNamedProperties() {
        Guid set = Guid.NewGuid();
        var properties = new[] {
            new MapiProperty(0x3001, MapiPropertyType.Unicode, "display"),
            new MapiProperty(0x8000, MapiPropertyType.Integer32, 42,
                name: new MapiNamedProperty(set, 0x1234)),
            new MapiProperty(0x8001, MapiPropertyType.Binary, new byte[] { 1, 2 },
                name: new MapiNamedProperty(set, "Custom"))
        };

        Assert.Equal("display", properties.GetMapiValue<string>(0x3001));
        Assert.Equal(42, properties.GetMapiValue<int>(set, 0x1234));
        Assert.Equal(new byte[] { 1, 2 }, properties.GetMapiValue<byte[]>(set, "custom"));
        Assert.Null(properties.GetMapiValue<string>(0x9999));
    }

    [Fact]
    public void PublicMapiProjectionProvidesOneSemanticOwnerForContainerReaders() {
        var document = new EmailDocument { Format = EmailFileFormat.Unknown };
        document.MapiProperties.Add(new MapiProperty(0x001A, MapiPropertyType.Unicode, "IPM.Task"));
        document.MapiProperties.Add(new MapiProperty(0x0037, MapiPropertyType.Unicode, "Projected task"));
        document.MapiProperties.Add(new MapiProperty(0x1000, MapiPropertyType.Unicode, "Task body"));
        document.MapiProperties.Add(new MapiProperty(0x8104, MapiPropertyType.Floating64, 0.5d,
            name: new MapiNamedProperty(new Guid("00062003-0000-0000-C000-000000000046"), 0x8102)));

        EmailReadResult result = EmailMapiProjection.Project(document, 1252, location: "store/message-1");

        Assert.Same(document, result.Document);
        Assert.False(result.HasErrors);
        Assert.Equal("Projected task", document.Subject);
        Assert.Equal("Task body", document.Body.Text);
        Assert.Equal(OutlookItemKind.Task, document.OutlookItemKind);
        Assert.NotNull(document.Task);
    }

    [Fact]
    public void PublicMapiProjectionUsesTheLastCanonicalPropertyValue() {
        var document = new EmailDocument();
        document.MapiProperties.Add(new MapiProperty(0x001A, MapiPropertyType.Unicode, "IPM.Appointment"));
        document.MapiProperties.Add(new MapiProperty(0x0037, MapiPropertyType.Unicode, "old subject"));
        document.MapiProperties.Add(new MapiProperty(0x0037, MapiPropertyType.String8, "current subject"));
        document.MapiProperties.Add(new MapiProperty(0x8200, MapiPropertyType.Unicode, "old room",
            name: new MapiNamedProperty(MapiPropertySets.Appointment, 0x8208)));
        document.MapiProperties.Add(new MapiProperty(0x8201, MapiPropertyType.Unicode, "current room",
            name: new MapiNamedProperty(MapiPropertySets.Appointment, 0x8208)));

        EmailMapiProjection.Project(document);

        Assert.Equal("current subject", document.Subject);
        Assert.Equal("current room", document.Appointment?.Location);
    }

    [Fact]
    public void DetectsTnefAndRequiresMsgDirectoryContractForCompoundFiles() {
        Assert.Equal(EmailFileFormat.Tnef, EmailDocumentReader.DetectFormat(new byte[] { 0x78, 0x9F, 0x3E, 0x22 }));
        Assert.Equal(EmailFileFormat.Unknown, EmailDocumentReader.DetectFormat(
            new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }));
        EmailDocument msg = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "detect" };
        byte[] bytes = new EmailDocumentWriter().ToBytes(msg, EmailFileFormat.OutlookMsg);
        Assert.Equal(EmailFileFormat.OutlookMsg, EmailDocumentReader.DetectFormat(bytes));
    }

    [Fact]
    public void DetectsMboxAndEmlFromContent() {
        Assert.Equal(EmailFileFormat.Mbox, EmailDocumentReader.DetectFormat(Encoding.ASCII.GetBytes("From sender@example.com now\n")));
        Assert.Equal(EmailFileFormat.Eml, EmailDocumentReader.DetectFormat(Encoding.ASCII.GetBytes("Subject: value\r\n\r\nbody")));
        Assert.Equal(EmailFileFormat.Eml, EmailDocumentReader.DetectFormat(Encoding.ASCII.GetBytes("X-Custom: value\r\n\r\nbody")));
        Assert.Equal(EmailFileFormat.Unknown, EmailDocumentReader.DetectFormat(Encoding.ASCII.GetBytes("{\"a\":1}")));
        Assert.Equal(EmailFileFormat.Unknown, EmailDocumentReader.DetectFormat(new byte[] { 1, 2, 3 }));
    }

    [Fact]
    public void EnforcesInputHeaderPartAndAttachmentLimits() {
        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentReader(
            new EmailReaderOptions(maxInputBytes: 3)).Read(new byte[] { 1, 2, 3, 4 }));

        string largeHeader = string.Concat("Subject: ", new string('a', 64), "\r\n\r\nbody");
        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentReader(
            new EmailReaderOptions(maxHeaderBytes: 16)).Read(Encoding.ASCII.GetBytes(largeHeader)));

        const string multipart = "Subject: x\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain\r\n\r\na\r\n" +
            "--x\r\nContent-Type: text/plain\r\n\r\nb\r\n--x--\r\n";
        EmailLimitExceededException partLimit = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(maxPartCount: 2))
                .Read(Encoding.ASCII.GetBytes(multipart)));
        Assert.Equal(nameof(EmailReaderOptions.MaxPartCount), partLimit.LimitName);
        Assert.Equal(3, partLimit.ActualValue);

        const string attachment = "Subject: x\r\nContent-Type: application/octet-stream\r\n\r\n12345";
        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentReader(
            new EmailReaderOptions(maxAttachmentBytes: 4)).Read(Encoding.ASCII.GetBytes(attachment)));

        const string base64Attachment = "Subject: x\r\nContent-Type: application/octet-stream\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nAQIDBAUG";
        EmailLimitExceededException decodedLimit = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(
                maxAttachmentBytes: 5, includeAttachmentContent: false))
                .Read(Encoding.ASCII.GetBytes(base64Attachment)));
        Assert.Equal(nameof(EmailReaderOptions.MaxAttachmentBytes), decodedLimit.LimitName);
        Assert.Equal(6, decodedLimit.ActualValue);
    }

    [Fact]
    public async Task AsyncReadUsesWholeSeekableArtifactAndRestoresPosition() {
        byte[] message = Encoding.ASCII.GetBytes("Subject: async\r\n\r\nbody");
        MemoryStream stream = new MemoryStream(message);
        stream.Position = 3;

        EmailReadResult result = await new EmailDocumentReader().ReadAsync(stream);

        Assert.Equal("async", result.Document.Subject);
        Assert.True(stream.CanRead);
        Assert.Equal(3, stream.Position);
        stream.Dispose();
    }

    [Fact]
    public void DetectFormatUsesWholeSeekableArtifactAndRestoresPosition() {
        byte[] message = Encoding.ASCII.GetBytes("Subject: detect-stream\r\n\r\nbody");
        using var stream = new MemoryStream(message);
        stream.Position = 5;

        EmailFileFormat format = EmailDocumentReader.DetectFormat(stream);

        Assert.Equal(EmailFileFormat.Eml, format);
        Assert.Equal(5, stream.Position);
    }

    [Fact]
    public async Task AsyncReadersAndWritersHonorCancellationBeforeCpuWork() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        byte[] source = Encoding.ASCII.GetBytes("Subject: cancelled\r\n\r\nbody");

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new EmailDocumentReader().ReadAsync(new MemoryStream(source), cancellation.Token));

        var document = new EmailDocument { Format = EmailFileFormat.Eml, Subject = "cancelled" };
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new EmailDocumentWriter().WriteAsync(document, new MemoryStream(), EmailFileFormat.Eml, cancellation.Token));

        var mailbox = new EmailMailbox();
        mailbox.Messages.Add(new EmailMailboxEntry(document));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new EmailMailboxWriter().WriteAsync(mailbox, new MemoryStream(), cancellation.Token));
    }

    [Fact]
    public async Task AsyncFileApisRoundTripMessageAndMailbox() {
        string directory = Path.Combine(Path.GetTempPath(), string.Concat("OfficeIMO-Email-Async-", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(directory);
        try {
            var document = new EmailDocument { Format = EmailFileFormat.Eml, Subject = "async-file" };
            document.Body.Text = "body";
            string emlPath = Path.Combine(directory, "message.eml");
            await new EmailDocumentWriter().WriteAsync(document, emlPath);
            EmailReadResult message = await new EmailDocumentReader().ReadAsync(emlPath);
            Assert.Equal("async-file", message.Document.Subject);

            var mailbox = new EmailMailbox();
            mailbox.Messages.Add(new EmailMailboxEntry(document));
            string mboxPath = Path.Combine(directory, "archive.mbox");
            await new EmailMailboxWriter().WriteAsync(mailbox, mboxPath);
            EmailMailboxReadResult archive = await new EmailMailboxReader().ReadAsync(mboxPath);
            Assert.Equal("async-file", Assert.Single(archive.Mailbox.Messages).Document.Subject);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void CanSkipAttachmentContentWhileKeepingMetadata() {
        const string eml = "Subject: x\r\nContent-Type: application/octet-stream; name=a.bin\r\n\r\n123";
        EmailReadResult result = new EmailDocumentReader(new EmailReaderOptions(includeAttachmentContent: false))
            .Read(Encoding.ASCII.GetBytes(eml));

        EmailAttachment attachment = Assert.Single(result.Document.Attachments);
        Assert.Equal(3, attachment.Length);
        Assert.Null(attachment.Content);
    }

    [Fact]
    public void EnforcesMapiPropertyAndWriterOutputLimits() {
        EmailDocument source = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "bounded" };
        byte[] msg = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);

        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentReader(
            new EmailReaderOptions(maxMapiPropertyCount: 1)).Read(msg));
        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentWriter(
            new EmailWriterOptions(maxOutputBytes: 10)).ToBytes(source, EmailFileFormat.OutlookMsg));
    }
}
