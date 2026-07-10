using OfficeIMO.Email;
using System.Threading;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailReaderContractTests {
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
    public void DetectsTnefAndRequiresMsgDirectoryContractForCompoundFiles() {
        Assert.Equal(EmailFileFormat.Tnef, EmailDocumentReader.DetectFormat(new byte[] { 0x78, 0x9F, 0x3E, 0x22 }));
        Assert.Equal(EmailFileFormat.Unknown, EmailDocumentReader.DetectFormat(
            new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }));
        EmailDocument msg = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "detect" };
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(msg, EmailFileFormat.OutlookMsg);
        Assert.Equal(EmailFileFormat.OutlookMsg, EmailDocumentReader.DetectFormat(bytes));
    }

    [Fact]
    public void DetectsMboxAndEmlFromContent() {
        Assert.Equal(EmailFileFormat.Mbox, EmailDocumentReader.DetectFormat(Encoding.ASCII.GetBytes("From sender@example.com now\n")));
        Assert.Equal(EmailFileFormat.Eml, EmailDocumentReader.DetectFormat(Encoding.ASCII.GetBytes("Subject: value\r\n\r\nbody")));
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
        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentReader(
            new EmailReaderOptions(maxPartCount: 2)).Read(Encoding.ASCII.GetBytes(multipart)));

        const string attachment = "Subject: x\r\nContent-Type: application/octet-stream\r\n\r\n12345";
        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentReader(
            new EmailReaderOptions(maxAttachmentBytes: 4)).Read(Encoding.ASCII.GetBytes(attachment)));
    }

    [Fact]
    public async Task AsyncReadUsesCurrentPositionAndLeavesStreamOpen() {
        byte[] prefix = { 1, 2, 3 };
        byte[] message = Encoding.ASCII.GetBytes("Subject: async\r\n\r\nbody");
        MemoryStream stream = new MemoryStream(prefix.Concat(message).ToArray());
        stream.Position = prefix.Length;

        EmailReadResult result = await new EmailDocumentReader().ReadAsync(stream);

        Assert.Equal("async", result.Document.Subject);
        Assert.True(stream.CanRead);
        stream.Dispose();
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
        byte[] msg = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentReader(
            new EmailReaderOptions(maxMapiPropertyCount: 1)).Read(msg));
        Assert.Throws<EmailLimitExceededException>(() => new EmailDocumentWriter(
            new EmailWriterOptions(maxOutputBytes: 10)).WriteToBytes(source, EmailFileFormat.OutlookMsg));
    }
}
