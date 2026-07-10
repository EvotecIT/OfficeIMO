using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailReaderContractTests {
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
