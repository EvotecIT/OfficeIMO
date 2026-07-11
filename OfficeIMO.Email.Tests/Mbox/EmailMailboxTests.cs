using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMailboxTests {
    [Fact]
    public void ReadsAndWritesMboxrdWithEnvelopeAndFromLineEscaping() {
        const string source = "From first@example.com Fri Jul 10 12:00:00 2026\n" +
            "From: first@example.com\nSubject: First\nContent-Type: text/plain; charset=utf-8\n\n" +
            ">From escaped separator\n>>From originally quoted\n" +
            "From second@example.com Fri Jul 10 13:00:00 2026\n" +
            "From: second@example.com\nSubject: Second\nContent-Type: text/plain; charset=utf-8\n\nsecond body\n";
        var messageOptions = new EmailReaderOptions(preserveRawSource: true);
        var reader = new EmailMailboxReader(new EmailMailboxReaderOptions(messageOptions));

        EmailMailboxReadResult result = reader.Read(Encoding.UTF8.GetBytes(source));

        Assert.Equal(2, result.Mailbox.Messages.Count);
        Assert.Equal("first@example.com", result.Mailbox.Messages[0].EnvelopeSender);
        Assert.Contains("From escaped separator", result.Mailbox.Messages[0].Document.Body.Text);
        Assert.Contains(">From originally quoted", result.Mailbox.Messages[0].Document.Body.Text);
        Assert.Equal("Second", result.Mailbox.Messages[1].Document.Subject);

        var writerOptions = new EmailMailboxWriterOptions(
            new EmailWriterOptions(usePreservedRawSource: true), MboxVariant.Mboxrd);
        byte[] rewritten = new EmailMailboxWriter(writerOptions).WriteToBytes(result.Mailbox);
        EmailMailboxReadResult reparsed = reader.Read(rewritten);
        Assert.Equal(2, reparsed.Mailbox.Messages.Count);
        Assert.Contains("From escaped separator", reparsed.Mailbox.Messages[0].Document.Body.Text);
        Assert.Contains(">From originally quoted", reparsed.Mailbox.Messages[0].Document.Body.Text);
    }

    [Fact]
    public async Task AsyncMailboxReadLeavesStreamOpenAndEnforcesMessageLimit() {
        byte[] source = Encoding.ASCII.GetBytes(
            "From a@example.com Fri Jul 10 12:00:00 2026\nSubject: A\n\nA\n" +
            "From b@example.com Fri Jul 10 13:00:00 2026\nSubject: B\n\nB\n");
        MemoryStream stream = new MemoryStream(source);

        EmailMailboxReadResult result = await new EmailMailboxReader().ReadAsync(stream);

        Assert.Equal(2, result.Mailbox.Messages.Count);
        Assert.True(stream.CanRead);
        Assert.Throws<EmailLimitExceededException>(() => new EmailMailboxReader(
            new EmailMailboxReaderOptions(maxMessageCount: 1)).Read(source));
        stream.Dispose();
    }

    [Fact]
    public void SingleDocumentReaderDirectsMboxToAggregateApi() {
        byte[] source = Encoding.ASCII.GetBytes("From a@example.com Fri Jul 10 12:00:00 2026\nSubject: A\n\nA\n");

        EmailReadResult result = new EmailDocumentReader().Read(source);

        Assert.Equal(EmailFileFormat.Mbox, result.Document.Format);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MBOX_REQUIRES_MAILBOX_READER");
    }

    [Fact]
    public void MailboxWriterReturnsPerMessageDiagnostics() {
        var document = new EmailDocument { Subject = "lossy attachment" };
        document.Attachments.Add(new EmailAttachment { FileName = "missing.bin", Length = 12 });
        var mailbox = new EmailMailbox();
        mailbox.Messages.Add(new EmailMailboxEntry(document));

        using var output = new MemoryStream();
        EmailWriteResult result = new EmailMailboxWriter().Write(mailbox, output);

        Assert.Contains(result.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE");
    }
}
