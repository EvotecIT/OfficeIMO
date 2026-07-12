using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class MsgResourceLimitTests {
    [Fact]
    public void RejectsByValueAttachmentBeforeBufferingItsCompoundStream() {
        var source = new EmailDocument { Subject = "bounded attachment" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "large.bin",
            Content = new byte[1000],
            Length = 1000
        });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(maxAttachmentBytes: 512)).Read(bytes));

        Assert.Equal(nameof(EmailReaderOptions.MaxAttachmentBytes), exception.LimitName);
        Assert.Equal(1000, exception.ActualValue);
    }

    [Fact]
    public void RejectsEmbeddedMsgStorageBeyondTheAttachmentLimit() {
        var embedded = new EmailDocument { Subject = "embedded" };
        embedded.Body.Text = new string('x', 5000);
        var source = new EmailDocument { Subject = "parent" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "embedded.msg",
            EmbeddedDocument = embedded
        });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(maxAttachmentBytes: 1024)).Read(bytes));

        Assert.Equal(nameof(EmailReaderOptions.MaxAttachmentBytes), exception.LimitName);
        Assert.True(exception.ActualValue > 1024);
    }

    [Fact]
    public void RejectsAggregateMsgPayloadStreamsBeforeBufferingTheSecondAttachment() {
        var source = new EmailDocument { Subject = "aggregate attachments" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "one.bin", Content = new byte[3000], Length = 3000
        });
        source.Attachments.Add(new EmailAttachment {
            FileName = "two.bin", Content = new byte[3000], Length = 3000
        });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailDocumentReader(new EmailReaderOptions(
                maxAttachmentBytes: 4000,
                maxTotalAttachmentBytes: 5000)).Read(bytes));

        Assert.Equal(nameof(EmailReaderOptions.MaxTotalAttachmentBytes), exception.LimitName);
        Assert.Equal(6000, exception.ActualValue);
    }

    [Fact]
    public void AttachmentPayloadDoesNotConsumeTheDecodedPropertyBudget() {
        var source = new EmailDocument { Subject = "separate budgets" };
        source.Attachments.Add(new EmailAttachment {
            FileName = "payload.bin",
            Content = new byte[1024 * 1024],
            Length = 1024 * 1024
        });
        byte[] bytes = new EmailDocumentWriter().WriteToBytes(source, EmailFileFormat.OutlookMsg);

        EmailReadResult result = new EmailDocumentReader(new EmailReaderOptions(
            maxAttachmentBytes: 2 * 1024 * 1024,
            maxTotalAttachmentBytes: 2 * 1024 * 1024,
            maxDecodedPropertyBytes: 64 * 1024)).Read(bytes);

        Assert.Equal(EmailFileFormat.OutlookMsg, result.Document.Format);
        Assert.Equal(1024 * 1024, Assert.Single(result.Document.Attachments).Content!.Length);
    }
}
