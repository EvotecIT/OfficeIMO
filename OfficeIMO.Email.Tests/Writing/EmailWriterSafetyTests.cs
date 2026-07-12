using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailWriterSafetyTests {
    [Theory]
    [InlineData(EmailFileFormat.Eml)]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.Tnef)]
    public void RejectsRetainedPayloadsBeforeFormatSerialization(EmailFileFormat format) {
        byte[] payload = new byte[4096];
        var document = new EmailDocument { Format = format, Subject = "bounded" };
        document.Attachments.Add(new EmailAttachment {
            FileName = "large.bin",
            ContentType = "application/octet-stream",
            Content = payload,
            Length = payload.Length
        });
        var writer = new EmailDocumentWriter(new EmailWriterOptions(maxOutputBytes: 1024));

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            writer.WriteToBytes(document, format));

        Assert.Equal(nameof(EmailWriterOptions.MaxOutputBytes), exception.LimitName);
        Assert.Equal(payload.LongLength, exception.ActualValue);
        Assert.Equal(1024, exception.MaximumValue);
    }

    [Fact]
    public void RejectsLargeBodyBeforeBase64Materialization() {
        var document = new EmailDocument { Format = EmailFileFormat.Eml };
        document.Body.Text = new string('x', 4096);
        var writer = new EmailDocumentWriter(new EmailWriterOptions(maxOutputBytes: 1024));

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            writer.WriteToBytes(document));

        Assert.Equal(nameof(EmailWriterOptions.MaxOutputBytes), exception.LimitName);
        Assert.Equal(4096, exception.ActualValue);
    }
}
