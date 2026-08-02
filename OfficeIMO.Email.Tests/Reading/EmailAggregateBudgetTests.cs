using System.Text;

namespace OfficeIMO.Email.Tests;

public sealed class EmailAggregateBudgetTests {
    [Theory]
    [InlineData(EmailFileFormat.Eml)]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.Tnef)]
    public void AggregateBudgetIsReportedForArtifactAndAttachment(EmailFileFormat format) {
        var document = new EmailDocument { Subject = "budget", OutlookItemKind = OutlookItemKind.Message };
        document.Attachments.Add(new EmailAttachment {
            FileName = "payload.bin", ContentType = "application/octet-stream",
            Content = Encoding.UTF8.GetBytes("aggregate payload")
        });
        byte[] bytes = new EmailDocumentWriter().ToBytes(document, format);

        EmailReadResult result = new EmailDocumentReader().Read(bytes);

        Assert.Equal(bytes.LongLength, result.ProcessingBudget.InputBytes);
        Assert.Equal(1, result.ProcessingBudget.AttachmentCount);
        Assert.Equal("aggregate payload".Length, result.ProcessingBudget.AttachmentBytes);
        if (format == EmailFileFormat.Eml) Assert.True(result.ProcessingBudget.PartCount >= 2);
        else Assert.True(result.ProcessingBudget.PropertyCount > 0);
    }

    [Fact]
    public void AggregateAttachmentLimitStopsAcrossMultipleMimeEntities() {
        const string eml = "MIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: application/octet-stream\r\nContent-Disposition: attachment; filename=a\r\n\r\n12345\r\n" +
            "--x\r\nContent-Type: application/octet-stream\r\nContent-Disposition: attachment; filename=b\r\n\r\n67890\r\n--x--\r\n";
        var reader = new EmailDocumentReader(new EmailReaderOptions(
            maxAttachmentBytes: 8, maxTotalAttachmentBytes: 8));

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            reader.Read(Encoding.ASCII.GetBytes(eml)));

        Assert.Equal(nameof(EmailReaderOptions.MaxTotalAttachmentBytes), exception.LimitName);
        EmailDiagnostic diagnostic = exception.Diagnostic;
        Assert.Equal(EmailDiagnosticDisposition.Stopped, diagnostic.Disposition);
        Assert.Equal(EmailDataLossRisk.None, diagnostic.DataLossRisk);
        Assert.False(string.IsNullOrWhiteSpace(diagnostic.SuggestedAction));
        Assert.Equal(exception.ActualValue, diagnostic.ActualValue);
    }
}
