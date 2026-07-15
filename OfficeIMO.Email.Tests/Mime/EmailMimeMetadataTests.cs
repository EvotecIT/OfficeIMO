using MimeKit;
using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMimeMetadataTests {
    [Fact]
    public void RoundTripsPortableMessageMetadataThroughStandardMimeHeaders() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "Metadata",
            From = new EmailAddress("sender@example.com")
        };
        source.MessageMetadata.Importance = EmailMessageImportance.High;
        source.MessageMetadata.Priority = EmailMessagePriority.Urgent;
        source.MessageMetadata.Sensitivity = 2;
        source.MessageMetadata.ReadReceiptRequested = true;
        source.MessageMetadata.ReadReceiptDestination = "read-receipts@example.com";
        source.MessageMetadata.DeliveryReceiptRequested = true;
        source.MessageMetadata.DeliveryReceiptDestination = "delivery-receipts@example.com";
        source.MessageMetadata.IsDraft = true;
        source.MessageMetadata.IsRead = true;
        source.MessageMetadata.Categories.Add("Blue");
        source.MessageMetadata.Categories.Add("Project X");
        source.MessageMetadata.Categories.Add("Project, X");

        EmailDocument result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml)).Document;

        Assert.Equal(EmailMessageImportance.High, result.MessageMetadata.Importance);
        Assert.Equal(EmailMessagePriority.Urgent, result.MessageMetadata.Priority);
        Assert.Equal(2, result.MessageMetadata.Sensitivity);
        Assert.True(result.MessageMetadata.ReadReceiptRequested);
        Assert.Equal("read-receipts@example.com", result.MessageMetadata.ReadReceiptDestination);
        Assert.True(result.MessageMetadata.DeliveryReceiptRequested);
        Assert.Equal("delivery-receipts@example.com", result.MessageMetadata.DeliveryReceiptDestination);
        Assert.True(result.MessageMetadata.IsDraft);
        Assert.True(result.MessageMetadata.IsRead);
        Assert.Equal(new[] { "Blue", "Project X", "Project, X" }, result.MessageMetadata.Categories);
    }

    [Fact]
    public void ReportsOpaqueMapiMetadataDuringEmlConversionWithoutBlockingCommonContent() {
        var source = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "Mapi" };
        source.MapiProperties.Add(new MapiProperty(0x66aa, MapiPropertyType.Binary, new byte[] { 1 }));

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);

        Assert.True(report.CanWrite);
        Assert.True(report.HasPotentialDataLoss);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_SOURCE_METADATA_NOT_REPRESENTED_IN_EML");
    }

    [Fact]
    public void FormatsReceiptDestinationsAsMailboxHeaders() {
        var source = new EmailDocument { Format = EmailFileFormat.OutlookMsg, Subject = "Receipts" };
        source.MessageMetadata.ReadReceiptRequested = true;
        source.MessageMetadata.ReadReceiptDestination = "Żaneta <read@example.com>";
        source.MessageMetadata.DeliveryReceiptRequested = true;
        source.MessageMetadata.DeliveryReceiptDestination = "Łukasz <delivery@example.com>";

        byte[] output = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml);
        using var stream = new MemoryStream(output);
        MimeMessage message = MimeMessage.Load(stream);
        string? readHeader = message.Headers["Disposition-Notification-To"];
        string? deliveryHeader = message.Headers["Return-Receipt-To"];
        Assert.NotNull(readHeader);
        Assert.NotNull(deliveryHeader);
        MailboxAddress read = Assert.IsType<MailboxAddress>(Assert.Single(
            InternetAddressList.Parse(readHeader!)));
        MailboxAddress delivery = Assert.IsType<MailboxAddress>(Assert.Single(
            InternetAddressList.Parse(deliveryHeader!)));

        Assert.Equal("Żaneta", read.Name);
        Assert.Equal("read@example.com", read.Address);
        Assert.Equal("Łukasz", delivery.Name);
        Assert.Equal("delivery@example.com", delivery.Address);
    }

    [Fact]
    public void RetainsMetadataHeadersThatCannotBeProjected() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Importance: critical\r\nPriority: immediate\r\nX-Unsent: maybe\r\n" +
            "Content-Type: text/plain; charset=utf-8\r\n\r\nBody\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        byte[] output = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml);
        using var stream = new MemoryStream(output);
        MimeMessage message = MimeMessage.Load(stream);

        Assert.Equal("critical", message.Headers["Importance"]);
        Assert.Equal("immediate", message.Headers["Priority"]);
        Assert.Equal("maybe", message.Headers["X-Unsent"]);
    }
}
