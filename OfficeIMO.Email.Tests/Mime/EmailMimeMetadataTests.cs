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
        Assert.Equal(new[] { "Blue", "Project X" }, result.MessageMetadata.Categories);
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
}
