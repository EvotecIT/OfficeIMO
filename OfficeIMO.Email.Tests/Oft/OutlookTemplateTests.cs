using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookTemplateTests {
    [Fact]
    public void TemplateUsesMsgCompoundPayloadAndSourceNameSemantics() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookTemplate,
            OutlookItemKind = OutlookItemKind.Message,
            Subject = "Reusable template"
        };
        source.Body.Html = "<html><body>Template body</body></html>";

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookTemplate);
        EmailDocument contentOnly = new EmailDocumentReader().Read(bytes).Document;
        EmailDocument template = new EmailDocumentReader().Read(bytes, "reusable.oft").Document;

        Assert.Equal(EmailFileFormat.OutlookMsg, EmailDocumentReader.DetectFormat(bytes));
        Assert.Equal(EmailFileFormat.OutlookMsg, contentOnly.Format);
        Assert.Equal(EmailFileFormat.OutlookTemplate, template.Format);
        Assert.Equal("Reusable template", template.Subject);
        Assert.Contains("Template body", template.Body.Html, StringComparison.Ordinal);
        Assert.True(template.MessageMetadata.IsDraft);
    }

    [Fact]
    public void DocumentSaveAndLoadInferOftFromTheFileName() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-oft-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "appointment.oft");
            var source = new EmailDocument {
                OutlookItemKind = OutlookItemKind.Appointment,
                Subject = "Appointment template",
                Appointment = new OutlookAppointment {
                    Start = new DateTimeOffset(2026, 8, 10, 8, 0, 0, TimeSpan.Zero),
                    End = new DateTimeOffset(2026, 8, 10, 9, 0, 0, TimeSpan.Zero),
                    Location = "Room A"
                }
            };

            source.Save(path);
            EmailDocument roundTrip = EmailDocument.Load(path);

            Assert.Equal(EmailFileFormat.OutlookTemplate, roundTrip.Format);
            Assert.Equal(OutlookItemKind.Appointment, roundTrip.OutlookItemKind);
            Assert.Equal("Room A", roundTrip.Appointment!.Location);
            Assert.True(roundTrip.MessageMetadata.IsDraft);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData(OutlookItemKind.Message)]
    [InlineData(OutlookItemKind.Appointment)]
    [InlineData(OutlookItemKind.Contact)]
    [InlineData(OutlookItemKind.Task)]
    [InlineData(OutlookItemKind.Journal)]
    [InlineData(OutlookItemKind.Note)]
    public void TemplateReusesTypedMapiProjection(OutlookItemKind itemKind) {
        var source = new EmailDocument { OutlookItemKind = itemKind, Subject = itemKind + " template" };

        byte[] bytes = source.ToBytes(EmailFileFormat.OutlookTemplate);
        EmailDocument roundTrip = new EmailDocumentReader().Read(bytes, "typed.oft").Document;

        Assert.Equal(EmailFileFormat.OutlookTemplate, roundTrip.Format);
        Assert.Equal(itemKind, roundTrip.OutlookItemKind);
    }

    [Fact]
    public void PreservedTemplateSourceCanBeReemittedByteForByte() {
        var source = new EmailDocument { Subject = "Preserved template" };
        byte[] bytes = source.ToBytes(EmailFileFormat.OutlookTemplate);
        var reader = new EmailDocumentReader(new EmailReaderOptions(preserveRawSource: true));

        EmailDocument document = reader.Read(bytes, "preserved.oft").Document;
        byte[] preserved = new EmailDocumentWriter(new EmailWriterOptions(usePreservedRawSource: true))
            .ToBytes(document, EmailFileFormat.OutlookTemplate);

        Assert.Equal(bytes, preserved);
    }
}
