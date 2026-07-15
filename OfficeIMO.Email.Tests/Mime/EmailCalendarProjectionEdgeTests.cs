using MimeKit;
using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailCalendarProjectionEdgeTests {
    [Fact]
    public void ProjectsVtodoAttendeesThroughStoreRecipients() {
        byte[] eml = Calendar(
            "BEGIN:VTODO\r\nUID:assigned@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ATTENDEE;CN=Assignee:mailto:assignee@example.com\r\nEND:VTODO\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Contains(document.Recipients, recipient => recipient.Address.Address == "assignee@example.com");
        Assert.Contains(roundTrip.Recipients, recipient => recipient.Address.Address == "assignee@example.com");
    }

    [Fact]
    public void PreservesVtodoPrivacyThroughStoreConversion() {
        byte[] eml = Calendar(
            "BEGIN:VTODO\r\nUID:private-task@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "CLASS:CONFIDENTIAL\r\nEND:VTODO\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument storeRoundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;
        string regenerated = CalendarText(new EmailDocumentWriter().ToBytes(storeRoundTrip, EmailFileFormat.Eml));

        Assert.Equal(3, document.MessageMetadata.Sensitivity);
        Assert.Equal(3, storeRoundTrip.MessageMetadata.Sensitivity);
        Assert.Contains("CLASS:CONFIDENTIAL", regenerated, StringComparison.Ordinal);
    }

    [Fact]
    public void DerivesVtodoDueDateFromStandardDuration() {
        DateTimeOffset start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero);
        byte[] eml = Calendar(
            "BEGIN:VTODO\r\nUID:duration-task@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "DURATION:PT2H\r\nEND:VTODO\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument storeRoundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal(start.AddHours(2), document.Task!.Due);
        Assert.Null(document.Task.EstimatedEffort);
        Assert.Equal(start.AddHours(2), storeRoundTrip.Task!.Due);
    }

    [Theory]
    [InlineData("VEVENT")]
    [InlineData("VTODO")]
    public void PreservesCalendarCategoriesThroughStoreConversion(string component) {
        byte[] eml = Calendar(
            "BEGIN:" + component + "\r\nUID:categories@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "CATEGORIES:Blue,Project\\, X\r\nEND:" + component + "\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument storeRoundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal(new[] { "Blue", "Project, X" }, document.MessageMetadata.Categories);
        Assert.Equal(new[] { "Blue", "Project, X" }, storeRoundTrip.MessageMetadata.Categories);
    }

    [Theory]
    [InlineData("TENTATIVE")]
    [InlineData("CANCELLED")]
    public void BlocksVeventStatusWhenTheStoreModelCannotRepresentItExactly(string status) {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:status@example.com\r\nDTSTART:20260801T100000Z\r\nSTATUS:" + status +
            "\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksMultipleSemanticMimeBodyPartsBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\nContent-Type: multipart/alternative; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:first@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n" +
            "--x\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:second@example.com\r\n" +
            "DTSTART:20260802T100000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n--x--\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.Equal(2, document.Attachments.Count(attachment => attachment.IsProjectedSemanticContent));
        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    private static byte[] Calendar(string component) => Encoding.ASCII.GetBytes(
        "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
        component + "END:VCALENDAR\r\n");

    private static string CalendarText(byte[] eml) {
        using var stream = new MemoryStream(eml);
        MimePart calendar = Assert.Single(MimeMessage.Load(stream).BodyParts.OfType<MimePart>(),
            part => part.ContentType.MimeType == "text/calendar");
        using var content = new MemoryStream();
        calendar.Content!.DecodeTo(content);
        return Encoding.UTF8.GetString(content.ToArray());
    }
}
