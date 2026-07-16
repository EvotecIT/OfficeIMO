using MimeKit;
using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailCalendarFidelityRegressionTests {
    [Theory]
    [InlineData("ROOM")]
    [InlineData("RESOURCE")]
    public void BlocksRoomAndResourceAttendeesWithoutExplicitNonParticipantRole(string calendarUserType) {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:room-role@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ATTENDEE;CUTYPE=" + calendarUserType + ":mailto:asset@example.com\r\nEND:VEVENT\r\n",
            "REQUEST");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksParameterizedCalendarLocationsBeforeStoreConversion() {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:location-parameter@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "LOCATION;ALTREP=\"cid:map\":Room\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.Equal("Room", document.Appointment!.Location);
        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void PreservesVtodoOrganizerDisplayNameThroughStoreConversion() {
        byte[] eml = Calendar(
            "BEGIN:VTODO\r\nUID:task-owner@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ORGANIZER;CN=Owner:mailto:owner@example.com\r\nEND:VTODO\r\n", "PUBLISH");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);
        EmailDocument stored = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;
        string regenerated = CalendarText(new EmailDocumentWriter().ToBytes(stored, EmailFileFormat.Eml));

        Assert.True(report.CanWrite);
        Assert.Contains("ORGANIZER;CN=\"Owner\":mailto:owner@example.com", regenerated,
            StringComparison.Ordinal);
    }

    [Fact]
    public void CalendarOnlyDescriptionDoesNotBecomeAnExtraMimeBody() {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:description-only@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "DESCRIPTION:Calendar description\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        byte[] rewritten = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml);
        using var stream = new MemoryStream(rewritten);
        MimeMessage message = MimeMessage.Load(stream);

        MimePart calendar = Assert.IsAssignableFrom<MimePart>(message.Body);
        Assert.Equal("text/calendar", calendar.ContentType.MimeType);
        Assert.DoesNotContain(message.BodyParts.OfType<MimePart>(),
            part => part.ContentType.MimeType == "text/plain");
    }

    [Fact]
    public void BlocksPartiallyUnmatchedAppointmentAttendeeDisplays() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero),
                RequiredAttendees = "Alice; Bob"
            }
        };
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("alice@example.com", "Alice")));

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_ATTENDEE_ADDRESS_REQUIRED");
    }

    private static byte[] Calendar(string component, string? method = null) => Encoding.ASCII.GetBytes(
        "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
        (string.IsNullOrWhiteSpace(method) ? string.Empty : "METHOD:" + method + "\r\n") +
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
