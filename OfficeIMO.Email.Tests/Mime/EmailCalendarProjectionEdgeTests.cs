using MimeKit;
using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailCalendarProjectionEdgeTests {
    [Fact]
    public void LegacyCalendarMimeProjectionPreservesLiteralCaretParameters() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:1.0\r\nPRODID:-//Legacy//EN\r\n" +
            "BEGIN:VEVENT\r\nUID:legacy@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ATTENDEE;CN=alpha^nbeta:mailto:attendee@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        EmailDocument document = new EmailDocumentReader().Read(eml).Document;
        EmailRecipient recipient = Assert.Single(document.Recipients);

        Assert.Equal("alpha^nbeta", recipient.Address.DisplayName);
        Assert.DoesNotContain("\n", recipient.Address.DisplayName!, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("text/calendar", OutlookItemKind.Appointment,
        "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:inline@example.com\r\n" +
        "DTSTART:20260801T100000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n")]
    [InlineData("text/vcard", OutlookItemKind.Contact,
        "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Ada Lovelace\r\nEMAIL:ada@example.com\r\nEND:VCARD\r\n")]
    public void ProjectsInlineSemanticMimeParts(string contentType, OutlookItemKind expectedKind, string content) {
        byte[] eml = Encoding.ASCII.GetBytes("Content-Type: " + contentType + "; charset=utf-8\r\n" +
            "Content-Disposition: inline\r\n\r\n" + content);

        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailAttachment semanticPart = Assert.Single(document.Attachments);
        Assert.Equal(expectedKind, document.OutlookItemKind);
        Assert.True(semanticPart.IsInline);
        Assert.True(semanticPart.IsMimeBodyPart);
        Assert.True(semanticPart.IsProjectedSemanticContent);
    }

    [Fact]
    public void ProjectsVtodoAttendeesThroughStoreRecipients() {
        byte[] eml = Calendar(
            "BEGIN:VTODO\r\nUID:assigned@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ATTENDEE;CN=Assignee:mailto:assignee@example.com\r\nEND:VTODO\r\n", "REQUEST");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;
        string regenerated = CalendarText(new EmailDocumentWriter().ToBytes(roundTrip, EmailFileFormat.Eml));

        Assert.Contains(document.Recipients, recipient => recipient.Address.Address == "assignee@example.com");
        Assert.Contains(roundTrip.Recipients, recipient => recipient.Address.Address == "assignee@example.com");
        Assert.Contains("ATTENDEE;ROLE=REQ-PARTICIPANT;CN=\"Assignee\":mailto:assignee@example.com",
            regenerated, StringComparison.Ordinal);
    }

    [Fact]
    public void BlocksFloatingCalendarTimesBeforeStoreConversion() {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:floating@example.com\r\nDTSTART:20260715T090000\r\nEND:VEVENT\r\n");
        EmailReadResult read = new EmailDocumentReader().Read(eml);
        EmailDocument document = read.Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(read.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_FLOATING_TIME");
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void DecodesPercentEncodedCalendarMailboxesBeforeStoreConversion() {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:encoded-address@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ORGANIZER:mailto:owner%2Bcalendar@example.com\r\n" +
            "ATTENDEE:mailto:alice%2Btag@example.com\r\nEND:VEVENT\r\n", "REQUEST");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal("owner+calendar@example.com", document.From!.Address);
        Assert.Equal("alice+tag@example.com", Assert.Single(document.Recipients).Address.Address);
        Assert.Equal("owner+calendar@example.com", roundTrip.From!.Address);
        Assert.Equal("alice+tag@example.com", Assert.Single(roundTrip.Recipients).Address.Address);
    }

    [Fact]
    public void PreservesLiteralPercentSequencesInCalendarMailboxes() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "Percent mailbox",
            Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero)
            }
        };
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("user%2Ctag@example.com")));

        byte[] eml = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml);
        EmailDocument roundTrip = new EmailDocumentReader().Read(eml).Document;

        Assert.Contains("mailto:user%252Ctag@example.com", CalendarText(eml), StringComparison.Ordinal);
        Assert.Contains(roundTrip.Recipients,
            recipient => recipient.Address.Address == "user%2Ctag@example.com");
    }

    [Fact]
    public void BlocksCalendarPriorityBeforeStoreConversion() {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:priority@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "PRIORITY:1\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("VEVENT")]
    [InlineData("VTODO")]
    public void BlocksCalendarUrlsBeforeStoreConversion(string component) {
        byte[] eml = Calendar(
            "BEGIN:" + component + "\r\nUID:url@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "URL:https://example.com/item\r\nEND:" + component + "\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("DTSTART")]
    [InlineData("DUE")]
    [InlineData("COMPLETED")]
    public void BlocksDateOnlyTaskDatesBeforeStoreConversion(string property) {
        byte[] eml = Calendar(
            "BEGIN:VTODO\r\nUID:date-only@example.com\r\n" + property +
            ";VALUE=DATE:20260715\r\nEND:VTODO\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("VEVENT")]
    [InlineData("VTODO")]
    public void BlocksCalendarTimestampsBeforeStoreConversion(string component) {
        byte[] eml = Calendar(
            "BEGIN:" + component + "\r\nUID:timestamp@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nDTSTAMP:20260715T080000Z\r\nEND:" + component + "\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("CANCEL")]
    [InlineData("REPLY")]
    [InlineData("ADD")]
    public void BlocksUnsupportedVtodoMethodsBeforeStoreConversion(string method) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; method=" + method + "; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:" + method + "\r\nBEGIN:VTODO\r\n" +
            "UID:task-method@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "END:VTODO\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksUnsupportedCalendarVersionsBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:1.0\r\n" +
            "BEGIN:VEVENT\r\nUID:version@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksVtodoSequenceBeforeStoreConversion() {
        byte[] eml = Calendar(
            "BEGIN:VTODO\r\nUID:task-sequence@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "SEQUENCE:3\r\nEND:VTODO\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("ORGANIZER:urn:uuid:owner")]
    [InlineData("ATTENDEE:urn:uuid:attendee")]
    public void BlocksNonMailtoCalendarAddressesBeforeStoreConversion(string property) {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:cal-address@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            property + "\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Null(document.From);
        Assert.Empty(document.Recipients);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("SENT-BY=mailto:assistant@example.com")]
    [InlineData("DIR=ldap://directory.example/owner")]
    public void BlocksUnsupportedOrganizerParametersBeforeStoreConversion(string parameter) {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:organizer-parameter@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ORGANIZER;" + parameter + ":mailto:owner@example.com\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("VEVENT", "CREATED:20260715T080000Z")]
    [InlineData("VTODO", "LAST-MODIFIED:20260715T090000Z")]
    [InlineData("VEVENT", "COMMENT:Bring documents")]
    [InlineData("VTODO", "RESOURCES:Conference room")]
    [InlineData("VEVENT", "GEO:52.2297;21.0122")]
    [InlineData("VEVENT", "CONTACT:calendar@example.com")]
    [InlineData("VTODO", "LOCATION:Warehouse")]
    public void BlocksUnprojectedCalendarAuditAndDetailFieldsBeforeStoreConversion(
        string component, string property) {
        byte[] eml = Calendar(
            "BEGIN:" + component + "\r\nUID:unprojected@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\n" + property + "\r\nEND:" + component + "\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("", "ATTENDEE:mailto:attendee@example.com\r\n")]
    [InlineData("To: recipient@example.com\r\n", "")]
    public void BlocksMethodlessCalendarsWithRecipientsBeforeStoreConversion(
        string transportHeaders, string calendarRecipients) {
        byte[] eml = Encoding.ASCII.GetBytes(
            transportHeaders + "Content-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\n" +
            "UID:methodless@example.com\r\nDTSTART:20260801T100000Z\r\n" + calendarRecipients +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void CalendarOptionalRoleOverridesEnvelopeRecipientKind() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "To: Optional <optional@example.com>\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:REQUEST\r\nBEGIN:VEVENT\r\nUID:optional@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\n" +
            "ATTENDEE;ROLE=OPT-PARTICIPANT;CN=Optional:mailto:optional@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal(EmailRecipientKind.Cc, Assert.Single(document.Recipients).Kind);
        Assert.Equal(EmailRecipientKind.Cc, Assert.Single(roundTrip.Recipients).Kind);
    }

    [Fact]
    public void UsesCalendarAttendeeNameWhenEnvelopeRecipientHasNoDisplayName() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "To: alice@example.com\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:REQUEST\r\nBEGIN:VEVENT\r\nUID:name@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nATTENDEE;CN=Alice:mailto:alice@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal("Alice", Assert.Single(document.Recipients).Address.DisplayName);
        Assert.Equal("Alice", Assert.Single(roundTrip.Recipients).Address.DisplayName);
    }

    [Fact]
    public void BlocksConflictingEnvelopeAndCalendarAttendeeNamesBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "To: Relay <alice@example.com>\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:name-conflict@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nATTENDEE;CN=Alice:mailto:alice@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Equal("Relay", Assert.Single(document.Recipients).Address.DisplayName);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("REQUEST", "IPM.Schedule.Meeting.Request")]
    [InlineData("CANCEL", "IPM.Schedule.Meeting.Canceled")]
    [InlineData("REPLY", "IPM.Schedule.Meeting.Resp.Pos")]
    public void UsesMimeCalendarMethodWhenPayloadMethodIsMissing(string method, string expectedMessageClass) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; method=" + method + "; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:method@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal(expectedMessageClass, document.MessageClass);
        Assert.Equal(expectedMessageClass, roundTrip.MessageClass);
    }

    [Fact]
    public void BlocksPublishCalendarsWhoseTransportRecipientsWouldBecomeAttendees() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "To: Distribution <distribution@example.com>\r\n" +
            "Content-Type: text/calendar; method=PUBLISH; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:PUBLISH\r\nBEGIN:VEVENT\r\n" +
            "UID:publish@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksConflictingMimeAndPayloadCalendarMethodsBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; method=REQUEST; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:CANCEL\r\nBEGIN:VEVENT\r\n" +
            "UID:method-conflict@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Equal("IPM.Schedule.Meeting.Canceled", document.MessageClass);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
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

    [Theory]
    [InlineData("ACTION:EMAIL\r\nATTENDEE:mailto:notify@example.com\r\n")]
    [InlineData("ACTION:DISPLAY\r\nATTACH:https://example.com/reminder.wav\r\n")]
    public void BlocksAlarmSemanticsThatOutlookReminderPropertiesCannotRepresent(string alarmProperties) {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:alarm@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "BEGIN:VALARM\r\n" + alarmProperties + "TRIGGER:-PT15M\r\n" +
            "END:VALARM\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("PARTSTAT=ACCEPTED")]
    [InlineData("RSVP=TRUE")]
    public void BlocksAttendeeParametersThatStoreRecipientsCannotRepresent(string parameter) {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:attendee-state@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ATTENDEE;" + parameter + ":mailto:alice@example.com\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksCustomAlarmDescriptionsThatStoreRemindersCannotRepresent() {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:alarm-description@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "SUMMARY:Meeting\r\nBEGIN:VALARM\r\nACTION:DISPLAY\r\n" +
            "DESCRIPTION:Bring the report\r\nTRIGGER:-PT15M\r\nEND:VALARM\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void PreservesAbsoluteAlarmTriggersThroughStoreConversion() {
        DateTimeOffset signal = new DateTimeOffset(2026, 8, 1, 9, 45, 0, TimeSpan.Zero);
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:absolute-alarm@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "SUMMARY:Reminder\r\nBEGIN:VALARM\r\nACTION:DISPLAY\r\nDESCRIPTION:Reminder\r\n" +
            "TRIGGER;VALUE=DATE-TIME:20260801T094500Z\r\nEND:VALARM\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);
        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.True(report.CanWrite);
        Assert.Equal(signal, document.Appointment!.ReminderSignalTime);
        Assert.Equal(signal, roundTrip.Appointment!.ReminderSignalTime);
    }

    [Fact]
    public void BlocksDistinctTransportMessageIdAndCalendarUidBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Message-ID: <transport@example.com>\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:event@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Equal("transport@example.com", document.MessageId);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksDistinctTransportFromAndCalendarOrganizerBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "From: Relay <relay@example.com>\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:event@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nORGANIZER;CN=Owner:mailto:owner@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Equal("relay@example.com", document.From!.Address);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void BlocksDistinctMimeBodyAndCalendarDescriptionBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\nContent-Type: multipart/alternative; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nWrapper text\r\n" +
            "--x\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:event@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nDESCRIPTION:Event notes\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n--x--\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Equal("Wrapper text", document.Body.Text!.Trim());
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void DecodesEscapedCalendarTextSequentially() {
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:escape@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "SUMMARY:Literal \\\\n value\r\nEND:VEVENT\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal("Literal \\n value", document.Subject);
        Assert.Equal("Literal \\n value", roundTrip.Subject);
    }

    [Fact]
    public void PreservesQuotedCalendarParameterBackslashes() {
        const string parameterName = "Alice \\\"A\\\" \\\\ Team";
        byte[] eml = Calendar(
            "BEGIN:VEVENT\r\nUID:parameter@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ATTENDEE;CN=\"" + parameterName + "\":mailto:alice@example.com\r\n" +
            "END:VEVENT\r\n", "REQUEST");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal("Alice \"A\" \\\\ Team", Assert.Single(document.Recipients).Address.DisplayName);
        Assert.Equal("Alice \"A\" \\\\ Team", Assert.Single(roundTrip.Recipients).Address.DisplayName);
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
