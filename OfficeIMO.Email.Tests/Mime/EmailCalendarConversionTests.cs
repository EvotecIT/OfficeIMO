using MimeKit;
using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailCalendarConversionTests {
    [Fact]
    public void ConvertsAppointmentThroughStandardsBasedEmlWithoutLosingCoreSemantics() {
        DateTimeOffset start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.FromHours(2));
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            MessageClass = "IPM.Schedule.Meeting.Request",
            Subject = "Planning",
            MessageId = "planning@example.com",
            From = new EmailAddress("organizer@example.com", "Organizer"),
            Appointment = new OutlookAppointment {
                Start = start,
                End = start.AddHours(2),
                Location = "Room 1",
                BusyStatus = 2,
                Sequence = 3,
                ReminderIsSet = true,
                ReminderDeltaMinutes = 20,
                ReminderTime = start.AddMinutes(-20),
                ReminderSignalTime = start.AddMinutes(-20)
            }
        };
        source.Body.Text = "Agenda";
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("attendee@example.com", "Attendee")));

        byte[] eml = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml);
        EmailReadResult read = new EmailDocumentReader().Read(eml);
        using var oracleStream = new MemoryStream(eml);
        MimeMessage oracle = MimeMessage.Load(oracleStream);

        Assert.Contains(oracle.BodyParts.OfType<MimePart>(), part => part.ContentType.MimeType == "text/calendar");
        Assert.Equal(OutlookItemKind.Appointment, read.Document.OutlookItemKind);
        Assert.Equal("Planning", read.Document.Subject);
        Assert.Equal(start.UtcDateTime, read.Document.Appointment!.Start!.Value.UtcDateTime);
        Assert.Equal(start.AddHours(2).UtcDateTime, read.Document.Appointment.End!.Value.UtcDateTime);
        Assert.Equal("Room 1", read.Document.Appointment.Location);
        Assert.Equal(3, read.Document.Appointment.Sequence);
        Assert.True(read.Document.Appointment.ReminderIsSet);
        Assert.Equal(20, read.Document.Appointment.ReminderDeltaMinutes);
        Assert.Equal(start.AddMinutes(-20).UtcDateTime,
            read.Document.Appointment.ReminderSignalTime!.Value.UtcDateTime);
        Assert.Equal("attendee@example.com", Assert.Single(read.Document.Recipients).Address.Address);
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void ProjectsEventPropertiesWithoutUsingTimezoneComponentValues() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VTIMEZONE\r\nTZID:Example\r\nBEGIN:STANDARD\r\nDTSTART:20000101T020000\r\n" +
            "END:STANDARD\r\nEND:VTIMEZONE\r\nBEGIN:VEVENT\r\nUID:event@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nDTEND:20260801T110000Z\r\nSUMMARY:Scoped event\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        Assert.Equal("Scoped event", document.Subject);
        Assert.Equal(new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero), document.Appointment!.Start);
        Assert.Equal(new DateTimeOffset(2026, 8, 1, 11, 0, 0, TimeSpan.Zero), document.Appointment.End);
    }

    [Fact]
    public void ProjectsUtcRecurrenceLimitsThroughMatchingEmbeddedTimeZoneRules() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VTIMEZONE\r\nTZID:Example/Eastern\r\n" +
            "BEGIN:DAYLIGHT\r\nDTSTART:20260308T020000\r\nTZOFFSETFROM:-0500\r\nTZOFFSETTO:-0400\r\n" +
            "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU\r\nEND:DAYLIGHT\r\n" +
            "BEGIN:STANDARD\r\nDTSTART:20261101T020000\r\nTZOFFSETFROM:-0400\r\nTZOFFSETTO:-0500\r\n" +
            "RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU\r\nEND:STANDARD\r\nEND:VTIMEZONE\r\n" +
            "BEGIN:VEVENT\r\nUID:zoned-series@example.com\r\n" +
            "DTSTART;TZID=Example/Eastern:20260701T230000\r\n" +
            "DTEND;TZID=Example/Eastern:20260702T000000\r\n" +
            "RRULE:FREQ=DAILY;UNTIL=20260703T030000Z\r\n" +
            "EXDATE:20260703T030000Z\r\nSUMMARY:Zoned series\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

        EmailReadResult read = new EmailDocumentReader().Read(eml);
        OutlookAppointment appointment = read.Document.Appointment!;
        OutlookRecurrence recurrence = appointment.Recurrence!;

        Assert.NotNull(appointment.RecurrenceTimeZone);
        Assert.Equal("Example/Eastern", appointment.RecurrenceTimeZone!.KeyName);
        Assert.Equal(new DateTime(2026, 7, 2), recurrence.EndDate);
        Assert.Equal(new DateTime(2026, 7, 2), Assert.Single(recurrence.DeletedOccurrenceDates));
        Assert.Equal(new DateTime(2026, 7, 2, 23, 0, 0),
            appointment.RecurrenceTimeZone.ConvertUtc(
                new DateTimeOffset(2026, 7, 3, 3, 0, 0, TimeSpan.Zero)).DateTime);
        Assert.DoesNotContain(read.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_RECURRENCE_TIMEZONE_REQUIRED");
    }

    [Fact]
    public void WithholdsTypedRecurrenceWhenUtcLimitsHaveNoMatchingTimeZoneRules() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:unresolved-series@example.com\r\n" +
            "DTSTART;TZID=Missing/Zone:20260701T230000\r\n" +
            "DTEND;TZID=Missing/Zone:20260702T000000\r\n" +
            "RRULE:FREQ=DAILY;UNTIL=20260703T030000Z\r\n" +
            "SUMMARY:Unresolved series\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

        EmailReadResult read = new EmailDocumentReader().Read(eml);

        Assert.Null(read.Document.Appointment!.Recurrence);
        Assert.Contains(read.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_RECURRENCE_TIMEZONE_REQUIRED");
    }

    [Fact]
    public void ParsesQuotedAttendeeParametersContainingDelimiters() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:event@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ATTENDEE;ROLE=REQ-PARTICIPANT;CN=\"Doe; John: Sr.\":mailto:john@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        EmailRecipient attendee = Assert.Single(new EmailDocumentReader().Read(eml).Document.Recipients);

        Assert.Equal("john@example.com", attendee.Address.Address);
        Assert.Equal("Doe; John: Sr.", attendee.Address.DisplayName);
    }

    [Fact]
    public void DecodesCalendarProjectionUsingTheDeclaredCharset() {
        byte[] prefix = Encoding.ASCII.GetBytes("Content-Type: text/calendar; charset=windows-1252\r\n\r\n");
        byte[] calendar = Encoding.ASCII.GetBytes(
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:event@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nSUMMARY:Caf#\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        for (int index = 0; index < calendar.Length; index++) {
            if (calendar[index] == (byte)'#') calendar[index] = 0xe9;
        }

        EmailDocument document = new EmailDocumentReader().Read(prefix.Concat(calendar).ToArray()).Document;
        EmailDocument roundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml)).Document;

        Assert.Equal("Café", document.Subject);
        Assert.Equal("Café", roundTrip.Subject);
        EmailAttachment retained = Assert.Single(roundTrip.Attachments,
            attachment => string.Equals(attachment.ContentType, "text/calendar", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("windows-1252", retained.ContentTypeParameters["charset"]);
    }

    [Fact]
    public void ParsesWeekFormEventAndReminderDurations() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:weekly@example.com\r\nDTSTART:20260801T100000Z\r\nDURATION:P1W\r\n" +
            "BEGIN:VALARM\r\nACTION:DISPLAY\r\nTRIGGER:-P1W\r\nEND:VALARM\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        OutlookAppointment appointment = new EmailDocumentReader().Read(eml).Document.Appointment!;

        Assert.Equal(new DateTimeOffset(2026, 8, 8, 10, 0, 0, TimeSpan.Zero), appointment.End);
        Assert.Equal(7 * 24 * 60, appointment.DurationMinutes);
        Assert.Equal(7 * 24 * 60, appointment.ReminderDeltaMinutes);
    }

    [Fact]
    public void ParsesExplicitlyPositiveEventAndReminderDurations() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:positive@example.com\r\nDTSTART:20260801T100000Z\r\nDURATION:+PT1H\r\n" +
            "BEGIN:VALARM\r\nACTION:DISPLAY\r\nTRIGGER:+PT15M\r\nEND:VALARM\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        OutlookAppointment appointment = new EmailDocumentReader().Read(eml).Document.Appointment!;

        Assert.Equal(new DateTimeOffset(2026, 8, 1, 11, 0, 0, TimeSpan.Zero), appointment.End);
        Assert.Equal(60, appointment.DurationMinutes);
        Assert.Equal(-15, appointment.ReminderDeltaMinutes);
    }

    [Fact]
    public void PreservesRoomAndResourceRecipientsThroughIcalendar() {
        DateTimeOffset start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero);
        var source = new EmailDocument {
            OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "Resource meeting",
            Appointment = new OutlookAppointment { Start = start, End = start.AddHours(1) }
        };
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Room,
            new EmailAddress("room@example.com", "Room 1")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Resource,
            new EmailAddress("projector@example.com", "Projector")));

        byte[] eml = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml);
        EmailDocument result = new EmailDocumentReader().Read(eml).Document;
        using var stream = new MemoryStream(eml);
        MimePart calendar = Assert.Single(MimeMessage.Load(stream).BodyParts.OfType<MimePart>(),
            part => part.ContentType.MimeType == "text/calendar");
        using var content = new MemoryStream();
        calendar.Content!.DecodeTo(content);
        string calendarText = Encoding.UTF8.GetString(content.ToArray());

        Assert.Contains("CUTYPE=ROOM", calendarText, StringComparison.Ordinal);
        Assert.Contains("CUTYPE=RESOURCE", calendarText, StringComparison.Ordinal);
        Assert.Contains(result.Recipients, recipient => recipient.Kind == EmailRecipientKind.Room &&
            recipient.Address.Address == "room@example.com");
        Assert.Contains(result.Recipients, recipient => recipient.Kind == EmailRecipientKind.Resource &&
            recipient.Address.Address == "projector@example.com");
    }

    [Fact]
    public void CalendarUserTypeOverridesTheEnvelopeRecipientKind() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "To: Room 1 <room@example.com>\r\nContent-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:room@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\n" +
            "ATTENDEE;CUTYPE=ROOM;ROLE=NON-PARTICIPANT:mailto:room@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        EmailRecipient recipient = Assert.Single(new EmailDocumentReader().Read(eml).Document.Recipients);

        Assert.Equal(EmailRecipientKind.Room, recipient.Kind);
    }

    [Theory]
    [InlineData("ACCEPTED", "IPM.Schedule.Meeting.Resp.Pos")]
    [InlineData("TENTATIVE", "IPM.Schedule.Meeting.Resp.Tent")]
    [InlineData("DECLINED", "IPM.Schedule.Meeting.Resp.Neg")]
    public void MapsIcalendarRepliesToOutlookResponseClasses(string participationStatus, string messageClass) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; method=REPLY; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:REPLY\r\nBEGIN:VEVENT\r\n" +
            "UID:reply@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ATTENDEE;PARTSTAT=" + participationStatus + ":mailto:person@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        Assert.Equal(messageClass, document.MessageClass);
        Assert.Equal("REPLY", IcalendarMethodAfterRoundTrip(document));
    }

    [Fact]
    public void MarksAttachedIcalendarPayloadsIncompleteForStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:attach@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "ATTACH:https://example.com/agenda.pdf\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void ReportsOversizedIcalendarDurationsWithoutThrowing() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:range@example.com\r\nDTSTART:00010101T000000Z\r\n" +
            "DTEND:99991231T000000Z\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

        EmailReadResult read = new EmailDocumentReader().Read(eml);
        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            read.Document, EmailFileFormat.OutlookMsg);

        Assert.Null(read.Document.Appointment!.DurationMinutes);
        Assert.Contains(read.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_DURATION_OUT_OF_RANGE");
        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void ReportsOversizedDurationAndReminderValuesWithoutThrowing() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:duration@example.com\r\nDTSTART:20260801T100000Z\r\nDURATION:P2000000D\r\n" +
            "BEGIN:VALARM\r\nACTION:DISPLAY\r\nTRIGGER:-P2000000D\r\nEND:VALARM\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");

        EmailReadResult read = new EmailDocumentReader().Read(eml);

        Assert.Equal(new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero).AddDays(2000000),
            read.Document.Appointment!.End);
        Assert.Null(read.Document.Appointment.DurationMinutes);
        Assert.Null(read.Document.Appointment.ReminderDeltaMinutes);
        Assert.Contains(read.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_DURATION_OUT_OF_RANGE");
    }

    [Fact]
    public void BlocksUnsupportedCalendarComponentsBeforeStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:event@example.com\r\nDTSTART:20260801T100000Z\r\nEND:VEVENT\r\n" +
            "BEGIN:VJOURNAL\r\nUID:journal@example.com\r\nSUMMARY:Journal\r\nEND:VJOURNAL\r\n" +
            "END:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void MarksSubMinuteDurationsAndTriggersIncomplete() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:seconds@example.com\r\nDTSTART:20260801T100000Z\r\nDURATION:PT90S\r\n" +
            "BEGIN:VALARM\r\nACTION:DISPLAY\r\nTRIGGER:-PT30S\r\nEND:VALARM\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailReadResult read = new EmailDocumentReader().Read(eml);

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            read.Document, EmailFileFormat.OutlookMsg);

        Assert.Equal(new DateTimeOffset(2026, 8, 1, 10, 1, 30, TimeSpan.Zero), read.Document.Appointment!.End);
        Assert.Equal(1, read.Document.Appointment.DurationMinutes);
        Assert.Equal(0, read.Document.Appointment.ReminderDeltaMinutes);
        Assert.Contains(read.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_DURATION_PRECISION_LOSS");
        Assert.False(report.CanWrite);
    }

    [Fact]
    public void MarksUnresolvedTimeZonesIncompleteForStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:timezone@example.com\r\n" +
            "DTSTART;TZID=OfficeIMO/Unknown:20260801T100000\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailReadResult read = new EmailDocumentReader().Read(eml);

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            read.Document, EmailFileFormat.OutlookMsg);

        Assert.Contains(read.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_TIMEZONE_UNRESOLVED");
        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void MarksEndRelativeAlarmsIncompleteForStoreConversion() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:end-alarm@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "DTEND:20260801T110000Z\r\nBEGIN:VALARM\r\nACTION:DISPLAY\r\n" +
            "TRIGGER;RELATED=END:-PT15M\r\nEND:VALARM\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Theory]
    [InlineData("TRANSPARENT", 0)]
    [InlineData("OPAQUE", 2)]
    public void MapsStandardTransparencyToOutlookBusyStatus(string transparency, int busyStatus) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:busy@example.com\r\nDTSTART:20260801T100000Z\r\nTRANSP:" + transparency +
            "\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        Assert.Equal(busyStatus, document.Appointment!.BusyStatus);
    }

    [Theory]
    [InlineData("PRIVATE", 2)]
    [InlineData("CONFIDENTIAL", 3)]
    public void PreservesEventPrivacyThroughStoreConversion(string calendarClass, int sensitivity) {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:private@example.com\r\nDTSTART:20260801T100000Z\r\nCLASS:" + calendarClass +
            "\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailDocument storeRoundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;
        byte[] regeneratedEml = new EmailDocumentWriter().ToBytes(storeRoundTrip, EmailFileFormat.Eml);
        using var stream = new MemoryStream(regeneratedEml);
        MimePart calendar = Assert.Single(MimeMessage.Load(stream).BodyParts.OfType<MimePart>(),
            part => part.ContentType.MimeType == "text/calendar");
        using var content = new MemoryStream();
        calendar.Content!.DecodeTo(content);

        Assert.Equal(sensitivity, document.MessageMetadata.Sensitivity);
        Assert.Equal(sensitivity, storeRoundTrip.MessageMetadata.Sensitivity);
        Assert.Contains("CLASS:" + calendarClass, Encoding.UTF8.GetString(content.ToArray()),
            StringComparison.Ordinal);
    }

    [Fact]
    public void ConvertsTaskThroughVtodo() {
        DateTimeOffset start = new DateTimeOffset(2026, 9, 3, 8, 0, 0, TimeSpan.Zero);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            MessageClass = "IPM.Task",
            Subject = "Ship release",
            Task = new OutlookTask {
                Start = start,
                Due = start.AddDays(2),
                PercentComplete = 0.5,
                Status = 1,
                Owner = "Release Manager",
                EstimatedEffort = TimeSpan.FromHours(6),
                ActualEffort = TimeSpan.FromHours(2),
                ReminderIsSet = true,
                ReminderDeltaMinutes = 30,
                BillingInformation = "Internal",
                Mileage = "12 km"
            }
        };
        source.Task.Contacts.Add("Ada Lovelace");
        source.Task.Companies.Add("Analytical Engines");

        EmailDocument result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml)).Document;

        Assert.Equal(OutlookItemKind.Task, result.OutlookItemKind);
        Assert.Equal("Ship release", result.Subject);
        Assert.Equal(start, result.Task!.Start);
        Assert.Equal(start.AddDays(2), result.Task.Due);
        Assert.Equal(0.5, result.Task.PercentComplete);
        Assert.Equal(1, result.Task.Status);
        Assert.Equal("Release Manager", result.Task.Owner);
        Assert.Equal(TimeSpan.FromHours(6), result.Task.EstimatedEffort);
        Assert.Equal(TimeSpan.FromHours(2), result.Task.ActualEffort);
        Assert.True(result.Task.ReminderIsSet);
        Assert.Equal(30, result.Task.ReminderDeltaMinutes);
        Assert.Equal("Internal", result.Task.BillingInformation);
        Assert.Equal("12 km", result.Task.Mileage);
        Assert.Equal("Ada Lovelace", Assert.Single(result.Task.Contacts));
        Assert.Equal("Analytical Engines", Assert.Single(result.Task.Companies));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(4)]
    public void PreservesNumericTaskStatusThroughVtodo(int status) {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            Subject = "Status",
            Task = new OutlookTask { Status = status }
        };

        EmailDocument result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml)).Document;

        Assert.Equal(status, result.Task!.Status);
    }

    [Fact]
    public void BlocksRecurringTaskWhenNoPortableRuleIsAvailable() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            Task = new OutlookTask { IsRecurring = true }
        };

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_OPAQUE_TASK_RECURRENCE");
    }

    [Fact]
    public void BlocksAddresslessTaskAssigneesBeforeEmlConversion() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            Task = new OutlookTask(),
            Subject = "Assigned task"
        };
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress(null, "Assignee without an address")));

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_ATTENDEE_ADDRESS_REQUIRED");
    }

    [Fact]
    public void BlocksOpaqueOutlookRecurrenceInsteadOfDroppingIt() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero),
                End = new DateTimeOffset(2026, 8, 1, 11, 0, 0, TimeSpan.Zero),
                RecurrenceState = new byte[] { 1, 2, 3 }
            }
        };

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_OPAQUE_RECURRENCE");
    }

    [Fact]
    public void BlocksAttendeeDisplayTextWhenNoCalendarAddressExists() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero),
                RequiredAttendees = "Person without an address"
            }
        };

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_ATTENDEE_ADDRESS_REQUIRED");
    }

    [Fact]
    public void BlocksAddresslessRecipientRowsAndNeverEmitsEmptyMailtoValues() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero)
            }
        };
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("valid@example.com", "Valid")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress(null, "No Address")));

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(source, EmailFileFormat.Eml);
        byte[] warned = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn))
            .ToBytes(source, EmailFileFormat.Eml);
        EmailAttachment calendar = Assert.Single(new EmailDocumentReader().Read(warned).Document.Attachments,
            attachment => string.Equals(attachment.ContentType, "text/calendar", StringComparison.OrdinalIgnoreCase));
        string text = Encoding.UTF8.GetString(Assert.IsType<byte[]>(calendar.Content));

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_ATTENDEE_ADDRESS_REQUIRED");
        Assert.Contains("mailto:valid@example.com", text, StringComparison.Ordinal);
        Assert.DoesNotContain("CN=\"No Address\"", text, StringComparison.Ordinal);
        Assert.DoesNotContain(":mailto:\r\n", text, StringComparison.Ordinal);
    }

    [Fact]
    public void KeepsRecurringIcalendarInEmlButBlocksIncompleteStoreProjection() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Subject: Recurring\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:r@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nDTEND:20260801T110000Z\r\n" +
            "RRULE:FREQ=WEEKLY;COUNT=4\r\nSUMMARY:Recurring\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;
        var writer = new EmailDocumentWriter();

        EmailConversionReport emlReport = writer.AnalyzeConversion(document, EmailFileFormat.Eml);
        EmailConversionReport msgReport = writer.AnalyzeConversion(document, EmailFileFormat.OutlookMsg);

        Assert.True(emlReport.CanWrite);
        Assert.NotEmpty(writer.ToBytes(document, EmailFileFormat.Eml));
        Assert.False(msgReport.CanWrite);
        Assert.Contains(msgReport.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");
    }

    [Fact]
    public void KeepsUnchangedCalendarWithoutDtStartInEml() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Subject: Undated\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:undated@example.com\r\n" +
            "SUMMARY:Undated\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;
        var writer = new EmailDocumentWriter();

        EmailConversionReport report = writer.AnalyzeConversion(document, EmailFileFormat.Eml);
        EmailAttachment calendar = Assert.Single(new EmailDocumentReader().Read(
            writer.ToBytes(document, EmailFileFormat.Eml)).Document.Attachments,
            attachment => string.Equals(attachment.ContentType, "text/calendar", StringComparison.OrdinalIgnoreCase));
        string rewritten = Encoding.ASCII.GetString(Assert.IsType<byte[]>(calendar.Content));

        Assert.True(report.CanWrite);
        Assert.Contains("UID:undated@example.com", rewritten, StringComparison.Ordinal);
        Assert.DoesNotContain("DTSTART", rewritten, StringComparison.Ordinal);
    }

    [Fact]
    public void RegeneratesEditedCalendarContentWhenLossIsAccepted() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Subject: Original\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: text/calendar; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:edit@example.com\r\n" +
            "DTSTART:20260801T100000Z\r\nSUMMARY:Original\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;
        document.Subject = "Edited";
        document.Appointment!.Start = new DateTimeOffset(2026, 8, 1, 12, 0, 0, TimeSpan.Zero);

        byte[] output = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn))
            .ToBytes(document, EmailFileFormat.Eml);
        EmailAttachment calendar = Assert.Single(new EmailDocumentReader().Read(output).Document.Attachments,
            attachment => string.Equals(attachment.ContentType, "text/calendar", StringComparison.OrdinalIgnoreCase));
        string rewritten = Encoding.UTF8.GetString(Assert.IsType<byte[]>(calendar.Content));

        Assert.Contains("DTSTART:20260801T120000Z", rewritten, StringComparison.Ordinal);
        Assert.Contains("SUMMARY:Edited", rewritten, StringComparison.Ordinal);
        Assert.DoesNotContain("DTSTART:20260801T100000Z", rewritten, StringComparison.Ordinal);
    }

    [Fact]
    public void RecomputesCalendarMethodWhenRegeneratingEditedContent() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Subject: Published\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: text/calendar; method=PUBLISH; charset=utf-8\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:PUBLISH\r\nBEGIN:VEVENT\r\n" +
            "UID:method@example.com\r\nDTSTART:20260801T100000Z\r\nSUMMARY:Published\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;
        document.Subject = "Requested";
        document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("attendee@example.com")));

        byte[] output = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn))
            .ToBytes(document, EmailFileFormat.Eml);
        EmailAttachment calendar = Assert.Single(new EmailDocumentReader().Read(output).Document.Attachments,
            attachment => string.Equals(attachment.ContentType, "text/calendar", StringComparison.OrdinalIgnoreCase));
        string rewritten = Encoding.UTF8.GetString(Assert.IsType<byte[]>(calendar.Content));

        Assert.Equal("REQUEST", calendar.ContentTypeParameters["method"]);
        Assert.Contains("METHOD:REQUEST", rewritten, StringComparison.Ordinal);
    }

    [Fact]
    public void KeepsAttachedCalendarAsAnAttachment() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Subject: Attached invitation\r\nMIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nPlease review the attachment.\r\n" +
            "--x\r\nContent-Type: text/calendar; method=PUBLISH; charset=utf-8\r\n" +
            "Content-Disposition: attachment; filename=invite.ics\r\n\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:PUBLISH\r\nBEGIN:VEVENT\r\n" +
            "UID:attached@example.com\r\nDTSTART:20260801T100000Z\r\nSUMMARY:Attached\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n--x--\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        byte[] output = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml);
        EmailDocument storeRoundTrip = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(document, EmailFileFormat.OutlookMsg)).Document;
        using var stream = new MemoryStream(output);
        MimeMessage message = MimeMessage.Load(stream);
        MimePart attachment = Assert.IsAssignableFrom<MimePart>(Assert.Single(message.Attachments));

        Assert.Equal(OutlookItemKind.Message, document.OutlookItemKind);
        Assert.Null(document.Appointment);
        Assert.Equal("Please review the attachment.", message.TextBody!.Trim());
        Assert.Equal("text/calendar", attachment.ContentType.MimeType);
        Assert.Equal("invite.ics", attachment.FileName);
        Assert.True(attachment.IsAttachment);
        Assert.Equal(OutlookItemKind.Message, storeRoundTrip.OutlookItemKind);
        Assert.Null(storeRoundTrip.Appointment);
        Assert.Equal("invite.ics", Assert.Single(storeRoundTrip.Attachments).FileName);
    }

    [Fact]
    public void KeepsOrdinaryCalendarAttachmentSeparateFromGeneratedAppointmentContent() {
        DateTimeOffset start = new DateTimeOffset(2026, 8, 1, 10, 0, 0, TimeSpan.Zero);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "Generated appointment",
            Appointment = new OutlookAppointment { Start = start, End = start.AddHours(1) }
        };
        source.Attachments.Add(new EmailAttachment {
            FileName = "ordinary.ics",
            ContentType = "text/calendar",
            Content = Encoding.ASCII.GetBytes(
                "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:ordinary@example.com\r\n" +
                "DTSTART:20260901T100000Z\r\nSUMMARY:Ordinary file\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n")
        });
        source.Attachments[0].Length = source.Attachments[0].Content!.LongLength;

        byte[] output = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml);
        using var stream = new MemoryStream(output);
        MimePart[] calendars = MimeMessage.Load(stream).BodyParts.OfType<MimePart>()
            .Where(part => part.ContentType.MimeType == "text/calendar").ToArray();
        MimePart attached = Assert.Single(calendars, part => part.IsAttachment);
        MimePart generated = Assert.Single(calendars, part => !part.IsAttachment);
        using var attachedContent = new MemoryStream();
        attached.Content!.DecodeTo(attachedContent);
        using var generatedContent = new MemoryStream();
        generated.Content!.DecodeTo(generatedContent);

        Assert.Equal("ordinary.ics", attached.FileName);
        Assert.Contains("SUMMARY:Ordinary file", Encoding.ASCII.GetString(attachedContent.ToArray()),
            StringComparison.Ordinal);
        Assert.Contains("SUMMARY:Generated appointment", Encoding.UTF8.GetString(generatedContent.ToArray()),
            StringComparison.Ordinal);
    }

    [Fact]
    public void AccumulatesIncompleteStateAcrossMultipleSemanticParts() {
        byte[] eml = Encoding.ASCII.GetBytes(
            "Subject: Multiple\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n" +
            "--x\r\nContent-Type: text/vcard\r\n\r\nBEGIN:VCARD\r\nVERSION:4.0\r\n" +
            "FN:Ada\r\nPHOTO:https://example.com/ada.jpg\r\nEND:VCARD\r\n" +
            "--x\r\nContent-Type: text/calendar\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
            "BEGIN:VEVENT\r\nUID:event@example.com\r\nDTSTART:20260801T100000Z\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n--x--\r\n");
        EmailDocument document = new EmailDocumentReader().Read(eml).Document;

        EmailConversionReport report = new EmailDocumentWriter().AnalyzeConversion(
            document, EmailFileFormat.OutlookMsg);

        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE");

        byte[] rewritten = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml);
        using var stream = new MemoryStream(rewritten);
        MimePart calendar = Assert.Single(MimeMessage.Load(stream).BodyParts.OfType<MimePart>(),
            part => part.ContentType.MimeType == "text/calendar");
        using var content = new MemoryStream();
        calendar.Content!.DecodeTo(content);
        string calendarText = Encoding.UTF8.GetString(content.ToArray());
        Assert.Contains("BEGIN:VCALENDAR", calendarText, StringComparison.Ordinal);
        Assert.DoesNotContain("BEGIN:VCARD", calendarText, StringComparison.Ordinal);
    }

    private static string? IcalendarMethodAfterRoundTrip(EmailDocument document) {
        byte[] eml = new EmailDocumentWriter().ToBytes(document, EmailFileFormat.Eml);
        using var stream = new MemoryStream(eml);
        MimePart calendar = Assert.Single(MimeMessage.Load(stream).BodyParts.OfType<MimePart>(),
            part => part.ContentType.MimeType == "text/calendar");
        return calendar.ContentType.Parameters["method"];
    }
}
