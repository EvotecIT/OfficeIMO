using System.Globalization;

namespace OfficeIMO.Email.Tests;

public sealed partial class NativeAppleExportTests {
    [Theory]
    [InlineData("native-calendar.ics")]
    [InlineData("native-olm-calendar.ics")]
    public void NativeCalendarRulesProjectTypedOccurrencesThroughMsgAndEml(string fixture) {
        VerifyFixtures("calendar/");
        using EmailReadResult source = ReadCalendar(File.ReadAllText(Path.Combine(FixtureRoot, "calendar", fixture)));
        AssertCalendarOccurrences(source.Document);
        Assert.DoesNotContain(source.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_ICALENDAR_TIMEZONE_PROJECTION_UNSUPPORTED");
        // Vendor fields are retained in the MIME source but not all fit Outlook;
        // explicitly choose Warn for this typed conversion.
        var writer = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn));
        using EmailReadResult msg = new EmailDocumentReader().Read(writer.ToBytes(source.Document, EmailFileFormat.OutlookMsg));
        AssertCalendarOccurrences(msg.Document);
        using EmailReadResult eml = new EmailDocumentReader().Read(writer.ToBytes(msg.Document, EmailFileFormat.Eml));
        AssertCalendarOccurrences(eml.Document);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EmbeddedRulesResolveEventInstantsWithoutAHostTimeZone(bool recurring) {
        string text = File.ReadAllText(Path.Combine(FixtureRoot, "calendar", "native-olm-calendar.ics"))
            .Replace("Europe/Warsaw", "OfficeIMO/SyntheticZone");
        if (!recurring) text = text.Replace("RRULE:FREQ=WEEKLY;COUNT=4;BYDAY=SU;WKST=SU\r\n", "");
        using EmailReadResult read = ReadCalendar(text);
        Assert.Equal(new DateTimeOffset(2026, 10, 18, 7, 0, 0, TimeSpan.Zero), read.Document.Appointment!.Start!.Value.ToUniversalTime());
        Assert.Equal(new DateTimeOffset(2026, 10, 18, 8, 0, 0, TimeSpan.Zero), read.Document.Appointment.End!.Value.ToUniversalTime());
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_TIMEZONE_UNRESOLVED");
        if (recurring) AssertCalendarOccurrences(read.Document);
        else Assert.Null(read.Document.Appointment.Recurrence);
    }

    [Theory]
    [InlineData("start")]
    [InlineData("exception")]
    [InlineData("expiry")]
    public void PartialObservanceHistoryRemainsUnprojectedAndBlocksLossyStoreConversion(string variant) {
        string text = File.ReadAllText(Path.Combine(FixtureRoot, "calendar", "native-olm-calendar.ics"));
        if (variant == "start") text = text.Replace("20261018T090000", "19961020T090000")
            .Replace("20261018T100000", "19961020T100000");
        else if (variant == "exception") text = text.Replace("SEQUENCE:0", "EXDATE:19961020T070000Z\r\nSEQUENCE:0");
        else text = text.Replace("BYDAY=-1SU", "BYDAY=-1SU;UNTIL=20271031T010000Z");
        using EmailReadResult read = ReadCalendar(text);
        Assert.Null(read.Document.Appointment!.RecurrenceTimeZone);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_TIMEZONE_PROJECTION_UNSUPPORTED");
        using var output = new MemoryStream();
        EmailWriteResult conversion = new EmailDocumentWriter().Write(read.Document, output, EmailFileFormat.OutlookMsg);
        Assert.Contains(conversion.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Error);
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false, "north", "20260329T023000", "2026-03-29T01:30:00Z")]
    [InlineData(true, "north", "20260329T023000", "2026-03-29T01:30:00Z")]
    [InlineData(false, "north", "20261025T023000", "2026-10-25T00:30:00Z")]
    [InlineData(true, "north", "20261025T023000", "2026-10-25T00:30:00Z")]
    [InlineData(false, "south", "20261004T023000", "2026-10-03T16:30:00Z")]
    [InlineData(true, "south", "20261004T023000", "2026-10-03T16:30:00Z")]
    [InlineData(false, "south", "20260405T023000", "2026-04-04T15:30:00Z")]
    [InlineData(true, "south", "20260405T023000", "2026-04-04T15:30:00Z")]
    [InlineData(false, "negative", "20261025T033000", "2026-10-25T02:30:00Z")]
    [InlineData(true, "negative", "20261025T033000", "2026-10-25T02:30:00Z")]
    public void ExplicitZonedEventAndTaskDatesUseIcalendarGapAndFoldRules(
        bool task, string zoneKind, string clock, string expectedUtc) {
        string zone = File.ReadAllText(Path.Combine(FixtureRoot, "calendar", "native-olm-calendar.ics"));
        zone = zone.Substring(zone.IndexOf("BEGIN:VTIMEZONE", StringComparison.Ordinal));
        zone = zone.Substring(0, zone.IndexOf("END:VTIMEZONE", StringComparison.Ordinal) + "END:VTIMEZONE".Length)
            .Replace("Europe/Warsaw", "OfficeIMO/SyntheticZone");
        if (zoneKind == "south") zone = zone.Replace("19961027T030000", "19960407T030000")
            .Replace("19880327T020000", "19881002T020000")
            .Replace("BYMONTH=10", "BYMONTH=4").Replace("BYMONTH=3", "BYMONTH=10")
            .Replace("BYDAY=-1SU", "BYDAY=1SU")
            .Replace("+0100", "+1000").Replace("+0200", "+1100");
        if (zoneKind == "negative") zone = zone.Replace("+0100", "OFFSET_SWAP")
            .Replace("+0200", "+0100").Replace("OFFSET_SWAP", "+0200");
        DateTime local = DateTime.ParseExact(clock, "yyyyMMdd'T'HHmmss", CultureInfo.InvariantCulture);
        DateTimeOffset expected = DateTimeOffset.Parse(expectedUtc, CultureInfo.InvariantCulture);
        foreach (bool atStart in new[] { false, true }) {
            string start = atStart ? clock : local.AddDays(-1).ToString("yyyyMMdd'T'HHmmss", CultureInfo.InvariantCulture);
            string end = atStart ? local.AddDays(1).ToString("yyyyMMdd'T'HHmmss", CultureInfo.InvariantCulture) : clock;
            string component = task ? "VTODO" : "VEVENT", endProperty = task ? "DUE" : "DTEND";
            string text = $"BEGIN:VCALENDAR\r\nVERSION:2.0\r\n{zone}\r\nBEGIN:{component}\r\nUID:explicit-transition\r\n" +
                $"DTSTART;TZID=OfficeIMO/SyntheticZone:{start}\r\n{endProperty};TZID=OfficeIMO/SyntheticZone:{end}\r\nEND:{component}\r\nEND:VCALENDAR\r\n";
            using EmailReadResult read = ReadCalendar(text);
            Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_DATE_INVALID" ||
                diagnostic.Code == "EMAIL_ICALENDAR_TIMEZONE_UNRESOLVED");
            Assert.Equal(expected.UtcDateTime, ProjectedDate(read.Document).UtcDateTime);
            var writer = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn));
            using EmailReadResult msg = new EmailDocumentReader().Read(writer.ToBytes(read.Document, EmailFileFormat.OutlookMsg));
            Assert.Equal(expected.UtcDateTime, ProjectedDate(msg.Document).UtcDateTime);
            using EmailReadResult eml = new EmailDocumentReader().Read(writer.ToBytes(msg.Document, EmailFileFormat.Eml));
            Assert.Equal(expected.UtcDateTime, ProjectedDate(eml.Document).UtcDateTime);

            DateTimeOffset ProjectedDate(EmailDocument document) => (task
                ? atStart ? document.Task!.Start : document.Task!.Due
                : atStart ? document.Appointment!.Start : document.Appointment!.End)!.Value;
        }
    }

    private static EmailReadResult ReadCalendar(string text) => new EmailDocumentReader().Read(
        Encoding.UTF8.GetBytes("Content-Type: text/calendar; charset=utf-8\r\n\r\n" + text));

    private static void AssertCalendarOccurrences(EmailDocument document) {
        OutlookAppointment appointment = Assert.IsType<OutlookAppointment>(document.Appointment);
        OutlookRecurrence recurrence = Assert.IsType<OutlookRecurrence>(appointment.Recurrence);
        Assert.Equal(new DateTime(2026, 10, 18, 9, 0, 0), recurrence.Start);
        Assert.Equal(TimeSpan.FromHours(1), recurrence.Duration);
        Assert.Equal(10, appointment.ReminderDeltaMinutes);
        Assert.True(appointment.ReminderIsSet);
        OutlookTimeZoneDefinition zone = Assert.IsType<OutlookTimeZoneDefinition>(appointment.RecurrenceTimeZone);
        OutlookRecurrenceExpansionResult expanded = OutlookRecurrenceExpander.Expand(recurrence);
        Assert.False(expanded.Truncated);
        Assert.Equal(new[] { 18, 25, 1, 8 }, expanded.Occurrences.Select(value => value.Start.Day));
        Assert.All(expanded.Occurrences, value => Assert.Equal(9, value.Start.Hour));
        Assert.Equal(new[] { 7, 8, 8, 8 }, expanded.Occurrences.Select(value => value.ResolveStart(zone).UtcDateTime.Hour));
    }
}
