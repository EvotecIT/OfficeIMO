using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class IcsDocumentTests {
    [Fact]
    public void Parse_MultipleCalendarsPreservesUnknownRecurrenceAndNestedAlarm() {
        const string source = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//One//EN\r\n" +
            "X-WR-CALNAME:First\r\nBEGIN:VEVENT\r\nUID:event-1\r\nDTSTART;TZID=Europe/Warsaw:20260717T090000\r\n" +
            "RRULE:FREQ=WEEKLY;COUNT=4\r\nRDATE;TZID=Europe/Warsaw:20260821T090000\r\nEXDATE;TZID=Europe/Warsaw:20260731T090000\r\n" +
            "X-VENDOR-STATE:opaque\r\nBEGIN:VALARM\r\nACTION:EMAIL\r\nTRIGGER:-PT30M\r\nEND:VALARM\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Two//EN\r\nBEGIN:VFREEBUSY\r\nUID:freebusy-1\r\nEND:VFREEBUSY\r\nEND:VCALENDAR\r\n";

        IcsDocument document = IcsDocument.Parse(source);
        ContentLineComponent appointment = document.GetComponents("VEVENT").Single();
        appointment.SetProperty("SUMMARY", "Updated meeting");

        IcsDocument reparsed = IcsDocument.Parse(document.Serialize());
        ContentLineComponent reparsedAppointment = reparsed.GetComponents("VEVENT").Single();

        Assert.Equal(2, reparsed.Calendars.Count);
        Assert.Equal("Updated meeting", reparsedAppointment.GetFirstProperty("SUMMARY")!.Value);
        Assert.Equal("FREQ=WEEKLY;COUNT=4", reparsedAppointment.GetFirstProperty("RRULE")!.Value);
        Assert.Equal("opaque", reparsedAppointment.GetFirstProperty("X-VENDOR-STATE")!.Value);
        Assert.Single(reparsedAppointment.GetComponents("VALARM"));
        Assert.Single(reparsed.GetComponents("VFREEBUSY"));
    }

    [Fact]
    public async Task SaveAndLoadAsync_UseStandaloneFileApi() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".ics");
        try {
            var document = new IcsDocument();
            ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
            appointment.AddProperty("UID", "standalone@example.com");
            appointment.AddProperty("DTSTART", "20260717").SetParameter("VALUE", "DATE");

            await document.SaveAsync(path);
            IcsDocument loaded = await IcsDocument.LoadAsync(path);

            Assert.Equal("20260717", loaded.GetComponents("VEVENT").Single()
                .GetFirstProperty("DTSTART")!.Value);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Parse_RejectsNonCalendarRoot() {
        Assert.Throws<InvalidDataException>(() => IcsDocument.Parse(
            "BEGIN:VCARD\r\nVERSION:4.0\r\nEND:VCARD\r\n"));
    }

    [Fact]
    public void TemporalHelpersPreserveDateFloatingUtcAndTzidForms() {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        ContentLineComponent timeZone = calendar.AddComponent("VTIMEZONE");
        timeZone.AddProperty("TZID", "Europe/Warsaw");
        ContentLineComponent appointment = calendar.AddComponent("VEVENT");
        appointment.AddProperty("UID", "temporal@example.com");
        appointment.AddProperty("DTSTAMP", "20260717T090000Z");
        appointment.SetTemporalValue("DTSTART", IcsTemporalValue.Zoned(
            new DateTime(2026, 7, 17, 11, 30, 0), "Europe/Warsaw"));
        appointment.SetTemporalValue("DTEND", IcsTemporalValue.Utc(
            new DateTimeOffset(2026, 7, 17, 12, 30, 0, TimeSpan.Zero)));
        ContentLineComponent task = calendar.AddComponent("VTODO");
        task.SetTemporalValue("DUE", IcsTemporalValue.Date(new DateTime(2026, 7, 20)));
        task.SetTemporalValue("DTSTART", IcsTemporalValue.Floating(new DateTime(2026, 7, 18, 8, 0, 0)));

        IcsDocument reparsed = IcsDocument.Parse(document.Serialize());
        ContentLineComponent reparsedAppointment = reparsed.GetComponents("VEVENT").Single();
        ContentLineComponent reparsedTask = reparsed.GetComponents("VTODO").Single();

        Assert.Equal(IcsTemporalValueKind.ZonedDateTime, reparsedAppointment.GetTemporalValue("DTSTART")!.Value.Kind);
        Assert.Equal("Europe/Warsaw", reparsedAppointment.GetTemporalValue("DTSTART")!.Value.TimeZoneId);
        Assert.Equal(IcsTemporalValueKind.UtcDateTime, reparsedAppointment.GetTemporalValue("DTEND")!.Value.Kind);
        Assert.Equal(IcsTemporalValueKind.Date, reparsedTask.GetTemporalValue("DUE")!.Value.Kind);
        Assert.Equal(IcsTemporalValueKind.FloatingDateTime, reparsedTask.GetTemporalValue("DTSTART")!.Value.Kind);
        Assert.DoesNotContain(reparsed.Validate(), issue => issue.Code == "ICAL_TIMEZONE_DEFINITION_MISSING");
    }

    [Fact]
    public void RecurrenceModelRetainsUnknownPartsAndValidationFindsConflicts() {
        IcsRecurrenceRule rule = IcsRecurrenceRule.Parse("FREQ=WEEKLY;COUNT=4;X-WORKDAY=YES;BYDAY=MO,WE");
        rule.Interval = 2;
        rule.SetValue("UNTIL", "20261231T235959Z");

        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("UID", "recurrence@example.com");
        appointment.AddProperty("DTSTAMP", "20260717T090000Z");
        appointment.AddRecurrenceRule(rule);

        IcsRecurrenceRule reparsed = IcsDocument.Parse(document.Serialize()).GetComponents("VEVENT")
            .Single().GetRecurrenceRules().Single();
        ContentLineValidationIssue issue = Assert.Single(document.Validate(), finding =>
            finding.Code == "ICAL_RRULE_COUNT_UNTIL_CONFLICT");

        Assert.Equal("YES", reparsed.GetValue("X-WORKDAY"));
        Assert.Equal(2, reparsed.Interval);
        Assert.Equal(ContentLineValidationSeverity.Error, issue.Severity);
    }

    [Fact]
    public void TimeZoneDefinitionsAreScopedToTheirContainingCalendar() {
        const string source = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//One//EN\r\n" +
            "BEGIN:VTIMEZONE\r\nTZID:Example/Zone\r\nEND:VTIMEZONE\r\nEND:VCALENDAR\r\n" +
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Two//EN\r\nBEGIN:VEVENT\r\n" +
            "UID:two\r\nDTSTAMP:20260717T090000Z\r\nDTSTART;TZID=Example/Zone:20260717T100000\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n";

        ContentLineValidationIssue issue = Assert.Single(IcsDocument.Parse(source).Validate(),
            finding => finding.Code == "ICAL_TIMEZONE_DEFINITION_MISSING");

        Assert.Equal("DTSTART", issue.PropertyName);
    }

    [Fact]
    public void TemporalParserRejectsConflictingValueAndTimeZoneParameters() {
        var date = new ContentLineProperty("DTSTART", "20260717");
        date.SetParameter("VALUE", "DATE").SetParameter("TZID", "Europe/Warsaw");
        var utc = new ContentLineProperty("DTSTART", "20260717T090000Z");
        utc.SetParameter("TZID", "Europe/Warsaw");
        var unsupported = new ContentLineProperty("DTSTART", "20260717T090000");
        unsupported.SetParameter("VALUE", "PERIOD");

        Assert.False(IcsTemporalValue.TryParse(date, out _));
        Assert.False(IcsTemporalValue.TryParse(utc, out _));
        Assert.False(IcsTemporalValue.TryParse(unsupported, out _));
    }

    [Fact]
    public void ValidationReportsInvalidRecurrencePartNamesInsteadOfThrowing() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("RRULE", "FREQ=DAILY;BAD_NAME=value");

        ContentLineValidationIssue issue = Assert.Single(document.Validate(), finding =>
            finding.Code == "ICAL_RRULE_INVALID");

        Assert.Equal(ContentLineValidationSeverity.Error, issue.Severity);
    }

    [Fact]
    public void ValidationAndSerializationRejectMissingOrMutatedCalendarRoots() {
        var empty = new IcsDocument();
        empty.Calendars.Clear();
        var mutated = new IcsDocument();
        mutated.Calendars.Single().Name = "VCARD";

        Assert.Contains(empty.Validate(), issue => issue.Code == "ICAL_ROOT_REQUIRED");
        Assert.Contains(mutated.Validate(), issue => issue.Code == "ICAL_ROOT_INVALID");
        Assert.Throws<InvalidDataException>(() => empty.ToBytes());
        Assert.Throws<InvalidDataException>(() => mutated.ToBytes());
    }
}
