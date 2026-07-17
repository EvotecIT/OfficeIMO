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
    public void ValidationChecksEveryRepeatedTimeZoneParameter() {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        ContentLineComponent timeZone = calendar.AddComponent("VTIMEZONE");
        timeZone.AddProperty("TZID", "Europe/Warsaw");
        ContentLineComponent appointment = calendar.AddComponent("VEVENT");
        appointment.AddProperty("UID", "timezones@example.test");
        appointment.AddProperty("DTSTAMP", "20260717T090000Z");
        ContentLineProperty start = appointment.AddProperty("DTSTART", "20260717T090000");
        start.Parameters.Add(new ContentLineParameter("TZID", "Europe/Warsaw"));
        start.Parameters.Add(new ContentLineParameter("TZID", "Missing/Zone"));

        ContentLineValidationIssue issue = Assert.Single(document.Validate(), finding =>
            finding.Code == "ICAL_TIMEZONE_DEFINITION_MISSING");

        Assert.Contains("Missing/Zone", issue.Message, StringComparison.Ordinal);
        Assert.Contains(document.Validate(), finding =>
            finding.Code == "ICAL_PARAMETER_CARDINALITY" && finding.PropertyName == "DTSTART");
    }

    [Theory]
    [InlineData("RDATE", "not-a-date", null)]
    [InlineData("RDATE", "20260717T090000Z,not-a-date", null)]
    [InlineData("EXDATE", "20260717T25AA00", null)]
    [InlineData("EXDATE", "20260717T090000Z/PT1H", "PERIOD")]
    [InlineData("RDATE", "20260717T090000Z/-PT1H", "PERIOD")]
    [InlineData("RDATE", "20260717T090000Z/P1Y", "PERIOD")]
    [InlineData("RDATE", "20260717T100000Z/20260717T090000Z", "PERIOD")]
    public void ValidationRejectsInvalidRecurrenceDateLists(
        string propertyName, string value, string? valueType) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        ContentLineProperty property = appointment.AddProperty(propertyName, value);
        if (valueType != null) property.SetParameter("VALUE", valueType);

        Assert.Contains(document.Validate(), finding =>
            finding.Code == "ICAL_TEMPORAL_VALUE_INVALID" &&
            finding.PropertyName == propertyName);
    }

    [Fact]
    public void ValidationAcceptsDateTimeDateAndPeriodRecurrenceLists() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("RDATE", "20260717T090000Z,20260718T090000Z");
        appointment.AddProperty("EXDATE", "20260719,20260720").SetParameter("VALUE", "DATE");
        appointment.AddProperty("RDATE", "20260721T090000Z/PT1H," +
                "20260722T090000Z/20260722T100000Z," +
                "20260723T090000Z/+PT30M," +
                "20260724T090000Z/P1D," +
                "20260725T090000Z/P1W," +
                "20260726T090000Z/PT1H30M30S")
            .SetParameter("VALUE", "PERIOD");

        Assert.DoesNotContain(document.Validate(), finding =>
            finding.Code == "ICAL_TEMPORAL_VALUE_INVALID");
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
    public void TemporalParserAndValidationRejectMalformedSingletonParameters() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        ContentLineProperty repeatedValue = appointment.AddProperty("DTSTART", "20260717T090000");
        repeatedValue.Parameters.Add(new ContentLineParameter("VALUE", "DATE-TIME"));
        repeatedValue.Parameters.Add(new ContentLineParameter("VALUE", "DATE-TIME"));
        ContentLineProperty emptyTimeZone = appointment.AddProperty("DTEND", "20260717T100000");
        emptyTimeZone.Parameters.Add(new ContentLineParameter("TZID", string.Empty));

        Assert.False(IcsTemporalValue.TryParse(repeatedValue, out _));
        Assert.False(IcsTemporalValue.TryParse(emptyTimeZone, out _));
        ContentLineValidationIssue[] issues = document.Validate().ToArray();
        Assert.Contains(issues, issue => issue.Code == "ICAL_TEMPORAL_VALUE_INVALID" &&
            issue.PropertyName == "DTSTART");
        Assert.Contains(issues, issue => issue.Code == "ICAL_TEMPORAL_VALUE_INVALID" &&
            issue.PropertyName == "DTEND");
        Assert.Contains(issues, issue => issue.Code == "ICAL_PARAMETER_CARDINALITY" &&
            issue.PropertyName == "DTEND");
    }

    [Theory]
    [InlineData("20260717T0930Z")]
    [InlineData("20260717T0930")]
    public void TemporalParserAndValidationRejectDateTimesWithoutSeconds(string text) {
        var property = new ContentLineProperty("DTSTART", text);
        Assert.False(IcsTemporalValue.TryParse(property, out _));

        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty(property.Name, property.Value);

        Assert.Single(document.Validate(), issue => issue.Code == "ICAL_TEMPORAL_VALUE_INVALID");
    }

    [Theory]
    [InlineData("19970630T235960Z", null, IcsTemporalValueKind.UtcDateTime)]
    [InlineData("19970630T235960", null, IcsTemporalValueKind.FloatingDateTime)]
    [InlineData("19970630T235960", "Europe/Warsaw", IcsTemporalValueKind.ZonedDateTime)]
    public void TemporalParserAcceptsLeapSecondsAndPreservesTheirLexicalForm(
        string text, string? timeZoneId, IcsTemporalValueKind expectedKind) {
        var property = new ContentLineProperty("DTSTART", text);
        if (timeZoneId != null) property.SetParameter("TZID", timeZoneId);

        Assert.True(IcsTemporalValue.TryParse(property, out IcsTemporalValue temporal));
        Assert.Equal(expectedKind, temporal.Kind);
        Assert.True(temporal.IsLeapSecond);
        var applied = new ContentLineProperty("DTSTART", string.Empty);
        temporal.ApplyTo(applied);
        Assert.Equal(text, applied.Value);
        Assert.Equal(timeZoneId, applied.GetParameter("TZID")?.Values.SingleOrDefault());
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.Properties.Add(property);

        Assert.DoesNotContain(document.Validate(), issue => issue.Code == "ICAL_TEMPORAL_VALUE_INVALID");
        Assert.Contains("DTSTART", document.Serialize(), StringComparison.Ordinal);
        Assert.Contains(text, document.Serialize(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(" 20260717T090000Z")]
    [InlineData("20260717T090000Z ")]
    [InlineData("20260717T090061Z")]
    public void TemporalParserRejectsWhitespaceAndSecondsBeyondLeapSecond(string text) {
        Assert.False(IcsTemporalValue.TryParse(new ContentLineProperty("DTSTART", text), out _));
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
    public void ValidationReportsEachDuplicatedRecurrencePart() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("RRULE", "FREQ=DAILY;FREQ=WEEKLY;COUNT=2;COUNT=3");

        ContentLineValidationIssue[] issues = document.Validate().Where(finding =>
            finding.Code == "ICAL_RRULE_PART_DUPLICATE").ToArray();

        Assert.Equal(2, issues.Length);
        Assert.Contains(issues, issue => issue.Message.Contains("FREQ", StringComparison.Ordinal));
        Assert.Contains(issues, issue => issue.Message.Contains("COUNT", StringComparison.Ordinal));
        Assert.Contains("FREQ=DAILY;FREQ=WEEKLY;COUNT=2;COUNT=3", document.Serialize(),
            StringComparison.Ordinal);
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

    [Fact]
    public void LegacyVcalendarParametersPreserveLiteralCaretSequences() {
        const string source = "BEGIN:VCALENDAR\r\nVERSION:1.0\r\nPRODID:-//Legacy//EN\r\n" +
            "BEGIN:VEVENT\r\nATTENDEE;X-LITERAL=alpha^nbeta:mailto:a@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n";

        IcsDocument document = IcsDocument.Parse(source);
        string value = document.GetComponents("VEVENT").Single().GetFirstProperty("ATTENDEE")!
            .GetParameter("X-LITERAL")!.Values.Single();
        string serialized = document.Serialize();

        Assert.Equal("alpha^nbeta", value);
        Assert.Contains("X-LITERAL=alpha^nbeta", serialized, StringComparison.Ordinal);
        Assert.DoesNotContain("X-LITERAL=alpha^^nbeta", serialized, StringComparison.Ordinal);
        Assert.Equal(value, IcsDocument.Parse(serialized).GetComponents("VEVENT").Single()
            .GetFirstProperty("ATTENDEE")!.GetParameter("X-LITERAL")!.Values.Single());
    }

    [Fact]
    public void MutableComponentCyclesAreReportedAndRejectedWithoutRecursiveOverflow() {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        calendar.Components.Add(calendar);

        Assert.Contains(document.Validate(), issue => issue.Code == "ICAL_COMPONENT_GRAPH_CYCLE");
        Assert.Throws<InvalidDataException>(() => document.GetComponents("VEVENT").ToArray());
        Assert.Throws<InvalidDataException>(() => document.ToBytes());
    }

    [Fact]
    public void ExcessivelyDeepMutableGraphsAreReportedAndRejected() {
        var document = new IcsDocument();
        ContentLineComponent current = document.Calendars.Single();
        for (int index = 0; index < 257; index++) current = current.AddComponent("XNODE");

        Assert.Contains(document.Validate(), issue => issue.Code == "ICAL_COMPONENT_DEPTH_EXCEEDED");
        Assert.Throws<InvalidDataException>(() => document.GetComponents("MISSING").ToArray());
        Assert.Throws<InvalidDataException>(() => document.ToBytes());
    }
}
