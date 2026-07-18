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
    public void ValidationRejectsPropertyGroupsWithoutMutatingRetainedContent() {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        ContentLineProperty version = calendar.GetFirstProperty("VERSION")!;
        ContentLineProperty productId = calendar.GetFirstProperty("PRODID")!;
        ContentLineComponent appointment = calendar.AddComponent("VEVENT");
        ContentLineProperty summary = appointment.AddProperty("SUMMARY", "Grouped content");
        version.Group = "X";
        productId.Group = "X";
        summary.Group = "X";

        ContentLineValidationIssue[] issues = document.Validate().Where(issue =>
            issue.Code == "ICAL_PROPERTY_GROUP_FORBIDDEN").ToArray();

        Assert.Equal(3, issues.Length);
        Assert.Contains(issues, issue => issue.PropertyName == "VERSION");
        Assert.Contains(issues, issue => issue.PropertyName == "PRODID");
        Assert.Contains(issues, issue => issue.PropertyName == "SUMMARY");
        Assert.Equal("X", version.Group);
        Assert.Equal("X", productId.Group);
        Assert.Equal("X", summary.Group);
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

    [Theory]
    [InlineData("20260718", "DATE", "20260719T090000Z", null)]
    [InlineData("20260718T090000Z", null, "20260719", "DATE")]
    public void ValidationAllowsRecurrenceDatesWithATypeDifferentFromStart(
        string start, string? startType, string recurrence, string? recurrenceType) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        ContentLineProperty startProperty = appointment.AddProperty("DTSTART", start);
        if (startType != null) startProperty.SetParameter("VALUE", startType);
        ContentLineProperty recurrenceProperty = appointment.AddProperty("RDATE", recurrence);
        if (recurrenceType != null) recurrenceProperty.SetParameter("VALUE", recurrenceType);

        Assert.DoesNotContain(document.Validate(), issue => issue.PropertyName == "RDATE" &&
            issue.Severity == ContentLineValidationSeverity.Error);
    }

    [Theory]
    [InlineData("20260718", "DATE", "20260719T090000Z", null, "ICAL_EXDATE_TYPE_MISMATCH")]
    [InlineData("20260718T090000Z", null, "20260719", "DATE", "ICAL_EXDATE_TYPE_MISMATCH")]
    [InlineData("20260718T090000Z", null, "20260719T090000", null,
        "ICAL_EXDATE_REPRESENTATION_MISMATCH")]
    [InlineData("20260718T090000", null, "20260719T090000Z", null,
        "ICAL_EXDATE_REPRESENTATION_MISMATCH")]
    public void ValidationMatchesExceptionDatesToStartRepresentation(string start, string? startType,
        string exception, string? exceptionType, string expectedCode) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        ContentLineProperty startProperty = appointment.AddProperty("DTSTART", start);
        if (startType != null) startProperty.SetParameter("VALUE", startType);
        ContentLineProperty exceptionProperty = appointment.AddProperty("EXDATE", exception);
        if (exceptionType != null) exceptionProperty.SetParameter("VALUE", exceptionType);

        Assert.Contains(document.Validate(), issue => issue.Code == expectedCode &&
            issue.PropertyName == "EXDATE");
    }

    [Fact]
    public void ValidationAcceptsMatchingExceptionDateListRepresentation() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260718T090000").SetParameter("TZID", "Europe/Warsaw");
        appointment.AddProperty("EXDATE", "20260719T090000,20260720T090000")
            .SetParameter("TZID", "Europe/Warsaw");

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code.StartsWith("ICAL_EXDATE_", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(null, "Europe/Warsaw")]
    [InlineData("Europe/Warsaw", null)]
    public void ValidationDistinguishesFloatingAndZonedExceptionDates(
        string? startTimeZone, string? exceptionTimeZone) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        ContentLineProperty start = appointment.AddProperty("DTSTART", "20260718T090000");
        if (startTimeZone != null) start.SetParameter("TZID", startTimeZone);
        ContentLineProperty exception = appointment.AddProperty("EXDATE", "20260719T090000");
        if (exceptionTimeZone != null) exception.SetParameter("TZID", exceptionTimeZone);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_EXDATE_REPRESENTATION_MISMATCH" && issue.PropertyName == "EXDATE");
    }

    [Fact]
    public void ValidationAcceptsZonedStartWithUtcExceptionDate() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260718T090000").SetParameter("TZID", "Europe/Warsaw");
        appointment.AddProperty("EXDATE", "20260719T070000Z");

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_EXDATE_REPRESENTATION_MISMATCH");
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
    public void ValidationRejectsDuplicateRecurrenceRuleProperties() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("RRULE", "FREQ=DAILY");
        appointment.AddProperty("RRULE", "FREQ=WEEKLY");

        Assert.Contains(document.Validate(), issue => issue.Code == "ICAL_PROPERTY_CARDINALITY" &&
            issue.ComponentName == "VEVENT" && issue.PropertyName == "RRULE");
    }

    [Theory]
    [InlineData("BYSECOND=61")]
    [InlineData("BYMINUTE=-1")]
    [InlineData("BYHOUR=24")]
    [InlineData("BYDAY=0MO")]
    [InlineData("BYDAY=54TU")]
    [InlineData("BYDAY=XX")]
    [InlineData("BYMONTHDAY=0")]
    [InlineData("BYYEARDAY=367")]
    [InlineData("BYWEEKNO=-54")]
    [InlineData("BYMONTH=13")]
    [InlineData("BYSETPOS=0")]
    [InlineData("WKST=MO,TU")]
    public void ValidationRejectsInvalidRegisteredRecurrenceSelectors(string selector) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("RRULE", "FREQ=YEARLY;" + selector);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_RRULE_PART_VALUE_INVALID" && issue.PropertyName == "RRULE");
    }

    [Fact]
    public void ValidationAcceptsRegisteredRecurrenceSelectorBoundariesAndExtensions() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("RRULE", "FREQ=YEARLY;BYSECOND=0,60;BYMINUTE=0,59;" +
            "BYHOUR=0,23;BYDAY=MO,SU,FR;BYMONTHDAY=-31,31;BYYEARDAY=-366,366;" +
            "BYWEEKNO=-53,53;BYMONTH=1,12;BYSETPOS=-366,366;WKST=SU;X-SELECTOR=opaque");

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_RRULE_PART_VALUE_INVALID");
    }

    [Theory]
    [InlineData("DAILY", "BYWEEKNO=1")]
    [InlineData("WEEKLY", "BYMONTHDAY=1")]
    [InlineData("DAILY", "BYYEARDAY=1")]
    [InlineData("DAILY", "BYDAY=1MO")]
    [InlineData("YEARLY", "BYWEEKNO=1;BYDAY=1MO")]
    [InlineData("MONTHLY", "BYSETPOS=1")]
    public void ValidationRejectsInvalidRecurrenceSelectorRelationships(
        string frequency, string selectors) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("RRULE", "FREQ=" + frequency + ";" + selectors);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_RRULE_PART_RELATION_INVALID" && issue.PropertyName == "RRULE");
    }

    [Fact]
    public void ValidationRejectsTimeSelectorsForDateRecurrences() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260717").SetParameter("VALUE", "DATE");
        appointment.AddProperty("RRULE", "FREQ=DAILY;BYHOUR=9");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_RRULE_PART_RELATION_INVALID" && issue.PropertyName == "RRULE");
    }

    [Theory]
    [InlineData("FREQ=MONTHLY;BYDAY=-5MO,+5FR;BYSETPOS=1")]
    [InlineData("FREQ=YEARLY;BYDAY=-53MO,+53FR")]
    public void ValidationAcceptsValidRecurrenceSelectorRelationships(string rule) {
        var document = new IcsDocument();
        document.Calendars.Single().AddComponent("VEVENT").AddProperty("RRULE", rule);

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_RRULE_PART_RELATION_INVALID");
    }

    [Theory]
    [InlineData("VEVENT", "DTSTART")]
    [InlineData("VEVENT", "DTEND")]
    [InlineData("VEVENT", "RECURRENCE-ID")]
    [InlineData("VEVENT", "CREATED")]
    [InlineData("VTODO", "COMPLETED")]
    [InlineData("VTODO", "DUE")]
    [InlineData("VJOURNAL", "LAST-MODIFIED")]
    [InlineData("VFREEBUSY", "DTSTAMP")]
    [InlineData("VTIMEZONE", "LAST-MODIFIED")]
    public void ValidationRejectsDuplicateSingletonTemporalProperties(
        string componentName, string propertyName) {
        var document = new IcsDocument();
        ContentLineComponent component = document.Calendars.Single().AddComponent(componentName);
        component.AddProperty(propertyName, "20260717T090000Z");
        component.AddProperty(propertyName, "20260718T090000Z");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_PROPERTY_CARDINALITY" && issue.PropertyName == propertyName);
    }

    [Theory]
    [InlineData("UID")]
    [InlineData("DTSTAMP")]
    public void ValidationRequiresFreeBusyIdentityProperties(string missingProperty) {
        var document = new IcsDocument();
        ContentLineComponent freeBusy = document.Calendars.Single().AddComponent("VFREEBUSY");
        if (missingProperty != "UID") freeBusy.AddProperty("UID", "freebusy@example.test");
        if (missingProperty != "DTSTAMP") freeBusy.AddProperty("DTSTAMP", "20260718T090000Z");

        Assert.Contains(document.Validate(), issue => issue.Code == "ICAL_PROPERTY_REQUIRED" &&
            issue.ComponentName == "VFREEBUSY" && issue.PropertyName == missingProperty);
    }

    [Fact]
    public void ValidationAcceptsFreeBusyIdentityPropertiesExactlyOnce() {
        var document = new IcsDocument();
        ContentLineComponent freeBusy = document.Calendars.Single().AddComponent("VFREEBUSY");
        freeBusy.AddProperty("UID", "freebusy@example.test");
        freeBusy.AddProperty("DTSTAMP", "20260718T090000Z");

        Assert.DoesNotContain(document.Validate(), issue =>
            (issue.Code == "ICAL_PROPERTY_REQUIRED" || issue.Code == "ICAL_PROPERTY_CARDINALITY") &&
            issue.ComponentName == "VFREEBUSY" &&
            (issue.PropertyName == "UID" || issue.PropertyName == "DTSTAMP"));
    }

    [Fact]
    public void ValidationRequiresAtLeastOneTimeZoneObservance() {
        var document = new IcsDocument();
        ContentLineComponent timeZone = document.Calendars.Single().AddComponent("VTIMEZONE");
        timeZone.AddProperty("TZID", "Europe/Warsaw");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_TIMEZONE_OBSERVANCE_REQUIRED" &&
            issue.ComponentName == "VTIMEZONE");

        ContentLineComponent standard = timeZone.AddComponent("STANDARD");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_PROPERTY_REQUIRED" && issue.ComponentName == "STANDARD" &&
            issue.PropertyName == "DTSTART");

        standard.AddProperty("DTSTART", "20261025T030000");
        standard.AddProperty("TZOFFSETFROM", "+0200");
        standard.AddProperty("TZOFFSETTO", "+0100");

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_TIMEZONE_OBSERVANCE_REQUIRED" ||
            issue.Code == "ICAL_PROPERTY_REQUIRED" && issue.ComponentName == "STANDARD");
    }

    [Theory]
    [InlineData("TZOFFSETFROM", "not-an-offset")]
    [InlineData("TZOFFSETTO", "+9999")]
    [InlineData("TZOFFSETFROM", "+0160")]
    [InlineData("TZOFFSETTO", "+010060")]
    [InlineData("TZOFFSETFROM", "-0000")]
    [InlineData("TZOFFSETTO", "-000000")]
    [InlineData("TZOFFSETFROM", "0100")]
    public void ValidationRejectsInvalidTimeZoneOffsets(string propertyName, string value) {
        var document = new IcsDocument();
        ContentLineComponent timeZone = document.Calendars.Single().AddComponent("VTIMEZONE");
        timeZone.AddProperty("TZID", "Custom/Invalid");
        ContentLineComponent standard = timeZone.AddComponent("STANDARD");
        standard.AddProperty("DTSTART", "20261025T030000");
        standard.AddProperty("TZOFFSETFROM", "+0200");
        standard.AddProperty("TZOFFSETTO", "+0100");
        standard.GetFirstProperty(propertyName)!.Value = value;

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_TIMEZONE_OFFSET_INVALID" && issue.ComponentName == "STANDARD" &&
            issue.PropertyName == propertyName);
    }

    [Theory]
    [InlineData("+0000")]
    [InlineData("+124530")]
    [InlineData("-0330")]
    [InlineData("-235959")]
    public void ValidationAcceptsValidTimeZoneOffsets(string value) {
        var document = new IcsDocument();
        ContentLineComponent timeZone = document.Calendars.Single().AddComponent("VTIMEZONE");
        timeZone.AddProperty("TZID", "Custom/Valid");
        ContentLineComponent standard = timeZone.AddComponent("STANDARD");
        standard.AddProperty("DTSTART", "20261025T030000");
        standard.AddProperty("TZOFFSETFROM", value);
        standard.AddProperty("TZOFFSETTO", "+0100");

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_TIMEZONE_OFFSET_INVALID");
    }

    [Theory]
    [InlineData("20261025T030000Z", null)]
    [InlineData("20261025T030000", "Europe/Warsaw")]
    [InlineData("20261025", "DATE")]
    public void ValidationRejectsNonFloatingTimeZoneObservanceStarts(
        string value, string? parameterValue) {
        var document = new IcsDocument();
        ContentLineComponent timeZone = document.Calendars.Single().AddComponent("VTIMEZONE");
        timeZone.AddProperty("TZID", "Europe/Warsaw");
        ContentLineComponent standard = timeZone.AddComponent("STANDARD");
        ContentLineProperty start = standard.AddProperty("DTSTART", value);
        if (string.Equals(parameterValue, "DATE", StringComparison.Ordinal))
            start.SetParameter("VALUE", parameterValue!);
        else if (parameterValue != null)
            start.SetParameter("TZID", parameterValue);
        standard.AddProperty("TZOFFSETFROM", "+0200");
        standard.AddProperty("TZOFFSETTO", "+0100");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_TIMEZONE_OBSERVANCE_START_INVALID" &&
            issue.ComponentName == "STANDARD" && issue.PropertyName == "DTSTART");
    }

    [Theory]
    [InlineData("20261025", "VALUE", "DATE")]
    [InlineData("20261025T030000Z", null, null)]
    [InlineData("20261025T030000", "TZID", "Europe/Warsaw")]
    [InlineData("20261025T030000/20261025T040000", "VALUE", "PERIOD")]
    public void ValidationRejectsNonFloatingTimeZoneObservanceRecurrenceDates(
        string value, string? parameterName, string? parameterValue) {
        var document = new IcsDocument();
        ContentLineComponent timeZone = document.Calendars.Single().AddComponent("VTIMEZONE");
        timeZone.AddProperty("TZID", "Europe/Warsaw");
        ContentLineComponent standard = timeZone.AddComponent("STANDARD");
        standard.AddProperty("DTSTART", "20261025T030000");
        standard.AddProperty("TZOFFSETFROM", "+0200");
        standard.AddProperty("TZOFFSETTO", "+0100");
        ContentLineProperty recurrenceDate = standard.AddProperty("RDATE", value);
        if (parameterName != null) recurrenceDate.SetParameter(parameterName, parameterValue!);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_TIMEZONE_OBSERVANCE_RDATE_INVALID" &&
            issue.ComponentName == "STANDARD" && issue.PropertyName == "RDATE");
    }

    [Fact]
    public void ValidationAcceptsFloatingTimeZoneObservanceRecurrenceDateLists() {
        var document = new IcsDocument();
        ContentLineComponent timeZone = document.Calendars.Single().AddComponent("VTIMEZONE");
        timeZone.AddProperty("TZID", "Europe/Warsaw");
        ContentLineComponent standard = timeZone.AddComponent("STANDARD");
        standard.AddProperty("DTSTART", "20261025T030000");
        standard.AddProperty("TZOFFSETFROM", "+0200");
        standard.AddProperty("TZOFFSETTO", "+0100");
        standard.AddProperty("RDATE", "20271031T030000,20281029T030000")
            .SetParameter("VALUE", "DATE-TIME");

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_TIMEZONE_OBSERVANCE_RDATE_INVALID");
    }

    [Theory]
    [InlineData("20260718T090000", "20260718T100000Z", "DTSTART",
        "ICAL_TEMPORAL_VALUE_UTC_REQUIRED")]
    [InlineData("20260718T090000Z", "20260718T100000", "DTEND",
        "ICAL_TEMPORAL_VALUE_UTC_REQUIRED")]
    [InlineData("20260718T100000Z", "20260718T090000Z", "DTEND",
        "ICAL_TEMPORAL_ENDPOINT_ORDER_INVALID")]
    public void ValidationRejectsInvalidFreeBusyWindows(string start, string end,
        string propertyName, string expectedCode) {
        var document = new IcsDocument();
        ContentLineComponent freeBusy = document.Calendars.Single().AddComponent("VFREEBUSY");
        freeBusy.AddProperty("DTSTART", start);
        freeBusy.AddProperty("DTEND", end);

        Assert.Contains(document.Validate(), issue => issue.Code == expectedCode &&
            issue.ComponentName == "VFREEBUSY" && issue.PropertyName == propertyName);
    }

    [Theory]
    [InlineData("not-a-period")]
    [InlineData("20260718T090000/20260718T100000Z")]
    [InlineData("20260718T090000Z/20260718T100000")]
    [InlineData("20260718T100000Z/20260718T090000Z")]
    [InlineData("20260718T090000Z/-PT1H")]
    [InlineData("20260718T090000Z/PT0S")]
    [InlineData("20260718T090000Z/PT1H,")]
    public void ValidationRejectsInvalidFreeBusyPeriods(string value) {
        var document = new IcsDocument();
        ContentLineComponent freeBusy = document.Calendars.Single().AddComponent("VFREEBUSY");
        freeBusy.AddProperty("FREEBUSY", value);

        Assert.Contains(document.Validate(), issue => issue.Code == "ICAL_FREEBUSY_PERIOD_INVALID" &&
            issue.ComponentName == "VFREEBUSY" && issue.PropertyName == "FREEBUSY");
    }

    [Fact]
    public void ValidationAcceptsUtcFreeBusyPeriodLists() {
        var document = new IcsDocument();
        ContentLineComponent freeBusy = document.Calendars.Single().AddComponent("VFREEBUSY");
        freeBusy.AddProperty("FREEBUSY", "20260718T090000Z/20260718T100000Z," +
            "20260718T140000Z/PT30M").SetParameter("FBTYPE", "BUSY");

        Assert.DoesNotContain(document.Validate(), issue => issue.Code == "ICAL_FREEBUSY_PERIOD_INVALID");
    }

    [Fact]
    public void ValidationRejectsValueParametersOnFixedTypeFreeBusyProperties() {
        var document = new IcsDocument();
        ContentLineComponent freeBusy = document.Calendars.Single().AddComponent("VFREEBUSY");
        freeBusy.AddProperty("FREEBUSY", "20260718T090000Z/PT1H").SetParameter("VALUE", "PERIOD");

        Assert.Contains(document.Validate(), issue => issue.Code == "ICAL_FREEBUSY_PERIOD_INVALID" &&
            issue.ComponentName == "VFREEBUSY" && issue.PropertyName == "FREEBUSY");
    }

    [Theory]
    [InlineData("20260718", "20260718T090000Z", "ICAL_RECURRENCE_ID_TYPE_MISMATCH")]
    [InlineData("20260718T090000Z", "20260718", "ICAL_RECURRENCE_ID_TYPE_MISMATCH")]
    [InlineData("20260718T090000", "20260718T090000Z",
        "ICAL_RECURRENCE_ID_REPRESENTATION_MISMATCH")]
    [InlineData("20260718T090000Z", "20260718T090000",
        "ICAL_RECURRENCE_ID_REPRESENTATION_MISMATCH")]
    public void ValidationMatchesRecurrenceIdToDtStart(string startValue,
        string recurrenceValue, string expectedCode) {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        ContentLineComponent master = calendar.AddComponent("VEVENT");
        master.AddProperty("UID", "recurrence-value@example.test");
        ContentLineProperty start = master.AddProperty("DTSTART", startValue);
        ContentLineComponent exception = calendar.AddComponent("VEVENT");
        exception.AddProperty("UID", "recurrence-value@example.test");
        ContentLineProperty exceptionStart = exception.AddProperty("DTSTART", startValue);
        ContentLineProperty recurrence = exception.AddProperty("RECURRENCE-ID", recurrenceValue);
        if (startValue.Length == 8) {
            start.SetParameter("VALUE", "DATE");
            exceptionStart.SetParameter("VALUE", "DATE");
        }
        if (recurrenceValue.Length == 8) recurrence.SetParameter("VALUE", "DATE");

        Assert.Contains(document.Validate(), issue => issue.Code == expectedCode &&
            issue.ComponentName == "VEVENT" && issue.PropertyName == "RECURRENCE-ID");
    }

    [Theory]
    [InlineData("Europe/Warsaw", null, "ICAL_RECURRENCE_ID_REPRESENTATION_MISMATCH")]
    [InlineData("Europe/Warsaw", "Europe/Berlin", "ICAL_RECURRENCE_ID_TIMEZONE_MISMATCH")]
    public void ValidationMatchesRecurrenceIdTimeZoneToDtStart(string startTimeZone,
        string? recurrenceTimeZone, string expectedCode) {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        ContentLineComponent master = calendar.AddComponent("VEVENT");
        master.AddProperty("UID", "recurrence-zone@example.test");
        master.AddProperty("DTSTART", "20260718T090000").SetParameter("TZID", startTimeZone);
        ContentLineComponent exception = calendar.AddComponent("VEVENT");
        exception.AddProperty("UID", "recurrence-zone@example.test");
        exception.AddProperty("DTSTART", "20260725T090000").SetParameter("TZID", "America/New_York");
        ContentLineProperty recurrence = exception.AddProperty("RECURRENCE-ID", "20260725T090000");
        if (recurrenceTimeZone != null) recurrence.SetParameter("TZID", recurrenceTimeZone);

        Assert.Contains(document.Validate(), issue => issue.Code == expectedCode &&
            issue.ComponentName == "VEVENT" && issue.PropertyName == "RECURRENCE-ID");
    }

    [Fact]
    public void ValidationAcceptsMatchingRecurrenceIdTimeZones() {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        ContentLineComponent master = calendar.AddComponent("VEVENT");
        master.AddProperty("UID", "rescheduled-zone@example.test");
        master.AddProperty("DTSTART", "20260718T090000").SetParameter("TZID", "Europe/Warsaw");
        ContentLineComponent exception = calendar.AddComponent("VEVENT");
        exception.AddProperty("UID", "rescheduled-zone@example.test");
        exception.AddProperty("DTSTART", "20260725T090000").SetParameter("TZID", "America/New_York");
        exception.AddProperty("RECURRENCE-ID", "20260725T090000")
            .SetParameter("TZID", "Europe/Warsaw");

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_RECURRENCE_ID_REPRESENTATION_MISMATCH" ||
            issue.Code == "ICAL_RECURRENCE_ID_TIMEZONE_MISMATCH");
    }

    [Theory]
    [InlineData("20260718", "20260719", true)]
    [InlineData("20260718T090000Z", "20260719T090000Z", false)]
    [InlineData("20260718T090000", "20260719T090000", false)]
    public void ValidationAcceptsMatchingRecurrenceIdValueForms(
        string startValue, string recurrenceValue, bool isDate) {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        ContentLineComponent master = calendar.AddComponent("VEVENT");
        master.AddProperty("UID", "matching-form@example.test");
        ContentLineProperty start = master.AddProperty("DTSTART", startValue);
        ContentLineComponent exception = calendar.AddComponent("VEVENT");
        exception.AddProperty("UID", "matching-form@example.test");
        ContentLineProperty exceptionStart = exception.AddProperty("DTSTART", startValue);
        ContentLineProperty recurrence = exception.AddProperty("RECURRENCE-ID", recurrenceValue);
        if (isDate) {
            start.SetParameter("VALUE", "DATE");
            exceptionStart.SetParameter("VALUE", "DATE");
            recurrence.SetParameter("VALUE", "DATE");
        }

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_RECURRENCE_ID_TYPE_MISMATCH" ||
            issue.Code == "ICAL_RECURRENCE_ID_REPRESENTATION_MISMATCH");
    }

    [Theory]
    [InlineData("VEVENT", "DTEND")]
    [InlineData("VTODO", "DUE")]
    public void ValidationRejectsEndpointValueTypesThatDoNotMatchDtStart(
        string componentName, string endpointName) {
        var document = new IcsDocument();
        ContentLineComponent component = document.Calendars.Single().AddComponent(componentName);
        component.AddProperty("DTSTART", "20260717").SetParameter("VALUE", "DATE");
        component.AddProperty(endpointName, "20260718T090000Z");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_TEMPORAL_ENDPOINT_TYPE_MISMATCH" &&
            issue.PropertyName == endpointName);
    }

    [Theory]
    [InlineData("VEVENT", "DTEND", "20260717T090000Z")]
    [InlineData("VTODO", "DUE", "20260717T085959Z")]
    public void ValidationRejectsEndpointsThatAreNotLaterThanDtStart(
        string componentName, string endpointName, string endpointValue) {
        var document = new IcsDocument();
        ContentLineComponent component = document.Calendars.Single().AddComponent(componentName);
        component.AddProperty("DTSTART", "20260717T090000Z");
        component.AddProperty(endpointName, endpointValue);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_TEMPORAL_ENDPOINT_ORDER_INVALID" &&
            issue.PropertyName == endpointName);
    }

    [Fact]
    public void ValidationRejectsUtcAndLocalEndpointRepresentationMismatch() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260717T100000")
            .SetParameter("TZID", "Europe/Warsaw");
        appointment.AddProperty("DTEND", "20260717T090000Z");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_TEMPORAL_ENDPOINT_REPRESENTATION_MISMATCH" &&
            issue.PropertyName == "DTEND");
    }

    [Theory]
    [InlineData("VEVENT", "VEVENT")]
    [InlineData("VCALENDAR", "STANDARD")]
    [InlineData("VEVENT", "DAYLIGHT")]
    public void ValidationRejectsKnownComponentsUnderInvalidParents(
        string parentName, string childName) {
        var document = new IcsDocument();
        ContentLineComponent parent = parentName == "VCALENDAR"
            ? document.Calendars.Single()
            : document.Calendars.Single().AddComponent(parentName);
        parent.AddComponent(childName);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_COMPONENT_PARENT_INVALID" && issue.ComponentName == childName);
    }

    [Fact]
    public void ValidationAllowsExtensionComponentsUnderKnownComponents() {
        var document = new IcsDocument();
        document.Calendars.Single().AddComponent("VEVENT").AddComponent("X-VENDOR-COMPONENT");

        Assert.DoesNotContain(document.Validate(), issue => issue.Code == "ICAL_COMPONENT_PARENT_INVALID");
    }

    [Theory]
    [InlineData("DTSTAMP", "not-a-timestamp", "ICAL_TEMPORAL_VALUE_INVALID")]
    [InlineData("CREATED", "20260717T250000Z", "ICAL_TEMPORAL_VALUE_INVALID")]
    [InlineData("LAST-MODIFIED", "20260717T090000", "ICAL_TEMPORAL_VALUE_UTC_REQUIRED")]
    [InlineData("COMPLETED", "20260717T090000", "ICAL_TEMPORAL_VALUE_UTC_REQUIRED")]
    public void ValidationChecksStandardUtcTimestampProperties(
        string propertyName, string value, string expectedCode) {
        var document = new IcsDocument();
        ContentLineComponent task = document.Calendars.Single().AddComponent("VTODO");
        task.AddProperty(propertyName, value);

        Assert.Contains(document.Validate(), issue => issue.Code == expectedCode &&
            issue.PropertyName == propertyName);
    }

    [Theory]
    [InlineData("20261340")]
    [InlineData("20260717T250000Z")]
    [InlineData("20260717T090000Zjunk")]
    public void ValidationRejectsMalformedRecurrenceUntilValues(string until) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260717T090000Z");
        appointment.AddProperty("RRULE", "FREQ=DAILY;UNTIL=" + until);

        Assert.Contains(document.Validate(), issue => issue.Code == "ICAL_RRULE_UNTIL_INVALID" &&
            issue.PropertyName == "RRULE");
    }

    [Theory]
    [InlineData("20260717", "DATE", "20260720T090000Z")]
    [InlineData("20260717T090000", "Europe/Warsaw", "20260720T090000")]
    [InlineData("20260717T090000", null, "20260720T090000Z")]
    public void ValidationRejectsRecurrenceUntilFormsThatDoNotMatchDtStart(
        string start, string? startParameter, string until) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        ContentLineProperty startProperty = appointment.AddProperty("DTSTART", start);
        if (startParameter == "DATE") startProperty.SetParameter("VALUE", "DATE");
        else if (startParameter != null) startProperty.SetParameter("TZID", startParameter);
        appointment.AddProperty("RRULE", "FREQ=DAILY;UNTIL=" + until);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_RRULE_UNTIL_TYPE_MISMATCH" && issue.PropertyName == "RRULE");
    }

    [Fact]
    public void ValidationAcceptsRecurrenceUntilFormsRequiredByDtStart() {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        ContentLineComponent date = calendar.AddComponent("VEVENT");
        date.AddProperty("DTSTART", "20260717").SetParameter("VALUE", "DATE");
        date.AddProperty("RRULE", "FREQ=DAILY;UNTIL=20260720");
        ContentLineComponent zoned = calendar.AddComponent("VEVENT");
        zoned.AddProperty("DTSTART", "20260717T090000").SetParameter("TZID", "Europe/Warsaw");
        zoned.AddProperty("RRULE", "FREQ=DAILY;UNTIL=20260720T090000Z");
        ContentLineComponent floating = calendar.AddComponent("VEVENT");
        floating.AddProperty("DTSTART", "20260717T090000");
        floating.AddProperty("RRULE", "FREQ=DAILY;UNTIL=20260720T090000");

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_RRULE_UNTIL_INVALID" ||
            issue.Code == "ICAL_RRULE_UNTIL_TYPE_MISMATCH");
    }

    [Theory]
    [InlineData("not-a-duration", null, null)]
    [InlineData("P1Y", null, null)]
    [InlineData("20260717T090000", "DATE-TIME", null)]
    [InlineData("20260717T090000Z", "DATE-TIME", "END")]
    [InlineData("-PT15M", "TEXT", null)]
    public void ValidationRejectsInvalidAlarmTriggers(
        string value, string? valueType, string? related) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260717T090000Z");
        appointment.AddProperty("DTEND", "20260717T100000Z");
        ContentLineComponent alarm = appointment.AddComponent("VALARM");
        alarm.AddProperty("ACTION", "DISPLAY");
        alarm.AddProperty("DESCRIPTION", "Reminder");
        ContentLineProperty trigger = alarm.AddProperty("TRIGGER", value);
        if (valueType != null) trigger.SetParameter("VALUE", valueType);
        if (related != null) trigger.SetParameter("RELATED", related);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_ALARM_TRIGGER_INVALID" && issue.PropertyName == "TRIGGER");
    }

    [Theory]
    [InlineData("-PT15M", null, null)]
    [InlineData("PT0S", "DURATION", "START")]
    [InlineData("+P1W", null, "END")]
    [InlineData("20260717T090000Z", "DATE-TIME", null)]
    public void ValidationAcceptsRfcAlarmTriggerForms(
        string value, string? valueType, string? related) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260717T090000Z");
        appointment.AddProperty("DTEND", "20260717T100000Z");
        ContentLineComponent alarm = appointment.AddComponent("VALARM");
        alarm.AddProperty("ACTION", "DISPLAY");
        alarm.AddProperty("DESCRIPTION", "Reminder");
        ContentLineProperty trigger = alarm.AddProperty("TRIGGER", value);
        if (valueType != null) trigger.SetParameter("VALUE", valueType);
        if (related != null) trigger.SetParameter("RELATED", related);

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_ALARM_TRIGGER_INVALID");
    }

    [Theory]
    [InlineData("VEVENT", "not-a-duration", null)]
    [InlineData("VEVENT", "P1Y", null)]
    [InlineData("VTODO", "-PT1H", null)]
    [InlineData("VTODO", "PT0S", null)]
    [InlineData("VEVENT", "PT1H", "DATE-TIME")]
    [InlineData("VEVENT", "PT1H", "DURATION")]
    public void ValidationRejectsInvalidEventAndTaskDurations(
        string componentName, string value, string? valueType) {
        var document = new IcsDocument();
        ContentLineComponent component = document.Calendars.Single().AddComponent(componentName);
        ContentLineProperty duration = component.AddProperty("DURATION", value);
        if (valueType != null) duration.SetParameter("VALUE", valueType);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_DURATION_INVALID" && issue.PropertyName == "DURATION");
    }

    [Theory]
    [InlineData("VEVENT", "P1W")]
    [InlineData("VEVENT", "+P2DT3H4M5S")]
    [InlineData("VTODO", "PT30M")]
    public void ValidationAcceptsPositiveEventAndTaskDurations(string componentName, string value) {
        var document = new IcsDocument();
        ContentLineComponent component = document.Calendars.Single().AddComponent(componentName);
        component.AddProperty("DTSTART", "20260717T090000Z");
        component.AddProperty("DURATION", value);

        Assert.DoesNotContain(document.Validate(), issue => issue.Code == "ICAL_DURATION_INVALID");
        Assert.DoesNotContain(document.Validate(), issue => issue.Code == "ICAL_DURATION_END_CONFLICT");
        Assert.DoesNotContain(document.Validate(), issue => issue.Code == "ICAL_DURATION_START_REQUIRED");
    }

    [Fact]
    public void ValidationRejectsDuplicateEventDurations() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DURATION", "PT1H");
        appointment.AddProperty("DURATION", "PT2H");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_PROPERTY_CARDINALITY" && issue.PropertyName == "DURATION");
    }

    [Theory]
    [InlineData("VEVENT", "DTEND")]
    [InlineData("VTODO", "DUE")]
    public void ValidationRejectsDurationTogetherWithAnExplicitEnd(
        string componentName, string endPropertyName) {
        var document = new IcsDocument();
        ContentLineComponent component = document.Calendars.Single().AddComponent(componentName);
        component.AddProperty("DTSTART", "20260717T090000Z");
        component.AddProperty(endPropertyName, "20260717T100000Z");
        component.AddProperty("DURATION", "PT1H");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_DURATION_END_CONFLICT" && issue.PropertyName == "DURATION");
    }

    [Fact]
    public void ValidationRequiresTaskStartWhenDurationIsPresent() {
        var document = new IcsDocument();
        document.Calendars.Single().AddComponent("VTODO").AddProperty("DURATION", "PT1H");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_DURATION_START_REQUIRED" && issue.PropertyName == "DURATION");
    }

    [Fact]
    public void ValidationRequiresStartForMethodlessEvent() {
        var document = new IcsDocument();
        document.Calendars.Single().AddComponent("VEVENT");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_PROPERTY_REQUIRED" && issue.PropertyName == "DTSTART");
    }

    [Fact]
    public void ValidationAllowsMissingEventStartWhenMethodIsPresent() {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        calendar.AddProperty("METHOD", "REQUEST");
        calendar.AddComponent("VEVENT");

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_PROPERTY_REQUIRED" && issue.PropertyName == "DTSTART");
    }

    [Theory]
    [InlineData("")]
    [InlineData("bad method")]
    [InlineData("PUBLISH!")]
    public void ValidationRejectsInvalidMethodsAndStillRequiresEventStart(string value) {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        calendar.AddProperty("METHOD", value);
        calendar.AddComponent("VEVENT");

        IReadOnlyList<ContentLineValidationIssue> issues = document.Validate();
        Assert.Contains(issues, issue => issue.Code == "ICAL_METHOD_INVALID" &&
            issue.PropertyName == "METHOD");
        Assert.Contains(issues, issue => issue.Code == "ICAL_PROPERTY_REQUIRED" &&
            issue.PropertyName == "DTSTART");
    }

    [Theory]
    [InlineData("VEVENT")]
    [InlineData("VTODO")]
    [InlineData("VJOURNAL")]
    public void ValidationRequiresStartForRecurringComponents(string componentName) {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        calendar.AddProperty("METHOD", "PUBLISH");
        calendar.AddComponent(componentName).AddProperty("RRULE", "FREQ=DAILY;COUNT=2");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_PROPERTY_REQUIRED" && issue.PropertyName == "DTSTART");
    }

    [Theory]
    [InlineData("not-an-integer")]
    [InlineData("-1")]
    [InlineData("2147483648")]
    public void ValidationRejectsInvalidSequenceValues(string value) {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        calendar.AddProperty("METHOD", "PUBLISH");
        ContentLineComponent appointment = calendar.AddComponent("VEVENT");
        appointment.AddProperty("SEQUENCE", value);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_SEQUENCE_INVALID" && issue.PropertyName == "SEQUENCE");
    }

    [Theory]
    [InlineData("0")]
    [InlineData("+3")]
    public void ValidationAcceptsNonNegativeSequenceValues(string value) {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        calendar.AddProperty("METHOD", "PUBLISH");
        calendar.AddComponent("VEVENT").AddProperty("SEQUENCE", value);

        Assert.DoesNotContain(document.Validate(), issue => issue.Code == "ICAL_SEQUENCE_INVALID");
    }

    [Fact]
    public void ValidationRejectsValueParameterOnSequence() {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        calendar.AddProperty("METHOD", "PUBLISH");
        calendar.AddComponent("VEVENT").AddProperty("SEQUENCE", "3").SetParameter("VALUE", "INTEGER");

        Assert.Contains(document.Validate(), issue => issue.Code == "ICAL_SEQUENCE_INVALID");
    }

    [Fact]
    public void ValidationRequiresWholeDayDurationForDateStart() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260717").SetParameter("VALUE", "DATE");
        appointment.AddProperty("DURATION", "PT1H");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_DURATION_DATE_START_INVALID" && issue.PropertyName == "DURATION");
    }

    [Theory]
    [InlineData("P2D")]
    [InlineData("P1W")]
    public void ValidationAcceptsWholeDayDurationsForDateStart(string duration) {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260717").SetParameter("VALUE", "DATE");
        appointment.AddProperty("DURATION", duration);

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_DURATION_DATE_START_INVALID" || issue.Code == "ICAL_DURATION_INVALID");
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void ValidationRequiresAlarmDurationAndRepeatTogether(bool addDuration, bool addRepeat) {
        ContentLineComponent alarm = CreateValidAlarm(out IcsDocument document);
        if (addDuration) alarm.AddProperty("DURATION", "PT5M");
        if (addRepeat) alarm.AddProperty("REPEAT", "3");

        Assert.Contains(document.Validate(), issue => issue.Code == "ICAL_ALARM_REPEAT_PAIR_REQUIRED");
    }

    [Theory]
    [InlineData("PT0S", "3", "ICAL_ALARM_DURATION_INVALID")]
    [InlineData("PT5M", "-1", "ICAL_ALARM_REPEAT_INVALID")]
    [InlineData("PT5M", "not-an-integer", "ICAL_ALARM_REPEAT_INVALID")]
    public void ValidationRejectsInvalidAlarmRepetition(
        string duration, string repeat, string expectedCode) {
        ContentLineComponent alarm = CreateValidAlarm(out IcsDocument document);
        alarm.AddProperty("DURATION", duration);
        alarm.AddProperty("REPEAT", repeat);

        Assert.Contains(document.Validate(), issue => issue.Code == expectedCode);
    }

    [Theory]
    [InlineData("0")]
    [InlineData("+3")]
    public void ValidationAcceptsValidAlarmRepetition(string repeat) {
        ContentLineComponent alarm = CreateValidAlarm(out IcsDocument document);
        alarm.AddProperty("DURATION", "PT5M");
        alarm.AddProperty("REPEAT", repeat);

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code.StartsWith("ICAL_ALARM_REPEAT", StringComparison.Ordinal) ||
            issue.Code == "ICAL_ALARM_DURATION_INVALID");
    }

    [Theory]
    [InlineData("DURATION", "ICAL_ALARM_DURATION_INVALID")]
    [InlineData("REPEAT", "ICAL_ALARM_REPEAT_INVALID")]
    public void ValidationRejectsValueParametersOnAlarmRepetitionProperties(
        string propertyName, string expectedCode) {
        ContentLineComponent alarm = CreateValidAlarm(out IcsDocument document);
        alarm.AddProperty("DURATION", "PT5M");
        alarm.AddProperty("REPEAT", "3");
        alarm.GetFirstProperty(propertyName)!.SetParameter("VALUE",
            propertyName == "DURATION" ? "DURATION" : "INTEGER");

        Assert.Contains(document.Validate(), issue => issue.Code == expectedCode);
    }

    [Theory]
    [InlineData("")]
    [InlineData("BAD ACTION")]
    [InlineData("DISPLAY!")]
    public void ValidationRejectsInvalidAlarmActionTokens(string action) {
        ContentLineComponent alarm = CreateValidAlarm(out IcsDocument document);
        alarm.GetFirstProperty("ACTION")!.Value = action;

        Assert.Contains(document.Validate(), issue => issue.Code == "ICAL_ALARM_ACTION_INVALID" &&
            issue.ComponentName == "VALARM" && issue.PropertyName == "ACTION");
    }

    [Fact]
    public void ValidationAcceptsExtensionAlarmActionTokens() {
        ContentLineComponent alarm = CreateValidAlarm(out IcsDocument document);
        alarm.GetFirstProperty("ACTION")!.Value = "X-CUSTOM";

        Assert.DoesNotContain(document.Validate(), issue => issue.Code == "ICAL_ALARM_ACTION_INVALID");
    }

    [Theory]
    [InlineData("DURATION")]
    [InlineData("REPEAT")]
    public void ValidationRejectsDuplicateAlarmRepetitionProperties(string propertyName) {
        ContentLineComponent alarm = CreateValidAlarm(out IcsDocument document);
        alarm.AddProperty("DURATION", "PT5M");
        alarm.AddProperty("REPEAT", "3");
        alarm.AddProperty(propertyName, propertyName == "DURATION" ? "PT10M" : "4");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "ICAL_PROPERTY_CARDINALITY" && issue.PropertyName == propertyName);
    }

    [Theory]
    [InlineData("START", "ICAL_ALARM_TRIGGER_START_REQUIRED")]
    [InlineData("END", "ICAL_ALARM_TRIGGER_END_REQUIRED")]
    public void ValidationRequiresRelativeAlarmParentAnchors(string relationship, string expectedCode) {
        var document = new IcsDocument();
        ContentLineComponent alarm = document.Calendars.Single().AddComponent("VEVENT")
            .AddComponent("VALARM");
        alarm.AddProperty("ACTION", "DISPLAY");
        alarm.AddProperty("DESCRIPTION", "Reminder");
        alarm.AddProperty("TRIGGER", "-PT15M").SetParameter("RELATED", relationship);

        Assert.Contains(document.Validate(), issue => issue.Code == expectedCode);
    }

    [Fact]
    public void ValidationAcceptsEndRelativeAlarmWithDerivedParentEnd() {
        var document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260717T090000Z");
        appointment.AddProperty("DURATION", "PT1H");
        ContentLineComponent alarm = appointment.AddComponent("VALARM");
        alarm.AddProperty("ACTION", "DISPLAY");
        alarm.AddProperty("DESCRIPTION", "Reminder");
        alarm.AddProperty("TRIGGER", "-PT15M").SetParameter("RELATED", "END");

        Assert.DoesNotContain(document.Validate(), issue =>
            issue.Code == "ICAL_ALARM_TRIGGER_END_REQUIRED");
    }

    private static ContentLineComponent CreateValidAlarm(out IcsDocument document) {
        document = new IcsDocument();
        ContentLineComponent appointment = document.Calendars.Single().AddComponent("VEVENT");
        appointment.AddProperty("DTSTART", "20260717T090000Z");
        ContentLineComponent alarm = appointment.AddComponent("VALARM");
        alarm.AddProperty("ACTION", "DISPLAY");
        alarm.AddProperty("DESCRIPTION", "Reminder");
        alarm.AddProperty("TRIGGER", "-PT15M");
        return alarm;
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
    public void SerializationRejectsLiteralLineBreaksInRawPropertyValues() {
        var document = new IcsDocument();
        document.Calendars.Single().AddProperty("X-INJECTED",
            "safe\r\nEND:VCALENDAR\r\nBEGIN:VCALENDAR");

        Assert.Throws<InvalidDataException>(() => document.Serialize());
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
