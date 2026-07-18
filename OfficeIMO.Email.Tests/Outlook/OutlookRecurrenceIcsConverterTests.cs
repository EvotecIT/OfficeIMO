using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookRecurrenceIcsConverterTests {
    private const string MicrosoftWeeklyWithException =
        "043004300B2001000000C0210000010000000000000032000000222000000C0000000000000001000000A096BC0C" +
        "01000000A096BC0C8020BC0C20ADBC0C0630000009300000580200007602000001003499BC0C5299BC0CF898BC0C" +
        "11002200210053696D706C6520526563757272656E6365207769746820657863657074696F6E730800070033342F" +
        "34313431000000000400000000000000000000003499BC0C5299BC0CF898BC0C2100530069006D0070006C006500" +
        "200052006500630075007200720065006E0063006500200077006900740068002000650078006300650070007400" +
        "69006F006E0073000700330034002F0034003100340031000000000000000000";

    private const string MicrosoftPacificRecurrenceDefinition =
        "0201300002001500500061006300690066006900630020005300740061006E0064006100720064002000540069006D00" +
        "6500020002013E000000D6070000000000000000000000000000E001000000000000C4FFFFFF00000A00000005000200" +
        "0000000000000000040000000100020000000000000002013E000300D7070000000000000000000000000000E00100" +
        "0000000000C4FFFFFF00000B0000000100020000000000000000000300000002000200000000000000";

    [Fact]
    public void ExportsMicrosoftSeriesWithoutTreatingModifiedOccurrenceAsDeleted() {
        OutlookRecurrence recurrence = OutlookRecurrenceBinary.DecodeAppointment(
            FromHex(MicrosoftWeeklyWithException));
        recurrence.TimeZoneId = "Pacific Standard Time";

        OutlookRecurrenceIcsExportResult result = OutlookRecurrenceIcsConverter.Export(recurrence,
            new OutlookRecurrenceIcsExportOptions {
                TimeZone = OutlookTimeZoneBinary.DecodeDefinition(
                    FromHex(MicrosoftPacificRecurrenceDefinition))
            });

        Assert.True(result.Report.Succeeded);
        Assert.True(result.Report.IsLossless);
        Assert.Equal("FREQ=WEEKLY;BYDAY=MO,TH,FR;COUNT=12;WKST=SU", result.Rule!.ToString());
        Assert.Empty(result.ExcludedDates);
        OutlookRecurrenceIcsException exception = Assert.Single(result.Exceptions);
        Assert.Equal(IcsTemporalValueKind.ZonedDateTime, exception.OriginalStart.Kind);
        Assert.Equal("Pacific Standard Time", exception.OriginalStart.TimeZoneId);
        Assert.Equal(new DateTime(2007, 4, 16, 10, 0, 0), exception.OriginalStart.Value);
        Assert.Equal(new DateTime(2007, 4, 16, 11, 0, 0), exception.Start.Value);
        Assert.Equal("Simple Recurrence with exceptions", exception.Subject);
    }

    [Fact]
    public void ImportsExportedRuleExceptionsAndZoneSemantics() {
        OutlookTimeZoneDefinition timeZone = OutlookTimeZoneBinary.DecodeDefinition(
            FromHex(MicrosoftPacificRecurrenceDefinition));
        OutlookRecurrence source = OutlookRecurrenceBinary.DecodeAppointment(
            FromHex(MicrosoftWeeklyWithException));
        source.TimeZoneId = timeZone.KeyName;
        OutlookRecurrenceIcsExportResult exported = OutlookRecurrenceIcsConverter.Export(source,
            new OutlookRecurrenceIcsExportOptions { TimeZone = timeZone });
        var options = new OutlookRecurrenceIcsImportOptions {
            Start = IcsTemporalValue.Zoned(source.Start, timeZone.KeyName!),
            Duration = source.Duration,
            TimeZone = timeZone
        };
        foreach (IcsTemporalValue value in exported.ExcludedDates) options.ExcludedDates.Add(value);
        foreach (OutlookRecurrenceIcsException value in exported.Exceptions) options.Exceptions.Add(value);

        OutlookRecurrenceIcsImportResult result = OutlookRecurrenceIcsConverter.Import(exported.Rule!, options);

        Assert.True(result.Report.Succeeded);
        Assert.Equal(source.Frequency, result.Recurrence!.Frequency);
        Assert.Equal(source.DaysOfWeek, result.Recurrence.DaysOfWeek);
        Assert.Equal(12, result.Recurrence.OccurrenceCount);
        Assert.Equal("Pacific Standard Time", result.Recurrence.TimeZoneId);
        Assert.Equal(new DateTime(2007, 4, 16, 11, 0, 0),
            Assert.Single(result.Recurrence.Exceptions).Start);
    }

    [Fact]
    public void ConvertsZonedEndDateToUtcUntil() {
        OutlookTimeZoneDefinition timeZone = OutlookTimeZoneBinary.DecodeDefinition(
            FromHex(MicrosoftPacificRecurrenceDefinition));
        var recurrence = new OutlookRecurrence {
            Frequency = OutlookRecurrenceFrequency.Daily,
            PatternKind = OutlookRecurrencePatternKind.Day,
            Start = new DateTime(2007, 7, 15, 10, 0, 0),
            Duration = TimeSpan.FromMinutes(30),
            RangeKind = OutlookRecurrenceRangeKind.EndDate,
            EndDate = new DateTime(2007, 7, 20),
            TimeZoneId = "Pacific Standard Time"
        };

        OutlookRecurrenceIcsExportResult result = OutlookRecurrenceIcsConverter.Export(recurrence,
            new OutlookRecurrenceIcsExportOptions { TimeZone = timeZone });

        Assert.True(result.Report.Succeeded);
        Assert.Equal("20070720T170000Z", result.Rule!.GetValue("UNTIL"));
    }

    [Fact]
    public void ReportsUnsupportedRulePartsAndRejectsUnrepresentablePattern() {
        IcsRecurrenceRule withExtension = IcsRecurrenceRule.Parse("FREQ=DAILY;X-TEST=1");
        var options = new OutlookRecurrenceIcsImportOptions {
            Start = IcsTemporalValue.Floating(new DateTime(2026, 1, 1, 9, 0, 0)),
            Duration = TimeSpan.FromHours(1)
        };

        OutlookRecurrenceIcsImportResult warning = OutlookRecurrenceIcsConverter.Import(withExtension, options);
        OutlookRecurrenceIcsImportResult error = OutlookRecurrenceIcsConverter.Import(
            IcsRecurrenceRule.Parse("FREQ=DAILY;BYDAY=MO"), options);

        Assert.True(warning.Report.Succeeded);
        Assert.False(warning.Report.IsLossless);
        Assert.Contains(warning.Report.Issues, issue => issue.Code == "ICAL_RRULE_PART_UNSUPPORTED");
        Assert.Null(error.Recurrence);
        Assert.False(error.Report.Succeeded);
        Assert.Contains(error.Report.Issues, issue => issue.Code == "ICAL_DAILY_FILTER_UNSUPPORTED");
    }

    [Fact]
    public void SupportsMultiYearOutlookPeriodAndExpansion() {
        var recurrence = new OutlookRecurrence {
            Frequency = OutlookRecurrenceFrequency.Yearly,
            PatternKind = OutlookRecurrencePatternKind.MonthDay,
            DayOfMonth = 6,
            Interval = 2,
            Start = new DateTime(2026, 7, 6, 9, 0, 0),
            Duration = TimeSpan.FromHours(1),
            RangeKind = OutlookRecurrenceRangeKind.OccurrenceCount,
            OccurrenceCount = 3
        };

        OutlookRecurrence decoded = OutlookRecurrenceBinary.DecodeTask(
            OutlookRecurrenceBinary.EncodeTask(recurrence));
        OutlookRecurrenceExpansionResult expanded = OutlookRecurrenceExpander.Expand(decoded,
            new OutlookRecurrenceExpansionOptions { MaxOccurrences = 10 });

        Assert.True(decoded.StateDecoded, decoded.DecodeError);
        Assert.Equal(2, decoded.Interval);
        Assert.Equal(new[] { 2026, 2028, 2030 }, expanded.Occurrences.Select(value => value.Start.Year));
    }

    [Fact]
    public void EmlWriterEmitsTimeZoneRuleExclusionsAndExceptionComponents() {
        OutlookRecurrence recurrence = OutlookRecurrenceBinary.DecodeAppointment(
            FromHex(MicrosoftWeeklyWithException));
        OutlookTimeZoneDefinition timeZone = OutlookTimeZoneBinary.DecodeDefinition(
            FromHex(MicrosoftPacificRecurrenceDefinition));
        recurrence.TimeZoneId = timeZone.KeyName;
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            MessageId = "series@example.com",
            Subject = "Recurring",
            Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2007, 3, 26, 10, 0, 0, TimeSpan.FromHours(-7)),
                End = new DateTimeOffset(2007, 3, 26, 10, 30, 0, TimeSpan.FromHours(-7)),
                Recurrence = recurrence,
                RecurrenceTimeZone = timeZone
            }
        };

        byte[] eml = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.Eml);
        EmailReadResult read = new EmailDocumentReader().Read(eml);
        EmailAttachment calendar = Assert.Single(read.Document.Attachments,
            attachment => string.Equals(attachment.ContentType, "text/calendar",
                StringComparison.OrdinalIgnoreCase));
        IcsDocument ics = IcsDocument.Load(Assert.IsType<byte[]>(calendar.Content));
        ContentLineComponent[] events = ics.GetComponents("VEVENT").ToArray();
        ContentLineComponent master = Assert.Single(events,
            component => component.GetFirstProperty("RECURRENCE-ID") == null);
        ContentLineComponent exception = Assert.Single(events,
            component => component.GetFirstProperty("RECURRENCE-ID") != null);

        Assert.Equal("FREQ=WEEKLY;BYDAY=MO,TH,FR;COUNT=12;WKST=SU",
            master.GetFirstProperty("RRULE")!.Value);
        Assert.Equal("Pacific Standard Time", master.GetFirstProperty("DTSTART")!
            .GetParameter("TZID")!.Values.Single());
        Assert.Equal("Simple Recurrence with exceptions",
            exception.GetFirstProperty("SUMMARY")!.Value);
        Assert.Single(ics.GetComponents("VTIMEZONE"));
        Assert.DoesNotContain(ics.Validate(), issue => issue.Severity == ContentLineValidationSeverity.Error);
        Assert.NotNull(read.Document.Appointment!.Recurrence);
        Assert.Equal(12, read.Document.Appointment.Recurrence!.OccurrenceCount);
        Assert.Equal(new DateTime(2007, 4, 16, 11, 0, 0),
            Assert.Single(read.Document.Appointment.Recurrence.Exceptions).Start);
    }

    private static byte[] FromHex(string value) {
        var bytes = new byte[value.Length / 2];
        for (int index = 0; index < bytes.Length; index++)
            bytes[index] = Convert.ToByte(value.Substring(index * 2, 2), 16);
        return bytes;
    }
}
