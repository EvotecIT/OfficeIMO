using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookRecurrenceTests {
    // MS-OXOCAL 4.1.1.1, Weekly Recurrence BLOB Without Exceptions.
    private const string MicrosoftWeeklyWithoutExceptions =
        "043004300B2001000000C0210000010000000000000032000000222000000C000000000000000000000000000000" +
        "8020BC0C20ADBC0C0630000009300000580200007602000000000000000000000000";

    // MS-OXOCAL 4.1.1.2, Weekly Recurrence BLOB with Exceptions.
    private const string MicrosoftWeeklyWithException =
        "043004300B2001000000C0210000010000000000000032000000222000000C0000000000000001000000A096BC0C" +
        "01000000A096BC0C8020BC0C20ADBC0C0630000009300000580200007602000001003499BC0C5299BC0CF898BC0C" +
        "11002200210053696D706C6520526563757272656E6365207769746820657863657074696F6E730800070033342F" +
        "34313431000000000400000000000000000000003499BC0C5299BC0CF898BC0C2100530069006D0070006C006500" +
        "200052006500630075007200720065006E0063006500200077006900740068002000650078006300650070007400" +
        "69006F006E0073000700330034002F0034003100340031000000000000000000";

    [Fact]
    public void DecodesAndPreservesMicrosoftWeeklyFixture() {
        byte[] bytes = FromHex(MicrosoftWeeklyWithoutExceptions);

        OutlookRecurrence recurrence = OutlookRecurrenceBinary.DecodeAppointment(bytes);

        Assert.True(recurrence.StateDecoded, recurrence.DecodeError);
        Assert.Equal(OutlookRecurrenceFrequency.Weekly, recurrence.Frequency);
        Assert.Equal(OutlookRecurrencePatternKind.Week, recurrence.PatternKind);
        Assert.Equal(1, recurrence.Interval);
        Assert.Equal(OutlookRecurrenceDays.Monday | OutlookRecurrenceDays.Thursday |
                     OutlookRecurrenceDays.Friday, recurrence.DaysOfWeek);
        Assert.Equal(OutlookRecurrenceRangeKind.OccurrenceCount, recurrence.RangeKind);
        Assert.Equal(12, recurrence.OccurrenceCount);
        Assert.Equal(DayOfWeek.Sunday, recurrence.FirstDayOfWeek);
        Assert.Equal(new DateTime(2007, 3, 26, 10, 0, 0), recurrence.Start);
        Assert.Equal(TimeSpan.FromMinutes(30), recurrence.Duration);
        Assert.Empty(recurrence.DeletedOccurrenceDates);
        Assert.Empty(recurrence.Exceptions);
        Assert.Equal(bytes, OutlookRecurrenceBinary.EncodeAppointment(recurrence));
    }

    [Fact]
    public void DecodesAndPreservesMicrosoftExceptionFixture() {
        byte[] bytes = FromHex(MicrosoftWeeklyWithException);

        OutlookRecurrence recurrence = OutlookRecurrenceBinary.DecodeAppointment(bytes);

        Assert.True(recurrence.StateDecoded, recurrence.DecodeError);
        Assert.Equal(new DateTime(2007, 4, 16), Assert.Single(recurrence.DeletedOccurrenceDates));
        OutlookRecurrenceException exception = Assert.Single(recurrence.Exceptions);
        Assert.Equal(new DateTime(2007, 4, 16, 10, 0, 0), exception.OriginalStart);
        Assert.Equal(new DateTime(2007, 4, 16, 11, 0, 0), exception.Start);
        Assert.Equal(new DateTime(2007, 4, 16, 11, 30, 0), exception.End);
        Assert.Equal("Simple Recurrence with exceptions", exception.Subject);
        Assert.Equal("34/4141", exception.Location);
        Assert.Equal(bytes, OutlookRecurrenceBinary.EncodeAppointment(recurrence));
    }

    [Fact]
    public void EncodesEditedExceptionAndDecodesItAgain() {
        OutlookRecurrence recurrence = OutlookRecurrenceBinary.DecodeAppointment(
            FromHex(MicrosoftWeeklyWithException));
        OutlookRecurrenceException exception = Assert.Single(recurrence.Exceptions);
        exception.Subject = "Żółw — moved";

        byte[] updated = OutlookRecurrenceBinary.EncodeAppointment(recurrence, 1250);
        OutlookRecurrence result = OutlookRecurrenceBinary.DecodeAppointment(updated, 1250);

        Assert.True(result.StateDecoded, result.DecodeError);
        Assert.Equal("Żółw — moved", Assert.Single(result.Exceptions).Subject);
        Assert.NotEqual(FromHex(MicrosoftWeeklyWithException), updated);
    }

    [Fact]
    public void ExpandsCountedSeriesWithMovedAndDeletedOccurrences() {
        OutlookRecurrence recurrence = OutlookRecurrenceBinary.DecodeAppointment(
            FromHex(MicrosoftWeeklyWithException));
        recurrence.DeletedOccurrenceDates.Add(new DateTime(2007, 4, 19));

        OutlookRecurrenceExpansionResult result = OutlookRecurrenceExpander.Expand(recurrence,
            new OutlookRecurrenceExpansionOptions {
                WindowStart = new DateTime(2007, 4, 13),
                WindowEnd = new DateTime(2007, 4, 21),
                MaxOccurrences = 20
            });

        Assert.False(result.Truncated);
        Assert.Equal(3, result.Occurrences.Count);
        Assert.Equal(new DateTime(2007, 4, 13, 10, 0, 0), result.Occurrences[0].Start);
        Assert.Equal(new DateTime(2007, 4, 16, 11, 0, 0), result.Occurrences[1].Start);
        Assert.True(result.Occurrences[1].IsException);
        Assert.Equal(new DateTime(2007, 4, 20, 10, 0, 0), result.Occurrences[2].Start);
    }

    [Fact]
    public void ReportsSafetyBoundInsteadOfExpandingForever() {
        var recurrence = new OutlookRecurrence {
            Frequency = OutlookRecurrenceFrequency.Daily,
            PatternKind = OutlookRecurrencePatternKind.Day,
            Start = new DateTime(2026, 1, 1, 9, 0, 0),
            Duration = TimeSpan.FromHours(1),
            RangeKind = OutlookRecurrenceRangeKind.NoEnd
        };

        OutlookRecurrenceExpansionResult result = OutlookRecurrenceExpander.Expand(recurrence,
            new OutlookRecurrenceExpansionOptions {
                WindowEnd = new DateTime(2030, 1, 1),
                MaxOccurrences = 3,
                MaxCandidateDays = 100
            });

        Assert.True(result.Truncated);
        Assert.Equal("MaxOccurrences was reached.", result.TruncationReason);
        Assert.Equal(3, result.Occurrences.Count);
    }

    [Fact]
    public void WeeklyExpansionCountsSkippedCalendarDatesAgainstCandidateBound() {
        var recurrence = new OutlookRecurrence {
            Frequency = OutlookRecurrenceFrequency.Weekly,
            PatternKind = OutlookRecurrencePatternKind.Week,
            Start = new DateTime(2026, 7, 6, 9, 0, 0),
            Duration = TimeSpan.FromHours(1),
            DaysOfWeek = OutlookRecurrenceDays.Monday,
            RangeKind = OutlookRecurrenceRangeKind.NoEnd
        };

        OutlookRecurrenceExpansionResult result = OutlookRecurrenceExpander.Expand(
            recurrence,
            new OutlookRecurrenceExpansionOptions {
                WindowEnd = new DateTime(2026, 8, 1),
                MaxOccurrences = 10,
                MaxCandidateDays = 3
            });

        Assert.True(result.Truncated);
        Assert.Equal("MaxCandidateDays was reached.", result.TruncationReason);
        Assert.Equal(3, result.CandidateDaysInspected);
        Assert.Single(result.Occurrences);
    }

    [Fact]
    public void CountedWeeklyExpansionStopsAfterItsFinalOccurrenceWithoutReportingCandidateTruncation() {
        var recurrence = new OutlookRecurrence {
            Frequency = OutlookRecurrenceFrequency.Weekly,
            PatternKind = OutlookRecurrencePatternKind.Week,
            Start = new DateTime(2026, 7, 6, 9, 0, 0),
            Duration = TimeSpan.FromHours(1),
            DaysOfWeek = OutlookRecurrenceDays.Monday,
            RangeKind = OutlookRecurrenceRangeKind.OccurrenceCount,
            OccurrenceCount = 1
        };

        OutlookRecurrenceExpansionResult result = OutlookRecurrenceExpander.Expand(
            recurrence,
            new OutlookRecurrenceExpansionOptions {
                MaxOccurrences = 10,
                MaxCandidateDays = 1
            });

        Assert.False(result.Truncated);
        Assert.Null(result.TruncationReason);
        Assert.Equal(1, result.CandidateDaysInspected);
        Assert.Single(result.Occurrences);

        recurrence.DeletedOccurrenceDates.Add(recurrence.Start.Date);
        OutlookRecurrenceExpansionResult deleted = OutlookRecurrenceExpander.Expand(
            recurrence,
            new OutlookRecurrenceExpansionOptions { MaxOccurrences = 10, MaxCandidateDays = 1 });
        Assert.False(deleted.Truncated);
        Assert.Empty(deleted.Occurrences);

        recurrence.DeletedOccurrenceDates.Clear();
        OutlookRecurrenceExpansionResult outsideWindow = OutlookRecurrenceExpander.Expand(
            recurrence,
            new OutlookRecurrenceExpansionOptions {
                WindowStart = recurrence.Start.Date.AddDays(1),
                WindowEnd = recurrence.Start.Date.AddDays(2),
                MaxOccurrences = 10,
                MaxCandidateDays = 1
            });
        Assert.False(outsideWindow.Truncated);
        Assert.Empty(outsideWindow.Occurrences);
    }

    [Fact]
    public void ExpansionUsesHalfOpenOverlapWhileRetainingPointEventsAtWindowStart() {
        var endingAtWindowStart = new OutlookRecurrence {
            Frequency = OutlookRecurrenceFrequency.Daily,
            PatternKind = OutlookRecurrencePatternKind.Day,
            Start = new DateTime(2026, 7, 6, 9, 0, 0),
            Duration = TimeSpan.FromHours(1),
            RangeKind = OutlookRecurrenceRangeKind.OccurrenceCount,
            OccurrenceCount = 1
        };
        var pointAtWindowStart = new OutlookRecurrence {
            Frequency = OutlookRecurrenceFrequency.Daily,
            PatternKind = OutlookRecurrencePatternKind.Day,
            Start = new DateTime(2026, 7, 6, 10, 0, 0),
            Duration = TimeSpan.Zero,
            RangeKind = OutlookRecurrenceRangeKind.OccurrenceCount,
            OccurrenceCount = 1
        };
        var options = new OutlookRecurrenceExpansionOptions {
            WindowStart = new DateTime(2026, 7, 6, 10, 0, 0),
            WindowEnd = new DateTime(2026, 7, 6, 11, 0, 0)
        };

        Assert.Empty(OutlookRecurrenceExpander.Expand(endingAtWindowStart, options).Occurrences);
        Assert.Single(OutlookRecurrenceExpander.Expand(pointAtWindowStart, options).Occurrences);
    }

    [Fact]
    public void EncodingRejectsIncompatibleFrequencyAndPatternCombinations() {
        OutlookRecurrence[] invalid = {
            new OutlookRecurrence {
                Frequency = OutlookRecurrenceFrequency.Daily,
                PatternKind = OutlookRecurrencePatternKind.Week,
                Start = new DateTime(2026, 7, 6),
                DaysOfWeek = OutlookRecurrenceDays.Monday
            },
            new OutlookRecurrence {
                Frequency = OutlookRecurrenceFrequency.Weekly,
                PatternKind = OutlookRecurrencePatternKind.Day,
                Start = new DateTime(2026, 7, 6)
            },
            new OutlookRecurrence {
                Frequency = OutlookRecurrenceFrequency.Monthly,
                PatternKind = OutlookRecurrencePatternKind.Day,
                Start = new DateTime(2026, 7, 6)
            },
            new OutlookRecurrence {
                Frequency = OutlookRecurrenceFrequency.Yearly,
                PatternKind = OutlookRecurrencePatternKind.Week,
                Start = new DateTime(2026, 7, 6),
                DaysOfWeek = OutlookRecurrenceDays.Monday
            }
        };

        foreach (OutlookRecurrence recurrence in invalid) {
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => OutlookRecurrenceBinary.EncodeTask(recurrence));
            Assert.Contains("incompatible", exception.Message, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void RetainsMalformedOrUnsupportedStateWithoutPretendingItDecoded() {
        byte[] unsupported = FromHex(MicrosoftWeeklyWithoutExceptions);
        unsupported[6] = 0x05;
        unsupported[7] = 0x00;

        OutlookRecurrence recurrence = OutlookRecurrenceBinary.DecodeAppointment(unsupported);

        Assert.False(recurrence.StateDecoded);
        Assert.NotNull(recurrence.DecodeError);
        Assert.Equal(unsupported, recurrence.RawState);
        Assert.Equal(unsupported, OutlookRecurrenceBinary.EncodeAppointment(recurrence));
    }

    [Fact]
    public void RoundTripsTypedAppointmentRecurrenceThroughMsg() {
        OutlookRecurrence recurrence = OutlookRecurrenceBinary.DecodeAppointment(
            FromHex(MicrosoftWeeklyWithException));
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            OutlookCodePage = 1250,
            Subject = "Recurring appointment",
            Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2007, 3, 26, 10, 0, 0, TimeSpan.Zero),
                End = new DateTimeOffset(2007, 3, 26, 10, 30, 0, TimeSpan.Zero),
                Recurrence = recurrence
            }
        };

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument result = new EmailDocumentReader().Read(bytes).Document;

        Assert.NotNull(result.Appointment!.Recurrence);
        Assert.True(result.Appointment.Recurrence!.StateDecoded, result.Appointment.Recurrence.DecodeError);
        Assert.Equal(12, result.Appointment.Recurrence.OccurrenceCount);
        Assert.Equal("Simple Recurrence with exceptions",
            Assert.Single(result.Appointment.Recurrence.Exceptions).Subject);
        Assert.True(result.Mapi.GetNullableValue(MapiKnownProperties.PidLid.Recurring));
        Assert.Equal(result.Appointment.RecurrenceState,
            result.Mapi.GetValueOrDefault(MapiKnownProperties.PidLid.AppointmentRecur));
    }

    [Fact]
    public void RoundTripsTaskRecurrenceUsingTaskPropertySetOnly() {
        var recurrence = new OutlookRecurrence {
            Frequency = OutlookRecurrenceFrequency.Weekly,
            PatternKind = OutlookRecurrencePatternKind.Week,
            Interval = 2,
            Start = new DateTime(2026, 7, 6),
            DaysOfWeek = OutlookRecurrenceDays.Monday,
            RangeKind = OutlookRecurrenceRangeKind.OccurrenceCount,
            OccurrenceCount = 5
        };
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            Subject = "Recurring task",
            Task = new OutlookTask { Recurrence = recurrence }
        };

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument result = new EmailDocumentReader().Read(bytes).Document;

        Assert.NotNull(result.Task!.Recurrence);
        Assert.True(result.Task.Recurrence!.StateDecoded, result.Task.Recurrence.DecodeError);
        Assert.Equal(2, result.Task.Recurrence.Interval);
        Assert.Equal(5, result.Task.Recurrence.OccurrenceCount);
        Assert.True(result.Task.IsRecurring);
        Assert.True(result.Mapi.Contains(MapiKnownProperties.PidLid.TaskRecurrence));
        Assert.True(result.Mapi.Contains(MapiKnownProperties.PidLid.TaskFRecurring));
        Assert.False(result.Mapi.Contains(MapiKnownProperties.PidLid.AppointmentRecur));
        Assert.False(result.Mapi.Contains(MapiKnownProperties.PidLid.Recurring));
    }

    [Fact]
    public void ExpansionStopsAtCalendarBoundaryWithoutOverflowing() {
        var recurrence = new OutlookRecurrence {
            Frequency = OutlookRecurrenceFrequency.Daily,
            PatternKind = OutlookRecurrencePatternKind.Day,
            Start = new DateTime(9999, 12, 31, 23, 30, 0),
            Duration = TimeSpan.FromHours(2),
            RangeKind = OutlookRecurrenceRangeKind.NoEnd
        };

        OutlookRecurrenceExpansionResult result = OutlookRecurrenceExpander.Expand(recurrence,
            new OutlookRecurrenceExpansionOptions { MaxOccurrences = 10, MaxCandidateDays = 10 });

        OutlookRecurrenceOccurrence occurrence = Assert.Single(result.Occurrences);
        Assert.Equal(DateTime.MaxValue, occurrence.End);
        Assert.False(result.Truncated);
    }

    private static byte[] FromHex(string value) {
        var bytes = new byte[value.Length / 2];
        for (int index = 0; index < bytes.Length; index++)
            bytes[index] = Convert.ToByte(value.Substring(index * 2, 2), 16);
        return bytes;
    }
}
