using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookTimeZoneTests {
    // MS-OXOCAL 4.1.5, Pacific Time PidLidTimeZoneStruct.
    private const string MicrosoftPacificStructure =
        "E001000000000000C4FFFFFF000000000B00000001000200000000000000000000000300000002000200000000000000";

    // MS-OXOCAL exception example, recurring-series TZDEFINITION with 2006 and 2007 Pacific rules.
    private const string MicrosoftPacificRecurrenceDefinition =
        "0201300002001500500061006300690066006900630020005300740061006E0064006100720064002000540069006D00" +
        "6500020002013E000000D6070000000000000000000000000000E001000000000000C4FFFFFF00000A00000005000200" +
        "0000000000000000040000000100020000000000000002013E000300D7070000000000000000000000000000E00100" +
        "0000000000C4FFFFFF00000B0000000100020000000000000000000300000002000200000000000000";

    [Fact]
    public void DecodesAndPreservesMicrosoftLegacyStructure() {
        byte[] bytes = FromHex(MicrosoftPacificStructure);

        OutlookTimeZoneStructure structure = OutlookTimeZoneBinary.DecodeStructure(bytes);

        Assert.True(structure.StateDecoded, structure.DecodeError);
        Assert.Equal(480, structure.Rule.BiasMinutes);
        Assert.Equal(0, structure.Rule.StandardBiasMinutes);
        Assert.Equal(-60, structure.Rule.DaylightBiasMinutes);
        Assert.Equal(11, structure.Rule.StandardTransition.Month);
        Assert.Equal(1, structure.Rule.StandardTransition.Day);
        Assert.Equal(3, structure.Rule.DaylightTransition.Month);
        Assert.Equal(2, structure.Rule.DaylightTransition.Day);
        Assert.Equal(bytes, OutlookTimeZoneBinary.EncodeStructure(structure));
    }

    [Fact]
    public void DecodesAndPreservesMicrosoftRecurrenceDefinition() {
        byte[] bytes = FromHex(MicrosoftPacificRecurrenceDefinition);

        OutlookTimeZoneDefinition definition = OutlookTimeZoneBinary.DecodeDefinition(bytes);

        Assert.True(definition.StateDecoded, definition.DecodeError);
        Assert.Equal("Pacific Standard Time", definition.KeyName);
        Assert.Equal(2, definition.Rules.Count);
        Assert.Equal((ushort)2006, definition.Rules[0].EffectiveYear);
        Assert.Equal((ushort)0, definition.Rules[0].Flags);
        Assert.Equal((ushort)2007, definition.Rules[1].EffectiveYear);
        Assert.Equal((ushort)3, definition.Rules[1].Flags);
        Assert.Same(definition.Rules[1], definition.GetRule(2026));
        Assert.Equal(bytes, OutlookTimeZoneBinary.EncodeDefinition(definition));
    }

    [Fact]
    public void ClassifiesPacificInvalidAmbiguousAndNormalLocalTimes() {
        OutlookTimeZoneDefinition definition = OutlookTimeZoneBinary.DecodeDefinition(
            FromHex(MicrosoftPacificRecurrenceDefinition));

        OutlookLocalTimeResolution invalid = definition.GetLocalTimeResolution(
            new DateTime(2007, 3, 11, 2, 30, 0));
        OutlookLocalTimeResolution ambiguous = definition.GetLocalTimeResolution(
            new DateTime(2007, 11, 4, 1, 30, 0));

        Assert.Equal(OutlookLocalTimeStatus.Invalid, invalid.Status);
        Assert.Empty(invalid.Offsets);
        Assert.Throws<InvalidOperationException>(() => invalid.Resolve());
        Assert.Equal(OutlookLocalTimeStatus.Ambiguous, ambiguous.Status);
        Assert.Equal(2, ambiguous.Offsets.Count);
        Assert.Equal(TimeSpan.FromHours(-7),
            ambiguous.Resolve(OutlookAmbiguousTimePolicy.EarlierUtc).Offset);
        Assert.Equal(TimeSpan.FromHours(-8),
            ambiguous.Resolve(OutlookAmbiguousTimePolicy.LaterUtc).Offset);
        Assert.Equal(TimeSpan.FromHours(-8), definition.ResolveLocal(new DateTime(2007, 1, 15, 9, 0, 0)).Offset);
        Assert.Equal(TimeSpan.FromHours(-7), definition.ResolveLocal(new DateTime(2007, 7, 15, 9, 0, 0)).Offset);
    }

    [Fact]
    public void EditedDefinitionUsesCanonicalEncoding() {
        OutlookTimeZoneDefinition definition = OutlookTimeZoneBinary.DecodeDefinition(
            FromHex(MicrosoftPacificRecurrenceDefinition));
        definition.KeyName = "Pacific Test Time";

        byte[] updated = OutlookTimeZoneBinary.EncodeDefinition(definition);
        OutlookTimeZoneDefinition result = OutlookTimeZoneBinary.DecodeDefinition(updated);

        Assert.True(result.StateDecoded, result.DecodeError);
        Assert.Equal("Pacific Test Time", result.KeyName);
        Assert.Equal(2, result.Rules.Count);
        Assert.NotEqual(FromHex(MicrosoftPacificRecurrenceDefinition), updated);
    }

    [Fact]
    public void RetainsMalformedDefinitionWithoutInventingRules() {
        byte[] bytes = FromHex(MicrosoftPacificRecurrenceDefinition);
        bytes[0] = 3;

        OutlookTimeZoneDefinition definition = OutlookTimeZoneBinary.DecodeDefinition(bytes);

        Assert.False(definition.StateDecoded);
        Assert.Empty(definition.Rules);
        Assert.Equal(bytes, definition.RawState);
        Assert.Equal(bytes, OutlookTimeZoneBinary.EncodeDefinition(definition));
    }

    [Fact]
    public void RoundTripsTypedTimeZonesThroughMsg() {
        byte[] legacyBytes = FromHex(MicrosoftPacificStructure);
        byte[] definitionBytes = FromHex(MicrosoftPacificRecurrenceDefinition);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Appointment,
            Subject = "Pacific appointment",
            Appointment = new OutlookAppointment {
                Start = new DateTimeOffset(2007, 7, 15, 9, 0, 0, TimeSpan.FromHours(-7)),
                End = new DateTimeOffset(2007, 7, 15, 10, 0, 0, TimeSpan.FromHours(-7)),
                TimeZoneDescription = "Pacific Standard Time",
                LegacyTimeZone = OutlookTimeZoneBinary.DecodeStructure(legacyBytes),
                StartTimeZone = OutlookTimeZoneBinary.DecodeDefinition(definitionBytes),
                EndTimeZone = OutlookTimeZoneBinary.DecodeDefinition(definitionBytes),
                RecurrenceTimeZone = OutlookTimeZoneBinary.DecodeDefinition(definitionBytes)
            }
        };

        byte[] msg = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument result = new EmailDocumentReader().Read(msg).Document;

        Assert.True(result.Appointment!.LegacyTimeZone!.StateDecoded,
            result.Appointment.LegacyTimeZone.DecodeError);
        Assert.Equal("Pacific Standard Time", result.Appointment.StartTimeZone!.KeyName);
        Assert.Equal("Pacific Standard Time", result.Appointment.EndTimeZone!.KeyName);
        Assert.Equal("Pacific Standard Time", result.Appointment.RecurrenceTimeZone!.KeyName);
        Assert.Equal(legacyBytes, result.Appointment.TimeZoneStructure);
        Assert.Equal(definitionBytes, result.Appointment.StartTimeZoneDefinition);
        Assert.Equal(definitionBytes, result.Appointment.EndTimeZoneDefinition);
        Assert.Equal(definitionBytes, result.Appointment.RecurrenceTimeZoneDefinition);
    }

    [Fact]
    public void ReportsLegacyDefinitionConsistencyAndFieldMismatch() {
        OutlookTimeZoneStructure legacy = OutlookTimeZoneBinary.DecodeStructure(
            FromHex(MicrosoftPacificStructure));
        OutlookTimeZoneDefinition definition = OutlookTimeZoneBinary.DecodeDefinition(
            FromHex(MicrosoftPacificRecurrenceDefinition));

        OutlookTimeZoneConsistencyReport consistent = OutlookTimeZoneConsistency.Compare(legacy, definition, 2007);
        definition.Rules[1].DaylightBiasMinutes = -30;
        OutlookTimeZoneConsistencyReport mismatch = OutlookTimeZoneConsistency.Compare(legacy, definition, 2007);

        Assert.True(consistent.IsConsistent);
        Assert.Equal(OutlookTimeZoneConsistencyStatus.Inconsistent, mismatch.Status);
        Assert.Contains(mismatch.Issues, issue => issue.Field == "DaylightBiasMinutes");
    }

    [Fact]
    public void ExpandedOccurrenceResolvesThroughEmbeddedRules() {
        OutlookTimeZoneDefinition definition = OutlookTimeZoneBinary.DecodeDefinition(
            FromHex(MicrosoftPacificRecurrenceDefinition));
        var recurrence = new OutlookRecurrence {
            Frequency = OutlookRecurrenceFrequency.Daily,
            PatternKind = OutlookRecurrencePatternKind.Day,
            Start = new DateTime(2007, 7, 15, 9, 0, 0),
            Duration = TimeSpan.FromHours(1),
            RangeKind = OutlookRecurrenceRangeKind.OccurrenceCount,
            OccurrenceCount = 1
        };

        OutlookRecurrenceOccurrence occurrence = Assert.Single(
            OutlookRecurrenceExpander.Expand(recurrence).Occurrences);

        Assert.Equal(TimeSpan.FromHours(-7), occurrence.ResolveStart(definition).Offset);
        Assert.Equal(new DateTime(2007, 7, 15, 16, 0, 0),
            occurrence.ResolveStart(definition).UtcDateTime);
    }

    [Theory]
    [InlineData(2020, 12, 31, 23, 30, 300, 360, 2020, 12, 31, 18, 30, -5)]
    [InlineData(2021, 1, 1, 5, 30, 300, 360, 2020, 12, 31, 23, 30, -6)]
    [InlineData(2020, 12, 31, 23, 30, 360, 300, 2020, 12, 31, 17, 30, -6)]
    [InlineData(2021, 1, 1, 5, 30, 360, 300, 2021, 1, 1, 0, 30, -5)]
    public void ConvertUtcUsesExplicitRuleIntervalsAcrossNewYear(int inputYear, int inputMonth,
        int inputDay, int inputHour, int inputMinute, int oldBias, int newBias,
        int expectedYear, int expectedMonth, int expectedDay, int expectedHour, int expectedMinute,
        int expectedOffsetHours) {
        var definition = new OutlookTimeZoneDefinition();
        definition.Rules.Add(new OutlookTimeZoneRule {
            EffectiveYear = 2020,
            BiasMinutes = oldBias
        });
        definition.Rules.Add(new OutlookTimeZoneRule {
            EffectiveYear = 2021,
            BiasMinutes = newBias
        });

        DateTimeOffset result = definition.ConvertUtc(
            new DateTimeOffset(inputYear, inputMonth, inputDay,
                inputHour, inputMinute, 0, TimeSpan.Zero));

        Assert.Equal(new DateTime(expectedYear, expectedMonth, expectedDay,
            expectedHour, expectedMinute, 0), result.DateTime);
        Assert.Equal(TimeSpan.FromHours(expectedOffsetHours), result.Offset);
    }

    [Fact]
    public void ConvertUtcUsesHistoricalDaylightTransitionsWithinEachRuleInterval() {
        OutlookTimeZoneDefinition definition = OutlookTimeZoneBinary.DecodeDefinition(
            FromHex(MicrosoftPacificRecurrenceDefinition));

        DateTimeOffset beforeDstLawChange = definition.ConvertUtc(
            new DateTimeOffset(2006, 3, 20, 12, 0, 0, TimeSpan.Zero));
        DateTimeOffset afterDstLawChange = definition.ConvertUtc(
            new DateTimeOffset(2007, 3, 20, 12, 0, 0, TimeSpan.Zero));

        Assert.Equal(TimeSpan.FromHours(-8), beforeDstLawChange.Offset);
        Assert.Equal(TimeSpan.FromHours(-7), afterDstLawChange.Offset);
    }

    [Fact]
    public void ConvertUtcIgnoresUnrepresentableFutureEffectiveYears() {
        var definition = new OutlookTimeZoneDefinition();
        definition.Rules.Add(new OutlookTimeZoneRule {
            EffectiveYear = 2020,
            BiasMinutes = 300
        });
        definition.Rules.Add(new OutlookTimeZoneRule {
            EffectiveYear = 10000,
            BiasMinutes = 360
        });

        DateTimeOffset result = definition.ConvertUtc(
            new DateTimeOffset(2021, 1, 1, 12, 0, 0, TimeSpan.Zero));

        Assert.Equal(new DateTime(2021, 1, 1, 7, 0, 0), result.DateTime);
        Assert.Equal(TimeSpan.FromHours(-5), result.Offset);
    }

    private static byte[] FromHex(string value) {
        var bytes = new byte[value.Length / 2];
        for (int index = 0; index < bytes.Length; index++)
            bytes[index] = Convert.ToByte(value.Substring(index * 2, 2), 16);
        return bytes;
    }
}
