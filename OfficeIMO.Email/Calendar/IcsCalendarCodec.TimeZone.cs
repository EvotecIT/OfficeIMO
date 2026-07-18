namespace OfficeIMO.Email;

internal static partial class IcsCalendarCodec {
    private static OutlookTimeZoneDefinition? ResolveEmbeddedRecurrenceTimeZone(IcsDocument calendar,
        IcsTemporalValue start, IList<EmailDiagnostic> diagnostics, string location, EmailDocument document) {
        if (start.Kind != IcsTemporalValueKind.ZonedDateTime || string.IsNullOrWhiteSpace(start.TimeZoneId))
            return null;
        ContentLineComponent[] matches = calendar.GetComponents("VTIMEZONE").Where(component =>
            string.Equals(component.GetFirstProperty("TZID")?.Value, start.TimeZoneId,
                StringComparison.OrdinalIgnoreCase)).ToArray();
        if (matches.Length == 0) return null;
        if (matches.Length != 1) {
            ReportTimeZoneProjectionFailure("The recurrence TZID has more than one matching VTIMEZONE definition.",
                diagnostics, location, document);
            return null;
        }
        if (TryConvertTimeZone(matches[0], out OutlookTimeZoneDefinition? definition, out string? error))
            return definition;
        ReportTimeZoneProjectionFailure(error ?? "The embedded VTIMEZONE cannot be represented by Outlook rules.",
            diagnostics, location, document);
        return null;
    }

    private static void ReportTimeZoneProjectionFailure(string message, IList<EmailDiagnostic> diagnostics,
        string location, EmailDocument document) {
        diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_TIMEZONE_PROJECTION_UNSUPPORTED", message,
            EmailDiagnosticSeverity.Warning, location));
        document.MimeSemanticProjectionIsIncomplete = true;
    }

    private static bool TryConvertTimeZone(ContentLineComponent component,
        out OutlookTimeZoneDefinition? definition, out string? error) {
        definition = null;
        error = null;
        string? timeZoneId = component.GetFirstProperty("TZID")?.Value;
        if (string.IsNullOrWhiteSpace(timeZoneId)) return Fail("VTIMEZONE requires one non-empty TZID.", out error);
        Observance[] observances;
        try {
            observances = component.Components.Where(child =>
                    string.Equals(child.Name, "STANDARD", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(child.Name, "DAYLIGHT", StringComparison.OrdinalIgnoreCase))
                .Select(ParseObservance).ToArray();
        } catch (Exception exception) when (exception is InvalidDataException || exception is FormatException ||
            exception is ArgumentException || exception is OverflowException) {
            return Fail(exception.Message, out error);
        }
        if (observances.Length == 0) return Fail("VTIMEZONE does not contain an Outlook-representable observance.", out error);

        var result = new OutlookTimeZoneDefinition { KeyName = timeZoneId };
        IGrouping<int, Observance>[] groups = observances.GroupBy(value => value.Start.Year)
            .OrderBy(value => value.Key).ToArray();
        for (int groupIndex = 0; groupIndex < groups.Length; groupIndex++) {
            IGrouping<int, Observance> group = groups[groupIndex];
            Observance[] standards = group.Where(value => !value.IsDaylight).ToArray();
            Observance[] daylights = group.Where(value => value.IsDaylight).ToArray();
            if (standards.Length != 1 || daylights.Length > 1)
                return Fail("Each VTIMEZONE rule period requires exactly one STANDARD and at most one DAYLIGHT observance.",
                    out error);
            Observance standard = standards[0];
            OutlookTimeZoneTransition standardTransition = OutlookTimeZoneRule.DisabledTransition;
            OutlookTimeZoneTransition daylightTransition = OutlookTimeZoneRule.DisabledTransition;
            TimeSpan daylightOffset = standard.OffsetTo;
            if (daylights.Length == 0) {
                if (standard.OffsetFrom != standard.OffsetTo || standard.Rule != null)
                    return Fail("A fixed-offset STANDARD observance must not declare a transition or changing offset.",
                        out error);
            } else {
                Observance daylight = daylights[0];
                bool hasUntil = standard.Rule?.GetValue("UNTIL") != null ||
                    daylight.Rule?.GetValue("UNTIL") != null;
                if (hasUntil && groupIndex == groups.Length - 1)
                    return Fail("The final VTIMEZONE rule period expires without a successor rule.", out error);
                if ((standard.Rule?.GetValue("UNTIL") == null) !=
                    (daylight.Rule?.GetValue("UNTIL") == null))
                    return Fail("Paired STANDARD and DAYLIGHT rules must use consistent UNTIL bounds.", out error);
                if (standard.OffsetFrom != daylight.OffsetTo || daylight.OffsetFrom != standard.OffsetTo)
                    return Fail("STANDARD and DAYLIGHT offsets do not form one reciprocal Outlook rule.", out error);
                if (!TryCreateTransition(standard, out standardTransition, out error) ||
                    !TryCreateTransition(daylight, out daylightTransition, out error)) return false;
                daylightOffset = daylight.OffsetTo;
            }
            int bias = checked(-(int)standard.OffsetTo.TotalMinutes);
            result.Rules.Add(new OutlookTimeZoneRule {
                EffectiveYear = checked((ushort)group.Key),
                BiasMinutes = bias,
                StandardBiasMinutes = 0,
                DaylightBiasMinutes = checked(-(int)daylightOffset.TotalMinutes - bias),
                StandardTransition = standardTransition,
                DaylightTransition = daylightTransition
            });
        }
        for (int index = 0; index < result.Rules.Count; index++)
            result.Rules[index].Flags = index == result.Rules.Count - 1 ? (ushort)0x0002 : (ushort)0;
        definition = result;
        return true;
    }

    private static Observance ParseObservance(ContentLineComponent component) {
        IcsTemporalValue? temporal = component.GetTemporalValue("DTSTART");
        if (!temporal.HasValue || temporal.Value.Kind != IcsTemporalValueKind.FloatingDateTime)
            throw new InvalidDataException("A VTIMEZONE observance requires a floating DTSTART.");
        TimeSpan offsetFrom = ParseUtcOffset(component.GetFirstProperty("TZOFFSETFROM")?.Value);
        TimeSpan offsetTo = ParseUtcOffset(component.GetFirstProperty("TZOFFSETTO")?.Value);
        ContentLineProperty[] rules = component.GetProperties("RRULE").ToArray();
        if (rules.Length > 1) throw new InvalidDataException("A VTIMEZONE observance cannot contain multiple RRULE values.");
        return new Observance(
            string.Equals(component.Name, "DAYLIGHT", StringComparison.OrdinalIgnoreCase),
            temporal.Value.Value, offsetFrom, offsetTo,
            rules.Length == 0 ? null : IcsRecurrenceRule.Parse(rules[0].Value));
    }

    private static TimeSpan ParseUtcOffset(string? value) {
        string text = value?.Trim() ?? string.Empty;
        if ((text.Length != 5 && text.Length != 7) || (text[0] != '+' && text[0] != '-'))
            throw new InvalidDataException("VTIMEZONE contains an invalid UTC offset.");
        if (!int.TryParse(text.Substring(1, 2), NumberStyles.None, CultureInfo.InvariantCulture, out int hours) ||
            !int.TryParse(text.Substring(3, 2), NumberStyles.None, CultureInfo.InvariantCulture, out int minutes) ||
            hours > 23 || minutes > 59)
            throw new InvalidDataException("VTIMEZONE contains an invalid UTC offset.");
        int seconds = 0;
        if (text.Length == 7 && (!int.TryParse(text.Substring(5, 2), NumberStyles.None,
                CultureInfo.InvariantCulture, out seconds) || seconds > 59))
            throw new InvalidDataException("VTIMEZONE contains an invalid UTC offset.");
        if (seconds != 0)
            throw new InvalidDataException("Outlook time-zone rules cannot represent sub-minute UTC offsets.");
        int totalMinutes = checked(hours * 60 + minutes);
        return TimeSpan.FromMinutes(text[0] == '-' ? -totalMinutes : totalMinutes);
    }

    private static bool TryCreateTransition(Observance observance,
        out OutlookTimeZoneTransition transition, out string? error) {
        transition = OutlookTimeZoneRule.DisabledTransition;
        error = null;
        if (observance.Rule == null ||
            !string.Equals(observance.Rule.Frequency, "YEARLY", StringComparison.OrdinalIgnoreCase))
            return Fail("Daylight VTIMEZONE observances require a yearly RRULE for Outlook projection.", out error);
        var allowed = new HashSet<string>(new[] { "FREQ", "BYMONTH", "BYDAY", "UNTIL" },
            StringComparer.OrdinalIgnoreCase);
        if (observance.Rule.Parts.Any(part => !allowed.Contains(part.Name)))
            return Fail("The VTIMEZONE RRULE contains a part that Outlook transition rules cannot represent.", out error);
        if (!int.TryParse(observance.Rule.GetValue("BYMONTH"), NumberStyles.None,
                CultureInfo.InvariantCulture, out int month) || month < 1 || month > 12)
            return Fail("The VTIMEZONE RRULE requires one valid BYMONTH value.", out error);
        string byDay = observance.Rule.GetValue("BYDAY") ?? string.Empty;
        if (!TryParseOrdinalDay(byDay, out int ordinal, out DayOfWeek dayOfWeek))
            return Fail("The VTIMEZONE RRULE requires one Outlook-representable ordinal BYDAY value.", out error);
        transition = new OutlookTimeZoneTransition(0, checked((ushort)month), checked((ushort)dayOfWeek),
            checked((ushort)(ordinal == -1 ? 5 : ordinal)), checked((ushort)observance.Start.Hour),
            checked((ushort)observance.Start.Minute), checked((ushort)observance.Start.Second),
            checked((ushort)observance.Start.Millisecond));
        if (transition.GetDateTime(observance.Start.Year)?.Date != observance.Start.Date)
            return Fail("The VTIMEZONE DTSTART does not match its yearly transition rule.", out error);
        return true;
    }

    private static bool TryParseOrdinalDay(string value, out int ordinal, out DayOfWeek dayOfWeek) {
        ordinal = 0;
        dayOfWeek = default;
        if (value.Length < 3) return false;
        string day = value.Substring(value.Length - 2).ToUpperInvariant();
        string[] days = { "SU", "MO", "TU", "WE", "TH", "FR", "SA" };
        int dayIndex = Array.IndexOf(days, day);
        if (dayIndex < 0 || !int.TryParse(value.Substring(0, value.Length - 2),
                NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out ordinal) ||
            ordinal != -1 && (ordinal < 1 || ordinal > 4)) return false;
        dayOfWeek = (DayOfWeek)dayIndex;
        return true;
    }

    private static bool Fail(string message, out string? error) {
        error = message;
        return false;
    }

    private sealed class Observance {
        internal Observance(bool isDaylight, DateTime start, TimeSpan offsetFrom, TimeSpan offsetTo,
            IcsRecurrenceRule? rule) {
            IsDaylight = isDaylight;
            Start = start;
            OffsetFrom = offsetFrom;
            OffsetTo = offsetTo;
            Rule = rule;
        }
        internal bool IsDaylight { get; }
        internal DateTime Start { get; }
        internal TimeSpan OffsetFrom { get; }
        internal TimeSpan OffsetTo { get; }
        internal IcsRecurrenceRule? Rule { get; }
    }
}
