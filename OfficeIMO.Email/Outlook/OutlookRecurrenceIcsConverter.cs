namespace OfficeIMO.Email;

/// <summary>Explicit, report-producing conversion between Outlook recurrence data and RFC 5545 RRULE data.</summary>
public static class OutlookRecurrenceIcsConverter {
    private static readonly string[] DayTokens = { "SU", "MO", "TU", "WE", "TH", "FR", "SA" };

    /// <summary>Exports one Outlook recurrence without silently discarding exceptions or time-zone constraints.</summary>
    public static OutlookRecurrenceIcsExportResult Export(OutlookRecurrence recurrence,
        OutlookRecurrenceIcsExportOptions? options = null) {
        if (recurrence == null) throw new ArgumentNullException(nameof(recurrence));
        OutlookRecurrenceIcsExportOptions effective = options ?? new OutlookRecurrenceIcsExportOptions();
        var report = new OutlookRecurrenceIcsConversionReport();
        var rule = new IcsRecurrenceRule();
        string? timeZoneId = effective.TimeZoneId ?? recurrence.TimeZoneId ?? effective.TimeZone?.KeyName;
        if (!TryWritePattern(recurrence, rule, report))
            return new OutlookRecurrenceIcsExportResult(null, Array.Empty<IcsTemporalValue>(),
                Array.Empty<OutlookRecurrenceIcsException>(), report);
        if (recurrence.Interval > 1) rule.Interval = recurrence.Interval;
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.OccurrenceCount) {
            if (!recurrence.OccurrenceCount.HasValue) Error(report, "ICAL_COUNT_REQUIRED",
                "The Outlook recurrence declares a counted range without OccurrenceCount.");
            else rule.Count = recurrence.OccurrenceCount.Value;
        } else if (recurrence.RangeKind == OutlookRecurrenceRangeKind.EndDate) {
            if (!recurrence.EndDate.HasValue) Error(report, "ICAL_UNTIL_REQUIRED",
                "The Outlook recurrence declares an end-date range without EndDate.");
            else WriteUntil(rule, recurrence, effective, timeZoneId, report);
        }
        rule.SetValue("WKST", DayTokens[(int)recurrence.FirstDayOfWeek]);

        var exceptionDates = new HashSet<DateTime>(recurrence.Exceptions.Select(value => value.OriginalStart.Date));
        IcsTemporalValue[] excluded = recurrence.DeletedOccurrenceDates.Select(value => value.Date).Distinct()
            .Where(value => !exceptionDates.Contains(value))
            .OrderBy(value => value)
            .Select(value => ToTemporal(value.Add(recurrence.Start.TimeOfDay), timeZoneId, effective.DateOnly))
            .ToArray();
        OutlookRecurrenceIcsException[] exceptions = recurrence.Exceptions.OrderBy(value => value.OriginalStart)
            .Select(value => ExportException(value, timeZoneId, effective.DateOnly, report)).ToArray();
        return new OutlookRecurrenceIcsExportResult(report.Succeeded ? rule : null, excluded, exceptions, report);
    }

    /// <summary>Imports one RRULE plus its EXDATE and exception-component data.</summary>
    public static OutlookRecurrenceIcsImportResult Import(IcsRecurrenceRule rule,
        OutlookRecurrenceIcsImportOptions options) {
        if (rule == null) throw new ArgumentNullException(nameof(rule));
        if (options == null) throw new ArgumentNullException(nameof(options));
        var report = new OutlookRecurrenceIcsConversionReport();
        DateTime start = ToLocal(options.Start, options.TimeZone, report, "DTSTART");
        var recurrence = new OutlookRecurrence { Start = start, Duration = options.Duration };
        recurrence.TimeZoneId = options.Start.TimeZoneId ?? options.TimeZone?.KeyName;
        ReadCommon(rule, recurrence, options, report);
        ReadPattern(rule, recurrence, start, report);
        if (!report.Succeeded) return new OutlookRecurrenceIcsImportResult(null, report);
        foreach (IcsTemporalValue value in options.ExcludedDates)
            recurrence.DeletedOccurrenceDates.Add(ToLocal(value, options.TimeZone, report, "EXDATE").Date);
        foreach (OutlookRecurrenceIcsException value in options.Exceptions) {
            recurrence.Exceptions.Add(new OutlookRecurrenceException {
                OriginalStart = ToLocal(value.OriginalStart, options.TimeZone, report, "RECURRENCE-ID"),
                Start = ToLocal(value.Start, options.TimeZone, report, "exception DTSTART"),
                End = ToLocal(value.End, options.TimeZone, report, "exception DTEND"),
                Subject = value.Subject,
                Location = value.Location,
                ReminderDeltaMinutes = value.ReminderDeltaMinutes,
                ReminderIsSet = value.ReminderIsSet,
                BusyStatus = value.BusyStatus,
                IsAllDay = value.IsAllDay
            });
        }
        return new OutlookRecurrenceIcsImportResult(report.Succeeded ? recurrence : null, report);
    }

    private static bool TryWritePattern(OutlookRecurrence source, IcsRecurrenceRule target,
        OutlookRecurrenceIcsConversionReport report) {
        switch (source.Frequency) {
            case OutlookRecurrenceFrequency.Daily:
                if (source.PatternKind != OutlookRecurrencePatternKind.Day) return UnsupportedCombination(report);
                target.Frequency = "DAILY";
                return true;
            case OutlookRecurrenceFrequency.Weekly:
                if (source.PatternKind != OutlookRecurrencePatternKind.Week ||
                    source.DaysOfWeek == OutlookRecurrenceDays.None) return UnsupportedCombination(report);
                target.Frequency = "WEEKLY";
                target.SetValue("BYDAY", FormatDays(source.DaysOfWeek));
                return true;
            case OutlookRecurrenceFrequency.Monthly:
                target.Frequency = "MONTHLY";
                return WriteMonthlyParts(source, target, report);
            case OutlookRecurrenceFrequency.Yearly:
                target.Frequency = "YEARLY";
                target.SetValue("BYMONTH", source.Start.Month.ToString(CultureInfo.InvariantCulture));
                return WriteMonthlyParts(source, target, report);
            default:
                Error(report, "ICAL_FREQUENCY_UNSUPPORTED", "The Outlook recurrence frequency is unsupported.");
                return false;
        }
    }

    private static bool WriteMonthlyParts(OutlookRecurrence source, IcsRecurrenceRule target,
        OutlookRecurrenceIcsConversionReport report) {
        if (source.PatternKind == OutlookRecurrencePatternKind.MonthEnd) {
            target.SetValue("BYMONTHDAY", "-1");
            return true;
        }
        if (source.PatternKind == OutlookRecurrencePatternKind.MonthDay && source.DayOfMonth.HasValue) {
            target.SetValue("BYMONTHDAY", source.DayOfMonth.Value.ToString(CultureInfo.InvariantCulture));
            return true;
        }
        if (source.PatternKind == OutlookRecurrencePatternKind.MonthNth &&
            source.DaysOfWeek != OutlookRecurrenceDays.None && source.WeekOrdinal.HasValue) {
            target.SetValue("BYDAY", FormatDays(source.DaysOfWeek));
            target.SetValue("BYSETPOS", source.WeekOrdinal == OutlookRecurrenceWeekOrdinal.Last ? "-1" :
                ((int)source.WeekOrdinal.Value).ToString(CultureInfo.InvariantCulture));
            return true;
        }
        return UnsupportedCombination(report);
    }

    private static bool UnsupportedCombination(OutlookRecurrenceIcsConversionReport report) {
        Error(report, "ICAL_PATTERN_UNSUPPORTED",
            "The Outlook frequency and pattern combination cannot be represented by the supported RRULE subset.");
        return false;
    }

    private static void WriteUntil(IcsRecurrenceRule rule, OutlookRecurrence recurrence,
        OutlookRecurrenceIcsExportOptions options, string? timeZoneId,
        OutlookRecurrenceIcsConversionReport report) {
        DateTime localUntil = recurrence.EndDate!.Value.Date.Add(recurrence.Start.TimeOfDay);
        if (options.DateOnly) {
            rule.SetValue("UNTIL", localUntil.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
            return;
        }
        if (options.TimeZone != null) {
            try {
                DateTimeOffset resolved = options.TimeZone.ResolveLocal(localUntil, options.AmbiguousTimePolicy);
                rule.SetValue("UNTIL", resolved.UtcDateTime.ToString("yyyyMMdd'T'HHmmss'Z'",
                    CultureInfo.InvariantCulture));
                return;
            } catch (InvalidOperationException exception) {
                Error(report, "ICAL_UNTIL_TIME_INVALID", exception.Message);
                return;
            }
        }
        rule.SetValue("UNTIL", localUntil.ToString("yyyyMMdd'T'HHmmss", CultureInfo.InvariantCulture));
        if (!string.IsNullOrWhiteSpace(timeZoneId)) Warning(report, "ICAL_UNTIL_TIMEZONE_UNRESOLVED",
            "TZID-local DTSTART requires a UTC UNTIL, but no Outlook time-zone rules were supplied.");
    }

    private static OutlookRecurrenceIcsException ExportException(OutlookRecurrenceException value,
        string? timeZoneId, bool dateOnly, OutlookRecurrenceIcsConversionReport report) {
        if (value.MeetingType.HasValue || value.HasAttachments.HasValue || value.AppointmentColor.HasValue ||
            value.HasExceptionalBody) Warning(report, "ICAL_EXCEPTION_EXTENSION_REQUIRED",
                "One or more Outlook-only exception fields require a non-standard extension or embedded item.");
        return new OutlookRecurrenceIcsException {
            OriginalStart = ToTemporal(value.OriginalStart, timeZoneId, dateOnly),
            Start = ToTemporal(value.Start, timeZoneId, dateOnly || value.IsAllDay == true),
            End = ToTemporal(value.End, timeZoneId, dateOnly || value.IsAllDay == true),
            Subject = value.Subject,
            Location = value.Location,
            ReminderDeltaMinutes = value.ReminderDeltaMinutes,
            ReminderIsSet = value.ReminderIsSet,
            BusyStatus = value.BusyStatus,
            IsAllDay = value.IsAllDay
        };
    }

    private static void ReadCommon(IcsRecurrenceRule rule, OutlookRecurrence recurrence,
        OutlookRecurrenceIcsImportOptions options, OutlookRecurrenceIcsConversionReport report) {
        recurrence.Interval = rule.Interval ?? 1;
        string? count = rule.GetValue("COUNT");
        string? until = rule.GetValue("UNTIL");
        if (count != null && until != null) {
            Error(report, "ICAL_COUNT_UNTIL_CONFLICT", "RRULE cannot contain both COUNT and UNTIL.");
        } else if (count != null) {
            if (!int.TryParse(count, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) || parsed <= 0)
                Error(report, "ICAL_COUNT_INVALID", "RRULE COUNT must be a positive integer.");
            else { recurrence.RangeKind = OutlookRecurrenceRangeKind.OccurrenceCount; recurrence.OccurrenceCount = parsed; }
        } else if (until != null) {
            var property = new ContentLineProperty("UNTIL", until);
            if (!IcsTemporalValue.TryParse(property, out IcsTemporalValue parsed)) {
                Error(report, "ICAL_UNTIL_INVALID", "RRULE UNTIL is not a supported DATE or DATE-TIME.");
            } else {
                recurrence.RangeKind = OutlookRecurrenceRangeKind.EndDate;
                recurrence.EndDate = ToLocal(parsed, options.TimeZone, report, "UNTIL").Date;
            }
        } else recurrence.RangeKind = OutlookRecurrenceRangeKind.NoEnd;
        string? weekStart = rule.GetValue("WKST");
        if (weekStart == null) recurrence.FirstDayOfWeek = DayOfWeek.Monday;
        else if (TryParseDay(weekStart, out DayOfWeek day)) recurrence.FirstDayOfWeek = day;
        else Error(report, "ICAL_WEEK_START_INVALID", "RRULE WKST is not a weekday token.");

        var allowed = new HashSet<string>(new[] { "FREQ", "INTERVAL", "COUNT", "UNTIL", "WKST", "BYDAY",
            "BYMONTHDAY", "BYSETPOS", "BYMONTH" }, StringComparer.OrdinalIgnoreCase);
        foreach (IcsRecurrencePart part in rule.Parts.Where(part => !allowed.Contains(part.Name)))
            Warning(report, "ICAL_RRULE_PART_UNSUPPORTED", "RRULE part " + part.Name + " is not represented by Outlook.");
    }

    private static void ReadPattern(IcsRecurrenceRule rule, OutlookRecurrence recurrence, DateTime start,
        OutlookRecurrenceIcsConversionReport report) {
        string frequency = rule.Frequency?.ToUpperInvariant() ?? string.Empty;
        string? byDay = rule.GetValue("BYDAY");
        string? byMonthDay = rule.GetValue("BYMONTHDAY");
        string? bySetPosition = rule.GetValue("BYSETPOS");
        if (frequency == "DAILY") {
            recurrence.Frequency = OutlookRecurrenceFrequency.Daily;
            recurrence.PatternKind = OutlookRecurrencePatternKind.Day;
            if (byDay != null || byMonthDay != null || bySetPosition != null)
                Error(report, "ICAL_DAILY_FILTER_UNSUPPORTED", "Filtered daily RRULE values are not an Outlook daily pattern.");
            return;
        }
        if (frequency == "WEEKLY") {
            recurrence.Frequency = OutlookRecurrenceFrequency.Weekly;
            recurrence.PatternKind = OutlookRecurrencePatternKind.Week;
            recurrence.DaysOfWeek = byDay == null ? (OutlookRecurrenceDays)(1 << (int)start.DayOfWeek) :
                ParseDays(byDay, allowOrdinal: false, report);
            if (byMonthDay != null || bySetPosition != null)
                Error(report, "ICAL_WEEKLY_FILTER_UNSUPPORTED", "The weekly RRULE contains unsupported filters.");
            return;
        }
        if (frequency != "MONTHLY" && frequency != "YEARLY") {
            Error(report, "ICAL_FREQUENCY_UNSUPPORTED", "RRULE FREQ is not supported by Outlook recurrence.");
            return;
        }
        recurrence.Frequency = frequency == "YEARLY" ? OutlookRecurrenceFrequency.Yearly :
            OutlookRecurrenceFrequency.Monthly;
        if (frequency == "YEARLY") {
            string? byMonth = rule.GetValue("BYMONTH");
            if (byMonth != null && (!int.TryParse(byMonth, NumberStyles.None, CultureInfo.InvariantCulture,
                out int month) || month != start.Month))
                Error(report, "ICAL_YEAR_MONTH_UNSUPPORTED",
                    "A yearly RRULE BYMONTH must match the DTSTART month for Outlook recurrence.");
        }
        if (byMonthDay != null && byDay == null && bySetPosition == null) {
            if (!int.TryParse(byMonthDay, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture,
                out int day) || day == 0 || day < -1 || day > 31)
                Error(report, "ICAL_MONTH_DAY_INVALID", "RRULE BYMONTHDAY is not representable by Outlook.");
            else if (day == -1) recurrence.PatternKind = OutlookRecurrencePatternKind.MonthEnd;
            else { recurrence.PatternKind = OutlookRecurrencePatternKind.MonthDay; recurrence.DayOfMonth = day; }
            return;
        }
        if (byDay != null && bySetPosition != null && byMonthDay == null) {
            recurrence.PatternKind = OutlookRecurrencePatternKind.MonthNth;
            recurrence.DaysOfWeek = ParseDays(byDay, allowOrdinal: false, report);
            if (!int.TryParse(bySetPosition, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture,
                out int position) || position == 0 || position < -1 || position > 4)
                Error(report, "ICAL_SET_POSITION_INVALID", "RRULE BYSETPOS must be 1 through 4 or -1 for Outlook.");
            else recurrence.WeekOrdinal = position == -1 ? OutlookRecurrenceWeekOrdinal.Last :
                (OutlookRecurrenceWeekOrdinal)position;
            return;
        }
        if (byMonthDay == null && byDay == null && bySetPosition == null) {
            recurrence.PatternKind = OutlookRecurrencePatternKind.MonthDay;
            recurrence.DayOfMonth = start.Day;
            return;
        }
        Error(report, "ICAL_MONTH_PATTERN_UNSUPPORTED", "The monthly/yearly RRULE filter combination is unsupported.");
    }

    private static OutlookRecurrenceDays ParseDays(string value, bool allowOrdinal,
        OutlookRecurrenceIcsConversionReport report) {
        OutlookRecurrenceDays result = OutlookRecurrenceDays.None;
        foreach (string token in value.Split(',')) {
            string dayToken = token;
            if (token.Length > 2) {
                if (!allowOrdinal) { Error(report, "ICAL_BYDAY_ORDINAL_UNSUPPORTED", "Ordinal BYDAY values require BYSETPOS for Outlook."); continue; }
                dayToken = token.Substring(token.Length - 2);
            }
            if (!TryParseDay(dayToken, out DayOfWeek day))
                Error(report, "ICAL_BYDAY_INVALID", "RRULE BYDAY contains an invalid weekday token.");
            else result |= (OutlookRecurrenceDays)(1 << (int)day);
        }
        if (result == OutlookRecurrenceDays.None) Error(report, "ICAL_BYDAY_REQUIRED", "The RRULE requires at least one weekday.");
        return result;
    }

    private static string FormatDays(OutlookRecurrenceDays days) => string.Join(",",
        Enumerable.Range(0, 7).Where(index => (days & (OutlookRecurrenceDays)(1 << index)) != 0)
            .Select(index => DayTokens[index]));

    private static bool TryParseDay(string value, out DayOfWeek day) {
        for (int index = 0; index < DayTokens.Length; index++) {
            if (!string.Equals(value, DayTokens[index], StringComparison.OrdinalIgnoreCase)) continue;
            day = (DayOfWeek)index;
            return true;
        }
        day = default;
        return false;
    }

    private static IcsTemporalValue ToTemporal(DateTime local, string? timeZoneId, bool dateOnly) {
        if (dateOnly) return IcsTemporalValue.Date(local);
        return string.IsNullOrWhiteSpace(timeZoneId) ? IcsTemporalValue.Floating(local) :
            IcsTemporalValue.Zoned(local, timeZoneId!);
    }

    private static DateTime ToLocal(IcsTemporalValue value, OutlookTimeZoneDefinition? timeZone,
        OutlookRecurrenceIcsConversionReport report, string field) {
        if (value.Kind == IcsTemporalValueKind.UtcDateTime) {
            if (timeZone == null) {
                Warning(report, "ICAL_UTC_TIMEZONE_UNRESOLVED", field + " is UTC but no Outlook time-zone rules were supplied.");
                return DateTime.SpecifyKind(value.Value, DateTimeKind.Unspecified);
            }
            return DateTime.SpecifyKind(timeZone.ConvertUtc(new DateTimeOffset(value.Value, TimeSpan.Zero)).DateTime,
                DateTimeKind.Unspecified);
        }
        if (value.Kind == IcsTemporalValueKind.ZonedDateTime && timeZone != null &&
            !string.Equals(value.TimeZoneId, timeZone.KeyName, StringComparison.OrdinalIgnoreCase))
            Warning(report, "ICAL_TIMEZONE_ID_MISMATCH", field + " TZID does not match the supplied Outlook definition key.");
        return DateTime.SpecifyKind(value.Value, DateTimeKind.Unspecified);
    }

    private static void Error(OutlookRecurrenceIcsConversionReport report, string code, string message) =>
        report.Add(code, message, OutlookRecurrenceIcsIssueSeverity.Error);
    private static void Warning(OutlookRecurrenceIcsConversionReport report, string code, string message) =>
        report.Add(code, message, OutlookRecurrenceIcsIssueSeverity.Warning);
}
