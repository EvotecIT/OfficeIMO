namespace OfficeIMO.Email;

/// <summary>Safety bounds and local-time window for recurrence expansion.</summary>
public sealed class OutlookRecurrenceExpansionOptions {
    /// <summary>Inclusive local-clock window start.</summary>
    public DateTime? WindowStart { get; set; }
    /// <summary>Exclusive local-clock window end.</summary>
    public DateTime? WindowEnd { get; set; }
    /// <summary>Maximum number of returned occurrences.</summary>
    public int MaxOccurrences { get; set; } = 1000;
    /// <summary>Maximum number of local calendar dates inspected.</summary>
    public int MaxCandidateDays { get; set; } = 1000000;
}

/// <summary>One expanded recurrence occurrence.</summary>
public sealed class OutlookRecurrenceOccurrence {
    internal OutlookRecurrenceOccurrence(int sequence, DateTime originalStart, DateTime start, DateTime end,
        OutlookRecurrenceException? exception) {
        Sequence = sequence;
        OriginalStart = originalStart;
        Start = start;
        End = end;
        Exception = exception;
    }

    /// <summary>One-based position in the base series, including deleted occurrences.</summary>
    public int Sequence { get; }
    /// <summary>Original local start before exception changes.</summary>
    public DateTime OriginalStart { get; }
    /// <summary>Effective local start.</summary>
    public DateTime Start { get; }
    /// <summary>Effective local end.</summary>
    public DateTime End { get; }
    /// <summary>Exception that produced the effective values.</summary>
    public OutlookRecurrenceException? Exception { get; }
    /// <summary>Whether this occurrence is modified.</summary>
    public bool IsException => Exception != null;

    /// <summary>Resolves the local start with a caller-selected time zone.</summary>
    public DateTimeOffset ResolveStart(TimeZoneInfo timeZone) => Resolve(Start, timeZone);
    /// <summary>Resolves the local end with a caller-selected time zone.</summary>
    public DateTimeOffset ResolveEnd(TimeZoneInfo timeZone) => Resolve(End, timeZone);

    /// <summary>Resolves the local start using embedded Outlook rules and an explicit ambiguity policy.</summary>
    public DateTimeOffset ResolveStart(OutlookTimeZoneDefinition timeZone,
        OutlookAmbiguousTimePolicy policy = OutlookAmbiguousTimePolicy.EarlierUtc) {
        if (timeZone == null) throw new ArgumentNullException(nameof(timeZone));
        return timeZone.ResolveLocal(Start, policy);
    }

    /// <summary>Resolves the local end using embedded Outlook rules and an explicit ambiguity policy.</summary>
    public DateTimeOffset ResolveEnd(OutlookTimeZoneDefinition timeZone,
        OutlookAmbiguousTimePolicy policy = OutlookAmbiguousTimePolicy.EarlierUtc) {
        if (timeZone == null) throw new ArgumentNullException(nameof(timeZone));
        return timeZone.ResolveLocal(End, policy);
    }

    private static DateTimeOffset Resolve(DateTime value, TimeZoneInfo timeZone) {
        if (timeZone == null) throw new ArgumentNullException(nameof(timeZone));
        DateTime local = DateTime.SpecifyKind(value, DateTimeKind.Unspecified);
        if (timeZone.IsInvalidTime(local)) {
            throw new ArgumentException("The occurrence falls in an invalid local time for the supplied time zone.",
                nameof(timeZone));
        }
        return new DateTimeOffset(local, timeZone.GetUtcOffset(local));
    }
}

/// <summary>Bounded recurrence expansion result.</summary>
public sealed class OutlookRecurrenceExpansionResult {
    internal OutlookRecurrenceExpansionResult(IReadOnlyList<OutlookRecurrenceOccurrence> occurrences,
        int candidateDaysInspected, bool truncated, string? truncationReason) {
        Occurrences = occurrences;
        CandidateDaysInspected = candidateDaysInspected;
        Truncated = truncated;
        TruncationReason = truncationReason;
    }

    /// <summary>Occurrences inside the requested local-time window.</summary>
    public IReadOnlyList<OutlookRecurrenceOccurrence> Occurrences { get; }
    /// <summary>Number of calendar dates inspected.</summary>
    public int CandidateDaysInspected { get; }
    /// <summary>Whether a configured safety bound stopped expansion.</summary>
    public bool Truncated { get; }
    /// <summary>Human-readable safety bound that stopped expansion.</summary>
    public string? TruncationReason { get; }
}

/// <summary>Deterministic, bounded expansion for supported Gregorian Outlook recurrences.</summary>
public static class OutlookRecurrenceExpander {
    /// <summary>Expands a recurrence in local-clock space without silently choosing a host time zone.</summary>
    public static OutlookRecurrenceExpansionResult Expand(OutlookRecurrence recurrence,
        OutlookRecurrenceExpansionOptions? options = null) {
        if (recurrence == null) throw new ArgumentNullException(nameof(recurrence));
        OutlookRecurrenceExpansionOptions effective = options ?? new OutlookRecurrenceExpansionOptions();
        Validate(recurrence, effective);

        DateTime windowStart = AsLocal(effective.WindowStart ?? DateTime.MinValue);
        DateTime windowEnd = AsLocal(effective.WindowEnd ?? DateTime.MaxValue);
        var occurrences = new List<OutlookRecurrenceOccurrence>();
        var deletedDates = new HashSet<DateTime>(recurrence.DeletedOccurrenceDates.Select(value => AsLocal(value).Date));
        var exceptions = recurrence.Exceptions.GroupBy(exception => AsLocal(exception.OriginalStart))
            .ToDictionary(group => group.Key, group => group.Last());
        DateTime latestRelevantOriginal = recurrence.Exceptions
            .Where(exception => InWindow(exception.Start, exception.End, windowStart, windowEnd))
            .Select(exception => AsLocal(exception.OriginalStart))
            .DefaultIfEmpty(DateTime.MinValue).Max();

        int sequence = 0;
        int inspected = 0;
        bool truncated = false;
        string? reason = null;
        foreach (RecurrenceDateCandidate candidate in EnumerateDates(recurrence)) {
            inspected++;
            if (inspected > effective.MaxCandidateDays) {
                truncated = true;
                reason = "MaxCandidateDays was reached.";
                break;
            }

            DateTime occurrenceDate = candidate.Date;
            if (recurrence.RangeKind == OutlookRecurrenceRangeKind.EndDate &&
                recurrence.EndDate.HasValue && occurrenceDate.Date > recurrence.EndDate.Value.Date) break;
            if (!candidate.IsOccurrence) {
                if (occurrenceDate > windowEnd.Date && occurrenceDate >= latestRelevantOriginal.Date &&
                    recurrence.RangeKind == OutlookRecurrenceRangeKind.NoEnd) break;
                continue;
            }
            DateTime originalStart = occurrenceDate.Date.Add(recurrence.Start.TimeOfDay);
            if (originalStart < recurrence.Start) continue;
            sequence++;
            if (recurrence.RangeKind == OutlookRecurrenceRangeKind.OccurrenceCount &&
                recurrence.OccurrenceCount.HasValue && sequence > recurrence.OccurrenceCount.Value) break;

            exceptions.TryGetValue(originalStart, out OutlookRecurrenceException? exception);
            bool deleted = deletedDates.Contains(occurrenceDate.Date) && exception == null;
            DateTime start = exception == null ? originalStart : AsLocal(exception.Start);
            DateTime end = exception == null ? AddClamped(originalStart, recurrence.Duration) : AsLocal(exception.End);
            if (!deleted && InWindow(start, end, windowStart, windowEnd)) {
                occurrences.Add(new OutlookRecurrenceOccurrence(sequence, originalStart, start, end, exception));
                if (occurrences.Count >= effective.MaxOccurrences) {
                    truncated = HasMorePotential(recurrence, sequence, occurrenceDate, windowEnd,
                        latestRelevantOriginal);
                    if (truncated) reason = "MaxOccurrences was reached.";
                    break;
                }
            }

            if (recurrence.RangeKind == OutlookRecurrenceRangeKind.OccurrenceCount &&
                recurrence.OccurrenceCount.HasValue && sequence >= recurrence.OccurrenceCount.Value) break;

            if (occurrenceDate > windowEnd.Date && occurrenceDate >= latestRelevantOriginal.Date &&
                recurrence.RangeKind == OutlookRecurrenceRangeKind.NoEnd) break;
        }
        return new OutlookRecurrenceExpansionResult(occurrences, Math.Min(inspected, effective.MaxCandidateDays),
            truncated, reason);
    }

    internal static HashSet<DateTime> FindBaseOccurrenceDates(OutlookRecurrence recurrence,
        IEnumerable<DateTime> requestedDates) {
        if (recurrence == null) throw new ArgumentNullException(nameof(recurrence));
        if (requestedDates == null) throw new ArgumentNullException(nameof(requestedDates));
        Validate(recurrence, new OutlookRecurrenceExpansionOptions());
        var requested = new HashSet<DateTime>(requestedDates.Select(value => AsLocal(value).Date));
        var matches = new HashSet<DateTime>();
        if (requested.Count == 0) return matches;

        DateTime lastRequested = requested.Max();
        int sequence = 0;
        foreach (RecurrenceDateCandidate candidate in EnumerateDates(recurrence)) {
            DateTime date = candidate.Date.Date;
            if (date > lastRequested) break;
            if (recurrence.RangeKind == OutlookRecurrenceRangeKind.EndDate &&
                recurrence.EndDate.HasValue && date > recurrence.EndDate.Value.Date) break;
            if (!candidate.IsOccurrence || date < recurrence.Start.Date) continue;
            sequence++;
            if (recurrence.RangeKind == OutlookRecurrenceRangeKind.OccurrenceCount &&
                recurrence.OccurrenceCount.HasValue && sequence > recurrence.OccurrenceCount.Value) break;
            if (requested.Contains(date)) matches.Add(date);
            if (matches.Count == requested.Count) break;
            if (recurrence.RangeKind == OutlookRecurrenceRangeKind.OccurrenceCount &&
                recurrence.OccurrenceCount.HasValue && sequence >= recurrence.OccurrenceCount.Value) break;
        }
        return matches;
    }

    private static IEnumerable<RecurrenceDateCandidate> EnumerateDates(OutlookRecurrence recurrence) {
        DateTime startDate = recurrence.Start.Date;
        if (recurrence.PatternKind == OutlookRecurrencePatternKind.Day) {
            int step = recurrence.Frequency == OutlookRecurrenceFrequency.Daily ? recurrence.Interval : 1;
            for (DateTime date = startDate; ;) {
                yield return new RecurrenceDateCandidate(date, isOccurrence: true);
                if (!TryAddDays(date, step, out date)) yield break;
            }
        } else if (recurrence.PatternKind == OutlookRecurrencePatternKind.Week) {
            DateTime anchor = StartOfWeek(startDate, recurrence.FirstDayOfWeek);
            for (DateTime date = startDate; ;) {
                long weeks = (long)(StartOfWeek(date, recurrence.FirstDayOfWeek) - anchor).TotalDays / 7;
                bool isOccurrence = weeks % recurrence.Interval == 0 &&
                    Includes(recurrence.DaysOfWeek, date.DayOfWeek);
                yield return new RecurrenceDateCandidate(date, isOccurrence);
                if (!TryAddDays(date, 1, out date)) yield break;
            }
        } else {
            DateTime month = new DateTime(startDate.Year, startDate.Month, 1);
            int monthInterval = recurrence.Frequency == OutlookRecurrenceFrequency.Yearly
                ? checked(12 * recurrence.Interval)
                : recurrence.Interval;
            for (;;) {
                DateTime? candidate = GetMonthlyCandidate(recurrence, month);
                if (candidate.HasValue && candidate.Value >= startDate)
                    yield return new RecurrenceDateCandidate(candidate.Value, isOccurrence: true);
                try {
                    month = month.AddMonths(monthInterval);
                } catch (ArgumentOutOfRangeException) {
                    yield break;
                }
            }
        }
    }

    private static DateTime? GetMonthlyCandidate(OutlookRecurrence recurrence, DateTime month) {
        int days = DateTime.DaysInMonth(month.Year, month.Month);
        if (recurrence.PatternKind == OutlookRecurrencePatternKind.MonthEnd)
            return new DateTime(month.Year, month.Month, days);
        if (recurrence.PatternKind == OutlookRecurrencePatternKind.MonthDay) {
            int day = recurrence.DayOfMonth.GetValueOrDefault();
            return day <= days ? new DateTime(month.Year, month.Month, day) : (DateTime?)null;
        }

        var matches = new List<DateTime>();
        for (int day = 1; day <= days; day++) {
            var candidate = new DateTime(month.Year, month.Month, day);
            if (Includes(recurrence.DaysOfWeek, candidate.DayOfWeek)) matches.Add(candidate);
        }
        if (matches.Count == 0 || !recurrence.WeekOrdinal.HasValue) return null;
        if (recurrence.WeekOrdinal == OutlookRecurrenceWeekOrdinal.Last) return matches[matches.Count - 1];
        int index = (int)recurrence.WeekOrdinal.Value - 1;
        return index < matches.Count ? matches[index] : (DateTime?)null;
    }

    private static bool HasMorePotential(OutlookRecurrence recurrence, int sequence, DateTime currentDate,
        DateTime windowEnd, DateTime latestRelevantOriginal) {
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.OccurrenceCount && recurrence.OccurrenceCount.HasValue)
            return sequence < recurrence.OccurrenceCount.Value;
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.EndDate && recurrence.EndDate.HasValue)
            return currentDate.Date < recurrence.EndDate.Value.Date;
        return currentDate.Date < windowEnd.Date || currentDate.Date < latestRelevantOriginal.Date;
    }

    private static bool InWindow(DateTime start, DateTime end, DateTime windowStart, DateTime windowEnd) =>
        end > start
            ? end > windowStart && start < windowEnd
            : start >= windowStart && start < windowEnd;

    private static DateTime StartOfWeek(DateTime date, DayOfWeek firstDay) {
        int offset = ((int)date.DayOfWeek - (int)firstDay + 7) % 7;
        return date.Date.Ticks < TimeSpan.FromDays(offset).Ticks
            ? DateTime.MinValue
            : date.Date.AddDays(-offset);
    }

    private static bool TryAddDays(DateTime value, int days, out DateTime result) {
        long delta;
        try {
            delta = checked(TimeSpan.TicksPerDay * (long)days);
        } catch (OverflowException) {
            result = default;
            return false;
        }
        if (delta > DateTime.MaxValue.Ticks - value.Ticks) {
            result = default;
            return false;
        }
        result = value.AddTicks(delta);
        return true;
    }

    private static DateTime AddClamped(DateTime value, TimeSpan duration) =>
        duration.Ticks > DateTime.MaxValue.Ticks - value.Ticks ? DateTime.MaxValue : value.Add(duration);

    private static bool Includes(OutlookRecurrenceDays days, DayOfWeek day) =>
        (days & (OutlookRecurrenceDays)(1 << (int)day)) != 0;

    private static void Validate(OutlookRecurrence recurrence, OutlookRecurrenceExpansionOptions options) {
        if (options.MaxOccurrences <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxOccurrences));
        if (options.MaxCandidateDays <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxCandidateDays));
        if (options.WindowStart.HasValue && options.WindowEnd.HasValue &&
            options.WindowEnd.Value <= options.WindowStart.Value)
            throw new ArgumentException("WindowEnd must be later than WindowStart.", nameof(options));
        if (recurrence.Start == default) throw new InvalidOperationException("A recurrence requires Start.");
        if (!OutlookRecurrenceBinary.IsValidFrequencyPattern(recurrence.Frequency, recurrence.PatternKind))
            throw new InvalidOperationException("The recurrence frequency and pattern kind are incompatible.");
        if (recurrence.CalendarType != 0 && recurrence.CalendarType != 1 && recurrence.CalendarType != 2)
            throw new NotSupportedException("Expansion currently supports Gregorian Outlook calendar types only.");
        if (recurrence.PatternKind == OutlookRecurrencePatternKind.Week && recurrence.DaysOfWeek == OutlookRecurrenceDays.None)
            throw new InvalidOperationException("A weekly recurrence requires at least one weekday.");
        if (recurrence.PatternKind == OutlookRecurrencePatternKind.MonthDay && !recurrence.DayOfMonth.HasValue)
            throw new InvalidOperationException("A month-day recurrence requires DayOfMonth.");
        if (recurrence.PatternKind == OutlookRecurrencePatternKind.MonthNth &&
            (recurrence.DaysOfWeek == OutlookRecurrenceDays.None || !recurrence.WeekOrdinal.HasValue))
            throw new InvalidOperationException("An ordinal-month recurrence requires weekdays and WeekOrdinal.");
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.OccurrenceCount && !recurrence.OccurrenceCount.HasValue)
            throw new InvalidOperationException("An occurrence-count range requires OccurrenceCount.");
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.EndDate && !recurrence.EndDate.HasValue)
            throw new InvalidOperationException("An end-date range requires EndDate.");
    }

    private static DateTime AsLocal(DateTime value) => DateTime.SpecifyKind(value, DateTimeKind.Unspecified);

    private readonly struct RecurrenceDateCandidate {
        internal RecurrenceDateCandidate(DateTime date, bool isOccurrence) {
            Date = date;
            IsOccurrence = isOccurrence;
        }

        internal DateTime Date { get; }
        internal bool IsOccurrence { get; }
    }
}
