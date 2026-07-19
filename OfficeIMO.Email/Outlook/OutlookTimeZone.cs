namespace OfficeIMO.Email;

/// <summary>Status of a local clock value in an Outlook time-zone rule.</summary>
public enum OutlookLocalTimeStatus {
    /// <summary>The local value maps to exactly one UTC instant.</summary>
    Unambiguous,
    /// <summary>The local value is skipped when the clock moves forward.</summary>
    Invalid,
    /// <summary>The local value occurs twice when the clock moves backward.</summary>
    Ambiguous
}

/// <summary>Policy used when an Outlook local time occurs twice.</summary>
public enum OutlookAmbiguousTimePolicy {
    /// <summary>Return the chronologically earlier UTC instant.</summary>
    EarlierUtc,
    /// <summary>Return the chronologically later UTC instant.</summary>
    LaterUtc,
    /// <summary>Reject ambiguous local clock values.</summary>
    Throw
}

/// <summary>Windows SYSTEMTIME transition used by Outlook time-zone BLOBs.</summary>
public sealed class OutlookTimeZoneTransition : IEquatable<OutlookTimeZoneTransition> {
    /// <summary>Creates a transition from its exact SYSTEMTIME fields.</summary>
    public OutlookTimeZoneTransition(ushort year, ushort month, ushort dayOfWeek, ushort day, ushort hour,
        ushort minute, ushort second = 0, ushort milliseconds = 0) {
        if (month > 12) throw new ArgumentOutOfRangeException(nameof(month));
        if (dayOfWeek > 6) throw new ArgumentOutOfRangeException(nameof(dayOfWeek));
        if (day > 31) throw new ArgumentOutOfRangeException(nameof(day));
        if (hour > 23) throw new ArgumentOutOfRangeException(nameof(hour));
        if (minute > 59) throw new ArgumentOutOfRangeException(nameof(minute));
        if (second > 59) throw new ArgumentOutOfRangeException(nameof(second));
        if (milliseconds > 999) throw new ArgumentOutOfRangeException(nameof(milliseconds));
        Year = year;
        Month = month;
        DayOfWeek = dayOfWeek;
        Day = day;
        Hour = hour;
        Minute = minute;
        Second = second;
        Milliseconds = milliseconds;
    }

    /// <summary>Absolute year, or zero for an annually relative transition.</summary>
    public ushort Year { get; }
    /// <summary>Month, or zero when daylight transitions are disabled.</summary>
    public ushort Month { get; }
    /// <summary>Sunday-based weekday for relative transitions.</summary>
    public ushort DayOfWeek { get; }
    /// <summary>Absolute day, or relative occurrence 1 through 5 where 5 means last.</summary>
    public ushort Day { get; }
    /// <summary>Transition hour.</summary>
    public ushort Hour { get; }
    /// <summary>Transition minute.</summary>
    public ushort Minute { get; }
    /// <summary>Transition second.</summary>
    public ushort Second { get; }
    /// <summary>Transition millisecond.</summary>
    public ushort Milliseconds { get; }
    /// <summary>Whether this transition is disabled.</summary>
    public bool IsDisabled => Month == 0;

    /// <summary>Resolves this absolute or relative transition for a calendar year.</summary>
    public DateTime? GetDateTime(int calendarYear) {
        if (IsDisabled) return null;
        int year = Year == 0 ? calendarYear : Year;
        if (year < 1 || year > 9999) return null;
        if (Year != 0) {
            if (Day < 1 || Day > DateTime.DaysInMonth(year, Month)) return null;
            return Create(year, Month, Day);
        }
        if (Day < 1 || Day > 5) return null;
        var first = new DateTime(year, Month, 1);
        int offset = ((int)DayOfWeek - (int)first.DayOfWeek + 7) % 7;
        int resolvedDay = 1 + offset + (Day - 1) * 7;
        int daysInMonth = DateTime.DaysInMonth(year, Month);
        if (Day == 5 && resolvedDay > daysInMonth) resolvedDay -= 7;
        if (resolvedDay > daysInMonth) return null;
        return Create(year, Month, resolvedDay);
    }

    /// <inheritdoc />
    public bool Equals(OutlookTimeZoneTransition? other) => other != null &&
        Year == other.Year && Month == other.Month && DayOfWeek == other.DayOfWeek && Day == other.Day &&
        Hour == other.Hour && Minute == other.Minute && Second == other.Second &&
        Milliseconds == other.Milliseconds;

    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as OutlookTimeZoneTransition);
    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = Year;
            hash = hash * 397 ^ Month;
            hash = hash * 397 ^ DayOfWeek;
            hash = hash * 397 ^ Day;
            hash = hash * 397 ^ Hour;
            hash = hash * 397 ^ Minute;
            hash = hash * 397 ^ Second;
            return hash * 397 ^ Milliseconds;
        }
    }

    private DateTime Create(int year, int month, int day) =>
        new DateTime(year, month, day, Hour, Minute, Second, Milliseconds, DateTimeKind.Unspecified);
}

/// <summary>One TZRule contained in an Outlook time-zone definition.</summary>
public sealed class OutlookTimeZoneRule {
    private Action? _changed;
    private ushort _flags;
    private ushort _effectiveYear;
    private int _biasMinutes;
    private int _standardBiasMinutes;
    private int _daylightBiasMinutes;
    private OutlookTimeZoneTransition _standardTransition = DisabledTransition;
    private OutlookTimeZoneTransition _daylightTransition = DisabledTransition;

    /// <summary>Raw TZRULE flags.</summary>
    public ushort Flags { get => _flags; set { _flags = value; Changed(); } }
    /// <summary>Year from which the rule is effective.</summary>
    public ushort EffectiveYear { get => _effectiveYear; set { _effectiveYear = value; Changed(); } }
    /// <summary>Minutes added to local time to obtain UTC.</summary>
    public int BiasMinutes { get => _biasMinutes; set { _biasMinutes = value; Changed(); } }
    /// <summary>Additional standard-time bias.</summary>
    public int StandardBiasMinutes { get => _standardBiasMinutes; set { _standardBiasMinutes = value; Changed(); } }
    /// <summary>Additional daylight-time bias.</summary>
    public int DaylightBiasMinutes { get => _daylightBiasMinutes; set { _daylightBiasMinutes = value; Changed(); } }
    /// <summary>Transition to standard time.</summary>
    public OutlookTimeZoneTransition StandardTransition { get => _standardTransition; set { _standardTransition = value ?? throw new ArgumentNullException(nameof(value)); Changed(); } }
    /// <summary>Transition to daylight time.</summary>
    public OutlookTimeZoneTransition DaylightTransition { get => _daylightTransition; set { _daylightTransition = value ?? throw new ArgumentNullException(nameof(value)); Changed(); } }
    /// <summary>UTC offset during standard time.</summary>
    public TimeSpan StandardUtcOffset => TimeSpan.FromMinutes(-(BiasMinutes + StandardBiasMinutes));
    /// <summary>UTC offset during daylight time.</summary>
    public TimeSpan DaylightUtcOffset => TimeSpan.FromMinutes(-(BiasMinutes + DaylightBiasMinutes));
    /// <summary>Whether this rule declares daylight transitions.</summary>
    public bool HasDaylightSaving => !StandardTransition.IsDisabled && !DaylightTransition.IsDisabled &&
        StandardUtcOffset != DaylightUtcOffset;

    internal static OutlookTimeZoneTransition DisabledTransition { get; } =
        new OutlookTimeZoneTransition(0, 0, 0, 0, 0, 0);
    internal void SetChangeTracker(Action changed) => _changed = changed;
    private void Changed() => _changed?.Invoke();
}

/// <summary>Evidence describing how a local clock value maps to UTC.</summary>
public sealed class OutlookLocalTimeResolution {
    internal OutlookLocalTimeResolution(DateTime localTime, OutlookLocalTimeStatus status,
        IReadOnlyList<TimeSpan> offsets) {
        LocalTime = DateTime.SpecifyKind(localTime, DateTimeKind.Unspecified);
        Status = status;
        Offsets = offsets;
    }

    /// <summary>Input local clock value.</summary>
    public DateTime LocalTime { get; }
    /// <summary>Mapping status.</summary>
    public OutlookLocalTimeStatus Status { get; }
    /// <summary>Valid UTC offsets. Ambiguous values contain two offsets; invalid values contain none.</summary>
    public IReadOnlyList<TimeSpan> Offsets { get; }

    /// <summary>Resolves to a UTC-bearing value using an explicit ambiguity policy.</summary>
    public DateTimeOffset Resolve(OutlookAmbiguousTimePolicy policy = OutlookAmbiguousTimePolicy.EarlierUtc) {
        if (Status == OutlookLocalTimeStatus.Invalid)
            throw new InvalidOperationException("The local time does not exist because the clock moves forward.");
        if (Status == OutlookLocalTimeStatus.Ambiguous && policy == OutlookAmbiguousTimePolicy.Throw)
            throw new InvalidOperationException("The local time is ambiguous because the clock moves backward.");
        IEnumerable<DateTimeOffset> candidates = Offsets.Select(offset => new DateTimeOffset(LocalTime, offset));
        return policy == OutlookAmbiguousTimePolicy.LaterUtc
            ? candidates.OrderBy(value => value.UtcDateTime).Last()
            : candidates.OrderBy(value => value.UtcDateTime).First();
    }
}

/// <summary>Typed Outlook TZDEFINITION with lossless native-state retention.</summary>
public sealed class OutlookTimeZoneDefinition {
    private readonly TimeZoneRuleList _rules;
    private string? _keyName;
    private bool _dirty = true;

    /// <summary>Creates an editable definition.</summary>
    public OutlookTimeZoneDefinition() => _rules = new TimeZoneRuleList(MarkDirty);

    /// <summary>Windows time-zone key carried by Outlook.</summary>
    public string? KeyName { get => _keyName; set { _keyName = value; MarkDirty(); } }
    /// <summary>Rules ordered by effective year.</summary>
    public IList<OutlookTimeZoneRule> Rules => _rules;
    /// <summary>Whether the native BLOB decoded completely.</summary>
    public bool StateDecoded { get; internal set; } = true;
    /// <summary>Decode failure for retained native input.</summary>
    public string? DecodeError { get; internal set; }
    /// <summary>Original TZDEFINITION bytes.</summary>
    public byte[]? RawState { get; internal set; }
    internal bool CanPreserveRawState => RawState != null && !_dirty;

    /// <summary>Returns the rule effective for a local year.</summary>
    public OutlookTimeZoneRule GetRule(int localYear) {
        if (_rules.Count == 0) throw new InvalidOperationException("The time-zone definition contains no rules.");
        OutlookTimeZoneRule? selected = _rules.Where(rule => rule.EffectiveYear <= localYear)
            .OrderBy(rule => rule.EffectiveYear).LastOrDefault();
        return selected ?? _rules.OrderBy(rule => rule.EffectiveYear).First();
    }

    /// <summary>Classifies a local clock value without consulting the host time zone.</summary>
    public OutlookLocalTimeResolution GetLocalTimeResolution(DateTime localTime) =>
        OutlookTimeZoneResolver.Resolve(GetRule(localTime.Year), localTime);

    /// <summary>Resolves a local clock value using an explicit ambiguity policy.</summary>
    public DateTimeOffset ResolveLocal(DateTime localTime,
        OutlookAmbiguousTimePolicy policy = OutlookAmbiguousTimePolicy.EarlierUtc) =>
        GetLocalTimeResolution(localTime).Resolve(policy);

    /// <summary>Converts a UTC instant using the embedded historical rules.</summary>
    public DateTimeOffset ConvertUtc(DateTimeOffset utcValue) {
        DateTimeOffset utc = utcValue.ToUniversalTime();
        OutlookTimeZoneRule rule = GetRuleForUtc(utc);
        if (OutlookTimeZoneResolver.TryConvertUtc(rule, utc, out DateTimeOffset result)) return result;
        throw new InvalidOperationException("The UTC instant could not be mapped by the Outlook time-zone rules.");
    }

    private OutlookTimeZoneRule GetRuleForUtc(DateTimeOffset utcValue) {
        if (_rules.Count == 0) throw new InvalidOperationException("The time-zone definition contains no rules.");
        // Outlook TZRULE effective years form UTC intervals beginning on January 1. A bias
        // cutover can legitimately map an instant into the preceding local calendar year.
        OutlookTimeZoneRule[] ordered = _rules.OrderBy(rule => rule.EffectiveYear).ToArray();
        for (int index = ordered.Length - 1; index >= 0; index--) {
            int effectiveYear = Math.Max(1, (int)ordered[index].EffectiveYear);
            if (utcValue.Year >= effectiveYear) return ordered[index];
        }
        return ordered[0];
    }

    internal void AcceptDecodedState(byte[] rawState) {
        RawState = (byte[])rawState.Clone();
        StateDecoded = true;
        DecodeError = null;
        _dirty = false;
    }

    internal void SetDecodeFailure(byte[] rawState, string error) {
        RawState = (byte[])rawState.Clone();
        StateDecoded = false;
        DecodeError = error;
        _dirty = false;
    }

    private void MarkDirty() => _dirty = true;
}

/// <summary>Typed legacy PidLidTimeZoneStruct value.</summary>
public sealed class OutlookTimeZoneStructure {
    private OutlookTimeZoneRule _rule = new OutlookTimeZoneRule();
    private ushort _standardYear;
    private ushort _daylightYear;
    private bool _dirty = true;

    /// <summary>Creates an editable legacy structure.</summary>
    public OutlookTimeZoneStructure() => _rule.SetChangeTracker(MarkDirty);
    /// <summary>Effective bias and transitions.</summary>
    public OutlookTimeZoneRule Rule { get => _rule; set { _rule = value ?? throw new ArgumentNullException(nameof(value)); _rule.SetChangeTracker(MarkDirty); MarkDirty(); } }
    /// <summary>Legacy standard-transition year.</summary>
    public ushort StandardYear { get => _standardYear; set { _standardYear = value; MarkDirty(); } }
    /// <summary>Legacy daylight-transition year.</summary>
    public ushort DaylightYear { get => _daylightYear; set { _daylightYear = value; MarkDirty(); } }
    /// <summary>Whether the native BLOB decoded completely.</summary>
    public bool StateDecoded { get; internal set; } = true;
    /// <summary>Decode failure for retained native input.</summary>
    public string? DecodeError { get; internal set; }
    /// <summary>Original PidLidTimeZoneStruct bytes.</summary>
    public byte[]? RawState { get; internal set; }
    internal bool CanPreserveRawState => RawState != null && !_dirty;

    /// <summary>Classifies a local clock value.</summary>
    public OutlookLocalTimeResolution GetLocalTimeResolution(DateTime localTime) =>
        OutlookTimeZoneResolver.Resolve(Rule, localTime);
    /// <summary>Resolves a local clock value using an explicit ambiguity policy.</summary>
    public DateTimeOffset ResolveLocal(DateTime localTime,
        OutlookAmbiguousTimePolicy policy = OutlookAmbiguousTimePolicy.EarlierUtc) =>
        GetLocalTimeResolution(localTime).Resolve(policy);
    /// <summary>Converts a UTC instant using the legacy rule.</summary>
    public DateTimeOffset ConvertUtc(DateTimeOffset utcValue) {
        if (OutlookTimeZoneResolver.TryConvertUtc(Rule, utcValue.ToUniversalTime(), out DateTimeOffset result))
            return result;
        throw new InvalidOperationException("The UTC instant could not be mapped by the Outlook time-zone rule.");
    }

    internal void AcceptDecodedState(byte[] rawState) { RawState = (byte[])rawState.Clone(); StateDecoded = true; DecodeError = null; _dirty = false; }
    internal void SetDecodeFailure(byte[] rawState, string error) { RawState = (byte[])rawState.Clone(); StateDecoded = false; DecodeError = error; _dirty = false; }
    private void MarkDirty() => _dirty = true;
}

internal static class OutlookTimeZoneResolver {
    internal static OutlookLocalTimeResolution Resolve(OutlookTimeZoneRule rule, DateTime value) {
        DateTime local = DateTime.SpecifyKind(value, DateTimeKind.Unspecified);
        TimeSpan standard = rule.StandardUtcOffset;
        if (!rule.HasDaylightSaving)
            return Result(local, OutlookLocalTimeStatus.Unambiguous, standard);
        DateTime? daylightTransition = rule.DaylightTransition.GetDateTime(local.Year);
        DateTime? standardTransition = rule.StandardTransition.GetDateTime(local.Year);
        if (!daylightTransition.HasValue || !standardTransition.HasValue)
            return Result(local, OutlookLocalTimeStatus.Unambiguous, standard);

        TimeSpan daylight = rule.DaylightUtcOffset;
        TimeSpan springDelta = daylight - standard;
        OutlookLocalTimeResolution? boundary = ClassifyBoundary(local, daylightTransition.Value, standard,
            daylight, springDelta);
        if (boundary != null) return boundary;
        TimeSpan autumnDelta = standard - daylight;
        boundary = ClassifyBoundary(local, standardTransition.Value, daylight, standard, autumnDelta);
        if (boundary != null) return boundary;

        bool isDaylight = daylightTransition < standardTransition
            ? local >= daylightTransition && local < standardTransition
            : local >= daylightTransition || local < standardTransition;
        return Result(local, OutlookLocalTimeStatus.Unambiguous, isDaylight ? daylight : standard);
    }

    internal static bool TryConvertUtc(OutlookTimeZoneRule rule, DateTimeOffset utcValue,
        out DateTimeOffset result) {
        DateTimeOffset utc = utcValue.ToUniversalTime();
        foreach (TimeSpan offset in new[] { rule.StandardUtcOffset, rule.DaylightUtcOffset }.Distinct()) {
            DateTime local;
            try {
                local = DateTime.SpecifyKind(utc.UtcDateTime.Add(offset), DateTimeKind.Unspecified);
            } catch (ArgumentOutOfRangeException) {
                continue;
            }
            OutlookLocalTimeResolution resolution = Resolve(rule, local);
            if (!resolution.Offsets.Contains(offset)) continue;
            var candidate = new DateTimeOffset(local, offset);
            if (candidate.UtcDateTime != utc.UtcDateTime) continue;
            result = candidate;
            return true;
        }
        result = default;
        return false;
    }

    private static OutlookLocalTimeResolution? ClassifyBoundary(DateTime local, DateTime transition,
        TimeSpan before, TimeSpan after, TimeSpan delta) {
        if (delta > TimeSpan.Zero && local >= transition && local < transition + delta)
            return new OutlookLocalTimeResolution(local, OutlookLocalTimeStatus.Invalid, Array.Empty<TimeSpan>());
        if (delta < TimeSpan.Zero && local >= transition + delta && local < transition)
            return new OutlookLocalTimeResolution(local, OutlookLocalTimeStatus.Ambiguous,
                new[] { before, after });
        return null;
    }

    private static OutlookLocalTimeResolution Result(DateTime local, OutlookLocalTimeStatus status,
        TimeSpan offset) => new OutlookLocalTimeResolution(local, status, new[] { offset });
}

internal sealed class TimeZoneRuleList : IList<OutlookTimeZoneRule> {
    private readonly List<OutlookTimeZoneRule> _items = new List<OutlookTimeZoneRule>();
    private readonly Action _changed;
    internal TimeZoneRuleList(Action changed) => _changed = changed;
    public IEnumerator<OutlookTimeZoneRule> GetEnumerator() => _items.GetEnumerator();
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    public void Add(OutlookTimeZoneRule item) { Attach(item); _items.Add(item); _changed(); }
    public void Clear() { if (_items.Count == 0) return; _items.Clear(); _changed(); }
    public bool Contains(OutlookTimeZoneRule item) => _items.Contains(item);
    public void CopyTo(OutlookTimeZoneRule[] array, int arrayIndex) => _items.CopyTo(array, arrayIndex);
    public bool Remove(OutlookTimeZoneRule item) { bool removed = _items.Remove(item); if (removed) _changed(); return removed; }
    public int Count => _items.Count;
    public bool IsReadOnly => false;
    public int IndexOf(OutlookTimeZoneRule item) => _items.IndexOf(item);
    public void Insert(int index, OutlookTimeZoneRule item) { Attach(item); _items.Insert(index, item); _changed(); }
    public void RemoveAt(int index) { _items.RemoveAt(index); _changed(); }
    public OutlookTimeZoneRule this[int index] { get => _items[index]; set { Attach(value); _items[index] = value; _changed(); } }
    private void Attach(OutlookTimeZoneRule item) { if (item == null) throw new ArgumentNullException(nameof(item)); item.SetChangeTracker(_changed); }
}
