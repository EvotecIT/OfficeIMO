namespace OfficeIMO.Email;

/// <summary>RFC 5545 temporal representation without collapsing floating or TZID-local values to UTC.</summary>
public enum IcsTemporalValueKind {
    /// <summary>Calendar date without a time of day.</summary>
    Date = 0,
    /// <summary>Floating local date-time without a time-zone association.</summary>
    FloatingDateTime = 1,
    /// <summary>UTC date-time ending in Z.</summary>
    UtcDateTime = 2,
    /// <summary>Local clock date-time associated with a TZID.</summary>
    ZonedDateTime = 3
}

/// <summary>A parsed iCalendar DATE or DATE-TIME value that retains its original time semantics.</summary>
public readonly struct IcsTemporalValue : IEquatable<IcsTemporalValue> {
    private IcsTemporalValue(DateTime value, IcsTemporalValueKind kind, string? timeZoneId) {
        Value = value;
        Kind = kind;
        TimeZoneId = timeZoneId;
    }

    /// <summary>Clock value. Inspect <see cref="Kind"/> before interpreting it.</summary>
    public DateTime Value { get; }
    /// <summary>Temporal representation kind.</summary>
    public IcsTemporalValueKind Kind { get; }
    /// <summary>TZID for <see cref="IcsTemporalValueKind.ZonedDateTime"/>.</summary>
    public string? TimeZoneId { get; }

    /// <summary>Creates a DATE value.</summary>
    public static IcsTemporalValue Date(DateTime value) => new IcsTemporalValue(value.Date,
        IcsTemporalValueKind.Date, null);

    /// <summary>Creates a floating DATE-TIME value.</summary>
    public static IcsTemporalValue Floating(DateTime value) => new IcsTemporalValue(
        DateTime.SpecifyKind(value, DateTimeKind.Unspecified), IcsTemporalValueKind.FloatingDateTime, null);

    /// <summary>Creates a UTC DATE-TIME value.</summary>
    public static IcsTemporalValue Utc(DateTimeOffset value) => new IcsTemporalValue(
        value.UtcDateTime, IcsTemporalValueKind.UtcDateTime, null);

    /// <summary>Creates a TZID-local DATE-TIME value without resolving the identifier through the host OS.</summary>
    public static IcsTemporalValue Zoned(DateTime localValue, string timeZoneId) {
        if (string.IsNullOrWhiteSpace(timeZoneId)) throw new ArgumentException("TZID cannot be empty.", nameof(timeZoneId));
        return new IcsTemporalValue(DateTime.SpecifyKind(localValue, DateTimeKind.Unspecified),
            IcsTemporalValueKind.ZonedDateTime, timeZoneId);
    }

    /// <summary>Parses a temporal content-line property.</summary>
    public static IcsTemporalValue Parse(ContentLineProperty property) {
        if (!TryParse(property, out IcsTemporalValue value))
            throw new FormatException("The property does not contain a supported iCalendar DATE or DATE-TIME value.");
        return value;
    }

    /// <summary>Attempts to parse a temporal property while preserving DATE, floating, UTC, and TZID forms.</summary>
    public static bool TryParse(ContentLineProperty? property, out IcsTemporalValue value) {
        value = default;
        if (property == null) return false;
        string text = property.Value;
        string? valueType = property.GetParameter("VALUE")?.Values.FirstOrDefault();
        string? timeZoneId = property.GetParameter("TZID")?.Values.FirstOrDefault();
        if (valueType != null && !string.Equals(valueType, "DATE", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(valueType, "DATE-TIME", StringComparison.OrdinalIgnoreCase)) return false;
        if (string.Equals(valueType, "DATE", StringComparison.OrdinalIgnoreCase)) {
            if (!string.IsNullOrWhiteSpace(timeZoneId)) return false;
            if (!TryParseDate(text, out DateTime date)) return false;
            value = Date(date);
            return true;
        }
        if (TryParseDateTime(text, utc: true, out DateTime utc)) {
            if (!string.IsNullOrWhiteSpace(timeZoneId)) return false;
            value = Utc(new DateTimeOffset(utc, TimeSpan.Zero));
            return true;
        }
        if (!TryParseDateTime(text, utc: false, out DateTime local)) return false;
        value = string.IsNullOrWhiteSpace(timeZoneId) ? Floating(local) : Zoned(local, timeZoneId!);
        return true;
    }

    /// <summary>Applies this value and its VALUE/TZID parameters to a property.</summary>
    public void ApplyTo(ContentLineProperty property) {
        if (property == null) throw new ArgumentNullException(nameof(property));
        RemoveParameter(property, "VALUE");
        RemoveParameter(property, "TZID");
        switch (Kind) {
            case IcsTemporalValueKind.Date:
                property.Value = Value.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                property.SetParameter("VALUE", "DATE");
                break;
            case IcsTemporalValueKind.UtcDateTime:
                property.Value = Value.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);
                break;
            case IcsTemporalValueKind.ZonedDateTime:
                if (string.IsNullOrWhiteSpace(TimeZoneId)) throw new InvalidOperationException("A zoned value requires TZID.");
                property.Value = Value.ToString("yyyyMMdd'T'HHmmss", CultureInfo.InvariantCulture);
                property.SetParameter("TZID", TimeZoneId!);
                break;
            default:
                property.Value = Value.ToString("yyyyMMdd'T'HHmmss", CultureInfo.InvariantCulture);
                break;
        }
    }

    /// <inheritdoc />
    public bool Equals(IcsTemporalValue other) => Value == other.Value && Kind == other.Kind &&
        string.Equals(TimeZoneId, other.TimeZoneId, StringComparison.Ordinal);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is IcsTemporalValue other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked { return ((Value.GetHashCode() * 397) ^ (int)Kind) * 397 ^ (TimeZoneId?.GetHashCode() ?? 0); }
    }

    private static void RemoveParameter(ContentLineProperty property, string name) {
        for (int index = property.Parameters.Count - 1; index >= 0; index--) {
            if (string.Equals(property.Parameters[index].Name, name, StringComparison.OrdinalIgnoreCase))
                property.Parameters.RemoveAt(index);
        }
    }

    private static bool TryParseDate(string text, out DateTime value) {
        value = default;
        if (text.Length != 8 || !TryReadNumber(text, 0, 4, out int year) ||
            !TryReadNumber(text, 4, 2, out int month) ||
            !TryReadNumber(text, 6, 2, out int day)) return false;
        try {
            value = new DateTime(year, month, day, 0, 0, 0, DateTimeKind.Unspecified);
            return true;
        } catch (ArgumentOutOfRangeException) {
            return false;
        }
    }

    private static bool TryParseDateTime(string text, bool utc, out DateTime value) {
        value = default;
        int expectedLength = utc ? 16 : 15;
        if (text.Length != expectedLength || text[8] != 'T' ||
            utc && text[15] != 'Z' ||
            !TryReadNumber(text, 0, 4, out int year) ||
            !TryReadNumber(text, 4, 2, out int month) ||
            !TryReadNumber(text, 6, 2, out int day) ||
            !TryReadNumber(text, 9, 2, out int hour) ||
            !TryReadNumber(text, 11, 2, out int minute) ||
            !TryReadNumber(text, 13, 2, out int second) || second > 60) return false;
        try {
            value = new DateTime(year, month, day, hour, minute, Math.Min(second, 59),
                utc ? DateTimeKind.Utc : DateTimeKind.Unspecified);
            if (second == 60) value = value.AddSeconds(1);
            return true;
        } catch (ArgumentOutOfRangeException) {
            return false;
        }
    }

    private static bool TryReadNumber(string text, int offset, int length, out int value) {
        value = 0;
        for (int index = offset; index < offset + length; index++) {
            char character = text[index];
            if (character < '0' || character > '9') return false;
            value = checked(value * 10 + character - '0');
        }
        return true;
    }
}
