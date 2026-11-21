namespace OfficeIMO.Toml;

/// <summary>
/// A TOML date, time or date-time with nanosecond precision. TOML has four kinds (<see cref="TomlDateTimeKind"/>); the
/// fields a kind does not have are zero. Fractional seconds beyond nine digits are truncated, as the specification
/// requires.
/// </summary>
/// <remarks>
/// Equality compares the kind and every field, including the offset, so <c>07:32:00Z</c> and <c>08:32:00+01:00</c> are
/// different values that describe the same instant; <see cref="ToDateTimeOffset"/> compares instants. The number of
/// fractional digits written in the source is kept for formatting and does not take part in equality.
/// </remarks>
public readonly struct TomlDateTime : IEquatable<TomlDateTime> {
    private readonly short _offsetMinutes;

    private TomlDateTime(TomlDateTimeKind kind, int year, int month, int day, int hour, int minute, int second, int nanosecond, int fractionDigits, int offsetMinutes) {
        Kind = kind;
        Year = year;
        Month = month;
        Day = day;
        Hour = hour;
        Minute = minute;
        Second = second;
        Nanosecond = nanosecond;
        FractionDigits = fractionDigits;
        _offsetMinutes = (short)offsetMinutes;
    }

    /// <summary>Gets which fields the value has.</summary>
    public TomlDateTimeKind Kind { get; }

    /// <summary>Gets the year, or 0 for a local time.</summary>
    public int Year { get; }

    /// <summary>Gets the month (1–12), or 0 for a local time.</summary>
    public int Month { get; }

    /// <summary>Gets the day of the month, or 0 for a local time.</summary>
    public int Day { get; }

    /// <summary>Gets the hour (0–23).</summary>
    public int Hour { get; }

    /// <summary>Gets the minute (0–59).</summary>
    public int Minute { get; }

    /// <summary>Gets the second (0–60; 60 is a leap second).</summary>
    public int Second { get; }

    /// <summary>Gets the fraction of the second in nanoseconds (0–999,999,999).</summary>
    public int Nanosecond { get; }

    /// <summary>Gets how many fractional-second digits the value is written with (0–9).</summary>
    public int FractionDigits { get; }

    /// <summary>Gets the UTC offset of an <see cref="TomlDateTimeKind.OffsetDateTime"/>, or <see langword="null"/> for the local kinds.</summary>
    public TimeSpan? Offset => Kind == TomlDateTimeKind.OffsetDateTime ? TimeSpan.FromMinutes(_offsetMinutes) : null;

    internal int OffsetMinutes => _offsetMinutes;

    /// <summary>Creates an offset date-time from a <see cref="DateTimeOffset"/>, keeping its sub-second ticks.</summary>
    public static TomlDateTime FromDateTimeOffset(DateTimeOffset value) {
        int nanos = (int)(value.Ticks % TimeSpan.TicksPerSecond) * 100;
        return new TomlDateTime(TomlDateTimeKind.OffsetDateTime, value.Year, value.Month, value.Day, value.Hour, value.Minute, value.Second,
            nanos, DigitsFor(nanos), (int)value.Offset.TotalMinutes);
    }

    /// <summary>
    /// Creates a value from a <see cref="DateTime"/>: an offset date-time with <c>Z</c> for <see cref="DateTimeKind.Utc"/>,
    /// an offset date-time with the machine's offset for <see cref="DateTimeKind.Local"/>, and a local date-time otherwise.
    /// </summary>
    public static TomlDateTime FromDateTime(DateTime value) {
        if (value.Kind == DateTimeKind.Utc) return FromDateTimeOffset(new DateTimeOffset(value, TimeSpan.Zero));
        if (value.Kind == DateTimeKind.Local) return FromDateTimeOffset(new DateTimeOffset(value));
        int nanos = (int)(value.Ticks % TimeSpan.TicksPerSecond) * 100;
        return new TomlDateTime(TomlDateTimeKind.LocalDateTime, value.Year, value.Month, value.Day, value.Hour, value.Minute, value.Second, nanos, DigitsFor(nanos), 0);
    }

    /// <summary>Creates a local date.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The fields do not form a valid date.</exception>
    public static TomlDateTime FromDate(int year, int month, int day) {
        if (!IsValidDate(year, month, day)) throw new ArgumentOutOfRangeException(nameof(day), "The fields do not form a valid date.");
        return new TomlDateTime(TomlDateTimeKind.LocalDate, year, month, day, 0, 0, 0, 0, 0, 0);
    }

    /// <summary>Creates a local time of day.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A field is out of range.</exception>
    public static TomlDateTime FromTime(int hour, int minute, int second, int nanosecond = 0) {
        if ((uint)hour > 23 || (uint)minute > 59 || (uint)second > 60 || (uint)nanosecond > 999_999_999) {
            throw new ArgumentOutOfRangeException(nameof(hour), "The fields do not form a valid time of day.");
        }
        return new TomlDateTime(TomlDateTimeKind.LocalTime, 0, 0, 0, hour, minute, second, nanosecond, DigitsFor(nanosecond), 0);
    }

    internal static TomlDateTime Create(TomlDateTimeKind kind, int year, int month, int day, int hour, int minute, int second, int nanosecond, int fractionDigits, int offsetMinutes) =>
        new TomlDateTime(kind, year, month, day, hour, minute, second, nanosecond, fractionDigits, offsetMinutes);

    /// <summary>
    /// Returns the instant of an offset date-time. Ticks are 100 nanoseconds, so finer digits are truncated, and a leap
    /// second (<c>:60</c>) becomes the first instant of the next minute.
    /// </summary>
    /// <exception cref="InvalidOperationException">The value is a local kind and has no offset.</exception>
    public DateTimeOffset ToDateTimeOffset() {
        if (Kind != TomlDateTimeKind.OffsetDateTime) throw new InvalidOperationException("A " + Kind + " has no offset and is not an instant.");
        return new DateTimeOffset(ToTicks(), TimeSpan.FromMinutes(_offsetMinutes));
    }

    /// <summary>
    /// Returns the value as a <see cref="DateTime"/>: the UTC instant (<see cref="DateTimeKind.Utc"/>) for an offset
    /// date-time, and the date and time with <see cref="DateTimeKind.Unspecified"/> for a local date-time or date.
    /// </summary>
    /// <exception cref="InvalidOperationException">The value is a local time without a date.</exception>
    public DateTime ToDateTime() {
        return Kind switch {
            TomlDateTimeKind.OffsetDateTime => ToDateTimeOffset().UtcDateTime,
            TomlDateTimeKind.LocalTime => throw new InvalidOperationException("A local time has no date."),
            _ => new DateTime(ToTicks(), DateTimeKind.Unspecified)
        };
    }

    /// <summary>Returns the time of day as the time elapsed since midnight; the date and offset are ignored.</summary>
    /// <exception cref="InvalidOperationException">The value is a local date without a time.</exception>
    public TimeSpan ToTimeOfDay() {
        if (Kind == TomlDateTimeKind.LocalDate) throw new InvalidOperationException("A local date has no time.");
        return new TimeSpan((Hour * 3600L + Minute * 60L + Second) * TimeSpan.TicksPerSecond + Nanosecond / 100);
    }

    private long ToTicks() {
        long days = new DateTime(Year, Month, Day).Ticks;
        return days + (Hour * 3600L + Minute * 60L + Second) * TimeSpan.TicksPerSecond + Nanosecond / 100;
    }

    /// <summary>Formats the value as TOML writes it, such as <c>1979-05-27T07:32:00.5-07:00</c>.</summary>
    public override string ToString() {
        var text = new StringBuilder(35);
        TomlFormatter.AppendDateTime(text, this, 'T');
        return text.ToString();
    }

    /// <inheritdoc />
    public bool Equals(TomlDateTime other) =>
        Kind == other.Kind && Year == other.Year && Month == other.Month && Day == other.Day && Hour == other.Hour
        && Minute == other.Minute && Second == other.Second && Nanosecond == other.Nanosecond && _offsetMinutes == other._offsetMinutes;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is TomlDateTime other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => unchecked(((((Year * 13 + Month) * 32 + Day) * 24 + Hour) * 60 + Minute) * 61 + Second ^ Nanosecond ^ _offsetMinutes << 20 ^ (int)Kind << 28);

    /// <summary>Compares two values for equality.</summary>
    public static bool operator ==(TomlDateTime left, TomlDateTime right) => left.Equals(right);

    /// <summary>Compares two values for inequality.</summary>
    public static bool operator !=(TomlDateTime left, TomlDateTime right) => !left.Equals(right);

    internal static bool IsValidDate(int year, int month, int day) =>
        year >= 1 && year <= 9999 && month >= 1 && month <= 12 && day >= 1 && day <= DateTime.DaysInMonth(year, month);

    private static int DigitsFor(int nanos) {
        if (nanos == 0) return 0;
        int digits = 9;
        while (nanos % 10 == 0) {
            nanos /= 10;
            digits--;
        }
        return digits;
    }
}
