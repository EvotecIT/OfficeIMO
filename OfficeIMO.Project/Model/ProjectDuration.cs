using System.Globalization;

namespace OfficeIMO.Project;

/// <summary>The display and conversion unit of a project duration.</summary>
public enum ProjectDurationUnit {
    /// <summary>Minutes.</summary>
    Minute,
    /// <summary>Hours.</summary>
    Hour,
    /// <summary>Days, interpreted using project settings for working durations.</summary>
    Day,
    /// <summary>Weeks, interpreted using project settings for working durations.</summary>
    Week,
    /// <summary>Months, interpreted using project settings for working durations.</summary>
    Month
}

/// <summary>A duration quantity that distinguishes working time from elapsed time and retains its display unit.</summary>
public readonly struct ProjectDuration : IEquatable<ProjectDuration> {
    // Keep the decoded quantity before display-unit division. A minute in days,
    // or a tick in minutes, may not have a finite decimal display representation.
    private readonly decimal _decodedQuantity, _decodedUnitFactor;
    private readonly bool _decodedTicks;
    /// <summary>Creates a duration. Negative values are reserved for dependency lag, not task duration.</summary>
    public ProjectDuration(decimal value, ProjectDurationUnit unit, bool elapsed = false, bool estimated = false) {
        this = default;
        if (!Enum.IsDefined(typeof(ProjectDurationUnit), unit)) throw new ArgumentOutOfRangeException(nameof(unit));
        Value = value; Unit = unit; IsElapsed = elapsed; IsEstimated = estimated;
    }
    private ProjectDuration(decimal value, ProjectDurationUnit unit, bool elapsed, bool estimated,
        decimal quantity, decimal factor, bool ticks) : this(value, unit, elapsed, estimated) {
        _decodedQuantity = quantity; _decodedUnitFactor = factor; _decodedTicks = ticks;
    }
    internal static ProjectDuration FromTicks(long ticks, ProjectDurationUnit unit, bool elapsed, bool estimated, decimal factor) {
        if (factor <= 0) throw new InvalidDataException("Project duration conversion requires positive working-time settings.");
        return new ProjectDuration(ticks / (decimal)TimeSpan.TicksPerMinute / factor, unit, elapsed, estimated, ticks, factor, true);
    }
    internal static ProjectDuration FromMinutes(decimal minutes, ProjectDurationUnit unit, bool elapsed, bool estimated, decimal factor) {
        if (factor <= 0) throw new InvalidDataException("Project duration conversion requires positive working-time settings.");
        decimal ticks = checked(minutes * TimeSpan.TicksPerMinute);
        decimal rounded = decimal.Round(ticks, 0, MidpointRounding.AwayFromZero);
        if (rounded >= long.MinValue && rounded <= long.MaxValue &&
            (ticks == rounded || rounded / TimeSpan.TicksPerMinute == minutes))
            return FromTicks((long)rounded, unit, elapsed, estimated, factor);
        return new ProjectDuration(minutes / factor, unit, elapsed, estimated, minutes, factor, false);
    }
    internal static ProjectDuration FromDecodedValue(decimal value, ProjectDurationUnit unit, bool elapsed, bool estimated, decimal factor) {
        var duration = new ProjectDuration(value, unit, elapsed, estimated);
        // MPX quantities can exceed TimeSpan's range; keep those readable and
        // let the destination format assessment diagnose representability.
        if (factor <= 0 || Math.Abs(value) > long.MaxValue / (decimal)TimeSpan.TicksPerMinute / factor) return duration;
        decimal ticks = duration.Ticks(factor);
        decimal rounded = decimal.Round(ticks, 0, MidpointRounding.AwayFromZero);
        if (rounded < long.MinValue || rounded > long.MaxValue) return duration;
        var exact = FromTicks((long)rounded, unit, elapsed, estimated, factor);
        // Only recover an integral quantity when its canonical decimal display is
        // exactly this lexical value. This does not accept arbitrary sub-tick input.
        return ticks == rounded || exact.Value == value ? exact : duration;
    }
    internal decimal Minutes(decimal factor) => _decodedUnitFactor == 0 ? checked(Value * factor) :
        _decodedTicks ? checked(_decodedQuantity * factor / _decodedUnitFactor) / TimeSpan.TicksPerMinute : checked(_decodedQuantity * factor / _decodedUnitFactor);
    internal decimal Ticks(decimal factor) => _decodedUnitFactor == 0 ? checked(Value * factor * TimeSpan.TicksPerMinute) :
        _decodedTicks ? checked(_decodedQuantity * factor / _decodedUnitFactor) : checked(_decodedQuantity * factor / _decodedUnitFactor * TimeSpan.TicksPerMinute);
    internal decimal ScaledMinutes(decimal factor, decimal scale) => _decodedTicks
        ? checked(_decodedQuantity * factor / _decodedUnitFactor * scale) / TimeSpan.TicksPerMinute : checked(Minutes(factor) * scale);
    /// <summary>Quantity in the selected unit.</summary>
    public decimal Value { get; }
    /// <summary>Display and conversion unit.</summary>
    public ProjectDurationUnit Unit { get; }
    /// <summary>True when all time counts, including non-working time.</summary>
    public bool IsElapsed { get; }
    /// <summary>True when the duration is marked as an estimate.</summary>
    public bool IsEstimated { get; }
    /// <summary>Creates a working-day duration.</summary>
    public static ProjectDuration WorkingDays(decimal value) => new ProjectDuration(value, ProjectDurationUnit.Day);
    /// <summary>Creates a working-hour duration.</summary>
    public static ProjectDuration WorkingHours(decimal value) => new ProjectDuration(value, ProjectDurationUnit.Hour);
    /// <summary>Creates a working-minute duration.</summary>
    public static ProjectDuration WorkingMinutes(decimal value) => new ProjectDuration(value, ProjectDurationUnit.Minute);
    /// <summary>Creates an elapsed-day duration.</summary>
    public static ProjectDuration ElapsedDays(decimal value) => new ProjectDuration(value, ProjectDurationUnit.Day, true);
    /// <summary>Creates an elapsed-hour duration.</summary>
    public static ProjectDuration ElapsedHours(decimal value) => new ProjectDuration(value, ProjectDurationUnit.Hour, true);
    /// <summary>Returns a copy with the estimate marker.</summary>
    public ProjectDuration Estimated(bool value = true) => new ProjectDuration(Value, Unit, IsElapsed, value, _decodedQuantity, _decodedUnitFactor, _decodedTicks);
    /// <summary>Tests quantity, unit, elapsed semantics, and estimate marker.</summary>
    public bool Equals(ProjectDuration other) => Value == other.Value && Unit == other.Unit && IsElapsed == other.IsElapsed && IsEstimated == other.IsEstimated;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ProjectDuration other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => Value.GetHashCode() ^ ((int)Unit << 4) ^ (IsElapsed ? 128 : 0) ^ (IsEstimated ? 256 : 0);
    /// <summary>Returns an invariant diagnostic representation, not a locale-specific Project input string.</summary>
    public override string ToString() => Value.ToString(CultureInfo.InvariantCulture) + " " + (IsElapsed ? "elapsed " : "working ") + Unit + (IsEstimated ? "?" : "");
}

/// <summary>Assignment work, measured in minutes independently of a task's duration.</summary>
public readonly struct ProjectWork : IEquatable<ProjectWork> {
    private readonly long? _decodedTicks;
    /// <summary>Creates an amount of work in minutes.</summary>
    public ProjectWork(decimal minutes) { if (minutes < 0) throw new ArgumentOutOfRangeException(nameof(minutes)); Minutes = minutes; _decodedTicks = null; }
    private ProjectWork(long ticks) : this(ticks / (decimal)TimeSpan.TicksPerMinute) { _decodedTicks = ticks; }
    internal static ProjectWork FromTicks(long ticks) => new ProjectWork(ticks);
    internal static ProjectWork FromMinutes(decimal minutes) {
        var work = new ProjectWork(minutes);
        if (minutes > long.MaxValue / (decimal)TimeSpan.TicksPerMinute) return work;
        decimal ticks = minutes * TimeSpan.TicksPerMinute;
        long rounded = checked((long)decimal.Round(ticks, 0, MidpointRounding.AwayFromZero));
        return ticks == rounded || rounded / (decimal)TimeSpan.TicksPerMinute == minutes ? FromTicks(rounded) : work;
    }
    internal decimal Ticks => _decodedTicks ?? checked(Minutes * TimeSpan.TicksPerMinute);
    internal decimal ScaledMinutes(decimal scale) => _decodedTicks.HasValue ? _decodedTicks.Value * scale / TimeSpan.TicksPerMinute : checked(Minutes * scale);
    /// <summary>Total work minutes.</summary>
    public decimal Minutes { get; }
    /// <summary>Creates work from hours.</summary>
    public static ProjectWork Hours(decimal value) => new ProjectWork(checked(value * 60));
    /// <inheritdoc />
    public bool Equals(ProjectWork other) => Minutes == other.Minutes;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ProjectWork other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => Minutes.GetHashCode();
}

/// <summary>Assignment allocation, where 100 percent is one full resource unit.</summary>
public readonly struct ProjectUnits : IEquatable<ProjectUnits> {
    private ProjectUnits(decimal value) { if (value < 0) throw new ArgumentOutOfRangeException(nameof(value)); Value = value; }
    /// <summary>Fractional units; 1 means 100 percent.</summary>
    public decimal Value { get; }
    /// <summary>Creates allocation from a percent such as 50 or 100.</summary>
    public static ProjectUnits Percent(decimal value) => new ProjectUnits(value / 100);
    /// <summary>Creates fractional resource units such as 0.5 or 1.</summary>
    public static ProjectUnits Fraction(decimal value) => new ProjectUnits(value);
    /// <inheritdoc />
    public bool Equals(ProjectUnits other) => Value == other.Value;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ProjectUnits other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => Value.GetHashCode();
}
