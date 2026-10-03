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
    private readonly DecodedQuantity? _decoded;
    private sealed class DecodedQuantity {
        internal readonly decimal Quantity, UnitFactor;
        internal readonly bool IsTicks;
        internal DecodedQuantity(decimal quantity, decimal factor, bool ticks) { Quantity = quantity; UnitFactor = factor; IsTicks = ticks; }
    }
    /// <summary>Creates a duration. Negative values are reserved for dependency lag, not task duration.</summary>
    public ProjectDuration(decimal value, ProjectDurationUnit unit, bool elapsed = false, bool estimated = false) {
        this = default;
        if (!Enum.IsDefined(typeof(ProjectDurationUnit), unit)) throw new ArgumentOutOfRangeException(nameof(unit));
        Value = value; Unit = unit; IsElapsed = elapsed; IsEstimated = estimated;
    }
    private ProjectDuration(decimal value, ProjectDurationUnit unit, bool elapsed, bool estimated,
        DecodedQuantity? decoded) : this(value, unit, elapsed, estimated) { _decoded = decoded; }
    internal static ProjectDuration FromTicks(long ticks, ProjectDurationUnit unit, bool elapsed, bool estimated, decimal factor) {
        if (factor <= 0) throw new InvalidDataException("Project duration conversion requires positive working-time settings.");
        var duration = new ProjectDuration(ticks / (decimal)TimeSpan.TicksPerMinute / factor, unit, elapsed, estimated);
        return duration.Ticks(factor) == ticks ? duration :
            new ProjectDuration(duration.Value, unit, elapsed, estimated, new DecodedQuantity(ticks, factor, true));
    }
    internal static ProjectDuration FromMinutes(decimal minutes, ProjectDurationUnit unit, bool elapsed, bool estimated, decimal factor) {
        if (factor <= 0) throw new InvalidDataException("Project duration conversion requires positive working-time settings.");
        decimal ticks = checked(minutes * TimeSpan.TicksPerMinute);
        decimal rounded = decimal.Round(ticks, 0, MidpointRounding.AwayFromZero);
        if (rounded >= long.MinValue && rounded <= long.MaxValue &&
            (ticks == rounded || rounded / TimeSpan.TicksPerMinute == minutes))
            return FromTicks((long)rounded, unit, elapsed, estimated, factor);
        var duration = new ProjectDuration(minutes / factor, unit, elapsed, estimated);
        return duration.Minutes(factor) == minutes ? duration :
            new ProjectDuration(duration.Value, unit, elapsed, estimated, new DecodedQuantity(minutes, factor, false));
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
    internal decimal Minutes(decimal factor) => _decoded == null ? checked(Value * factor) :
        _decoded.IsTicks ? checked(_decoded.Quantity * factor / _decoded.UnitFactor) / TimeSpan.TicksPerMinute : checked(_decoded.Quantity * factor / _decoded.UnitFactor);
    internal decimal Ticks(decimal factor) => _decoded == null ? checked(Value * factor * TimeSpan.TicksPerMinute) :
        _decoded.IsTicks ? checked(_decoded.Quantity * factor / _decoded.UnitFactor) : checked(_decoded.Quantity * factor / _decoded.UnitFactor * TimeSpan.TicksPerMinute);
    internal decimal ScaledMinutes(decimal factor, decimal scale) => _decoded?.IsTicks == true
        ? checked(_decoded.Quantity * factor / _decoded.UnitFactor * scale) / TimeSpan.TicksPerMinute : checked(Minutes(factor) * scale);
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
    public ProjectDuration Estimated(bool value = true) => new ProjectDuration(Value, Unit, IsElapsed, value, _decoded);
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
    private readonly DecodedTicks? _decoded;
    private sealed class DecodedTicks {
        internal readonly long Value;
        internal DecodedTicks(long ticks) { Value = ticks; }
    }
    /// <summary>Creates an amount of work in minutes.</summary>
    public ProjectWork(decimal minutes) { if (minutes < 0) throw new ArgumentOutOfRangeException(nameof(minutes)); Minutes = minutes; _decoded = null; }
    private ProjectWork(decimal minutes, long ticks) : this(minutes) { _decoded = new DecodedTicks(ticks); }
    internal static ProjectWork FromTicks(long ticks) {
        var work = new ProjectWork(ticks / (decimal)TimeSpan.TicksPerMinute);
        return work.Ticks == ticks ? work : new ProjectWork(work.Minutes, ticks);
    }
    internal static ProjectWork FromMinutes(decimal minutes) {
        var work = new ProjectWork(minutes);
        if (minutes > long.MaxValue / (decimal)TimeSpan.TicksPerMinute) return work;
        decimal ticks = minutes * TimeSpan.TicksPerMinute;
        long rounded = checked((long)decimal.Round(ticks, 0, MidpointRounding.AwayFromZero));
        return ticks == rounded || rounded / (decimal)TimeSpan.TicksPerMinute == minutes ? FromTicks(rounded) : work;
    }
    internal static ProjectWork Sum(IEnumerable<ProjectWork> values) {
        decimal minutes = 0, ticks = 0; bool tickRange = true;
        foreach (var value in values) {
            minutes = checked(minutes + value.Minutes);
            if (tickRange) {
                try { ticks = checked(ticks + value.Ticks); }
                catch (OverflowException) { tickRange = false; }
            }
        }
        return tickRange ? FromQuantities(minutes, ticks) : new ProjectWork(minutes);
    }
    internal static ProjectWork Subtract(ProjectWork total, ProjectWork part) {
        decimal minutes = checked(total.Minutes - part.Minutes);
        try { return FromQuantities(minutes, checked(total.Ticks - part.Ticks)); }
        catch (OverflowException) { return new ProjectWork(minutes); }
    }
    internal static ProjectWork Add(ProjectWork first, ProjectWork second) {
        decimal minutes = checked(first.Minutes + second.Minutes);
        try { return FromQuantities(minutes, checked(first.Ticks + second.Ticks)); }
        catch (OverflowException) { return new ProjectWork(minutes); }
    }
    internal ProjectWork Scale(decimal factor) {
        return MultiplyDivide(factor, 1);
    }
    internal ProjectWork MultiplyDivide(decimal numerator, decimal denominator) {
        decimal minutes = checked(Minutes * numerator / denominator);
        try { return FromQuantities(minutes, checked(Ticks * numerator / denominator)); }
        catch (OverflowException) { return new ProjectWork(minutes); }
    }
    /// <summary>Computes a dimensionless ratio from stored ticks when they fit the decimal range.</summary>
    internal decimal Ratio(ProjectWork total, decimal scale = 1) {
        try { return checked(Ticks * scale / total.Ticks); }
        catch (OverflowException) { return checked(Minutes * scale / total.Minutes); }
    }
    private static ProjectWork FromQuantities(decimal minutes, decimal ticks) =>
        ticks >= 0 && ticks <= long.MaxValue && ticks == decimal.Truncate(ticks) ? FromTicks((long)ticks) : new ProjectWork(minutes);
    internal decimal Ticks => _decoded == null ? checked(Minutes * TimeSpan.TicksPerMinute) : _decoded.Value;
    internal decimal ScaledMinutes(decimal scale) => _decoded != null ? _decoded.Value * scale / TimeSpan.TicksPerMinute : checked(Minutes * scale);
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
