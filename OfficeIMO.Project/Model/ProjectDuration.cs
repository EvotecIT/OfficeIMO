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
    /// <summary>Creates a duration. Negative values are reserved for dependency lag, not task duration.</summary>
    public ProjectDuration(decimal value, ProjectDurationUnit unit, bool elapsed = false, bool estimated = false) {
        if (!Enum.IsDefined(typeof(ProjectDurationUnit), unit)) throw new ArgumentOutOfRangeException(nameof(unit));
        Value = value; Unit = unit; IsElapsed = elapsed; IsEstimated = estimated;
    }
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
    public ProjectDuration Estimated(bool value = true) => new ProjectDuration(Value, Unit, IsElapsed, value);
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
    /// <summary>Creates an amount of work in minutes.</summary>
    public ProjectWork(decimal minutes) { if (minutes < 0) throw new ArgumentOutOfRangeException(nameof(minutes)); Minutes = minutes; }
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
