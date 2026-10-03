namespace OfficeIMO.Project;

/// <summary>Shared working/elapsed unit conversion for codecs and explicit calculation.</summary>
internal static class ProjectTimeUnits {
    internal static decimal AddMinutes(decimal first, decimal second) => ProjectWork.Add(ProjectWork.FromMinutes(first), ProjectWork.FromMinutes(second)).Minutes;
    internal static decimal SubtractMinutes(decimal total, decimal part) => total < part ? total - part : ProjectWork.Subtract(ProjectWork.FromMinutes(total), ProjectWork.FromMinutes(part)).Minutes;
    internal static decimal MultiplyDivideMinutes(decimal minutes, decimal numerator, decimal denominator = 1) => ProjectWork.FromMinutes(minutes).MultiplyDivide(numerator, denominator).Minutes;
    /// <summary>Scales time by a ratio of time quantities, preserving canonical ticks before division.</summary>
    internal static decimal ScaleMinutesByRatio(decimal minutes, decimal partMinutes, decimal totalMinutes) {
        try { return MultiplyDivideMinutes(minutes, ProjectWork.FromMinutes(partMinutes).Ticks, ProjectWork.FromMinutes(totalMinutes).Ticks); }
        catch (OverflowException) { return MultiplyDivideMinutes(minutes, partMinutes, totalMinutes); }
    }
    internal static int Percentage(ProjectWork actual, ProjectWork total) => total.Minutes <= 0 ? 0 :
        (int)decimal.Round(Math.Min(100, actual.Ratio(total, 100)), 0, MidpointRounding.AwayFromZero);
    internal static decimal Minutes(ProjectDuration duration, ProjectSettings settings) => duration.Minutes(MinutesPerUnit(duration.Unit, duration.IsElapsed, settings));
    internal static decimal Ticks(ProjectDuration duration, ProjectSettings settings) => duration.Ticks(MinutesPerUnit(duration.Unit, duration.IsElapsed, settings));
    internal static decimal ScaledMinutes(ProjectDuration duration, ProjectSettings settings, decimal scale) => duration.ScaledMinutes(MinutesPerUnit(duration.Unit, duration.IsElapsed, settings), scale);
    internal static decimal MinutesPerUnit(ProjectDurationUnit unit, bool elapsed, ProjectSettings settings) =>
        MinutesPerUnit(unit, elapsed, settings.MinutesPerDay, settings.MinutesPerWeek, settings.DaysPerMonth);

    internal static decimal MinutesPerUnit(ProjectDurationUnit unit, bool elapsed, int? minutesPerDay, int? minutesPerWeek, int? daysPerMonth = 20) => unit switch {
        ProjectDurationUnit.Minute => 1,
        ProjectDurationUnit.Hour => 60,
        ProjectDurationUnit.Day => elapsed ? 1440 : minutesPerDay ?? 480,
        ProjectDurationUnit.Week => elapsed ? 10080 : minutesPerWeek ?? 2400,
        ProjectDurationUnit.Month => elapsed ? 43200 : checked((minutesPerDay ?? 480) * (daysPerMonth ?? 20)),
        _ => throw new InvalidDataException("Unknown duration unit.")
    };
}
