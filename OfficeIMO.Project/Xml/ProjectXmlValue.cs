using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Project;

internal static class ProjectXmlValue {
    private static readonly string[] LocalDateFormats = { "yyyy-MM-dd", "yyyy-MM-dd'T'HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF" };
    private static readonly string[] LocalClockFormats = { "HH:mm:ss", "HH:mm:ss.FFFFFFF" };
    internal static string? Text(string? value) => value;
    internal static string? Integer(int? value) => value?.ToString(CultureInfo.InvariantCulture);
    internal static string? Number(decimal? value) => value?.ToString(CultureInfo.InvariantCulture);
    internal static string? Money(decimal? value) => value.HasValue ? Number(checked(value.Value * 100)) : null;
    internal static string? Boolean(bool? value) => value.HasValue ? (value.Value ? "1" : "0") : null;
    internal static string? Identifier(Guid? value) => value?.ToString("D").ToUpperInvariant();
    internal static string? Date(DateTime? value) {
        if (!value.HasValue) return null;
        if (value.Value.Kind != DateTimeKind.Unspecified) throw new InvalidDataException("Project dates must use DateTimeKind.Unspecified; timezone conversion must be explicit.");
        return value.Value.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture);
    }
    internal static string? Clock(TimeSpan? value) {
        if (!value.HasValue) return null;
        if (value < TimeSpan.Zero || value >= TimeSpan.FromDays(1)) throw new InvalidDataException("A project clock time must be within one day.");
        return DateTime.MinValue.Add(value.Value).ToString("HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture);
    }
    internal static string? Work(ProjectWork? value) => value.HasValue ? Span(MinutesToSpan(value.Value.Minutes)) : null;
    internal static string? Units(ProjectUnits? value) => value.HasValue ? Number(value.Value.Value) : null;
    internal static string Span(TimeSpan value) {
        // Project's importer expects hours/minutes/seconds even above 24 hours. The
        // otherwise valid XML duration P1DT16H imports as zero in Project 2024.
        decimal ticks = value.Ticks;
        string sign = ticks < 0 ? "-" : "";
        ticks = Math.Abs(ticks);
        decimal hours = decimal.Truncate(ticks / TimeSpan.TicksPerHour);
        ticks %= TimeSpan.TicksPerHour;
        decimal minutes = decimal.Truncate(ticks / TimeSpan.TicksPerMinute);
        decimal seconds = (ticks % TimeSpan.TicksPerMinute) / TimeSpan.TicksPerSecond;
        return sign + "PT" + hours.ToString(CultureInfo.InvariantCulture) + "H" + minutes.ToString(CultureInfo.InvariantCulture) + "M" + seconds.ToString("0.#######", CultureInfo.InvariantCulture) + "S";
    }
    internal static int ParseInt(string value) => XmlConvert.ToInt32(value);
    internal static decimal ParseNumber(string value) => XmlConvert.ToDecimal(value);
    internal static decimal ParseMoney(string value) => ParseNumber(value) / 100;
    internal static bool ParseBool(string value) => XmlConvert.ToBoolean(value);
    internal static Guid ParseGuid(string value) => System.Guid.Parse(value);
    internal static DateTime ParseDate(string value) {
        if (!DateTime.TryParseExact(value, LocalDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
            throw new InvalidDataException("Project dates require a local ISO value without a timezone suffix.");
        return DateTime.SpecifyKind(parsed, DateTimeKind.Unspecified);
    }
    internal static TimeSpan ParseClock(string value) {
        if (!DateTime.TryParseExact(value, LocalClockFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
            throw new InvalidDataException("Project clock times require a local ISO value without a timezone suffix.");
        return parsed.TimeOfDay;
    }
    internal static ProjectWork ParseWork(string value) => new ProjectWork((decimal)XmlConvert.ToTimeSpan(value).Ticks / TimeSpan.TicksPerMinute);
    internal static ProjectUnits ParseUnits(string value) => ProjectUnits.Fraction(ParseNumber(value));
    internal static TimeSpan MinutesToSpan(decimal minutes) => TimeSpan.FromTicks(checked((long)decimal.Round(minutes * TimeSpan.TicksPerMinute, 0, MidpointRounding.AwayFromZero)));

    internal static decimal MinutesPerUnit(ProjectDurationUnit unit, bool elapsed, ProjectDocument document) => ProjectTimeUnits.MinutesPerUnit(unit, elapsed, document.Settings);
    internal static string? Duration(ProjectDuration? duration, ProjectDocument document) => duration.HasValue
        ? Span(MinutesToSpan(checked(duration.Value.Value * MinutesPerUnit(duration.Value.Unit, duration.Value.IsElapsed, document)))) : null;
    internal static int DurationFormat(ProjectDuration duration) => 3 + (int)duration.Unit * 2 + (duration.IsElapsed ? 1 : 0) + (duration.IsEstimated ? 32 : 0);

    internal static bool TryTaskDurationFormat(ProjectTask task, out int? format) {
        format = null;
        foreach (var duration in new[] { task.Duration, task.ActualDuration, task.RemainingDuration }) {
            if (!duration.HasValue) continue;
            int current = DurationFormat(duration.Value);
            if (format.HasValue && format.Value != current) return false;
            format = current;
        }
        return true;
    }
    internal static ProjectDuration ParseDuration(string text, int? format, ProjectDocument document) {
        int code = format ?? 7;
        bool estimated = code >= 32;
        if (estimated) code -= 32;
        // Project 2024 emits 21/53 ("null format") even for ordinary automatic tasks.
        if (code == 21) code = 7;
        if (code < 3 || code > 12) throw new InvalidDataException("Unsupported duration format " + format + ".");
        bool elapsed = code % 2 == 0;
        var unit = (ProjectDurationUnit)((code - 3) / 2);
        decimal minutes = (decimal)XmlConvert.ToTimeSpan(text).Ticks / TimeSpan.TicksPerMinute;
        decimal factor = MinutesPerUnit(unit, elapsed, document);
        if (factor <= 0) throw new InvalidDataException("Project duration conversion requires positive working-time settings.");
        return new ProjectDuration(minutes / factor, unit, elapsed, estimated);
    }

    internal static string Location(XElement element) {
        var uid = element.Element(element.Name.Namespace + "UID")?.Value;
        return "/" + element.Name.LocalName + (uid == null ? "" : "[UID=" + uid + "]");
    }
}
