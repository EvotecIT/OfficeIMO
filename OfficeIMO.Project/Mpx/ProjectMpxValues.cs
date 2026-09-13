using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Project;

/// <summary>Explicit MPX locale and unit interpretation; never uses the host's regional settings or current date.</summary>
internal sealed class ProjectMpxValues {
    internal string Currency = "$", DecimalSeparator = ".", GroupSeparator = ",";
    internal int DateOrder, DefaultDurationUnit = 2, DefaultWorkUnit = 1;
    internal string DateSeparator = "/", TimeSeparator = ":", Am = "AM", Pm = "PM";
    internal int MinutesPerDay = 480, MinutesPerWeek = 2400;
    internal TimeSpan DefaultTime = TimeSpan.FromHours(8);
    internal Action<string>? OnLoss;
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;
    internal static bool Empty(string value) => value.Length == 0 || value.Equals("NA", StringComparison.OrdinalIgnoreCase);
    internal static string Get(string[] record, int index) => index < record.Length ? record[index] : "";
    internal static int Integer(string value) {
        if (!int.TryParse(value, NumberStyles.Integer, Invariant, out int number)) throw new InvalidDataException("Invalid MPX integer: " + value);
        return number;
    }
    internal decimal Number(string value) {
        if (Currency.Length != 0) value = value.Replace(Currency, "").Trim();
        var format = (NumberFormatInfo)Invariant.NumberFormat.Clone();
        format.NumberDecimalSeparator = DecimalSeparator; format.NumberGroupSeparator = GroupSeparator;
        if (!decimal.TryParse(value, NumberStyles.Number | NumberStyles.AllowParentheses, format, out decimal number))
            throw new InvalidDataException("Invalid MPX number: " + value);
        return number;
    }
    internal int Percent(string value) {
        decimal number = Number(value.TrimEnd('%'));
        if (number < 0 || number > 100) throw new InvalidDataException("MPX progress must be between 0 and 100 percent.");
        if (number != decimal.Truncate(number)) OnLoss?.Invoke("Fractional progress is rounded to the model's whole percentage; the original lexical value remains in the MPX source bytes.");
        return (int)decimal.Round(number, 0, MidpointRounding.AwayFromZero);
    }
    internal bool Flag(string value) => value.ToUpperInvariant() switch {
        "1" or "YES" or "TRUE" => true, "0" or "NO" or "FALSE" => false,
        _ => throw new InvalidDataException("Invalid MPX flag: " + value)
    };
    internal ProjectDuration Duration(string value, int? defaultUnit = null) {
        var match = Regex.Match(value, @"^\s*([+-]?[\d.,]+)\s*(e?)(min(?:ute)?s?|m|h(?:our)?s?|d(?:ay)?s?|w(?:eek)?s?|mo(?:nth)?s?)?\s*(\?)?\s*$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
        if (!match.Success) throw new InvalidDataException("Invalid MPX duration: " + value);
        string unit = match.Groups[3].Value.ToLowerInvariant();
        var kind = unit.Length == 0 ? (ProjectDurationUnit)(defaultUnit ?? DefaultDurationUnit) :
            unit.StartsWith("mo", StringComparison.Ordinal) ? ProjectDurationUnit.Month : unit[0] switch {
                'm' => ProjectDurationUnit.Minute, 'h' => ProjectDurationUnit.Hour, 'd' => ProjectDurationUnit.Day, 'w' => ProjectDurationUnit.Week,
                _ => throw new InvalidDataException("Unknown MPX duration unit.")
            };
        return new ProjectDuration(Number(match.Groups[1].Value), kind, match.Groups[2].Success && match.Groups[2].Length != 0, match.Groups[4].Success);
    }
    internal decimal Minutes(ProjectDuration value) => checked(value.Value * ProjectTimeUnits.MinutesPerUnit(value.Unit, value.IsElapsed, MinutesPerDay, MinutesPerWeek));
    internal ProjectWork Work(string value) => new ProjectWork(Minutes(Duration(value, DefaultWorkUnit)));
    internal ProjectUnits Units(string value) => value.EndsWith("%", StringComparison.Ordinal)
        ? ProjectUnits.Percent(Number(value.Substring(0, value.Length - 1))) : ProjectUnits.Fraction(Number(value));
    internal decimal Rate(string value) {
        int slash = value.LastIndexOf('/');
        if (slash < 0) return Number(value);
        decimal minutes = Minutes(Duration("1" + value.Substring(slash + 1)));
        if (minutes <= 0) throw new InvalidDataException("Invalid MPX rate unit.");
        return checked(Number(value.Substring(0, slash)) * 60 / minutes);
    }
    private DateTimeFormatInfo DateFormat() {
        var format = (DateTimeFormatInfo)Invariant.DateTimeFormat.Clone();
        format.DateSeparator = DateSeparator; format.TimeSeparator = TimeSeparator;
        format.AMDesignator = Am; format.PMDesignator = Pm;
        format.Calendar = new GregorianCalendar { TwoDigitYearMax = 2029 };
        return format;
    }
    internal DateTime Date(string value) {
        string date = DateOrder switch { 0 => "M/d/", 1 => "d/M/", 2 => "", _ => throw new InvalidDataException("Invalid MPX date order.") };
        var numeric = DateOrder == 2 ? new[] { "yyyy/M/d", "yy/M/d" } : new[] { date + "yyyy", date + "yy" };
        var baseDates = numeric.Concat(new[] { "d MMMM yyyy", "d MMM yyyy", "d MMMM yy", "d MMM yy", "MMMM d yyyy", "MMM d yyyy", "d-MMM-yyyy", "d-MMM-yy" });
        var dates = baseDates.SelectMany(pattern => new[] { pattern, "ddd " + pattern, "dddd " + pattern, "ddd, " + pattern, "dddd, " + pattern });
        var format = DateFormat();
        foreach (string pattern in dates) {
            // A date-only record retains midnight. The declared default clock remains a
            // separate setting; loading does not apply application scheduling defaults.
            if (DateTime.TryParseExact(value, pattern, format, DateTimeStyles.AllowWhiteSpaces, out var result)) return result.Date;
            foreach (string time in new[] { " H:m", " H:m:s", " h:m tt", " h:m:s tt", " h tt" })
                if (DateTime.TryParseExact(value, pattern + time, format, DateTimeStyles.AllowWhiteSpaces, out result)) return DateTime.SpecifyKind(result, DateTimeKind.Unspecified);
        }
        throw new InvalidDataException("Invalid MPX date for the declared date order: " + value);
    }
    internal TimeSpan Time(string value) {
        if (DateTime.TryParseExact(value, new[] { "H:m", "H:m:s", "h:m tt", "h:m:s tt", "h tt" }, DateFormat(), DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault, out var result))
            return result.TimeOfDay;
        throw new InvalidDataException("Invalid MPX clock time: " + value);
    }
    internal static string Text(object? value) => value switch {
        null => "", bool flag => flag ? "Yes" : "No", DateTime date => date.ToString("yyyy/MM/dd HH:mm", Invariant),
        TimeSpan time => time.ToString(@"hh\:mm", Invariant),
        ProjectDuration duration => duration.Value.ToString(Invariant) + (duration.IsElapsed ? "e" : "") +
            (duration.Unit switch { ProjectDurationUnit.Minute => "m", ProjectDurationUnit.Hour => "h", ProjectDurationUnit.Day => "d", ProjectDurationUnit.Week => "w", _ => "mo" }) + (duration.IsEstimated ? "?" : ""),
        ProjectWork work => work.Minutes.ToString(Invariant) + "m", ProjectUnits units => units.Value.ToString(Invariant),
        IFormattable number => number.ToString(null, Invariant), _ => value.ToString() ?? ""
    };
}
