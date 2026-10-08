using System.Text.RegularExpressions;

namespace OfficeIMO.OpenDocument;

// Parse civil components explicitly: offset-less values must never acquire the host's time zone.
internal readonly struct OdfDateTimeFieldValue {
    private static readonly Regex ValuePattern = new(@"\A(?:(?<date>[0-9]{4}-[0-9]{2}-[0-9]{2})(?:T(?<clock>[0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?))?|(?<clock>[0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?))(?<zone>Z|[+-][0-9]{2}:[0-9]{2})?\z", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
    private static readonly Regex DurationPattern = new(@"\A(?<negative>-)?P(?:(?<years>[0-9]+)Y)?(?:(?<months>[0-9]+)M)?(?:(?<days>[0-9]+)D)?(?:T(?:(?<hours>[0-9]+)H)?(?:(?<minutes>[0-9]+)M)?(?:(?<seconds>[0-9]+(?:\.[0-9]+)?)S)?)?\z", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
    internal OdfDateTimeFieldValue(DateTime civil, TimeSpan clock) { Civil = civil; Clock = clock; }
    internal DateTime Civil { get; }
    internal TimeSpan Clock { get; }
    internal static OdfDateTimeFieldValue FromTimestamp(DateTimeOffset timestamp) => new(timestamp.DateTime, timestamp.TimeOfDay);

    internal static bool TryParse(string? lexical, OdfTextFieldKind kind, out OdfDateTimeFieldValue value) {
        value = default;
        if (lexical == null || lexical.Length > 128) return false;
        Match match = ValuePattern.Match(lexical.Trim());
        if (!match.Success || kind == OdfTextFieldKind.Date && !match.Groups["date"].Success ||
            kind == OdfTextFieldKind.Time && !match.Groups["clock"].Success) return false;
        string zone = match.Groups["zone"].Value;
        if (zone.Length > 1) {
            int hours = int.Parse(zone.Substring(1, 2), CultureInfo.InvariantCulture);
            int minutes = int.Parse(zone.Substring(4, 2), CultureInfo.InvariantCulture);
            if (hours > 14 || minutes > 59 || hours == 14 && minutes != 0) return false;
        }
        try {
            DateTime date = match.Groups["date"].Success
                ? DateTime.ParseExact(match.Groups["date"].Value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None)
                : new DateTime(2000, 1, 1);
            TimeSpan clock = TimeSpan.Zero;
            if (match.Groups["clock"].Success) {
                string raw = match.Groups["clock"].Value;
                int hour = int.Parse(raw.Substring(0, 2), CultureInfo.InvariantCulture);
                int minute = int.Parse(raw.Substring(3, 2), CultureInfo.InvariantCulture);
                int second = int.Parse(raw.Substring(6, 2), CultureInfo.InvariantCulture);
                string fraction = raw.Length > 8 ? raw.Substring(9) : string.Empty;
                if (hour > 24 || minute > 59 || second > 59 || fraction.Length > 7 ||
                    hour == 24 && (minute != 0 || second != 0 || fraction.Any(character => character != '0'))) return false;
                long ticks = fraction.Length == 0 ? 0 : long.Parse(fraction.PadRight(7, '0'), CultureInfo.InvariantCulture);
                clock = new TimeSpan(0, hour, minute, second).Add(TimeSpan.FromTicks(ticks));
            }
            value = new OdfDateTimeFieldValue(date.Add(clock), clock); return true;
        } catch (Exception exception) when (exception is FormatException or ArgumentOutOfRangeException or OverflowException) { return false; }
    }

    internal bool TryAdjust(string? lexical, OdfTextFieldKind kind, out OdfDateTimeFieldValue value) {
        value = this;
        if (lexical == null) return true;
        if (!TryReadAdjustment(lexical, kind, out int years, out int months, out TimeSpan delta)) return false;
        try {
            DateTime adjusted = Civil.AddMonths(checked(years * 12 + months)).Add(delta);
            value = new OdfDateTimeFieldValue(adjusted, Clock.Add(delta)); return true;
        } catch (Exception exception) when (exception is ArgumentOutOfRangeException or OverflowException) { return false; }
    }
    internal static bool TryReadAdjustment(string lexical, OdfTextFieldKind kind, out int years, out int months, out TimeSpan delta) {
        years = 0; months = 0; delta = default;
        if (lexical.Length > 128) return false;
        string raw = lexical.Trim(); Match match = DurationPattern.Match(raw);
        string[] parts = { "years", "months", "days", "hours", "minutes", "seconds" };
        if (!match.Success || !parts.Any(part => match.Groups[part].Success) || raw.EndsWith("T", StringComparison.Ordinal)) return false;
        try {
            string secondsRaw = match.Groups["seconds"].Value;
            int separator = secondsRaw.IndexOf('.');
            if (separator >= 0 && kind == OdfTextFieldKind.Date && secondsRaw.Length - separator - 1 > 7) return false;
            decimal[] amounts = parts.Select(part => {
                if (!match.Groups[part].Success) return 0M;
                string token = match.Groups[part].Value;
                if (part == "seconds" && kind == OdfTextFieldKind.Time && separator >= 0) token = token.Substring(0, separator);
                return decimal.Parse(token, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture);
            }).ToArray();
            int sign = match.Groups["negative"].Success ? -1 : 1;
            if (kind == OdfTextFieldKind.Time && (amounts[0] != 0 || amounts[1] != 0)) return false;
            decimal seconds = amounts[2] * 86400 + amounts[3] * 3600 + amounts[4] * 60 + amounts[5];
            if (kind == OdfTextFieldKind.Time) seconds = decimal.Truncate(seconds / 60) * 60;
            decimal ticks = seconds * TimeSpan.TicksPerSecond;
            if (ticks != decimal.Truncate(ticks)) return false;
            delta = TimeSpan.FromTicks(checked((long)ticks * sign));
            years = checked((int)amounts[0] * sign); months = checked((int)amounts[1] * sign); return true;
        } catch (Exception exception) when (exception is FormatException or OverflowException) { return false; }
    }
}
