namespace OfficeIMO.Email;

internal static class IcsDurationCodec {
    internal static TimeSpan? Parse(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string normalized = value!.Trim().ToUpperInvariant();
        bool negative = normalized.StartsWith("-", StringComparison.Ordinal);
        if (negative || normalized.StartsWith("+", StringComparison.Ordinal)) normalized = normalized.Substring(1);
        if (normalized.Length > 2 && normalized[0] == 'P' && normalized[normalized.Length - 1] == 'W' &&
            long.TryParse(normalized.Substring(1, normalized.Length - 2), NumberStyles.None,
                CultureInfo.InvariantCulture, out long weeks)) {
            try {
                long ticks = checked(weeks * 7L * TimeSpan.TicksPerDay);
                return TimeSpan.FromTicks(negative ? checked(-ticks) : ticks);
            } catch (OverflowException) {
                return null;
            }
        }
        try {
            string xmlValue = negative ? string.Concat("-", normalized) : normalized;
            return System.Xml.XmlConvert.ToTimeSpan(xmlValue);
        } catch (FormatException) {
            return null;
        } catch (OverflowException) {
            return null;
        }
    }

    internal static int? ToWholeMinutes(TimeSpan value, IList<EmailDiagnostic> diagnostics,
        string location, ref bool incomplete, bool invert = false) {
        if (value.Ticks % TimeSpan.TicksPerMinute != 0) {
            AddDiagnosticOnce(diagnostics, "EMAIL_ICALENDAR_DURATION_PRECISION_LOSS",
                "An iCalendar duration contains sub-minute precision that the Outlook model cannot represent exactly.",
                location);
            incomplete = true;
        }
        double minutes = value.TotalMinutes;
        if (invert) minutes = -minutes;
        if (minutes >= int.MinValue && minutes <= int.MaxValue) return (int)minutes;
        AddDiagnosticOnce(diagnostics, "EMAIL_ICALENDAR_DURATION_OUT_OF_RANGE",
            "An iCalendar duration exceeds the supported whole-minute range and was retained only in the semantic source.",
            location);
        incomplete = true;
        return null;
    }

    internal static void ReportOutOfRange(IList<EmailDiagnostic> diagnostics, string location) =>
        AddDiagnosticOnce(diagnostics, "EMAIL_ICALENDAR_DURATION_OUT_OF_RANGE",
            "An iCalendar duration exceeds the supported whole-minute range and was retained only in the semantic source.",
            location);

    internal static string Format(TimeSpan value) {
        string sign = value < TimeSpan.Zero ? "-" : string.Empty;
        TimeSpan absolute = value.Duration();
        var result = new StringBuilder(sign).Append('P');
        if (absolute.Days > 0) result.Append(absolute.Days.ToString(CultureInfo.InvariantCulture)).Append('D');
        if (absolute.Hours > 0 || absolute.Minutes > 0 || absolute.Seconds > 0 || absolute.Days == 0) {
            result.Append('T');
            if (absolute.Hours > 0) result.Append(absolute.Hours.ToString(CultureInfo.InvariantCulture)).Append('H');
            if (absolute.Minutes > 0) result.Append(absolute.Minutes.ToString(CultureInfo.InvariantCulture)).Append('M');
            if (absolute.Seconds > 0 || absolute == TimeSpan.Zero) {
                result.Append(absolute.Seconds.ToString(CultureInfo.InvariantCulture)).Append('S');
            }
        }
        return result.ToString();
    }

    private static void AddDiagnosticOnce(IList<EmailDiagnostic> diagnostics, string code, string message,
        string location) {
        if (diagnostics.Any(diagnostic => diagnostic.Code == code &&
            string.Equals(diagnostic.Location, location, StringComparison.Ordinal))) return;
        diagnostics.Add(new EmailDiagnostic(code, message, EmailDiagnosticSeverity.Warning, location));
    }
}
