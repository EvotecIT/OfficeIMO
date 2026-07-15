namespace OfficeIMO.Email;

internal static partial class IcsCalendarCodec {
    private static List<IcsProperty> ParseProperties(string text) {
        string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var unfolded = new List<string>();
        foreach (string line in normalized.Split('\n')) {
            if (line.Length > 0 && (line[0] == ' ' || line[0] == '\t') && unfolded.Count > 0) {
                unfolded[unfolded.Count - 1] += line.Substring(1);
            } else {
                unfolded.Add(line);
            }
        }

        var result = new List<IcsProperty>();
        foreach (string line in unfolded) {
            int colon = FindUnquotedSeparator(line, ':');
            if (colon <= 0) continue;
            IReadOnlyList<string> tokens = SplitUnquoted(line.Substring(0, colon), ';');
            var property = new IcsProperty(tokens[0].Trim().ToUpperInvariant(), line.Substring(colon + 1));
            for (int index = 1; index < tokens.Count; index++) {
                int equals = FindUnquotedSeparator(tokens[index], '=');
                if (equals > 0) property.Parameters[tokens[index].Substring(0, equals).Trim()] =
                    tokens[index].Substring(equals + 1).Trim().Trim('"');
            }
            result.Add(property);
        }
        return result;
    }

    private static IReadOnlyList<IcsProperty> SelectActiveComponentProperties(
        IReadOnlyList<IcsProperty> properties, string componentName) {
        var selected = new List<IcsProperty>();
        var components = new List<string>();
        int activeDepth = -1;
        bool selectedComponent = false;
        foreach (IcsProperty property in properties) {
            if (property.Name == "BEGIN") {
                components.Add(property.Value.Trim().ToUpperInvariant());
                if (!selectedComponent && property.Value.Equals(componentName, StringComparison.OrdinalIgnoreCase)) {
                    activeDepth = components.Count;
                    selectedComponent = true;
                }
                continue;
            }
            if (property.Name == "END") {
                if (components.Count == activeDepth &&
                    property.Value.Equals(componentName, StringComparison.OrdinalIgnoreCase)) activeDepth = -1;
                if (components.Count > 0) components.RemoveAt(components.Count - 1);
                continue;
            }

            bool calendarProperty = components.Count == 1 &&
                components[0].Equals("VCALENDAR", StringComparison.OrdinalIgnoreCase);
            bool componentProperty = activeDepth > 0 && components.Count == activeDepth;
            bool alarmTrigger = activeDepth > 0 && components.Count == activeDepth + 1 &&
                components[components.Count - 1].Equals("VALARM", StringComparison.OrdinalIgnoreCase) &&
                property.Name == "TRIGGER";
            if (calendarProperty || componentProperty || alarmTrigger) selected.Add(property);
        }
        return selected;
    }

    private static int FindUnquotedSeparator(string value, char separator) {
        bool quoted = false;
        bool escaped = false;
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (escaped) escaped = false;
            else if (character == '\\') escaped = true;
            else if (character == '"') quoted = !quoted;
            else if (!quoted && character == separator) return index;
        }
        return -1;
    }

    private static IReadOnlyList<string> SplitUnquoted(string value, char separator) {
        var result = new List<string>();
        int start = 0;
        while (start <= value.Length) {
            int relative = FindUnquotedSeparator(value.Substring(start), separator);
            if (relative < 0) {
                result.Add(value.Substring(start));
                break;
            }
            result.Add(value.Substring(start, relative));
            start += relative + 1;
        }
        return result;
    }

    private static DateTimeOffset? ParseDate(IcsProperty? property, IList<EmailDiagnostic> diagnostics,
        string location, out bool isDateOnly) {
        isDateOnly = property != null && property.Parameters.TryGetValue("VALUE", out string? valueType) &&
            string.Equals(valueType, "DATE", StringComparison.OrdinalIgnoreCase);
        if (property == null || string.IsNullOrWhiteSpace(property.Value)) return null;
        string value = property.Value.Trim();
        if (isDateOnly && DateTime.TryParseExact(value, "yyyyMMdd", CultureInfo.InvariantCulture,
            DateTimeStyles.None, out DateTime date)) return new DateTimeOffset(date, TimeSpan.Zero);
        if (DateTimeOffset.TryParseExact(value, new[] { "yyyyMMdd'T'HHmmss'Z'", "yyyyMMdd'T'HHmm'Z'" },
            CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
            out DateTimeOffset utc)) return utc;
        if (DateTime.TryParseExact(value, new[] { "yyyyMMdd'T'HHmmss", "yyyyMMdd'T'HHmm" },
            CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime local)) {
            if (property.Parameters.TryGetValue("TZID", out string? timeZoneId)) {
                try {
                    TimeZoneInfo zone = TimeZoneInfo.FindSystemTimeZoneById(timeZoneId);
                    return new DateTimeOffset(local, zone.GetUtcOffset(local));
                } catch (TimeZoneNotFoundException) {
                } catch (InvalidTimeZoneException) {
                }
                diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_TIMEZONE_UNRESOLVED",
                    string.Concat("The iCalendar time zone '", timeZoneId,
                        "' could not be resolved and the local clock value was interpreted as UTC."),
                    EmailDiagnosticSeverity.Warning, location));
                return new DateTimeOffset(local, TimeSpan.Zero);
            }
            diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_FLOATING_TIME",
                "A floating iCalendar time could not be associated with a known time zone and was interpreted as UTC.",
                EmailDiagnosticSeverity.Warning, location));
            return new DateTimeOffset(local, TimeSpan.Zero);
        }
        diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_DATE_INVALID",
            string.Concat("The iCalendar value '", value, "' could not be parsed."),
            EmailDiagnosticSeverity.Warning, location));
        return null;
    }
}
