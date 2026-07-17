namespace OfficeIMO.Email;

internal static partial class IcsCalendarCodec {
    private static List<IcsProperty> ParseProperties(string text) {
        var result = new List<IcsProperty>();
        foreach (ContentLineComponent root in ContentLineCodec.Parse(text, ContentLineReaderOptions.Default))
            FlattenComponent(root, result);
        return result;
    }

    private static void FlattenComponent(ContentLineComponent component, ICollection<IcsProperty> result) {
        result.Add(new IcsProperty("BEGIN", component.Name));
        foreach (ContentLineProperty source in component.Properties) {
            string name = source.Group == null ? source.Name : string.Concat(source.Group, ".", source.Name);
            var property = new IcsProperty(name.ToUpperInvariant(), source.Value);
            foreach (ContentLineParameter parameter in source.Parameters) {
                if (parameter.Values.Count > 0)
                    property.Parameters[parameter.Name] = string.Join(",", parameter.Values);
            }
            result.Add(property);
        }
        foreach (ContentLineComponent child in component.Components) FlattenComponent(child, result);
        result.Add(new IcsProperty("END", component.Name));
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

    private static IReadOnlyList<IcsProperty> SelectActiveAlarmProperties(
        IReadOnlyList<IcsProperty> properties, string componentName) {
        var selected = new List<IcsProperty>();
        var components = new List<string>();
        int activeComponentDepth = -1;
        int activeAlarmDepth = -1;
        bool selectedComponent = false;
        foreach (IcsProperty property in properties) {
            if (property.Name == "BEGIN") {
                components.Add(property.Value.Trim().ToUpperInvariant());
                if (!selectedComponent && property.Value.Equals(componentName, StringComparison.OrdinalIgnoreCase)) {
                    activeComponentDepth = components.Count;
                    selectedComponent = true;
                } else if (activeComponentDepth > 0 && components.Count == activeComponentDepth + 1 &&
                           property.Value.Equals("VALARM", StringComparison.OrdinalIgnoreCase)) {
                    activeAlarmDepth = components.Count;
                }
                continue;
            }
            if (property.Name == "END") {
                if (components.Count == activeAlarmDepth &&
                    property.Value.Equals("VALARM", StringComparison.OrdinalIgnoreCase)) activeAlarmDepth = -1;
                if (components.Count == activeComponentDepth &&
                    property.Value.Equals(componentName, StringComparison.OrdinalIgnoreCase)) activeComponentDepth = -1;
                if (components.Count > 0) components.RemoveAt(components.Count - 1);
                continue;
            }
            if (activeAlarmDepth > 0 && components.Count == activeAlarmDepth) selected.Add(property);
        }
        return selected;
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
