namespace OfficeIMO.Email;

internal static partial class IcsCalendarCodec {
    private static bool HasIncompleteStoreProjection(IEnumerable<IcsProperty> properties,
        IEnumerable<IcsProperty> activeProperties, IReadOnlyList<IcsProperty> alarmProperties, bool isEvent,
        string? envelopeSubject) {
        int calendarItems = properties.Count(property => property.Name == "BEGIN" &&
            (property.Value.Equals("VEVENT", StringComparison.OrdinalIgnoreCase) ||
             property.Value.Equals("VTODO", StringComparison.OrdinalIgnoreCase)));
        int calendarRoots = properties.Count(property => property.Name == "BEGIN" &&
            property.Value.Equals("VCALENDAR", StringComparison.OrdinalIgnoreCase));
        int alarms = properties.Count(property => property.Name == "BEGIN" &&
            property.Value.Equals("VALARM", StringComparison.OrdinalIgnoreCase));
        bool hasTimeZone = properties.Any(property => property.Name == "BEGIN" &&
            property.Value.Equals("VTIMEZONE", StringComparison.OrdinalIgnoreCase));
        bool hasUnsupportedComponent = properties.Any(property => property.Name == "BEGIN" &&
            !IsSupportedComponent(property.Value));
        string? calendarSummary = Unescape(GetValue(activeProperties, "SUMMARY"));
        string? reminderDescription = string.IsNullOrWhiteSpace(calendarSummary)
            ? envelopeSubject
            : calendarSummary;
        return calendarRoots != 1 || calendarItems > 1 || alarms > 1 || hasTimeZone || hasUnsupportedComponent ||
            HasIncompleteAlarmProjection(alarms, alarmProperties, reminderDescription) ||
            activeProperties.Any(property =>
                property.Name == "RRULE" || property.Name == "RDATE" ||
                property.Name == "EXDATE" || property.Name == "RECURRENCE-ID" ||
                property.Name == "CLASS" && !ParseCalendarSensitivity(property.Value).HasValue ||
                isEvent && property.Name == "STATUS" ||
                property.Name == "PRIORITY" || property.Name == "URL" ||
                !isEvent && IsDateOnlyTaskProperty(property) ||
                property.Name == "ATTENDEE" && HasIncompleteAttendeeProjection(property) ||
                property.Name == "ATTACH" || property.Name == "TRIGGER" &&
                property.Parameters.TryGetValue("RELATED", out string? related) &&
                related.Equals("END", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsDateOnlyTaskProperty(IcsProperty property) {
        if (property.Name != "DTSTART" && property.Name != "DUE" && property.Name != "COMPLETED") return false;
        if (property.Parameters.TryGetValue("VALUE", out string? valueType) &&
            valueType.Equals("DATE", StringComparison.OrdinalIgnoreCase)) return true;
        string value = property.Value.Trim();
        return value.Length == 8 && value.All(character => character >= '0' && character <= '9');
    }

    private static bool HasIncompleteAlarmProjection(int alarmCount,
        IReadOnlyList<IcsProperty> alarmProperties, string? reminderDescription) {
        if (alarmCount == 0) return false;
        if (alarmCount != 1) return true;
        IcsProperty[] actions = alarmProperties.Where(property => property.Name == "ACTION").ToArray();
        IcsProperty[] triggers = alarmProperties.Where(property => property.Name == "TRIGGER").ToArray();
        if (actions.Length != 1 || !actions[0].Value.Trim().Equals("DISPLAY", StringComparison.OrdinalIgnoreCase) ||
            triggers.Length != 1) return true;
        bool absolute = triggers[0].Parameters.TryGetValue("VALUE", out string? valueType) &&
            valueType.Equals("DATE-TIME", StringComparison.OrdinalIgnoreCase);
        if (absolute) {
            if (string.IsNullOrWhiteSpace(triggers[0].Value) || triggers[0].Parameters.Any(parameter =>
                    !parameter.Key.Equals("VALUE", StringComparison.OrdinalIgnoreCase))) return true;
        } else {
            if (!IcsDurationCodec.Parse(triggers[0].Value).HasValue || triggers[0].Parameters.Any(parameter =>
                    parameter.Key.Equals("RELATED", StringComparison.OrdinalIgnoreCase)
                        ? !parameter.Value.Equals("START", StringComparison.OrdinalIgnoreCase)
                        : !parameter.Key.Equals("VALUE", StringComparison.OrdinalIgnoreCase) ||
                          !parameter.Value.Equals("DURATION", StringComparison.OrdinalIgnoreCase))) return true;
        }
        IcsProperty[] descriptions = alarmProperties.Where(property => property.Name == "DESCRIPTION").ToArray();
        string expectedDescription = string.IsNullOrWhiteSpace(reminderDescription)
            ? "Reminder"
            : reminderDescription!;
        return descriptions.Length > 1 || descriptions.Length == 1 &&
                !string.Equals(Unescape(descriptions[0].Value), expectedDescription, StringComparison.Ordinal) ||
            alarmProperties.Any(property => property.Name != "ACTION" && property.Name != "TRIGGER" &&
                property.Name != "DESCRIPTION");
    }

    private static bool HasIncompleteAttendeeProjection(IcsProperty attendee) {
        foreach (KeyValuePair<string, string> parameter in attendee.Parameters) {
            if (parameter.Key.Equals("CN", StringComparison.OrdinalIgnoreCase)) continue;
            if (parameter.Key.Equals("CUTYPE", StringComparison.OrdinalIgnoreCase)) {
                if (parameter.Value.Equals("INDIVIDUAL", StringComparison.OrdinalIgnoreCase) ||
                    parameter.Value.Equals("ROOM", StringComparison.OrdinalIgnoreCase) ||
                    parameter.Value.Equals("RESOURCE", StringComparison.OrdinalIgnoreCase)) continue;
                return true;
            }
            if (!parameter.Key.Equals("ROLE", StringComparison.OrdinalIgnoreCase)) return true;
            bool roomOrResource = attendee.Parameters.TryGetValue("CUTYPE", out string? calendarUserType) &&
                (calendarUserType.Equals("ROOM", StringComparison.OrdinalIgnoreCase) ||
                 calendarUserType.Equals("RESOURCE", StringComparison.OrdinalIgnoreCase));
            if (roomOrResource
                    ? parameter.Value.Equals("NON-PARTICIPANT", StringComparison.OrdinalIgnoreCase)
                    : parameter.Value.Equals("REQ-PARTICIPANT", StringComparison.OrdinalIgnoreCase) ||
                      parameter.Value.Equals("OPT-PARTICIPANT", StringComparison.OrdinalIgnoreCase)) continue;
            return true;
        }
        return false;
    }

    private static bool IsSupportedComponent(string value) =>
        value.Equals("VCALENDAR", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("VEVENT", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("VTODO", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("VTIMEZONE", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("VALARM", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("STANDARD", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("DAYLIGHT", StringComparison.OrdinalIgnoreCase);
}
