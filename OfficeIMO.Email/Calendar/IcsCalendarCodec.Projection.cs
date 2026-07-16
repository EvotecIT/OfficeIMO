namespace OfficeIMO.Email;

internal static partial class IcsCalendarCodec {
    private static readonly HashSet<string> RepeatableProjectedCalendarProperties = new HashSet<string>(
        StringComparer.OrdinalIgnoreCase) {
            "ATTENDEE", "CATEGORIES", "CONTACT", "X-OFFICEIMO-COMPANY"
        };

    private static readonly HashSet<string> ProjectedCalendarExtensions = new HashSet<string>(
        StringComparer.OrdinalIgnoreCase) {
            "X-MICROSOFT-CDO-BUSYSTATUS", "X-MICROSOFT-DISALLOW-COUNTER",
            "X-OFFICEIMO-ACCEPTANCE-STATE", "X-OFFICEIMO-ACTUAL-EFFORT", "X-OFFICEIMO-ASSIGNER",
            "X-OFFICEIMO-BILLING-INFORMATION", "X-OFFICEIMO-CLIENT-INTENT", "X-OFFICEIMO-COMMON-END",
            "X-OFFICEIMO-COMMON-START", "X-OFFICEIMO-COMPANY", "X-OFFICEIMO-ESTIMATED-EFFORT",
            "X-OFFICEIMO-MEETING-STATUS", "X-OFFICEIMO-MILEAGE", "X-OFFICEIMO-ORDINAL",
            "X-OFFICEIMO-OWNERSHIP", "X-OFFICEIMO-REMINDER-SET", "X-OFFICEIMO-REMINDER-SIGNAL-TIME",
            "X-OFFICEIMO-REMINDER-TIME", "X-OFFICEIMO-RESPONSE-STATUS",
            "X-OFFICEIMO-SEND-STATUS-ON-COMPLETE", "X-OFFICEIMO-SEND-UPDATES", "X-OFFICEIMO-TASK-MODE",
            "X-OFFICEIMO-TASK-OWNER", "X-OFFICEIMO-TASK-STATE", "X-OFFICEIMO-TASK-VERSION",
            "X-OFFICEIMO-TEAM-TASK", "X-OFFICEIMO-TIMEZONE-DESCRIPTION",
            "X-OFFICEIMO-TODO-ORDINAL-DATE", "X-OFFICEIMO-TODO-SUBORDINAL"
        };

    private static bool HasIncompleteStoreProjection(IEnumerable<IcsProperty> properties,
        IEnumerable<IcsProperty> activeProperties, IReadOnlyList<IcsProperty> alarmProperties, bool isEvent,
        string? envelopeSubject, EmailAddress? envelopeFrom) {
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
        IcsProperty[] versions = properties.Where(property => property.Name == "VERSION").ToArray();
        bool hasUnsupportedVersion = versions.Length != 1 ||
            !versions[0].Value.Trim().Equals("2.0", StringComparison.OrdinalIgnoreCase);
        bool missingEventStart = isEvent && !activeProperties.Any(property => property.Name == "DTSTART");
        string? calendarSummary = Unescape(GetValue(activeProperties, "SUMMARY"));
        string? reminderDescription = string.IsNullOrWhiteSpace(calendarSummary)
            ? envelopeSubject
            : calendarSummary;
        bool hasConflictingEnvelopeSubject = !string.IsNullOrWhiteSpace(envelopeSubject) &&
            !string.Equals(envelopeSubject, calendarSummary, StringComparison.Ordinal);
        bool wouldSynthesizeOrganizer = isEvent && !string.IsNullOrWhiteSpace(envelopeFrom?.Address) &&
            !activeProperties.Any(property => property.Name == "ORGANIZER");
        return calendarRoots != 1 || calendarItems > 1 || alarms > 1 || hasTimeZone || hasUnsupportedComponent ||
            hasUnsupportedVersion || missingEventStart || hasConflictingEnvelopeSubject || wouldSynthesizeOrganizer ||
            HasDuplicateAttendeeAddresses(activeProperties) ||
            HasDuplicateCalendarSingletons(activeProperties) ||
            activeProperties.Any(property => property.Name.IndexOf('.') >= 0) ||
            activeProperties.Any(HasUnpreservedPropertyParameters) ||
            activeProperties.Any(property => property.Name.StartsWith("X-", StringComparison.OrdinalIgnoreCase) &&
                !ProjectedCalendarExtensions.Contains(property.Name)) ||
            HasIncompleteAlarmProjection(alarms, alarmProperties, reminderDescription) ||
            activeProperties.Any(property =>
                property.Name == "RRULE" || property.Name == "RDATE" ||
                property.Name == "EXDATE" || property.Name == "RECURRENCE-ID" ||
                property.Name == "RELATED-TO" || property.Name == "REQUEST-STATUS" ||
                property.Name == "PRODID" &&
                !property.Value.Trim().Equals("-//Evotec//OfficeIMO.Email//EN", StringComparison.Ordinal) ||
                property.Name == "CREATED" || property.Name == "LAST-MODIFIED" ||
                property.Name == "COMMENT" || property.Name == "RESOURCES" || property.Name == "GEO" ||
                isEvent && property.Name == "CONTACT" ||
                property.Name == "CLASS" && !ParseCalendarSensitivity(property.Value).HasValue ||
                isEvent && property.Name == "STATUS" ||
                property.Name == "PRIORITY" || property.Name == "URL" ||
                property.Name == "LOCATION" && (!isEvent || property.Parameters.Count > 0) ||
                !isEvent && property.Name == "SEQUENCE" ||
                !isEvent && IsDateOnlyTaskProperty(property) ||
                property.Name == "ATTENDEE" && HasIncompleteAttendeeProjection(property) ||
                property.Name == "ORGANIZER" && HasIncompleteOrganizerProjection(property) ||
                (property.Name == "ORGANIZER" || property.Name == "ATTENDEE") &&
                IsUnprojectableCalendarAddress(property.Value) ||
                property.Name == "ATTACH" || property.Name == "TRIGGER" &&
                property.Parameters.TryGetValue("RELATED", out string? related) &&
                related.Equals("END", StringComparison.OrdinalIgnoreCase));
    }

    private static bool HasUnpreservedPropertyParameters(IcsProperty property) {
        if (property.Parameters.Count == 0) return false;
        if (property.Name == "ATTENDEE") return HasIncompleteAttendeeProjection(property);
        if (property.Name == "ORGANIZER") return HasIncompleteOrganizerProjection(property);
        if (property.Name == "TRIGGER") return false;
        if (!IsCalendarDateProperty(property.Name)) return true;
        return property.Parameters.Keys.Any(parameter =>
            !parameter.Equals("VALUE", StringComparison.OrdinalIgnoreCase) &&
            !parameter.Equals("TZID", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsCalendarDateProperty(string name) => name == "DTSTART" || name == "DTEND" ||
        name == "DUE" || name == "COMPLETED" || name == "DTSTAMP" ||
        name == "X-OFFICEIMO-COMMON-START" || name == "X-OFFICEIMO-COMMON-END" ||
        name == "X-OFFICEIMO-TODO-ORDINAL-DATE" || name == "X-OFFICEIMO-REMINDER-TIME" ||
        name == "X-OFFICEIMO-REMINDER-SIGNAL-TIME";

    private static bool HasIncompleteTimestampProjection(IEnumerable<IcsProperty> activeProperties,
        EmailDocument document, bool isEvent, IList<EmailDiagnostic> diagnostics, string location) {
        IcsProperty? timestamp = GetProperty(activeProperties, "DTSTAMP");
        if (timestamp == null) return false;
        DateTimeOffset? parsed = ParseDate(timestamp, diagnostics, location, out bool isDateOnly);
        if (!parsed.HasValue || isDateOnly) return true;
        DateTimeOffset expected = document.Date ?? (isEvent
            ? document.Appointment?.Start
            : document.Task?.Start ?? document.Task?.Due) ?? DeterministicEpoch;
        return !string.Equals(FormatUtc(parsed.Value), FormatUtc(expected), StringComparison.Ordinal);
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
        bool roomOrResource = attendee.Parameters.TryGetValue("CUTYPE", out string? calendarUserType) &&
            (calendarUserType.Equals("ROOM", StringComparison.OrdinalIgnoreCase) ||
             calendarUserType.Equals("RESOURCE", StringComparison.OrdinalIgnoreCase));
        foreach (KeyValuePair<string, string> parameter in attendee.Parameters) {
            if (parameter.Key.Equals("CN", StringComparison.OrdinalIgnoreCase)) continue;
            if (parameter.Key.Equals("CUTYPE", StringComparison.OrdinalIgnoreCase)) {
                if (parameter.Value.Equals("INDIVIDUAL", StringComparison.OrdinalIgnoreCase) ||
                    parameter.Value.Equals("ROOM", StringComparison.OrdinalIgnoreCase) ||
                    parameter.Value.Equals("RESOURCE", StringComparison.OrdinalIgnoreCase)) continue;
                return true;
            }
            if (!parameter.Key.Equals("ROLE", StringComparison.OrdinalIgnoreCase)) return true;
            if (roomOrResource
                    ? parameter.Value.Equals("NON-PARTICIPANT", StringComparison.OrdinalIgnoreCase)
                    : parameter.Value.Equals("REQ-PARTICIPANT", StringComparison.OrdinalIgnoreCase) ||
                      parameter.Value.Equals("OPT-PARTICIPANT", StringComparison.OrdinalIgnoreCase)) continue;
            return true;
        }
        return roomOrResource && (!attendee.Parameters.TryGetValue("ROLE", out string? role) ||
            !role.Equals("NON-PARTICIPANT", StringComparison.OrdinalIgnoreCase));
    }

    private static bool HasIncompleteOrganizerProjection(IcsProperty organizer) =>
        organizer.Parameters.Keys.Any(parameter =>
            !parameter.Equals("CN", StringComparison.OrdinalIgnoreCase));

    private static bool HasDuplicateAttendeeAddresses(IEnumerable<IcsProperty> properties) {
        var addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IcsProperty attendee in properties.Where(property => property.Name == "ATTENDEE")) {
            string? address = StripMailTo(attendee.Value);
            if (!string.IsNullOrWhiteSpace(address) && !addresses.Add(address!)) return true;
        }
        return false;
    }

    private static bool HasDuplicateCalendarSingletons(IEnumerable<IcsProperty> properties) =>
        properties.Where(property => !RepeatableProjectedCalendarProperties.Contains(property.Name))
            .GroupBy(property => property.Name, StringComparer.OrdinalIgnoreCase)
            .Any(group => group.Skip(1).Any());

    private static bool IsUnprojectableCalendarAddress(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return true;
        string address = value!.Trim();
        if (!address.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) return true;
        string mailbox = address.Substring(7);
        return mailbox.Length == 0 || mailbox.IndexOf('?') >= 0 || mailbox.IndexOf('#') >= 0;
    }

    private static bool IsStoreProjectableTaskMethod(string? method) => string.IsNullOrWhiteSpace(method) ||
        string.Equals(method, "PUBLISH", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(method, "REQUEST", StringComparison.OrdinalIgnoreCase);

    private static bool IsSupportedComponent(string value) =>
        value.Equals("VCALENDAR", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("VEVENT", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("VTODO", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("VTIMEZONE", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("VALARM", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("STANDARD", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("DAYLIGHT", StringComparison.OrdinalIgnoreCase);
}
