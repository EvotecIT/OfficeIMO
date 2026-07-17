namespace OfficeIMO.Email;

public sealed partial class IcsDocument {
    private static readonly HashSet<string> RecurrenceFrequencies = new HashSet<string>(
        new[] { "SECONDLY", "MINUTELY", "HOURLY", "DAILY", "WEEKLY", "MONTHLY", "YEARLY" },
        StringComparer.OrdinalIgnoreCase);

    /// <summary>Validates stable RFC 5545 structure without rejecting unknown or extension properties.</summary>
    public IReadOnlyList<ContentLineValidationIssue> Validate() {
        var issues = new List<ContentLineValidationIssue>();
        if (_calendars.Count == 0) {
            issues.Add(new ContentLineValidationIssue("ICAL_ROOT_REQUIRED",
                "The iCalendar document must contain at least one VCALENDAR root.",
                ContentLineValidationSeverity.Error, "VCALENDAR"));
        }
        foreach (ContentLineComponent calendar in _calendars) {
            if (!string.Equals(calendar.Name, "VCALENDAR", StringComparison.OrdinalIgnoreCase)) {
                issues.Add(new ContentLineValidationIssue("ICAL_ROOT_INVALID",
                    "Every iCalendar document root must be VCALENDAR.",
                    ContentLineValidationSeverity.Error, calendar.Name));
            }
            var definedTimeZones = new HashSet<string>(StringComparer.Ordinal);
            ValidateSingle(calendar, "VERSION", required: true, issues);
            ValidateSingle(calendar, "PRODID", required: true, issues);
            ValidateSingle(calendar, "CALSCALE", required: false, issues);
            ValidateSingle(calendar, "METHOD", required: false, issues);
            ContentLineProperty? version = calendar.GetFirstProperty("VERSION");
            if (version != null && !string.Equals(version.Value, "2.0", StringComparison.Ordinal)) {
                issues.Add(Issue("ICAL_VERSION_UNSUPPORTED", "VCALENDAR VERSION must be 2.0.",
                    ContentLineValidationSeverity.Error, calendar, version));
            }
            foreach (ContentLineComponent timeZone in calendar.Components.Where(component =>
                string.Equals(component.Name, "VTIMEZONE", StringComparison.OrdinalIgnoreCase))) {
                ContentLineProperty? timeZoneId = timeZone.GetFirstProperty("TZID");
                if (timeZoneId != null && !string.IsNullOrWhiteSpace(timeZoneId.Value) &&
                    !definedTimeZones.Add(timeZoneId.Value)) {
                    issues.Add(Issue("ICAL_TIMEZONE_ID_DUPLICATE",
                        "VTIMEZONE TZID must be unique within one VCALENDAR.",
                        ContentLineValidationSeverity.Error, timeZone, timeZoneId));
                }
            }
            var active = new HashSet<ContentLineComponent> { calendar };
            foreach (ContentLineComponent component in calendar.Components)
                ValidateComponent(component, calendar.Name, issues, active, depth: 2);
            ValidateTimeZoneReferences(calendar, definedTimeZones, issues,
                new HashSet<ContentLineComponent>(), depth: 1);
        }
        return issues.AsReadOnly();
    }

    private static void ValidateComponent(ContentLineComponent component, string parentName,
        ICollection<ContentLineValidationIssue> issues, ISet<ContentLineComponent> active, int depth) {
        if (depth > ContentLineComponent.MaximumTraversalDepth) {
            issues.Add(Issue("ICAL_COMPONENT_DEPTH_EXCEEDED",
                "The mutable iCalendar component graph exceeds the supported nesting depth.",
                ContentLineValidationSeverity.Error, component));
            return;
        }
        if (!active.Add(component)) {
            issues.Add(Issue("ICAL_COMPONENT_GRAPH_CYCLE",
                "The mutable iCalendar component graph contains a cycle.",
                ContentLineValidationSeverity.Error, component));
            return;
        }
        try {
            string name = component.Name.ToUpperInvariant();
            if (name == "VEVENT" || name == "VTODO" || name == "VJOURNAL") {
                WarnWhenMissing(component, "UID", issues);
                WarnWhenMissing(component, "DTSTAMP", issues);
                ValidateSingle(component, "UID", required: false, issues);
                ValidateSingle(component, "DTSTAMP", required: false, issues);
                ValidateSingle(component, "SEQUENCE", required: false, issues);
            } else if (name == "VTIMEZONE") {
                ValidateSingle(component, "TZID", required: true, issues);
            } else if (name == "VALARM") {
                if (parentName != "VEVENT" && parentName != "VTODO") {
                    issues.Add(Issue("ICAL_ALARM_PARENT_INVALID", "VALARM must be nested in VEVENT or VTODO.",
                        ContentLineValidationSeverity.Error, component));
                }
                ValidateSingle(component, "ACTION", required: true, issues);
                ValidateSingle(component, "TRIGGER", required: true, issues);
                string? action = component.GetFirstProperty("ACTION")?.Value;
                if (string.Equals(action, "DISPLAY", StringComparison.OrdinalIgnoreCase))
                    ValidateSingle(component, "DESCRIPTION", required: true, issues);
                if (string.Equals(action, "EMAIL", StringComparison.OrdinalIgnoreCase)) {
                    ValidateSingle(component, "DESCRIPTION", required: true, issues);
                    ValidateSingle(component, "SUMMARY", required: true, issues);
                    if (!component.GetProperties("ATTENDEE").Any())
                        issues.Add(Issue("ICAL_ALARM_ATTENDEE_REQUIRED", "An EMAIL VALARM requires ATTENDEE.",
                            ContentLineValidationSeverity.Error, component, propertyName: "ATTENDEE"));
                }
            }

            foreach (ContentLineProperty ruleProperty in component.GetProperties("RRULE")) {
                try {
                    IcsRecurrenceRule rule = IcsRecurrenceRule.Parse(ruleProperty.Value);
                    if (!RecurrenceFrequencies.Contains(rule.Frequency!))
                        issues.Add(Issue("ICAL_RRULE_FREQUENCY_INVALID", "RRULE FREQ is not a registered frequency.",
                            ContentLineValidationSeverity.Error, component, ruleProperty));
                    if (rule.GetValue("COUNT") != null && rule.GetValue("UNTIL") != null)
                        issues.Add(Issue("ICAL_RRULE_COUNT_UNTIL_CONFLICT", "RRULE cannot contain both COUNT and UNTIL.",
                            ContentLineValidationSeverity.Error, component, ruleProperty));
                    if (rule.GetValue("COUNT") != null && !rule.Count.HasValue)
                        issues.Add(Issue("ICAL_RRULE_COUNT_INVALID", "RRULE COUNT must be a positive integer.",
                            ContentLineValidationSeverity.Error, component, ruleProperty));
                    if (rule.GetValue("INTERVAL") != null && !rule.Interval.HasValue)
                        issues.Add(Issue("ICAL_RRULE_INTERVAL_INVALID", "RRULE INTERVAL must be a positive integer.",
                            ContentLineValidationSeverity.Error, component, ruleProperty));
                } catch (FormatException exception) {
                    issues.Add(Issue("ICAL_RRULE_INVALID", exception.Message, ContentLineValidationSeverity.Error,
                        component, ruleProperty));
                }
            }

            foreach (string temporalName in new[] { "DTSTART", "DTEND", "DUE", "RECURRENCE-ID", "RDATE", "EXDATE" }) {
                foreach (ContentLineProperty property in component.GetProperties(temporalName)) {
                    if (temporalName == "RDATE" || temporalName == "EXDATE") continue;
                    if (!IcsTemporalValue.TryParse(property, out _))
                        issues.Add(Issue("ICAL_TEMPORAL_VALUE_INVALID", "The temporal property value is invalid.",
                            ContentLineValidationSeverity.Error, component, property));
                }
            }

            foreach (ContentLineComponent child in component.Components)
                ValidateComponent(child, name, issues, active, depth + 1);
        } finally {
            active.Remove(component);
        }
    }

    private static void ValidateTimeZoneReferences(ContentLineComponent component, ISet<string> definedTimeZones,
        ICollection<ContentLineValidationIssue> issues, ISet<ContentLineComponent> active, int depth) {
        if (depth > ContentLineComponent.MaximumTraversalDepth || !active.Add(component)) return;
        try {
            foreach (ContentLineProperty property in component.Properties) {
                ContentLineParameter? timeZone = property.GetParameter("TZID");
                foreach (string timeZoneId in timeZone?.Values ?? Array.Empty<string>()) {
                    if (!definedTimeZones.Contains(timeZoneId))
                        issues.Add(Issue("ICAL_TIMEZONE_DEFINITION_MISSING",
                            "TZID '" + timeZoneId + "' has no matching VTIMEZONE definition in this VCALENDAR.",
                            ContentLineValidationSeverity.Warning, component, property));
                }
            }
            foreach (ContentLineComponent child in component.Components)
                ValidateTimeZoneReferences(child, definedTimeZones, issues, active, depth + 1);
        } finally {
            active.Remove(component);
        }
    }

    private static void ValidateSingle(ContentLineComponent component, string propertyName, bool required,
        ICollection<ContentLineValidationIssue> issues) {
        ContentLineProperty[] properties = component.GetProperties(propertyName).ToArray();
        if (required && properties.Length == 0)
            issues.Add(Issue("ICAL_PROPERTY_REQUIRED", propertyName + " is required.",
                ContentLineValidationSeverity.Error, component, propertyName: propertyName));
        if (properties.Length > 1)
            issues.Add(Issue("ICAL_PROPERTY_CARDINALITY", propertyName + " must not occur more than once.",
                ContentLineValidationSeverity.Error, component, properties[1]));
    }

    private static void WarnWhenMissing(ContentLineComponent component, string propertyName,
        ICollection<ContentLineValidationIssue> issues) {
        if (component.GetFirstProperty(propertyName) == null)
            issues.Add(Issue("ICAL_INTEROPERABILITY_PROPERTY_MISSING",
                propertyName + " is normally required for interoperable scheduling and storage.",
                ContentLineValidationSeverity.Warning, component, propertyName: propertyName));
    }

    private static ContentLineValidationIssue Issue(string code, string message,
        ContentLineValidationSeverity severity, ContentLineComponent component,
        ContentLineProperty? property = null, string? propertyName = null) =>
        new ContentLineValidationIssue(code, message, severity, component.Name,
            property?.Name ?? propertyName);
}
