namespace OfficeIMO.Email;

public sealed partial class IcsDocument {
    private static readonly HashSet<string> RecurrenceFrequencies = new HashSet<string>(
        new[] { "SECONDLY", "MINUTELY", "HOURLY", "DAILY", "WEEKLY", "MONTHLY", "YEARLY" },
        StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> UtcTimestampPropertyNames = new HashSet<string>(
        new[] { "DTSTAMP", "CREATED", "LAST-MODIFIED", "COMPLETED" },
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
                ValidateAlarmTriggers(component, issues);
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
                    var seenParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var reportedDuplicateParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (IcsRecurrencePart part in rule.Parts) {
                        if (!seenParts.Add(part.Name) && reportedDuplicateParts.Add(part.Name)) {
                            issues.Add(Issue("ICAL_RRULE_PART_DUPLICATE",
                                "RRULE " + part.Name.ToUpperInvariant() + " must not occur more than once.",
                                ContentLineValidationSeverity.Error, component, ruleProperty));
                        }
                    }
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
                    ValidateRecurrenceUntil(component, ruleProperty, rule, issues);
                } catch (FormatException exception) {
                    issues.Add(Issue("ICAL_RRULE_INVALID", exception.Message, ContentLineValidationSeverity.Error,
                        component, ruleProperty));
                }
            }

            if (name == "VEVENT" || name == "VTODO") {
                ValidateDurations(component, issues);
            }

            foreach (string temporalName in new[] { "DTSTART", "DTEND", "DUE", "RECURRENCE-ID",
                         "RDATE", "EXDATE", "DTSTAMP", "CREATED", "LAST-MODIFIED", "COMPLETED" }) {
                foreach (ContentLineProperty property in component.GetProperties(temporalName)) {
                    IcsTemporalValue temporalValue = default;
                    bool valid;
                    if (temporalName == "RDATE" || temporalName == "EXDATE") {
                        valid = ValidateRecurrenceDateValues(property,
                            allowPeriod: temporalName == "RDATE");
                    } else {
                        valid = IcsTemporalValue.TryParse(property, out temporalValue);
                    }
                    if (!valid) {
                        issues.Add(Issue("ICAL_TEMPORAL_VALUE_INVALID", "The temporal property value is invalid.",
                            ContentLineValidationSeverity.Error, component, property));
                    } else if (UtcTimestampPropertyNames.Contains(temporalName) &&
                               temporalValue.Kind != IcsTemporalValueKind.UtcDateTime) {
                        issues.Add(Issue("ICAL_TEMPORAL_VALUE_UTC_REQUIRED",
                            temporalName + " must contain a UTC DATE-TIME value.",
                            ContentLineValidationSeverity.Error, component, property));
                    }
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
                ContentLineParameter[] timeZones = property.Parameters.Where(parameter =>
                    string.Equals(parameter.Name, "TZID", StringComparison.OrdinalIgnoreCase)).ToArray();
                if (timeZones.Length > 1 || timeZones.Any(parameter => parameter.Values.Count != 1 ||
                    string.IsNullOrWhiteSpace(parameter.Values[0]))) {
                    issues.Add(Issue("ICAL_PARAMETER_CARDINALITY",
                        "TZID must occur at most once and contain exactly one non-empty value.",
                        ContentLineValidationSeverity.Error, component, property));
                }
                foreach (ContentLineParameter timeZone in timeZones) {
                    foreach (string timeZoneId in timeZone.Values) {
                        if (string.IsNullOrWhiteSpace(timeZoneId)) continue;
                        if (!definedTimeZones.Contains(timeZoneId))
                            issues.Add(Issue("ICAL_TIMEZONE_DEFINITION_MISSING",
                                "TZID '" + timeZoneId + "' has no matching VTIMEZONE definition in this VCALENDAR.",
                                ContentLineValidationSeverity.Warning, component, property));
                    }
                }
            }
            foreach (ContentLineComponent child in component.Components)
                ValidateTimeZoneReferences(child, definedTimeZones, issues, active, depth + 1);
        } finally {
            active.Remove(component);
        }
    }

    private static bool ValidateRecurrenceDateValues(ContentLineProperty property, bool allowPeriod) {
        ContentLineParameter[] valueParameters = property.Parameters.Where(parameter =>
            string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (valueParameters.Length > 1 || valueParameters.Any(parameter => parameter.Values.Count != 1 ||
            string.IsNullOrWhiteSpace(parameter.Values[0])))
            return false;
        ContentLineParameter[] timeZoneParameters = property.Parameters.Where(parameter =>
            string.Equals(parameter.Name, "TZID", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (timeZoneParameters.Length > 1 || timeZoneParameters.Any(parameter =>
            parameter.Values.Count != 1 || string.IsNullOrWhiteSpace(parameter.Values[0]))) return false;
        string? valueType = valueParameters.FirstOrDefault()?.Values[0];
        bool isPeriod = string.Equals(valueType, "PERIOD", StringComparison.OrdinalIgnoreCase);
        if (isPeriod && !allowPeriod) return false;
        string[] values = property.Value.Split(',');
        if (values.Length == 0) return false;
        foreach (string value in values) {
            if (value.Length == 0) return false;
            if (isPeriod) {
                if (!ValidatePeriodValue(property, value)) return false;
                continue;
            }
            var candidate = new ContentLineProperty(property.Name, value);
            foreach (ContentLineParameter parameter in property.Parameters)
                candidate.Parameters.Add(parameter);
            if (!IcsTemporalValue.TryParse(candidate, out _)) return false;
        }
        return true;
    }

    private static void ValidateRecurrenceUntil(ContentLineComponent component,
        ContentLineProperty ruleProperty, IcsRecurrenceRule rule,
        ICollection<ContentLineValidationIssue> issues) {
        string? untilText = rule.GetValue("UNTIL");
        if (untilText == null) return;
        if (!TryParseRecurrenceUntil(untilText, out IcsTemporalValue until)) {
            issues.Add(Issue("ICAL_RRULE_UNTIL_INVALID",
                "RRULE UNTIL must contain a valid DATE or DATE-TIME value.",
                ContentLineValidationSeverity.Error, component, ruleProperty));
            return;
        }

        ContentLineProperty? startProperty = component.GetFirstProperty("DTSTART");
        if (startProperty == null || !IcsTemporalValue.TryParse(startProperty, out IcsTemporalValue start)) return;
        IcsTemporalValueKind expectedKind;
        if (start.Kind == IcsTemporalValueKind.Date) {
            expectedKind = IcsTemporalValueKind.Date;
        } else if (string.Equals(component.Name, "STANDARD", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(component.Name, "DAYLIGHT", StringComparison.OrdinalIgnoreCase) ||
                   start.Kind == IcsTemporalValueKind.UtcDateTime ||
                   start.Kind == IcsTemporalValueKind.ZonedDateTime) {
            expectedKind = IcsTemporalValueKind.UtcDateTime;
        } else {
            expectedKind = IcsTemporalValueKind.FloatingDateTime;
        }
        if (until.Kind != expectedKind) {
            issues.Add(Issue("ICAL_RRULE_UNTIL_TYPE_MISMATCH",
                "RRULE UNTIL must use the DATE or DATE-TIME form required by DTSTART.",
                ContentLineValidationSeverity.Error, component, ruleProperty));
        }
    }

    private static bool TryParseRecurrenceUntil(string value, out IcsTemporalValue result) {
        var property = new ContentLineProperty("UNTIL", value);
        if (value.Length == 8) property.SetParameter("VALUE", "DATE");
        return IcsTemporalValue.TryParse(property, out result);
    }

    private static void ValidateAlarmTriggers(ContentLineComponent component,
        ICollection<ContentLineValidationIssue> issues) {
        foreach (ContentLineProperty trigger in component.GetProperties("TRIGGER")) {
            ContentLineParameter[] valueParameters = trigger.Parameters.Where(parameter =>
                string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase)).ToArray();
            ContentLineParameter[] relatedParameters = trigger.Parameters.Where(parameter =>
                string.Equals(parameter.Name, "RELATED", StringComparison.OrdinalIgnoreCase)).ToArray();
            bool validParameters = valueParameters.Length <= 1 && relatedParameters.Length <= 1 &&
                valueParameters.All(parameter => parameter.Values.Count == 1 &&
                    !string.IsNullOrWhiteSpace(parameter.Values[0])) &&
                relatedParameters.All(parameter => parameter.Values.Count == 1 &&
                    (string.Equals(parameter.Values[0], "START", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(parameter.Values[0], "END", StringComparison.OrdinalIgnoreCase)));
            string valueType = valueParameters.FirstOrDefault()?.Values.FirstOrDefault() ?? "DURATION";
            bool valid;
            if (!validParameters) {
                valid = false;
            } else if (string.Equals(valueType, "DATE-TIME", StringComparison.OrdinalIgnoreCase)) {
                valid = relatedParameters.Length == 0 &&
                    IcsTemporalValue.TryParse(trigger, out IcsTemporalValue absolute) &&
                    absolute.Kind == IcsTemporalValueKind.UtcDateTime;
            } else if (string.Equals(valueType, "DURATION", StringComparison.OrdinalIgnoreCase)) {
                valid = ValidateDuration(trigger.Value, allowNegative: true, requireNonZero: false);
            } else {
                valid = false;
            }
            if (!valid) {
                issues.Add(Issue("ICAL_ALARM_TRIGGER_INVALID",
                    "VALARM TRIGGER must contain a duration or a UTC DATE-TIME value.",
                    ContentLineValidationSeverity.Error, component, trigger));
            }
        }
    }

    private static void ValidateDurations(ContentLineComponent component,
        ICollection<ContentLineValidationIssue> issues) {
        foreach (ContentLineProperty duration in component.GetProperties("DURATION")) {
            ContentLineParameter[] valueParameters = duration.Parameters.Where(parameter =>
                string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase)).ToArray();
            bool validParameters = valueParameters.Length <= 1 &&
                valueParameters.All(parameter => parameter.Values.Count == 1 &&
                    string.Equals(parameter.Values[0], "DURATION", StringComparison.OrdinalIgnoreCase));
            if (!validParameters || !ValidatePositiveDuration(duration.Value)) {
                issues.Add(Issue("ICAL_DURATION_INVALID",
                    component.Name + " DURATION must contain a positive RFC duration value.",
                    ContentLineValidationSeverity.Error, component, duration));
            }
        }
    }

    private static bool ValidatePeriodValue(ContentLineProperty property, string value) {
        int separator = value.IndexOf('/');
        if (separator <= 0 || separator != value.LastIndexOf('/') || separator == value.Length - 1)
            return false;
        string startText = value.Substring(0, separator);
        string endText = value.Substring(separator + 1);
        var startProperty = CreatePeriodDateTimeProperty(property, startText);
        if (!IcsTemporalValue.TryParse(startProperty, out IcsTemporalValue start) ||
            start.Kind == IcsTemporalValueKind.Date) return false;
        if (endText[0] == 'P' || endText.StartsWith("+P", StringComparison.Ordinal))
            return ValidatePositiveDuration(endText);
        var endProperty = CreatePeriodDateTimeProperty(property, endText);
        return IcsTemporalValue.TryParse(endProperty, out IcsTemporalValue end) &&
            end.Kind != IcsTemporalValueKind.Date && end.Kind == start.Kind &&
            string.Equals(end.TimeZoneId, start.TimeZoneId, StringComparison.Ordinal) &&
            end.Value > start.Value;
    }

    private static ContentLineProperty CreatePeriodDateTimeProperty(
        ContentLineProperty source, string value) {
        var result = new ContentLineProperty(source.Name, value);
        foreach (ContentLineParameter parameter in source.Parameters) {
            if (!string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase))
                result.Parameters.Add(parameter);
        }
        return result;
    }

    private static bool ValidatePositiveDuration(string value) {
        return ValidateDuration(value, allowNegative: false, requireNonZero: true);
    }

    private static bool ValidateDuration(string value, bool allowNegative, bool requireNonZero) {
        if (value.Length < 2) return false;
        int index = 0;
        if (value[index] == '+' || value[index] == '-') {
            if (value[index] == '-' && !allowNegative) return false;
            index++;
        }
        if (index >= value.Length || value[index++] != 'P') return false;
        bool nonZero = false;
        bool hasLeadingNumber = ReadDurationNumber(value, ref index, ref nonZero);
        if (hasLeadingNumber && index < value.Length && value[index] == 'W')
            return index + 1 == value.Length && (!requireNonZero || nonZero);
        bool hasDays = false;
        if (hasLeadingNumber) {
            if (index >= value.Length || value[index] != 'D') return false;
            index++;
            hasDays = true;
        }
        if (index == value.Length) return hasDays && (!requireNonZero || nonZero);
        if (value[index] != 'T' || ++index == value.Length) return false;
        if (!ReadDurationNumber(value, ref index, ref nonZero) || index >= value.Length) return false;
        char designator = value[index++];
        if (designator == 'H') {
            if (index < value.Length) {
                if (!ReadDurationNumber(value, ref index, ref nonZero) || index >= value.Length) return false;
                designator = value[index++];
                if (designator != 'M') return false;
                if (index < value.Length) {
                    if (!ReadDurationNumber(value, ref index, ref nonZero) || index >= value.Length ||
                        value[index++] != 'S') return false;
                }
            }
        } else if (designator == 'M') {
            if (index < value.Length && (!ReadDurationNumber(value, ref index, ref nonZero) ||
                index >= value.Length || value[index++] != 'S')) return false;
        } else if (designator != 'S') return false;
        return index == value.Length && (!requireNonZero || nonZero);
    }

    private static bool ReadDurationNumber(string value, ref int index, ref bool nonZero) {
        int start = index;
        while (index < value.Length && value[index] >= '0' && value[index] <= '9') {
            if (value[index] != '0') nonZero = true;
            index++;
        }
        return index > start;
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
