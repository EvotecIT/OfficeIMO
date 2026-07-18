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
                ValidateComponent(component, calendar, issues, active, depth: 2);
            ValidateTimeZoneReferences(calendar, definedTimeZones, issues,
                new HashSet<ContentLineComponent>(), depth: 1);
        }
        return issues.AsReadOnly();
    }

    private static void ValidateComponent(ContentLineComponent component, ContentLineComponent parent,
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
            ValidateKnownComponentParent(component, parent, name, issues);
            ValidateTemporalCardinality(component, name, issues);
            if (name == "VEVENT" || name == "VTODO" || name == "VJOURNAL") {
                WarnWhenMissing(component, "UID", issues);
                WarnWhenMissing(component, "DTSTAMP", issues);
                ValidateSingle(component, "UID", required: false, issues);
                ValidateSingle(component, "DTSTAMP", required: false, issues);
                ValidateSingle(component, "SEQUENCE", required: false, issues);
            } else if (name == "VTIMEZONE") {
                ValidateSingle(component, "TZID", required: true, issues);
                if (!component.Components.Any(child =>
                    string.Equals(child.Name, "STANDARD", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(child.Name, "DAYLIGHT", StringComparison.OrdinalIgnoreCase))) {
                    issues.Add(Issue("ICAL_TIMEZONE_OBSERVANCE_REQUIRED",
                        "VTIMEZONE requires at least one STANDARD or DAYLIGHT observance.",
                        ContentLineValidationSeverity.Error, component));
                }
            } else if (name == "VALARM") {
                ValidateSingle(component, "ACTION", required: true, issues);
                ValidateSingle(component, "TRIGGER", required: true, issues);
                ValidateAlarmTriggers(component, parent, issues);
                ValidateAlarmRepetition(component, issues);
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
                    ValidateRegisteredRecurrenceParts(component, ruleProperty, rule, issues);
                    ValidateRecurrenceUntil(component, ruleProperty, rule, issues);
                } catch (FormatException exception) {
                    issues.Add(Issue("ICAL_RRULE_INVALID", exception.Message, ContentLineValidationSeverity.Error,
                        component, ruleProperty));
                }
            }

            if (name == "VEVENT" || name == "VTODO") {
                ValidateDurations(component, issues);
                ValidateTemporalEndpoint(component, name == "VEVENT" ? "DTEND" : "DUE", issues);
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
                ValidateComponent(child, component, issues, active, depth + 1);
        } finally {
            active.Remove(component);
        }
    }

    private static void ValidateKnownComponentParent(ContentLineComponent component,
        ContentLineComponent parent, string name, ICollection<ContentLineValidationIssue> issues) {
        string parentName = parent.Name.ToUpperInvariant();
        bool valid = name == "VEVENT" || name == "VTODO" || name == "VJOURNAL" ||
                     name == "VFREEBUSY" || name == "VTIMEZONE"
            ? parentName == "VCALENDAR"
            : name == "STANDARD" || name == "DAYLIGHT"
                ? parentName == "VTIMEZONE"
                : name == "VALARM"
                    ? parentName == "VEVENT" || parentName == "VTODO"
                    : name != "VCALENDAR";
        if (!valid) {
            string code = name == "VALARM"
                ? "ICAL_ALARM_PARENT_INVALID"
                : "ICAL_COMPONENT_PARENT_INVALID";
            string message = name == "VALARM"
                ? "VALARM must be nested in VEVENT or VTODO."
                : name + " is nested under an invalid parent component.";
            issues.Add(Issue(code, message, ContentLineValidationSeverity.Error, component));
        }
    }

    private static void ValidateTemporalCardinality(ContentLineComponent component, string name,
        ICollection<ContentLineValidationIssue> issues) {
        if (name == "VEVENT") {
            ValidateSingle(component, "DTSTART", required: false, issues);
            ValidateSingle(component, "DTEND", required: false, issues);
            ValidateSingle(component, "RECURRENCE-ID", required: false, issues);
            ValidateSingle(component, "CREATED", required: false, issues);
            ValidateSingle(component, "LAST-MODIFIED", required: false, issues);
        } else if (name == "VTODO") {
            ValidateSingle(component, "DTSTART", required: false, issues);
            ValidateSingle(component, "DUE", required: false, issues);
            ValidateSingle(component, "RECURRENCE-ID", required: false, issues);
            ValidateSingle(component, "COMPLETED", required: false, issues);
            ValidateSingle(component, "CREATED", required: false, issues);
            ValidateSingle(component, "LAST-MODIFIED", required: false, issues);
        } else if (name == "VJOURNAL") {
            ValidateSingle(component, "DTSTART", required: false, issues);
            ValidateSingle(component, "RECURRENCE-ID", required: false, issues);
            ValidateSingle(component, "CREATED", required: false, issues);
            ValidateSingle(component, "LAST-MODIFIED", required: false, issues);
        } else if (name == "VFREEBUSY") {
            ValidateSingle(component, "DTSTART", required: false, issues);
            ValidateSingle(component, "DTEND", required: false, issues);
            ValidateSingle(component, "DTSTAMP", required: false, issues);
        } else if (name == "VTIMEZONE") {
            ValidateSingle(component, "LAST-MODIFIED", required: false, issues);
        } else if (name == "STANDARD" || name == "DAYLIGHT") {
            ValidateSingle(component, "DTSTART", required: true, issues);
            ValidateSingle(component, "TZOFFSETFROM", required: true, issues);
            ValidateSingle(component, "TZOFFSETTO", required: true, issues);
            ValidateUtcOffset(component, "TZOFFSETFROM", issues);
            ValidateUtcOffset(component, "TZOFFSETTO", issues);
        }
    }

    private static void ValidateUtcOffset(ContentLineComponent component, string propertyName,
        ICollection<ContentLineValidationIssue> issues) {
        foreach (ContentLineProperty property in component.GetProperties(propertyName)) {
            if (IsValidUtcOffset(property.Value)) continue;
            issues.Add(Issue("ICAL_TIMEZONE_OFFSET_INVALID",
                propertyName + " must contain an RFC 5545 UTC-OFFSET value.",
                ContentLineValidationSeverity.Error, component, property));
        }
    }

    private static bool IsValidUtcOffset(string value) {
        if (value == null || (value.Length != 5 && value.Length != 7) ||
            (value[0] != '+' && value[0] != '-')) return false;
        for (int index = 1; index < value.Length; index++) {
            if (value[index] < '0' || value[index] > '9') return false;
        }
        int hours = (value[1] - '0') * 10 + value[2] - '0';
        int minutes = (value[3] - '0') * 10 + value[4] - '0';
        int seconds = value.Length == 7 ? (value[5] - '0') * 10 + value[6] - '0' : 0;
        if (hours > 23 || minutes > 59 || seconds > 59) return false;
        return value[0] != '-' || hours != 0 || minutes != 0 || seconds != 0;
    }

    private static void ValidateTemporalEndpoint(ContentLineComponent component, string endpointName,
        ICollection<ContentLineValidationIssue> issues) {
        ContentLineProperty? startProperty = component.GetFirstProperty("DTSTART");
        ContentLineProperty? endpointProperty = component.GetFirstProperty(endpointName);
        if (startProperty == null || endpointProperty == null ||
            !IcsTemporalValue.TryParse(startProperty, out IcsTemporalValue start) ||
            !IcsTemporalValue.TryParse(endpointProperty, out IcsTemporalValue endpoint)) return;

        bool startIsDate = start.Kind == IcsTemporalValueKind.Date;
        bool endpointIsDate = endpoint.Kind == IcsTemporalValueKind.Date;
        if (startIsDate != endpointIsDate) {
            issues.Add(Issue("ICAL_TEMPORAL_ENDPOINT_TYPE_MISMATCH",
                endpointName + " must use the same DATE or DATE-TIME value type as DTSTART.",
                ContentLineValidationSeverity.Error, component, endpointProperty));
            return;
        }

        bool startIsUtc = start.Kind == IcsTemporalValueKind.UtcDateTime;
        bool endpointIsUtc = endpoint.Kind == IcsTemporalValueKind.UtcDateTime;
        if (startIsUtc != endpointIsUtc) {
            issues.Add(Issue("ICAL_TEMPORAL_ENDPOINT_REPRESENTATION_MISMATCH",
                endpointName + " must use UTC if and only if DTSTART uses UTC.",
                ContentLineValidationSeverity.Error, component, endpointProperty));
            return;
        }

        bool comparable = start.Kind == endpoint.Kind &&
            (start.Kind != IcsTemporalValueKind.ZonedDateTime ||
             string.Equals(start.TimeZoneId, endpoint.TimeZoneId, StringComparison.Ordinal));
        if (comparable && endpoint.Value <= start.Value) {
            issues.Add(Issue("ICAL_TEMPORAL_ENDPOINT_ORDER_INVALID",
                endpointName + " must be later than DTSTART.",
                ContentLineValidationSeverity.Error, component, endpointProperty));
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

    private static void ValidateAlarmTriggers(ContentLineComponent component, ContentLineComponent parent,
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
                if (valid) ValidateRelativeAlarmAnchor(component, parent, trigger, relatedParameters, issues);
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

    private static void ValidateRelativeAlarmAnchor(ContentLineComponent alarm, ContentLineComponent parent,
        ContentLineProperty trigger, IReadOnlyList<ContentLineParameter> relatedParameters,
        ICollection<ContentLineValidationIssue> issues) {
        string relationship = relatedParameters.FirstOrDefault()?.Values.FirstOrDefault() ?? "START";
        if (string.Equals(relationship, "START", StringComparison.OrdinalIgnoreCase)) {
            if (parent.GetFirstProperty("DTSTART") == null) {
                issues.Add(Issue("ICAL_ALARM_TRIGGER_START_REQUIRED",
                    "A START-relative VALARM TRIGGER requires DTSTART on its parent component.",
                    ContentLineValidationSeverity.Error, alarm, trigger));
            }
            return;
        }

        bool hasExplicitEnd = string.Equals(parent.Name, "VEVENT", StringComparison.OrdinalIgnoreCase)
            ? parent.GetFirstProperty("DTEND") != null
            : parent.GetFirstProperty("DUE") != null;
        bool hasDerivedEnd = parent.GetFirstProperty("DTSTART") != null &&
            parent.GetFirstProperty("DURATION") != null;
        if (!hasExplicitEnd && !hasDerivedEnd) {
            issues.Add(Issue("ICAL_ALARM_TRIGGER_END_REQUIRED",
                "An END-relative VALARM TRIGGER requires an explicit end or DTSTART with DURATION on its parent component.",
                ContentLineValidationSeverity.Error, alarm, trigger));
        }
    }

    private static void ValidateAlarmRepetition(ContentLineComponent component,
        ICollection<ContentLineValidationIssue> issues) {
        ContentLineProperty[] durations = component.GetProperties("DURATION").ToArray();
        ContentLineProperty[] repeats = component.GetProperties("REPEAT").ToArray();
        ValidateSingle(component, "DURATION", required: false, issues);
        ValidateSingle(component, "REPEAT", required: false, issues);

        if ((durations.Length > 0) != (repeats.Length > 0)) {
            string missingProperty = durations.Length > 0 ? "REPEAT" : "DURATION";
            issues.Add(Issue("ICAL_ALARM_REPEAT_PAIR_REQUIRED",
                "VALARM DURATION and REPEAT must occur together.",
                ContentLineValidationSeverity.Error, component, propertyName: missingProperty));
        }

        foreach (ContentLineProperty duration in durations) {
            bool hasForbiddenValueParameter = duration.Parameters.Any(parameter =>
                string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase));
            if (hasForbiddenValueParameter || !ValidatePositiveDuration(duration.Value)) {
                issues.Add(Issue("ICAL_ALARM_DURATION_INVALID",
                    "VALARM DURATION must contain a positive RFC duration and cannot declare VALUE.",
                    ContentLineValidationSeverity.Error, component, duration));
            }
        }

        foreach (ContentLineProperty repeat in repeats) {
            bool hasForbiddenValueParameter = repeat.Parameters.Any(parameter =>
                string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase));
            bool validInteger = int.TryParse(repeat.Value,
                System.Globalization.NumberStyles.AllowLeadingSign,
                System.Globalization.CultureInfo.InvariantCulture, out int count) && count >= 0;
            if (hasForbiddenValueParameter || !validInteger) {
                issues.Add(Issue("ICAL_ALARM_REPEAT_INVALID",
                    "VALARM REPEAT must contain a non-negative integer and cannot declare VALUE.",
                    ContentLineValidationSeverity.Error, component, repeat));
            }
        }
    }

    private static void ValidateDurations(ContentLineComponent component,
        ICollection<ContentLineValidationIssue> issues) {
        ContentLineProperty[] durations = component.GetProperties("DURATION").ToArray();
        if (durations.Length > 1) {
            issues.Add(Issue("ICAL_PROPERTY_CARDINALITY", "DURATION must not occur more than once.",
                ContentLineValidationSeverity.Error, component, durations[1]));
        }
        if (durations.Length > 0) {
            string componentName = component.Name.ToUpperInvariant();
            if ((componentName == "VEVENT" && component.GetFirstProperty("DTEND") != null) ||
                (componentName == "VTODO" && component.GetFirstProperty("DUE") != null)) {
                issues.Add(Issue("ICAL_DURATION_END_CONFLICT",
                    component.Name + " cannot contain both DURATION and " +
                    (componentName == "VEVENT" ? "DTEND." : "DUE."),
                    ContentLineValidationSeverity.Error, component, durations[0]));
            }
            if (componentName == "VTODO" && component.GetFirstProperty("DTSTART") == null) {
                issues.Add(Issue("ICAL_DURATION_START_REQUIRED",
                    "VTODO DURATION requires DTSTART.",
                    ContentLineValidationSeverity.Error, component, durations[0]));
            }
        }
        foreach (ContentLineProperty duration in durations) {
            bool hasForbiddenValueParameter = duration.Parameters.Any(parameter =>
                string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase));
            if (hasForbiddenValueParameter || !ValidatePositiveDuration(duration.Value)) {
                issues.Add(Issue("ICAL_DURATION_INVALID",
                    component.Name + " DURATION must contain a positive RFC duration and cannot declare VALUE.",
                    ContentLineValidationSeverity.Error, component, duration));
            }
            ContentLineProperty? startProperty = component.GetFirstProperty("DTSTART");
            if (startProperty != null &&
                IcsTemporalValue.TryParse(startProperty, out IcsTemporalValue start) &&
                start.Kind == IcsTemporalValueKind.Date && duration.Value.IndexOf('T') >= 0) {
                issues.Add(Issue("ICAL_DURATION_DATE_START_INVALID",
                    component.Name + " with a DATE DTSTART requires a day- or week-based DURATION.",
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
