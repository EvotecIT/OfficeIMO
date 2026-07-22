namespace OfficeIMO.Email;

public sealed partial class IcsDocument {
    private static void ValidateCalendarScaleValues(ContentLineComponent calendar,
        ICollection<ContentLineValidationIssue> issues) {
        foreach (ContentLineProperty calendarScale in calendar.GetProperties("CALSCALE")) {
            if (IsValidTokenValue(calendarScale)) continue;
            issues.Add(Issue("ICAL_CALSCALE_INVALID",
                "CALSCALE must contain a non-empty iana-token value.",
                ContentLineValidationSeverity.Error, calendar, calendarScale));
        }
    }

    private static void ValidateMethodValues(ContentLineComponent calendar,
        ICollection<ContentLineValidationIssue> issues) {
        foreach (ContentLineProperty method in calendar.GetProperties("METHOD")) {
            if (IsValidTokenValue(method)) continue;
            issues.Add(Issue("ICAL_METHOD_INVALID",
                "METHOD must contain a non-empty iana-token or x-name value.",
                ContentLineValidationSeverity.Error, calendar, method));
        }
    }

    private static void ValidateAlarmActions(ContentLineComponent alarm,
        ICollection<ContentLineValidationIssue> issues) {
        foreach (ContentLineProperty action in alarm.GetProperties("ACTION")) {
            if (IsValidTokenValue(action)) continue;
            issues.Add(Issue("ICAL_ALARM_ACTION_INVALID",
                "ACTION must contain a non-empty iana-token or x-name value.",
                ContentLineValidationSeverity.Error, alarm, action));
        }
    }

    private static bool IsValidTokenValue(ContentLineProperty property) {
        return ContentLineSyntax.IsToken(property.Value);
    }

    private static void ValidateSequenceValues(ContentLineComponent component,
        ICollection<ContentLineValidationIssue> issues) {
        foreach (ContentLineProperty sequence in component.GetProperties("SEQUENCE")) {
            bool hasForbiddenValueParameter = sequence.Parameters.Any(parameter =>
                string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase));
            bool validInteger = int.TryParse(sequence.Value,
                System.Globalization.NumberStyles.AllowLeadingSign,
                System.Globalization.CultureInfo.InvariantCulture, out int value) && value >= 0;
            if (!hasForbiddenValueParameter && validInteger) continue;
            issues.Add(Issue("ICAL_SEQUENCE_INVALID",
                "SEQUENCE must contain a non-negative integer and cannot declare VALUE.",
                ContentLineValidationSeverity.Error, component, sequence));
        }
    }

    private static void ValidateExceptionDateRepresentation(ContentLineComponent component,
        ContentLineProperty property, ICollection<ContentLineValidationIssue> issues) {
        ContentLineProperty? startProperty = component.GetFirstProperty("DTSTART");
        if (!IcsTemporalValue.TryParse(startProperty, out IcsTemporalValue start)) return;

        foreach (string value in property.Value.Split(',')) {
            var candidate = new ContentLineProperty(property.Name, value);
            foreach (ContentLineParameter parameter in property.Parameters) candidate.Parameters.Add(parameter);
            if (!IcsTemporalValue.TryParse(candidate, out IcsTemporalValue exception)) continue;

            bool startIsDate = start.Kind == IcsTemporalValueKind.Date;
            bool exceptionIsDate = exception.Kind == IcsTemporalValueKind.Date;
            if (startIsDate != exceptionIsDate) {
                issues.Add(Issue("ICAL_EXDATE_TYPE_MISMATCH",
                    "EXDATE must use the same DATE or DATE-TIME value type as DTSTART.",
                    ContentLineValidationSeverity.Error, component, property));
                return;
            }
            if (startIsDate) continue;

            bool startIsFloating = start.Kind == IcsTemporalValueKind.FloatingDateTime;
            bool exceptionIsFloating = exception.Kind == IcsTemporalValueKind.FloatingDateTime;
            if (startIsFloating != exceptionIsFloating) {
                issues.Add(Issue("ICAL_EXDATE_REPRESENTATION_MISMATCH",
                    "EXDATE must use floating local time if and only if DTSTART uses floating local time.",
                    ContentLineValidationSeverity.Error, component, property));
                return;
            }
        }
    }

    private static void ValidateFreeBusyWindow(ContentLineComponent component,
        ICollection<ContentLineValidationIssue> issues) {
        ContentLineProperty? startProperty = component.GetFirstProperty("DTSTART");
        ContentLineProperty? endProperty = component.GetFirstProperty("DTEND");
        bool validStart = IcsTemporalValue.TryParse(startProperty, out IcsTemporalValue start);
        bool validEnd = IcsTemporalValue.TryParse(endProperty, out IcsTemporalValue end);
        if (validStart && start.Kind != IcsTemporalValueKind.UtcDateTime) {
            issues.Add(Issue("ICAL_TEMPORAL_VALUE_UTC_REQUIRED",
                "VFREEBUSY DTSTART must contain a UTC DATE-TIME value.",
                ContentLineValidationSeverity.Error, component, startProperty));
        }
        if (validEnd && end.Kind != IcsTemporalValueKind.UtcDateTime) {
            issues.Add(Issue("ICAL_TEMPORAL_VALUE_UTC_REQUIRED",
                "VFREEBUSY DTEND must contain a UTC DATE-TIME value.",
                ContentLineValidationSeverity.Error, component, endProperty));
        }
        if (validStart && validEnd && start.Kind == IcsTemporalValueKind.UtcDateTime &&
            end.Kind == IcsTemporalValueKind.UtcDateTime && end.CompareClockTo(start) <= 0) {
            issues.Add(Issue("ICAL_TEMPORAL_ENDPOINT_ORDER_INVALID",
                "VFREEBUSY DTEND must be later than DTSTART.",
                ContentLineValidationSeverity.Error, component, endProperty));
        }

        foreach (ContentLineProperty freeBusy in component.GetProperties("FREEBUSY")) {
            if (ValidateFreeBusyValues(freeBusy)) continue;
            issues.Add(Issue("ICAL_FREEBUSY_PERIOD_INVALID",
                "FREEBUSY must contain one or more valid UTC PERIOD values.",
                ContentLineValidationSeverity.Error, component, freeBusy));
        }
    }

    private static bool ValidateFreeBusyValues(ContentLineProperty property) {
        ContentLineParameter[] valueParameters = property.Parameters.Where(parameter =>
            string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (valueParameters.Length != 0) return false;
        string[] values = property.Value.Split(',');
        return values.Length > 0 && values.All(value => value.Length > 0 &&
            ValidateFreeBusyPeriod(property, value));
    }

    private static bool ValidateFreeBusyPeriod(ContentLineProperty property, string value) {
        int separator = value.IndexOf('/');
        if (separator <= 0 || separator != value.LastIndexOf('/') || separator == value.Length - 1)
            return false;
        var startProperty = CreatePeriodDateTimeProperty(property, value.Substring(0, separator));
        if (!IcsTemporalValue.TryParse(startProperty, out IcsTemporalValue start) ||
            start.Kind != IcsTemporalValueKind.UtcDateTime) return false;
        string endText = value.Substring(separator + 1);
        if (endText[0] == 'P' || endText.StartsWith("+P", StringComparison.Ordinal))
            return ValidatePositiveDuration(endText);
        var endProperty = CreatePeriodDateTimeProperty(property, endText);
        return IcsTemporalValue.TryParse(endProperty, out IcsTemporalValue end) &&
            end.Kind == IcsTemporalValueKind.UtcDateTime && end.CompareClockTo(start) > 0;
    }

    private static void ValidateRecurrenceIdentifier(ContentLineComponent component,
        IReadOnlyDictionary<(string Name, string Uid), ContentLineComponent> recurrenceMasters,
        ICollection<ContentLineValidationIssue> issues) {
        ContentLineProperty? recurrenceProperty = component.GetFirstProperty("RECURRENCE-ID");
        if (recurrenceProperty == null) return;
        ValidateRecurrenceIdentifierRange(component, recurrenceProperty, issues);
        ContentLineProperty? uidProperty = component.GetFirstProperty("UID");
        if (uidProperty == null || string.IsNullOrWhiteSpace(uidProperty.Value)) return;
        recurrenceMasters.TryGetValue(
            (component.Name.ToUpperInvariant(), uidProperty.Value),
            out ContentLineComponent? master);
        ContentLineProperty? masterStartProperty = master?.GetFirstProperty("DTSTART");
        if (!IcsTemporalValue.TryParse(masterStartProperty, out IcsTemporalValue start) ||
            !IcsTemporalValue.TryParse(recurrenceProperty, out IcsTemporalValue recurrence)) return;
        bool startIsDate = start.Kind == IcsTemporalValueKind.Date;
        bool recurrenceIsDate = recurrence.Kind == IcsTemporalValueKind.Date;
        if (startIsDate != recurrenceIsDate) {
            issues.Add(Issue("ICAL_RECURRENCE_ID_TYPE_MISMATCH",
                "RECURRENCE-ID must use the same DATE or DATE-TIME value type as DTSTART.",
                ContentLineValidationSeverity.Error, component, recurrenceProperty));
            return;
        }
        if (start.Kind != recurrence.Kind) {
            issues.Add(Issue("ICAL_RECURRENCE_ID_REPRESENTATION_MISMATCH",
                "RECURRENCE-ID must use the same floating, UTC, or zoned representation as DTSTART.",
                ContentLineValidationSeverity.Error, component, recurrenceProperty));
            return;
        }
        if (start.Kind == IcsTemporalValueKind.ZonedDateTime &&
            !string.Equals(start.TimeZoneId, recurrence.TimeZoneId, StringComparison.Ordinal)) {
            issues.Add(Issue("ICAL_RECURRENCE_ID_TIMEZONE_MISMATCH",
                "A zoned RECURRENCE-ID must use the same TZID as DTSTART.",
                ContentLineValidationSeverity.Error, component, recurrenceProperty));
        }
    }

    private static void ValidateRecurrenceIdentifierRange(ContentLineComponent component,
        ContentLineProperty recurrenceProperty, ICollection<ContentLineValidationIssue> issues) {
        ContentLineParameter[] rangeParameters = recurrenceProperty.Parameters.Where(parameter =>
            string.Equals(parameter.Name, "RANGE", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (rangeParameters.Length == 0) return;
        if (rangeParameters.Length == 1 && rangeParameters[0].Values.Count == 1 &&
            string.Equals(rangeParameters[0].Values[0], "THISANDFUTURE", StringComparison.OrdinalIgnoreCase))
            return;
        issues.Add(Issue("ICAL_RECURRENCE_ID_RANGE_INVALID",
            "RECURRENCE-ID RANGE must occur at most once with the value THISANDFUTURE.",
            ContentLineValidationSeverity.Error, component, recurrenceProperty));
    }

    private static void ValidateTimeZoneObservanceTimes(ContentLineComponent component,
        ICollection<ContentLineValidationIssue> issues) {
        ContentLineProperty? startProperty = component.GetFirstProperty("DTSTART");
        if (IcsTemporalValue.TryParse(startProperty, out IcsTemporalValue start) &&
            start.Kind != IcsTemporalValueKind.FloatingDateTime) {
            issues.Add(Issue("ICAL_TIMEZONE_OBSERVANCE_START_INVALID",
                "STANDARD and DAYLIGHT DTSTART must contain a floating local DATE-TIME value.",
                ContentLineValidationSeverity.Error, component, startProperty));
        }
        foreach (ContentLineProperty recurrenceDate in component.GetProperties("RDATE")) {
            if (ValidateObservanceRecurrenceDates(recurrenceDate)) continue;
            issues.Add(Issue("ICAL_TIMEZONE_OBSERVANCE_RDATE_INVALID",
                "STANDARD and DAYLIGHT RDATE must contain floating local DATE-TIME values.",
                ContentLineValidationSeverity.Error, component, recurrenceDate));
        }
    }

    private static bool ValidateObservanceRecurrenceDates(ContentLineProperty property) {
        ContentLineParameter[] valueParameters = property.Parameters.Where(parameter =>
            string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (valueParameters.Length > 1 || valueParameters.Any(parameter => parameter.Values.Count != 1 ||
            !string.Equals(parameter.Values[0], "DATE-TIME", StringComparison.OrdinalIgnoreCase))) return false;
        if (property.Parameters.Any(parameter =>
            string.Equals(parameter.Name, "TZID", StringComparison.OrdinalIgnoreCase))) return false;
        string[] values = property.Value.Split(',');
        if (values.Length == 0) return false;
        foreach (string value in values) {
            if (value.Length == 0) return false;
            var candidate = new ContentLineProperty(property.Name, value);
            foreach (ContentLineParameter parameter in property.Parameters) candidate.Parameters.Add(parameter);
            if (!IcsTemporalValue.TryParse(candidate, out IcsTemporalValue temporal) ||
                temporal.Kind != IcsTemporalValueKind.FloatingDateTime) return false;
        }
        return true;
    }
}
