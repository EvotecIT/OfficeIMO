namespace OfficeIMO.Email;

internal static partial class IcsCalendarCodec {
    private static void TryProjectTypedRecurrence(string text, EmailDocument document, bool isEvent,
        IList<EmailDiagnostic> diagnostics, string location) {
        try {
            IcsDocument calendar = IcsDocument.Parse(text.TrimStart('\uFEFF'));
            string componentName = isEvent ? "VEVENT" : "VTODO";
            ContentLineComponent[] components = calendar.GetComponents(componentName).ToArray();
            ContentLineComponent? master = components.FirstOrDefault();
            if (master == null || master.GetFirstProperty("RECURRENCE-ID") != null ||
                master.GetFirstProperty("RRULE") == null) return;
            ContentLineProperty[] rules = master.GetProperties("RRULE").ToArray();
            if (rules.Length != 1) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_RRULE_CARDINALITY_INVALID",
                    "A component must contain exactly one RRULE for Outlook recurrence projection.",
                    EmailDiagnosticSeverity.Warning, location));
                document.MimeSemanticProjectionIsIncomplete = true;
                return;
            }
            IcsTemporalValue? start = master.GetTemporalValue("DTSTART");
            if (!start.HasValue) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_RECURRENCE_START_REQUIRED",
                    "A recurring component requires a supported DTSTART.", EmailDiagnosticSeverity.Warning,
                    location));
                document.MimeSemanticProjectionIsIncomplete = true;
                return;
            }
            TimeSpan duration = GetRecurrenceDuration(master, start.Value, isEvent);
            OutlookTimeZoneDefinition? timeZone = ResolveEmbeddedRecurrenceTimeZone(
                calendar, start.Value, diagnostics, location, document);
            if (start.Value.Kind == IcsTemporalValueKind.ZonedDateTime && timeZone == null &&
                RequiresUtcRecurrenceConversion(master, components)) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_RECURRENCE_TIMEZONE_REQUIRED",
                    "UTC recurrence limits or exceptions cannot be projected without the matching embedded VTIMEZONE rules.",
                    EmailDiagnosticSeverity.Warning, location));
                document.MimeSemanticProjectionIsIncomplete = true;
                return;
            }
            var options = new OutlookRecurrenceIcsImportOptions {
                Start = start.Value,
                Duration = duration,
                TimeZone = timeZone
            };
            AddExcludedDates(master, options, diagnostics, location, document);
            if (isEvent) AddExceptionComponents(components, master, options, diagnostics, location, document);
            if (master.GetFirstProperty("RDATE") != null) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_RDATE_UNSUPPORTED",
                    "RDATE values are retained in the source calendar but are not part of Outlook recurrence projection.",
                    EmailDiagnosticSeverity.Warning, location));
                document.MimeSemanticProjectionIsIncomplete = true;
            }

            OutlookRecurrenceIcsImportResult result = OutlookRecurrenceIcsConverter.Import(
                IcsRecurrenceRule.Parse(rules[0].Value), options);
            AddConversionDiagnostics(result.Report, diagnostics, location, document);
            if (result.Recurrence == null) return;
            if (isEvent && document.Appointment != null) {
                document.Appointment.Recurrence = result.Recurrence;
                if (timeZone != null) document.Appointment.RecurrenceTimeZone = timeZone;
                document.Appointment.IsRecurring = true;
                document.Appointment.RecurrencePattern = rules[0].Value;
            } else if (!isEvent && document.Task != null) {
                document.Task.Recurrence = result.Recurrence;
                document.Task.IsRecurring = true;
            }
        } catch (Exception exception) when (exception is InvalidDataException || exception is FormatException ||
            exception is ArgumentException || exception is InvalidOperationException || exception is OverflowException) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_RECURRENCE_PROJECTION_INVALID", exception.Message,
                EmailDiagnosticSeverity.Warning, location));
            document.MimeSemanticProjectionIsIncomplete = true;
        }
    }

    private static bool RequiresUtcRecurrenceConversion(ContentLineComponent master,
        IEnumerable<ContentLineComponent> components) {
        string? until = master.GetFirstProperty("RRULE") == null
            ? null
            : IcsRecurrenceRule.Parse(master.GetFirstProperty("RRULE")!.Value).GetValue("UNTIL");
        if (until != null && IsUtcTemporal(new ContentLineProperty("UNTIL", until))) return true;
        foreach (ContentLineProperty property in master.GetProperties("EXDATE")) {
            foreach (string raw in property.Value.Split(',')) {
                if (IsUtcTemporal(CloneTemporalProperty(property, raw))) return true;
            }
        }
        string? masterUid = master.GetFirstProperty("UID")?.Value;
        if (string.IsNullOrWhiteSpace(masterUid)) return false;
        foreach (ContentLineComponent component in components.Where(component =>
                     component.GetFirstProperty("RECURRENCE-ID") != null &&
                     string.Equals(component.GetFirstProperty("UID")?.Value, masterUid,
                         StringComparison.Ordinal))) {
            foreach (string name in new[] { "RECURRENCE-ID", "DTSTART", "DTEND" }) {
                ContentLineProperty? property = component.GetFirstProperty(name);
                if (property != null && IsUtcTemporal(property)) return true;
            }
        }
        return false;
    }

    private static bool IsUtcTemporal(ContentLineProperty property) =>
        IcsTemporalValue.TryParse(property, out IcsTemporalValue value) &&
        value.Kind == IcsTemporalValueKind.UtcDateTime;

    private static TimeSpan GetRecurrenceDuration(ContentLineComponent component, IcsTemporalValue start,
        bool isEvent) {
        string endName = isEvent ? "DTEND" : "DUE";
        IcsTemporalValue? end = component.GetTemporalValue(endName);
        if (end.HasValue && end.Value.Kind == start.Kind &&
            string.Equals(end.Value.TimeZoneId, start.TimeZoneId, StringComparison.OrdinalIgnoreCase))
            return end.Value.Value - start.Value;
        TimeSpan? duration = IcsDurationCodec.Parse(component.GetFirstProperty("DURATION")?.Value);
        if (duration.HasValue) return duration.Value;
        return start.Kind == IcsTemporalValueKind.Date ? TimeSpan.FromDays(1) : TimeSpan.Zero;
    }

    private static void AddExcludedDates(ContentLineComponent master, OutlookRecurrenceIcsImportOptions options,
        IList<EmailDiagnostic> diagnostics, string location, EmailDocument document) {
        foreach (ContentLineProperty property in master.GetProperties("EXDATE")) {
            foreach (string raw in property.Value.Split(',')) {
                ContentLineProperty valueProperty = CloneTemporalProperty(property, raw);
                if (IcsTemporalValue.TryParse(valueProperty, out IcsTemporalValue parsed)) {
                    options.ExcludedDates.Add(parsed);
                } else {
                    diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_EXDATE_INVALID",
                        "An EXDATE value could not be projected to Outlook recurrence.",
                        EmailDiagnosticSeverity.Warning, location));
                    document.MimeSemanticProjectionIsIncomplete = true;
                }
            }
        }
    }

    private static void AddExceptionComponents(IEnumerable<ContentLineComponent> components,
        ContentLineComponent master, OutlookRecurrenceIcsImportOptions options,
        IList<EmailDiagnostic> diagnostics, string location, EmailDocument document) {
        ContentLineComponent[] exceptions = components.Where(component =>
            component.GetFirstProperty("RECURRENCE-ID") != null).ToArray();
        if (exceptions.Length == 0) return;
        string? masterUid = master.GetFirstProperty("UID")?.Value;
        if (string.IsNullOrWhiteSpace(masterUid)) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_EXCEPTION_UID_REQUIRED",
                "A recurring component without UID cannot safely associate recurrence exceptions; they were retained but not projected.",
                EmailDiagnosticSeverity.Warning, location));
            document.MimeSemanticProjectionIsIncomplete = true;
            return;
        }
        foreach (ContentLineComponent component in exceptions) {
            string? uid = component.GetFirstProperty("UID")?.Value;
            if (!string.Equals(uid, masterUid, StringComparison.Ordinal)) continue;
            ContentLineProperty recurrenceId = component.GetFirstProperty("RECURRENCE-ID")!;
            if (recurrenceId.Parameters.Any(parameter =>
                    string.Equals(parameter.Name, "RANGE", StringComparison.OrdinalIgnoreCase) &&
                    parameter.Values.Any(value => string.Equals(value, "THISANDFUTURE",
                        StringComparison.OrdinalIgnoreCase)))) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_RECURRENCE_RANGE_UNSUPPORTED",
                    "RECURRENCE-ID;RANGE=THISANDFUTURE cannot be represented by one Outlook recurrence exception.",
                    EmailDiagnosticSeverity.Warning, location));
                document.MimeSemanticProjectionIsIncomplete = true;
                continue;
            }
            IcsTemporalValue? original = component.GetTemporalValue("RECURRENCE-ID");
            bool isCancelled = string.Equals(component.GetFirstProperty("STATUS")?.Value, "CANCELLED",
                StringComparison.OrdinalIgnoreCase);
            if (isCancelled && original.HasValue) {
                options.ExcludedDates.Add(original.Value);
                continue;
            }
            IcsTemporalValue? start = component.GetTemporalValue("DTSTART");
            IcsTemporalValue? end = component.GetTemporalValue("DTEND");
            if (!original.HasValue || !start.HasValue) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_EXCEPTION_TIME_INVALID",
                    "A recurrence exception requires supported RECURRENCE-ID and DTSTART values.",
                    EmailDiagnosticSeverity.Warning, location));
                document.MimeSemanticProjectionIsIncomplete = true;
                continue;
            }
            TimeSpan duration = end.HasValue ? end.Value.Value - start.Value.Value : options.Duration;
            IcsTemporalValue effectiveEnd = end ?? CreateMatchingTemporal(start.Value,
                start.Value.Value.Add(duration));
            options.Exceptions.Add(new OutlookRecurrenceIcsException {
                OriginalStart = original.Value,
                Start = start.Value,
                End = effectiveEnd,
                Subject = Unescape(component.GetFirstProperty("SUMMARY")?.Value),
                Location = Unescape(component.GetFirstProperty("LOCATION")?.Value),
                BusyStatus = ParseBusyStatus(component.GetFirstProperty("X-MICROSOFT-CDO-BUSYSTATUS")?.Value) ??
                    ParseTransparency(component.GetFirstProperty("TRANSP")?.Value),
                IsAllDay = start.Value.Kind == IcsTemporalValueKind.Date
            });
        }
    }

    private static ContentLineProperty CloneTemporalProperty(ContentLineProperty source, string value) {
        var result = new ContentLineProperty(source.Name, value);
        foreach (ContentLineParameter parameter in source.Parameters)
            result.Parameters.Add(new ContentLineParameter(parameter.Name, parameter.Values.ToArray()));
        return result;
    }

    private static IcsTemporalValue CreateMatchingTemporal(IcsTemporalValue template, DateTime value) {
        if (template.Kind == IcsTemporalValueKind.Date) return IcsTemporalValue.Date(value);
        if (template.Kind == IcsTemporalValueKind.UtcDateTime)
            return IcsTemporalValue.Utc(new DateTimeOffset(DateTime.SpecifyKind(value, DateTimeKind.Utc)));
        if (template.Kind == IcsTemporalValueKind.ZonedDateTime)
            return IcsTemporalValue.Zoned(value, template.TimeZoneId!);
        return IcsTemporalValue.Floating(value);
    }

    private static void AddConversionDiagnostics(OutlookRecurrenceIcsConversionReport report,
        IList<EmailDiagnostic> diagnostics, string location, EmailDocument document) {
        foreach (OutlookRecurrenceIcsIssue issue in report.Issues) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_" + issue.Code, issue.Message,
                issue.Severity == OutlookRecurrenceIcsIssueSeverity.Information
                    ? EmailDiagnosticSeverity.Information
                    : EmailDiagnosticSeverity.Warning, location));
        }
        if (!report.IsLossless) document.MimeSemanticProjectionIsIncomplete = true;
    }
}
