namespace OfficeIMO.Email;

internal static partial class IcsCalendarCodec {
    internal static bool TryProject(string text, EmailDocument document, IList<EmailDiagnostic> diagnostics,
        string location, string? mimeMethod = null) {
        int projectionDiagnosticStart = diagnostics.Count;
        List<IcsProperty> properties;
        try {
            properties = ParseProperties(text.TrimStart('\uFEFF'));
        } catch (InvalidDataException exception) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_PARSE_INVALID", exception.Message,
                EmailDiagnosticSeverity.Warning, location));
            document.MimeSemanticProjectionIsIncomplete = true;
            return false;
        }
        IcsProperty? activeComponent = properties.FirstOrDefault(property => property.Name == "BEGIN" &&
            (property.Value.Equals("VEVENT", StringComparison.OrdinalIgnoreCase) ||
             property.Value.Equals("VTODO", StringComparison.OrdinalIgnoreCase)));
        if (activeComponent == null) {
            document.MimeSemanticProjectionIsIncomplete = true;
            return false;
        }
        bool isEvent = activeComponent.Value.Equals("VEVENT", StringComparison.OrdinalIgnoreCase);

        IReadOnlyList<IcsProperty> activeProperties = SelectActiveComponentProperties(
            properties, activeComponent.Value);
        IReadOnlyList<IcsProperty> alarmProperties = SelectActiveAlarmProperties(
            properties, activeComponent.Value);
        document.MimeSemanticProjectionIsIncomplete |= HasIncompleteStoreProjection(
            properties, activeProperties, alarmProperties, isEvent, document.Subject, document.From);
        string? calendarMethod = GetValue(activeProperties, "METHOD");
        string? effectiveMethod = calendarMethod ?? mimeMethod;
        bool hasMethodConflict = !string.IsNullOrWhiteSpace(calendarMethod) &&
            !string.IsNullOrWhiteSpace(mimeMethod) &&
            !string.Equals(calendarMethod, mimeMethod, StringComparison.OrdinalIgnoreCase);
        bool hasTransportRecipients = HasCalendarRecipients(document);
        bool hasCalendarAttendees = activeProperties.Any(property => property.Name == "ATTENDEE");
        bool hasCalendarRecipients = hasTransportRecipients || hasCalendarAttendees;
        bool requestWouldSynthesizeAttendees =
            string.Equals(effectiveMethod, "REQUEST", StringComparison.OrdinalIgnoreCase) &&
            hasTransportRecipients && !hasCalendarAttendees;
        bool methodWouldChange =
            string.IsNullOrWhiteSpace(effectiveMethod) && hasCalendarRecipients ||
            string.Equals(effectiveMethod, "PUBLISH", StringComparison.OrdinalIgnoreCase) &&
            hasCalendarRecipients || !isEvent &&
            string.Equals(effectiveMethod, "REQUEST", StringComparison.OrdinalIgnoreCase) &&
            !hasCalendarRecipients;
        if (hasMethodConflict || requestWouldSynthesizeAttendees ||
            isEvent && !IsStoreProjectableMethod(effectiveMethod) ||
            !isEvent && !IsStoreProjectableTaskMethod(effectiveMethod) || methodWouldChange) {
            document.MimeSemanticProjectionIsIncomplete = true;
        }
        if (isEvent) ProjectEvent(activeProperties, document, diagnostics, location, effectiveMethod);
        else ProjectTask(activeProperties, document, diagnostics, location);
        TryProjectTypedRecurrence(text, document, isEvent, diagnostics, location);
        document.MimeSemanticProjectionIsIncomplete |= HasIncompleteTimestampProjection(
            activeProperties, document, isEvent, diagnostics, location);
        document.MimeSemanticProjectionIsIncomplete |= diagnostics.Skip(projectionDiagnosticStart).Any(diagnostic =>
            diagnostic.Code == "EMAIL_ICALENDAR_TIMEZONE_UNRESOLVED" ||
            diagnostic.Code == "EMAIL_ICALENDAR_FLOATING_TIME" ||
            diagnostic.Code == "EMAIL_ICALENDAR_DATE_INVALID");
        return true;
    }

    private static void ProjectEvent(IReadOnlyList<IcsProperty> properties, EmailDocument document,
        IList<EmailDiagnostic> diagnostics, string location, string? method) {
        var appointment = document.Appointment ?? new OutlookAppointment();
        document.OutlookItemKind = OutlookItemKind.Appointment;
        document.MessageClass = MessageClassForMethod(method, properties);
        document.Appointment = appointment;
        ApplyCommon(properties, document);

        appointment.Start = ParseDate(GetProperty(properties, "DTSTART"), diagnostics, location, out bool allDay);
        appointment.End = ParseDate(GetProperty(properties, "DTEND"), diagnostics, location, out _);
        appointment.IsAllDay = allDay;
        appointment.Location = Unescape(GetValue(properties, "LOCATION"));
        appointment.Sequence = ParseInt(GetValue(properties, "SEQUENCE"));
        appointment.IsRecurring = GetProperty(properties, "RRULE") != null;
        appointment.RecurrencePattern = GetValue(properties, "RRULE");
        TimeSpan? duration = IcsDurationCodec.Parse(GetValue(properties, "DURATION"));
        bool incompleteDuration = false;
        TimeSpan? effectiveDuration = appointment.Start.HasValue && appointment.End.HasValue
            ? appointment.End.Value - appointment.Start.Value
            : duration;
        appointment.DurationMinutes = effectiveDuration.HasValue
            ? IcsDurationCodec.ToWholeMinutes(effectiveDuration.Value, diagnostics, location, ref incompleteDuration)
            : null;
        if (!appointment.End.HasValue && appointment.Start.HasValue && duration.HasValue) {
            try {
                appointment.End = appointment.Start.Value.Add(duration.Value);
            } catch (ArgumentOutOfRangeException) {
                IcsDurationCodec.ReportOutOfRange(diagnostics, location);
                incompleteDuration = true;
            }
        }
        string? busy = GetValue(properties, "X-MICROSOFT-CDO-BUSYSTATUS");
        appointment.BusyStatus = ParseBusyStatus(busy) ?? ParseTransparency(GetValue(properties, "TRANSP"));
        appointment.NotAllowPropose = ParseBoolean(GetValue(properties, "X-MICROSOFT-DISALLOW-COUNTER"));
        appointment.MeetingStatus = ParseInt(GetValue(properties, "X-OFFICEIMO-MEETING-STATUS"));
        appointment.ResponseStatus = ParseInt(GetValue(properties, "X-OFFICEIMO-RESPONSE-STATUS"));
        appointment.ClientIntentFlags = ParseInt(GetValue(properties, "X-OFFICEIMO-CLIENT-INTENT"));
        appointment.TimeZoneDescription = UnescapeOrNull(GetValue(properties, "X-OFFICEIMO-TIMEZONE-DESCRIPTION"));
        incompleteDuration |= ApplyReminder(properties, appointment, diagnostics, location);
        document.MimeSemanticProjectionIsIncomplete |= incompleteDuration;

        string[] required = GetAttendees(properties, optional: false).ToArray();
        string[] optional = GetAttendees(properties, optional: true).ToArray();
        appointment.RequiredAttendees = required.Length == 0 ? null : string.Join("; ", required);
        appointment.OptionalAttendees = optional.Length == 0 ? null : string.Join("; ", optional);
        appointment.AllAttendees = required.Concat(optional).Any()
            ? string.Join("; ", required.Concat(optional))
            : null;
        AddAttendeeRecipients(properties, document);
    }

    private static void ProjectTask(IReadOnlyList<IcsProperty> properties, EmailDocument document,
        IList<EmailDiagnostic> diagnostics, string location) {
        var task = document.Task ?? new OutlookTask();
        document.OutlookItemKind = OutlookItemKind.Task;
        document.MessageClass = "IPM.Task";
        document.Task = task;
        ApplyCommon(properties, document);
        task.Start = ParseDate(GetProperty(properties, "DTSTART"), diagnostics, location, out _);
        task.Due = ParseDate(GetProperty(properties, "DUE"), diagnostics, location, out _);
        task.CompletedAt = ParseDate(GetProperty(properties, "COMPLETED"), diagnostics, location, out _);
        if (double.TryParse(GetValue(properties, "PERCENT-COMPLETE"), NumberStyles.Float,
            CultureInfo.InvariantCulture, out double percent)) task.PercentComplete = percent / 100d;
        string? status = GetValue(properties, "STATUS");
        task.IsComplete = string.Equals(status, "COMPLETED", StringComparison.OrdinalIgnoreCase);
        task.Status = ParseTaskStatus(status);
        task.Owner = UnescapeOrNull(GetValue(properties, "X-OFFICEIMO-TASK-OWNER")) ?? GetOrganizer(properties);
        task.IsRecurring = GetProperty(properties, "RRULE") != null;
        TimeSpan? standardDuration = IcsDurationCodec.Parse(GetValue(properties, "DURATION"));
        task.EstimatedEffort = IcsDurationCodec.Parse(GetValue(properties, "X-OFFICEIMO-ESTIMATED-EFFORT"));
        task.ActualEffort = IcsDurationCodec.Parse(GetValue(properties, "X-OFFICEIMO-ACTUAL-EFFORT"));
        bool incompleteEffort = false;
        if (task.EstimatedEffort.HasValue) IcsDurationCodec.ToWholeMinutes(task.EstimatedEffort.Value,
            diagnostics, location, ref incompleteEffort);
        if (task.ActualEffort.HasValue) IcsDurationCodec.ToWholeMinutes(task.ActualEffort.Value,
            diagnostics, location, ref incompleteEffort);
        if (!task.Due.HasValue && standardDuration.HasValue) {
            if (task.Start.HasValue) {
                try {
                    task.Due = task.Start.Value.Add(standardDuration.Value);
                } catch (ArgumentOutOfRangeException) {
                    IcsDurationCodec.ReportOutOfRange(diagnostics, location);
                    incompleteEffort = true;
                }
            } else {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_TASK_DURATION_START_REQUIRED",
                    "A VTODO duration cannot be projected to a due date without DTSTART.",
                    EmailDiagnosticSeverity.Warning, location));
                incompleteEffort = true;
            }
        }
        task.SendUpdates = ParseBoolean(GetValue(properties, "X-OFFICEIMO-SEND-UPDATES"));
        task.SendStatusOnComplete = ParseBoolean(GetValue(properties, "X-OFFICEIMO-SEND-STATUS-ON-COMPLETE"));
        task.Ownership = ParseInt(GetValue(properties, "X-OFFICEIMO-OWNERSHIP"));
        task.AcceptanceState = ParseInt(GetValue(properties, "X-OFFICEIMO-ACCEPTANCE-STATE"));
        task.Version = ParseInt(GetValue(properties, "X-OFFICEIMO-TASK-VERSION"));
        task.State = ParseInt(GetValue(properties, "X-OFFICEIMO-TASK-STATE"));
        task.Assigner = UnescapeOrNull(GetValue(properties, "X-OFFICEIMO-ASSIGNER"));
        task.IsTeamTask = ParseBoolean(GetValue(properties, "X-OFFICEIMO-TEAM-TASK"));
        task.Ordinal = ParseInt(GetValue(properties, "X-OFFICEIMO-ORDINAL"));
        task.CommonStart = ParseDate(GetProperty(properties, "X-OFFICEIMO-COMMON-START"), diagnostics, location, out _);
        task.CommonEnd = ParseDate(GetProperty(properties, "X-OFFICEIMO-COMMON-END"), diagnostics, location, out _);
        task.Mode = ParseInt(GetValue(properties, "X-OFFICEIMO-TASK-MODE"));
        task.ToDoOrdinalDate = ParseDate(GetProperty(properties, "X-OFFICEIMO-TODO-ORDINAL-DATE"),
            diagnostics, location, out _);
        task.ToDoSubOrdinal = UnescapeOrNull(GetValue(properties, "X-OFFICEIMO-TODO-SUBORDINAL"));
        task.BillingInformation = UnescapeOrNull(GetValue(properties, "X-OFFICEIMO-BILLING-INFORMATION"));
        task.Mileage = UnescapeOrNull(GetValue(properties, "X-OFFICEIMO-MILEAGE"));
        foreach (IcsProperty property in properties.Where(property => property.Name == "CONTACT")) {
            task.Contacts.Add(Unescape(property.Value));
        }
        foreach (IcsProperty property in properties.Where(property => property.Name == "X-OFFICEIMO-COMPANY")) {
            task.Companies.Add(Unescape(property.Value));
        }
        AddAttendeeRecipients(properties, document);
        document.MimeSemanticProjectionIsIncomplete |= incompleteEffort |
            ApplyReminder(properties, task, diagnostics, location);
    }

    private static void ApplyCommon(IReadOnlyList<IcsProperty> properties, EmailDocument document) {
        string? summary = Unescape(GetValue(properties, "SUMMARY"));
        string? description = Unescape(GetValue(properties, "DESCRIPTION"));
        string? uid = Unescape(GetValue(properties, "UID"));
        document.MimeSemanticSourceHasTextBody |= !string.IsNullOrWhiteSpace(description);
        if (!string.IsNullOrWhiteSpace(summary)) document.Subject = summary;
        if (!string.IsNullOrWhiteSpace(description)) {
            if (document.Body.Text == null) document.Body.Text = description;
            else if (!string.Equals(document.Body.Text, description, StringComparison.Ordinal)) {
                document.MimeSemanticProjectionIsIncomplete = true;
            }
        }
        if (!string.IsNullOrWhiteSpace(uid)) {
            if (document.MessageId == null) document.MessageId = uid;
            else if (!string.Equals(document.MessageId.Trim().Trim('<', '>'), uid,
                         StringComparison.Ordinal)) document.MimeSemanticProjectionIsIncomplete = true;
        }
        int? calendarSensitivity = ParseCalendarSensitivity(GetValue(properties, "CLASS"));
        if (calendarSensitivity.HasValue) document.MessageMetadata.Sensitivity = calendarSensitivity;
        foreach (IcsProperty property in properties.Where(property => property.Name == "CATEGORIES")) {
            foreach (string category in SplitEscapedValues(property.Value, ',')) {
                if (!string.IsNullOrWhiteSpace(category) && !document.MessageMetadata.Categories.Any(existing =>
                    string.Equals(existing, category, StringComparison.OrdinalIgnoreCase))) {
                    document.MessageMetadata.Categories.Add(category);
                }
            }
        }
        string? organizer = GetOrganizer(properties);
        if (!string.IsNullOrWhiteSpace(organizer)) {
            IcsProperty? organizerProperty = GetProperty(properties, "ORGANIZER");
            string? organizerName = null;
            organizerProperty?.Parameters.TryGetValue("CN", out organizerName);
            if (document.From == null) document.From = new EmailAddress(organizer!, organizerName);
            else if (!string.Equals(document.From.Address, organizer, StringComparison.OrdinalIgnoreCase)) {
                document.MimeSemanticProjectionIsIncomplete = true;
            } else if (string.IsNullOrWhiteSpace(document.From.DisplayName)) {
                document.From.DisplayName = organizerName;
            } else if (!string.IsNullOrWhiteSpace(organizerName) &&
                       !string.Equals(document.From.DisplayName, organizerName, StringComparison.Ordinal)) {
                document.MimeSemanticProjectionIsIncomplete = true;
            }
        }
    }
}
