namespace OfficeIMO.Email;

internal static partial class IcsCalendarCodec {
    private static readonly DateTimeOffset DeterministicEpoch =
        new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero);

    internal static bool TryProject(string text, EmailDocument document, IList<EmailDiagnostic> diagnostics,
        string location, string? mimeMethod = null) {
        int projectionDiagnosticStart = diagnostics.Count;
        List<IcsProperty> properties;
        try {
            properties = ParseProperties(text.TrimStart('\uFEFF'));
        } catch (InvalidDataException exception) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_PARSE_INVALID", exception.Message,
                EmailDiagnosticSeverity.Warning, location));
            return false;
        }
        IcsProperty? activeComponent = properties.FirstOrDefault(property => property.Name == "BEGIN" &&
            (property.Value.Equals("VEVENT", StringComparison.OrdinalIgnoreCase) ||
             property.Value.Equals("VTODO", StringComparison.OrdinalIgnoreCase)));
        if (activeComponent == null) return false;
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
        document.MimeSemanticProjectionIsIncomplete |= HasIncompleteTimestampProjection(
            activeProperties, document, isEvent, diagnostics, location);
        document.MimeSemanticProjectionIsIncomplete |= diagnostics.Skip(projectionDiagnosticStart).Any(diagnostic =>
            diagnostic.Code == "EMAIL_ICALENDAR_TIMEZONE_UNRESOLVED" ||
            diagnostic.Code == "EMAIL_ICALENDAR_FLOATING_TIME" ||
            diagnostic.Code == "EMAIL_ICALENDAR_DATE_INVALID");
        return true;
    }

    internal static EmailAttachment? FindSemanticAttachment(EmailDocument document) {
        if (document.OutlookItemKind != OutlookItemKind.Appointment &&
            document.OutlookItemKind != OutlookItemKind.Task) return null;
        return document.Attachments.FirstOrDefault(attachment => attachment.IsProjectedSemanticContent &&
            string.Equals(attachment.ContentType, "text/calendar", StringComparison.OrdinalIgnoreCase));
    }

    internal static bool ShouldWriteAsAttachment(EmailAttachment attachment) => !attachment.IsMimeBodyPart;

    internal static EmailAttachment CreateRegeneratedAttachment(EmailDocument document, EmailAttachment source) {
        byte[] content = Create(document);
        var attachment = new EmailAttachment {
            FileName = source.FileName,
            ContentType = "text/calendar",
            ContentId = source.ContentId,
            ContentLocation = source.ContentLocation,
            IsInline = source.IsInline,
            IsHidden = source.IsHidden,
            RenderingPosition = source.RenderingPosition,
            CreatedDate = source.CreatedDate,
            ModifiedDate = source.ModifiedDate,
            Content = content,
            Length = content.LongLength,
            IsProjectedSemanticContent = true,
            IsMimeAttachment = source.IsMimeAttachment,
            IsMimeBodyPart = false
        };
        foreach (KeyValuePair<string, string> parameter in source.ContentTypeParameters) {
            attachment.ContentTypeParameters[parameter.Key] = parameter.Value;
        }
        attachment.ContentTypeParameters["charset"] = "utf-8";
        attachment.ContentTypeParameters["method"] = GetMethod(document);
        return attachment;
    }

    internal static byte[] Create(EmailDocument document) {
        var output = new StringBuilder();
        AppendLine(output, "BEGIN:VCALENDAR");
        AppendLine(output, "PRODID:-//Evotec//OfficeIMO.Email//EN");
        AppendLine(output, "VERSION:2.0");
        AppendLine(output, string.Concat("METHOD:", GetMethod(document)));
        if (document.OutlookItemKind == OutlookItemKind.Task) WriteTask(output, document);
        else WriteEvent(output, document);
        AppendLine(output, "END:VCALENDAR");
        return Encoding.UTF8.GetBytes(output.ToString());
    }

    internal static bool HasOpaqueAppointmentState(OutlookAppointment appointment) {
        return appointment.RecurrenceState != null || appointment.TimeZoneStructure != null ||
            appointment.StartTimeZoneDefinition != null || appointment.EndTimeZoneDefinition != null ||
            appointment.RecurrenceTimeZoneDefinition != null || appointment.IsRecurring == true ||
            appointment.RecurrenceType.HasValue || !string.IsNullOrWhiteSpace(appointment.RecurrencePattern);
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

    private static void WriteEvent(StringBuilder output, EmailDocument document) {
        OutlookAppointment appointment = document.Appointment ?? new OutlookAppointment();
        AppendLine(output, "BEGIN:VEVENT");
        WriteCommon(output, document, appointment.Start);
        if (appointment.Start.HasValue) {
            if (appointment.IsAllDay == true) {
                AppendLine(output, string.Concat("DTSTART;VALUE=DATE:", FormatDate(appointment.Start.Value)));
                DateTimeOffset end = appointment.End ?? appointment.Start.Value.AddDays(1);
                AppendLine(output, string.Concat("DTEND;VALUE=DATE:", FormatDate(end)));
            } else {
                AppendLine(output, string.Concat("DTSTART:", FormatUtc(appointment.Start.Value)));
                if (appointment.End.HasValue) AppendLine(output, string.Concat("DTEND:", FormatUtc(appointment.End.Value)));
                else if (appointment.DurationMinutes.HasValue) AppendLine(output,
                    string.Concat("DURATION:", IcsDurationCodec.Format(
                        TimeSpan.FromMinutes(appointment.DurationMinutes.Value))));
            }
        }
        AppendText(output, "LOCATION", appointment.Location);
        if (appointment.Sequence.HasValue) AppendLine(output, string.Concat("SEQUENCE:",
            appointment.Sequence.Value.ToString(CultureInfo.InvariantCulture)));
        if (appointment.BusyStatus.HasValue) {
            string busy = FormatBusyStatus(appointment.BusyStatus.Value);
            AppendLine(output, string.Concat("X-MICROSOFT-CDO-BUSYSTATUS:", busy));
            AppendLine(output, appointment.BusyStatus.Value == 0 ? "TRANSP:TRANSPARENT" : "TRANSP:OPAQUE");
        }
        if (appointment.NotAllowPropose.HasValue) AppendLine(output,
            string.Concat("X-MICROSOFT-DISALLOW-COUNTER:", appointment.NotAllowPropose.Value ? "TRUE" : "FALSE"));
        AppendInteger(output, "X-OFFICEIMO-MEETING-STATUS", appointment.MeetingStatus);
        AppendInteger(output, "X-OFFICEIMO-RESPONSE-STATUS", appointment.ResponseStatus);
        AppendInteger(output, "X-OFFICEIMO-CLIENT-INTENT", appointment.ClientIntentFlags);
        AppendText(output, "X-OFFICEIMO-TIMEZONE-DESCRIPTION", appointment.TimeZoneDescription);
        WriteReminderMetadata(output, appointment.ReminderTime, appointment.ReminderSignalTime);
        WriteOrganizerAndAttendees(output, document);
        WriteAlarm(output, appointment.ReminderIsSet, appointment.ReminderDeltaMinutes,
            appointment.ReminderSignalTime ?? appointment.ReminderTime, document.Subject);
        AppendLine(output, "END:VEVENT");
    }

    private static void WriteTask(StringBuilder output, EmailDocument document) {
        OutlookTask task = document.Task ?? new OutlookTask();
        AppendLine(output, "BEGIN:VTODO");
        WriteCommon(output, document, task.Start ?? task.Due);
        if (task.Start.HasValue) AppendLine(output, string.Concat("DTSTART:", FormatUtc(task.Start.Value)));
        if (task.Due.HasValue) AppendLine(output, string.Concat("DUE:", FormatUtc(task.Due.Value)));
        if (task.CompletedAt.HasValue) AppendLine(output, string.Concat("COMPLETED:", FormatUtc(task.CompletedAt.Value)));
        if (task.PercentComplete.HasValue) AppendLine(output, string.Concat("PERCENT-COMPLETE:",
            Math.Max(0, Math.Min(100, (int)Math.Round(task.PercentComplete.Value * 100d))).ToString(CultureInfo.InvariantCulture)));
        if (task.IsComplete == true || task.Status == 2) AppendLine(output, "STATUS:COMPLETED");
        else if (task.Status == 4) AppendLine(output, "STATUS:CANCELLED");
        else if (task.Status == 1) AppendLine(output, "STATUS:IN-PROCESS");
        else AppendLine(output, "STATUS:NEEDS-ACTION");
        if (task.EstimatedEffort.HasValue) AppendLine(output,
            string.Concat("X-OFFICEIMO-ESTIMATED-EFFORT:", IcsDurationCodec.Format(task.EstimatedEffort.Value)));
        if (!string.IsNullOrWhiteSpace(task.Owner)) {
            if (task.Owner!.IndexOf('@') >= 0) {
                EmailAddress organizer = document.From != null && string.Equals(document.From.Address, task.Owner,
                    StringComparison.OrdinalIgnoreCase) ? document.From : new EmailAddress(task.Owner);
                WriteOrganizer(output, organizer);
            }
            else AppendText(output, "X-OFFICEIMO-TASK-OWNER", task.Owner);
        }
        if (task.ActualEffort.HasValue) AppendLine(output,
            string.Concat("X-OFFICEIMO-ACTUAL-EFFORT:", IcsDurationCodec.Format(task.ActualEffort.Value)));
        AppendBoolean(output, "X-OFFICEIMO-SEND-UPDATES", task.SendUpdates);
        AppendBoolean(output, "X-OFFICEIMO-SEND-STATUS-ON-COMPLETE", task.SendStatusOnComplete);
        AppendInteger(output, "X-OFFICEIMO-OWNERSHIP", task.Ownership);
        AppendInteger(output, "X-OFFICEIMO-ACCEPTANCE-STATE", task.AcceptanceState);
        AppendInteger(output, "X-OFFICEIMO-TASK-VERSION", task.Version);
        AppendInteger(output, "X-OFFICEIMO-TASK-STATE", task.State);
        AppendText(output, "X-OFFICEIMO-ASSIGNER", task.Assigner);
        AppendBoolean(output, "X-OFFICEIMO-TEAM-TASK", task.IsTeamTask);
        AppendInteger(output, "X-OFFICEIMO-ORDINAL", task.Ordinal);
        AppendDateTime(output, "X-OFFICEIMO-COMMON-START", task.CommonStart);
        AppendDateTime(output, "X-OFFICEIMO-COMMON-END", task.CommonEnd);
        AppendInteger(output, "X-OFFICEIMO-TASK-MODE", task.Mode);
        AppendDateTime(output, "X-OFFICEIMO-TODO-ORDINAL-DATE", task.ToDoOrdinalDate);
        AppendText(output, "X-OFFICEIMO-TODO-SUBORDINAL", task.ToDoSubOrdinal);
        AppendText(output, "X-OFFICEIMO-BILLING-INFORMATION", task.BillingInformation);
        AppendText(output, "X-OFFICEIMO-MILEAGE", task.Mileage);
        foreach (string contact in task.Contacts) AppendText(output, "CONTACT", contact);
        foreach (string company in task.Companies) AppendText(output, "X-OFFICEIMO-COMPANY", company);
        WriteReminderMetadata(output, task.ReminderTime, task.ReminderSignalTime);
        WriteAttendees(output, document);
        WriteAlarm(output, task.ReminderIsSet, task.ReminderDeltaMinutes,
            task.ReminderSignalTime ?? task.ReminderTime, document.Subject);
        AppendLine(output, "END:VTODO");
    }

    private static void WriteCommon(StringBuilder output, EmailDocument document, DateTimeOffset? fallbackDate) {
        string uid = string.IsNullOrWhiteSpace(document.MessageId)
            ? CreateDeterministicUid(document, fallbackDate)
            : document.MessageId!.Trim().Trim('<', '>');
        AppendText(output, "UID", uid);
        AppendLine(output, string.Concat("DTSTAMP:", FormatUtc(document.Date ?? fallbackDate ?? DeterministicEpoch)));
        AppendText(output, "SUMMARY", document.Subject);
        AppendText(output, "DESCRIPTION", document.Body.Text);
        WriteCalendarSensitivity(output, document.MessageMetadata.Sensitivity);
        if (document.MessageMetadata.Categories.Count > 0) AppendLine(output, string.Concat("CATEGORIES:",
            string.Join(",", document.MessageMetadata.Categories.Where(category =>
                !string.IsNullOrWhiteSpace(category)).Select(EscapeText))));
    }

    private static void WriteCalendarSensitivity(StringBuilder output, int? sensitivity) {
        if (sensitivity == 0) AppendLine(output, "CLASS:PUBLIC");
        else if (sensitivity == 3) AppendLine(output, "CLASS:CONFIDENTIAL");
        else if (sensitivity == 1 || sensitivity == 2) AppendLine(output, "CLASS:PRIVATE");
    }

    private static void WriteOrganizerAndAttendees(StringBuilder output, EmailDocument document) {
        WriteOrganizer(output, document.From);
        WriteAttendees(output, document);
    }

    private static void WriteOrganizer(StringBuilder output, EmailAddress? address) {
        if (!HasPortableMailtoAddress(address)) return;
        string organizer = string.Concat("ORGANIZER");
        if (!string.IsNullOrWhiteSpace(address!.DisplayName)) organizer += string.Concat(";CN=\"",
            EscapeParameter(address.DisplayName!), "\"");
        AppendLine(output, string.Concat(organizer, ":mailto:", EscapeUriValue(address.Address!)));
    }

    private static void WriteAttendees(StringBuilder output, EmailDocument document) {
        foreach (EmailRecipient recipient in document.Recipients.Where(recipient =>
            (recipient.Kind == EmailRecipientKind.To || recipient.Kind == EmailRecipientKind.Cc ||
             recipient.Kind == EmailRecipientKind.Room || recipient.Kind == EmailRecipientKind.Resource) &&
            HasPortableMailtoAddress(recipient.Address))) {
            string role = recipient.Kind == EmailRecipientKind.Cc ? "OPT-PARTICIPANT" :
                recipient.Kind == EmailRecipientKind.Room || recipient.Kind == EmailRecipientKind.Resource
                    ? "NON-PARTICIPANT"
                    : "REQ-PARTICIPANT";
            string attendee = string.Concat("ATTENDEE;ROLE=", role);
            if (recipient.Kind == EmailRecipientKind.Room) attendee += ";CUTYPE=ROOM";
            else if (recipient.Kind == EmailRecipientKind.Resource) attendee += ";CUTYPE=RESOURCE";
            if (!string.IsNullOrWhiteSpace(recipient.Address.DisplayName)) attendee += string.Concat(";CN=\"",
                EscapeParameter(recipient.Address.DisplayName!), "\"");
            AppendLine(output, string.Concat(attendee, ":mailto:", EscapeUriValue(recipient.Address.Address!)));
        }
    }

    private static IcsProperty? GetProperty(IEnumerable<IcsProperty> properties, string name) =>
        properties.FirstOrDefault(property => property.Name == name);

    private static string? GetValue(IEnumerable<IcsProperty> properties, string name) => GetProperty(properties, name)?.Value;

    private static string? GetOrganizer(IEnumerable<IcsProperty> properties) {
        string? value = GetValue(properties, "ORGANIZER");
        return StripMailTo(value);
    }

    private static IEnumerable<string> GetAttendees(IEnumerable<IcsProperty> properties, bool optional) {
        foreach (IcsProperty property in properties.Where(property => property.Name == "ATTENDEE")) {
            bool isOptional = property.Parameters.TryGetValue("ROLE", out string? role) &&
                string.Equals(role, "OPT-PARTICIPANT", StringComparison.OrdinalIgnoreCase);
            if (isOptional != optional) continue;
            yield return property.Parameters.TryGetValue("CN", out string? name) && !string.IsNullOrWhiteSpace(name)
                ? name
                : StripMailTo(property.Value) ?? property.Value;
        }
    }

    private static void AddAttendeeRecipients(IEnumerable<IcsProperty> properties, EmailDocument document) {
        foreach (IcsProperty property in properties.Where(property => property.Name == "ATTENDEE")) {
            string? address = StripMailTo(property.Value);
            if (string.IsNullOrWhiteSpace(address)) continue;
            bool optional = property.Parameters.TryGetValue("ROLE", out string? role) &&
                string.Equals(role, "OPT-PARTICIPANT", StringComparison.OrdinalIgnoreCase);
            EmailRecipientKind kind = optional ? EmailRecipientKind.Cc : EmailRecipientKind.To;
            if (property.Parameters.TryGetValue("CUTYPE", out string? calendarUserType)) {
                if (calendarUserType.Equals("ROOM", StringComparison.OrdinalIgnoreCase)) kind = EmailRecipientKind.Room;
                else if (calendarUserType.Equals("RESOURCE", StringComparison.OrdinalIgnoreCase)) {
                    kind = EmailRecipientKind.Resource;
                }
            }
            property.Parameters.TryGetValue("CN", out string? name);
            EmailRecipient? existing = document.Recipients.FirstOrDefault(recipient =>
                string.Equals(recipient.Address.Address, address, StringComparison.OrdinalIgnoreCase));
            if (existing == null) {
                document.Recipients.Add(new EmailRecipient(kind, new EmailAddress(address!, name)));
            } else {
                existing.Kind = kind;
                if (string.IsNullOrWhiteSpace(existing.Address.DisplayName)) {
                    existing.Address.DisplayName = name;
                } else if (!string.IsNullOrWhiteSpace(name) &&
                           !string.Equals(existing.Address.DisplayName, name, StringComparison.Ordinal)) {
                    document.MimeSemanticProjectionIsIncomplete = true;
                }
            }
        }
    }

    private static string? StripMailTo(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string result = value!.Trim();
        if (IsUnprojectableCalendarAddress(result)) return null;
        if (!result.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) return result;
        string address = result.Substring(7);
        try {
            return Uri.UnescapeDataString(address);
        } catch (UriFormatException) {
            return address;
        }
    }

    private static int? ParseInt(string? value) => int.TryParse(value, NumberStyles.Integer,
        CultureInfo.InvariantCulture, out int result) ? result : (int?)null;

    private static bool? ParseBoolean(string? value) => string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase)
        ? true : string.Equals(value, "FALSE", StringComparison.OrdinalIgnoreCase) ? false : (bool?)null;

    private static int? ParseBusyStatus(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (value!.Equals("FREE", StringComparison.OrdinalIgnoreCase)) return 0;
        if (value.Equals("TENTATIVE", StringComparison.OrdinalIgnoreCase)) return 1;
        if (value.Equals("BUSY", StringComparison.OrdinalIgnoreCase)) return 2;
        if (value.Equals("OOF", StringComparison.OrdinalIgnoreCase)) return 3;
        if (value.Equals("WORKINGELSEWHERE", StringComparison.OrdinalIgnoreCase)) return 4;
        return null;
    }

    private static int? ParseTransparency(string? value) =>
        string.Equals(value, "TRANSPARENT", StringComparison.OrdinalIgnoreCase) ? 0 :
        string.Equals(value, "OPAQUE", StringComparison.OrdinalIgnoreCase) ? 2 : (int?)null;

    private static int? ParseCalendarSensitivity(string? value) =>
        string.Equals(value, "PUBLIC", StringComparison.OrdinalIgnoreCase) ? 0 :
        string.Equals(value, "PRIVATE", StringComparison.OrdinalIgnoreCase) ? 2 :
        string.Equals(value, "CONFIDENTIAL", StringComparison.OrdinalIgnoreCase) ? 3 : (int?)null;

    private static int? ParseTaskStatus(string? value) {
        if (string.Equals(value, "NOT-STARTED", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "NEEDS-ACTION", StringComparison.OrdinalIgnoreCase)) return 0;
        if (string.Equals(value, "IN-PROCESS", StringComparison.OrdinalIgnoreCase)) return 1;
        if (string.Equals(value, "COMPLETED", StringComparison.OrdinalIgnoreCase)) return 2;
        if (string.Equals(value, "CANCELLED", StringComparison.OrdinalIgnoreCase)) return 4;
        return null;
    }

    private static string FormatBusyStatus(int value) => value == 0 ? "FREE" : value == 1 ? "TENTATIVE" :
        value == 3 ? "OOF" : value == 4 ? "WORKINGELSEWHERE" : "BUSY";

    private static string MessageClassForMethod(string? method, IEnumerable<IcsProperty> properties) {
        if (string.Equals(method, "REQUEST", StringComparison.OrdinalIgnoreCase)) {
            return "IPM.Schedule.Meeting.Request";
        }
        if (string.Equals(method, "CANCEL", StringComparison.OrdinalIgnoreCase)) {
            return "IPM.Schedule.Meeting.Canceled";
        }
        if (!string.Equals(method, "REPLY", StringComparison.OrdinalIgnoreCase)) return "IPM.Appointment";
        string? participationStatus = properties.Where(property => property.Name == "ATTENDEE")
            .Select(property => property.Parameters.TryGetValue("PARTSTAT", out string? value) ? value : null)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (string.Equals(participationStatus, "TENTATIVE", StringComparison.OrdinalIgnoreCase)) {
            return "IPM.Schedule.Meeting.Resp.Tent";
        }
        if (string.Equals(participationStatus, "DECLINED", StringComparison.OrdinalIgnoreCase)) {
            return "IPM.Schedule.Meeting.Resp.Neg";
        }
        return "IPM.Schedule.Meeting.Resp.Pos";
    }

    private static bool IsStoreProjectableMethod(string? method) => string.IsNullOrWhiteSpace(method) ||
        string.Equals(method, "PUBLISH", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(method, "REQUEST", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(method, "REPLY", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(method, "CANCEL", StringComparison.OrdinalIgnoreCase);

    private static bool HasCalendarRecipients(EmailDocument document) => document.Recipients.Any(recipient =>
        recipient.Kind == EmailRecipientKind.To || recipient.Kind == EmailRecipientKind.Cc ||
        recipient.Kind == EmailRecipientKind.Room || recipient.Kind == EmailRecipientKind.Resource);

    /// <summary>Determines whether an address can be represented safely as an iCalendar mailto URI.</summary>
    internal static bool HasPortableMailtoAddress(EmailAddress? address) =>
        !string.IsNullOrWhiteSpace(address?.Address) &&
        (string.IsNullOrWhiteSpace(address!.AddressType) ||
         string.Equals(address.AddressType, "SMTP", StringComparison.OrdinalIgnoreCase));

    internal static string GetMethod(EmailDocument document) {
        string messageClass = document.MessageClass ?? string.Empty;
        if (messageClass.IndexOf("Canceled", StringComparison.OrdinalIgnoreCase) >= 0) return "CANCEL";
        if (messageClass.IndexOf("Request", StringComparison.OrdinalIgnoreCase) >= 0) return "REQUEST";
        if (messageClass.IndexOf("Resp", StringComparison.OrdinalIgnoreCase) >= 0) return "REPLY";
        return HasCalendarRecipients(document) ? "REQUEST" : "PUBLISH";
    }

    private static string FormatUtc(DateTimeOffset value) =>
        value.UtcDateTime.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);

    private static string FormatDate(DateTimeOffset value) =>
        value.Date.ToString("yyyyMMdd", CultureInfo.InvariantCulture);

    private static bool ApplyReminder(IReadOnlyList<IcsProperty> properties, OutlookAppointment appointment,
        IList<EmailDiagnostic> diagnostics, string location) {
        IcsProperty? trigger = GetProperty(properties, "TRIGGER");
        appointment.ReminderIsSet = trigger == null ? ParseBoolean(GetValue(properties, "X-OFFICEIMO-REMINDER-SET")) : true;
        appointment.ReminderDeltaMinutes = ParseRelativeTriggerMinutes(trigger, diagnostics, location,
            out bool incomplete);
        appointment.ReminderTime = ParseDate(GetProperty(properties, "X-OFFICEIMO-REMINDER-TIME"),
            diagnostics, location, out _);
        appointment.ReminderSignalTime = ParseDate(GetProperty(properties, "X-OFFICEIMO-REMINDER-SIGNAL-TIME"),
            diagnostics, location, out _);
        if (!appointment.ReminderSignalTime.HasValue) {
            appointment.ReminderSignalTime = ParseAbsoluteTrigger(trigger, diagnostics, location);
        }
        return incomplete;
    }

    private static bool ApplyReminder(IReadOnlyList<IcsProperty> properties, OutlookTask task,
        IList<EmailDiagnostic> diagnostics, string location) {
        IcsProperty? trigger = GetProperty(properties, "TRIGGER");
        task.ReminderIsSet = trigger == null ? ParseBoolean(GetValue(properties, "X-OFFICEIMO-REMINDER-SET")) : true;
        task.ReminderDeltaMinutes = ParseRelativeTriggerMinutes(trigger, diagnostics, location, out bool incomplete);
        task.ReminderTime = ParseDate(GetProperty(properties, "X-OFFICEIMO-REMINDER-TIME"),
            diagnostics, location, out _);
        task.ReminderSignalTime = ParseDate(GetProperty(properties, "X-OFFICEIMO-REMINDER-SIGNAL-TIME"),
            diagnostics, location, out _);
        if (!task.ReminderSignalTime.HasValue) {
            task.ReminderSignalTime = ParseAbsoluteTrigger(trigger, diagnostics, location);
        }
        return incomplete;
    }

    private static int? ParseRelativeTriggerMinutes(IcsProperty? trigger, IList<EmailDiagnostic> diagnostics,
        string location, out bool incomplete) {
        incomplete = false;
        if (trigger == null || trigger.Parameters.TryGetValue("VALUE", out string? valueType) &&
            valueType.Equals("DATE-TIME", StringComparison.OrdinalIgnoreCase)) return null;
        TimeSpan? value = IcsDurationCodec.Parse(trigger.Value);
        return value.HasValue
            ? IcsDurationCodec.ToWholeMinutes(value.Value, diagnostics, location, ref incomplete, invert: true)
            : null;
    }

    private static DateTimeOffset? ParseAbsoluteTrigger(IcsProperty? trigger,
        IList<EmailDiagnostic> diagnostics, string location) {
        if (trigger == null || !trigger.Parameters.TryGetValue("VALUE", out string? valueType) ||
            !valueType.Equals("DATE-TIME", StringComparison.OrdinalIgnoreCase)) return null;
        return ParseDate(trigger, diagnostics, location, out _);
    }

    private static void WriteReminderMetadata(StringBuilder output, DateTimeOffset? reminderTime,
        DateTimeOffset? reminderSignalTime) {
        AppendDateTime(output, "X-OFFICEIMO-REMINDER-TIME", reminderTime);
        AppendDateTime(output, "X-OFFICEIMO-REMINDER-SIGNAL-TIME", reminderSignalTime);
    }

    private static void WriteAlarm(StringBuilder output, bool? isSet, int? deltaMinutes,
        DateTimeOffset? absoluteTime, string? subject) {
        if (isSet != true) {
            AppendBoolean(output, "X-OFFICEIMO-REMINDER-SET", isSet);
            return;
        }
        AppendLine(output, "BEGIN:VALARM");
        AppendLine(output, "ACTION:DISPLAY");
        AppendText(output, "DESCRIPTION", string.IsNullOrWhiteSpace(subject) ? "Reminder" : subject);
        if (deltaMinutes.HasValue) AppendLine(output, string.Concat("TRIGGER:",
            IcsDurationCodec.Format(TimeSpan.FromMinutes(-deltaMinutes.Value))));
        else if (absoluteTime.HasValue) AppendLine(output,
            string.Concat("TRIGGER;VALUE=DATE-TIME:", FormatUtc(absoluteTime.Value)));
        else AppendLine(output, "TRIGGER:PT0S");
        AppendLine(output, "END:VALARM");
    }

    private static void AppendInteger(StringBuilder output, string name, int? value) {
        if (value.HasValue) AppendLine(output,
            string.Concat(name, ":", value.Value.ToString(CultureInfo.InvariantCulture)));
    }

    private static void AppendBoolean(StringBuilder output, string name, bool? value) {
        if (value.HasValue) AppendLine(output, string.Concat(name, ":", value.Value ? "TRUE" : "FALSE"));
    }

    private static void AppendDateTime(StringBuilder output, string name, DateTimeOffset? value) {
        if (value.HasValue) AppendLine(output, string.Concat(name, ":", FormatUtc(value.Value)));
    }

    private static string CreateDeterministicUid(EmailDocument document, DateTimeOffset? date) {
        string input = string.Concat(document.Subject, "|", (date ?? DeterministicEpoch).UtcTicks.ToString(CultureInfo.InvariantCulture));
        ulong hash = 14695981039346656037UL;
        foreach (byte value in Encoding.UTF8.GetBytes(input)) {
            hash ^= value;
            hash *= 1099511628211UL;
        }
        return string.Concat(hash.ToString("x16", CultureInfo.InvariantCulture), "@officeimo.local");
    }

    private static void AppendText(StringBuilder output, string name, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) AppendLine(output, string.Concat(name, ":", EscapeText(value!)));
    }

    private static void AppendLine(StringBuilder output, string line) {
        const int maximumOctets = 75;
        var current = new StringBuilder();
        int octets = 0;
        for (int index = 0; index < line.Length;) {
            int length = char.IsHighSurrogate(line[index]) && index + 1 < line.Length &&
                char.IsLowSurrogate(line[index + 1]) ? 2 : 1;
            string character = line.Substring(index, length);
            int bytes = Encoding.UTF8.GetByteCount(character);
            if (current.Length > 0 && octets + bytes > maximumOctets) {
                output.Append(current).Append("\r\n ");
                current.Clear();
                octets = 1;
            }
            current.Append(character);
            octets += bytes;
            index += length;
        }
        output.Append(current).Append("\r\n");
    }

    private static string EscapeText(string value) => value.Replace("\\", "\\\\").Replace(";", "\\;")
        .Replace(",", "\\,").Replace("\r\n", "\\n").Replace("\r", "\\n").Replace("\n", "\\n");

    private static IEnumerable<string> SplitEscapedValues(string value, char separator) {
        var current = new StringBuilder();
        bool escaped = false;
        foreach (char character in value) {
            if (escaped) {
                current.Append('\\').Append(character);
                escaped = false;
            } else if (character == '\\') {
                escaped = true;
            } else if (character == separator) {
                yield return Unescape(current.ToString());
                current.Clear();
            } else {
                current.Append(character);
            }
        }
        if (escaped) current.Append('\\');
        yield return Unescape(current.ToString());
    }

    private static string Unescape(string? value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        var result = new StringBuilder(value!.Length);
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (character != '\\' || index + 1 >= value.Length) {
                result.Append(character);
                continue;
            }
            char escaped = value[++index];
            if (escaped == 'n' || escaped == 'N') result.Append('\n');
            else if (escaped == ',' || escaped == ';' || escaped == '\\') result.Append(escaped);
            else result.Append('\\').Append(escaped);
        }
        return result.ToString();
    }

    private static string? UnescapeOrNull(string? value) => string.IsNullOrWhiteSpace(value)
        ? null
        : Unescape(value);

    private static string EscapeParameter(string value) => value.Replace("^", "^^")
        .Replace("\r\n", "^n").Replace("\r", "^n").Replace("\n", "^n").Replace("\"", "^'");

    private static string EscapeUriValue(string value) => value.Replace("\r", string.Empty).Replace("\n", string.Empty)
        .Replace("%", "%25").Replace(";", "%3B").Replace(",", "%2C");

    private sealed class IcsProperty {
        internal IcsProperty(string name, string value) { Name = name; Value = value; }
        internal string Name { get; }
        internal string Value { get; }
        internal IDictionary<string, string> Parameters { get; } =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }
}
