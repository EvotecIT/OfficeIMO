namespace OfficeIMO.Email;

internal static class IcsCalendarCodec {
    private static readonly DateTimeOffset DeterministicEpoch =
        new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero);

    internal static bool TryProject(string text, EmailDocument document, IList<EmailDiagnostic> diagnostics,
        string location) {
        List<IcsProperty> properties = ParseProperties(text.TrimStart('\uFEFF'));
        IcsProperty? activeComponent = properties.FirstOrDefault(property => property.Name == "BEGIN" &&
            (property.Value.Equals("VEVENT", StringComparison.OrdinalIgnoreCase) ||
             property.Value.Equals("VTODO", StringComparison.OrdinalIgnoreCase)));
        if (activeComponent == null) return false;
        bool isEvent = activeComponent.Value.Equals("VEVENT", StringComparison.OrdinalIgnoreCase);

        IReadOnlyList<IcsProperty> activeProperties = SelectActiveComponentProperties(
            properties, activeComponent.Value);
        document.MimeSemanticProjectionIsIncomplete |= HasIncompleteStoreProjection(properties, activeProperties);
        if (isEvent) ProjectEvent(activeProperties, document, diagnostics, location);
        else ProjectTask(activeProperties, document, diagnostics, location);
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

    private static bool HasIncompleteStoreProjection(IEnumerable<IcsProperty> properties,
        IEnumerable<IcsProperty> activeProperties) {
        int calendarItems = properties.Count(property => property.Name == "BEGIN" &&
            (property.Value.Equals("VEVENT", StringComparison.OrdinalIgnoreCase) ||
             property.Value.Equals("VTODO", StringComparison.OrdinalIgnoreCase)));
        bool hasTimeZone = properties.Any(property => property.Name == "BEGIN" &&
            property.Value.Equals("VTIMEZONE", StringComparison.OrdinalIgnoreCase));
        return calendarItems > 1 || hasTimeZone || activeProperties.Any(property =>
            property.Name == "RRULE" || property.Name == "RDATE" ||
            property.Name == "EXDATE" || property.Name == "RECURRENCE-ID" ||
            property.Name == "ATTACH");
    }

    private static void ProjectEvent(IReadOnlyList<IcsProperty> properties, EmailDocument document,
        IList<EmailDiagnostic> diagnostics, string location) {
        var appointment = document.Appointment ?? new OutlookAppointment();
        document.OutlookItemKind = OutlookItemKind.Appointment;
        document.MessageClass = MessageClassForMethod(GetValue(properties, "METHOD"), properties);
        document.Appointment = appointment;
        ApplyCommon(properties, document);

        appointment.Start = ParseDate(GetProperty(properties, "DTSTART"), diagnostics, location, out bool allDay);
        appointment.End = ParseDate(GetProperty(properties, "DTEND"), diagnostics, location, out _);
        appointment.IsAllDay = allDay;
        appointment.Location = Unescape(GetValue(properties, "LOCATION"));
        appointment.Sequence = ParseInt(GetValue(properties, "SEQUENCE"));
        appointment.IsRecurring = GetProperty(properties, "RRULE") != null;
        appointment.RecurrencePattern = GetValue(properties, "RRULE");
        TimeSpan? duration = ParseDuration(GetValue(properties, "DURATION"));
        bool incompleteDuration = false;
        TimeSpan? effectiveDuration = appointment.Start.HasValue && appointment.End.HasValue
            ? appointment.End.Value - appointment.Start.Value
            : duration;
        appointment.DurationMinutes = effectiveDuration.HasValue
            ? ConvertDurationMinutes(effectiveDuration.Value, diagnostics, location, ref incompleteDuration)
            : null;
        if (!appointment.End.HasValue && appointment.Start.HasValue && duration.HasValue) {
            try {
                appointment.End = appointment.Start.Value.Add(duration.Value);
            } catch (ArgumentOutOfRangeException) {
                ReportDurationOutOfRange(diagnostics, location);
                incompleteDuration = true;
            }
        }
        string? busy = GetValue(properties, "X-MICROSOFT-CDO-BUSYSTATUS");
        appointment.BusyStatus = ParseBusyStatus(busy);
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
        task.EstimatedEffort = ParseDuration(GetValue(properties, "X-OFFICEIMO-ESTIMATED-EFFORT")) ??
            ParseDuration(GetValue(properties, "DURATION"));
        task.ActualEffort = ParseDuration(GetValue(properties, "X-OFFICEIMO-ACTUAL-EFFORT"));
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
        document.MimeSemanticProjectionIsIncomplete |= ApplyReminder(properties, task, diagnostics, location);
    }

    private static void ApplyCommon(IReadOnlyList<IcsProperty> properties, EmailDocument document) {
        string? summary = Unescape(GetValue(properties, "SUMMARY"));
        string? description = Unescape(GetValue(properties, "DESCRIPTION"));
        string? uid = Unescape(GetValue(properties, "UID"));
        if (!string.IsNullOrWhiteSpace(summary)) document.Subject = summary;
        if (!string.IsNullOrWhiteSpace(description) && document.Body.Text == null) document.Body.Text = description;
        if (!string.IsNullOrWhiteSpace(uid) && document.MessageId == null) document.MessageId = uid;
        string? organizer = GetOrganizer(properties);
        if (document.From == null && !string.IsNullOrWhiteSpace(organizer)) document.From = new EmailAddress(organizer!);
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
                    string.Concat("DURATION:", FormatDuration(TimeSpan.FromMinutes(appointment.DurationMinutes.Value))));
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
            string.Concat("X-OFFICEIMO-ESTIMATED-EFFORT:", FormatDuration(task.EstimatedEffort.Value)));
        if (!string.IsNullOrWhiteSpace(task.Owner)) {
            if (task.Owner!.IndexOf('@') >= 0) AppendLine(output,
                string.Concat("ORGANIZER:mailto:", EscapeUriValue(task.Owner)));
            else AppendText(output, "X-OFFICEIMO-TASK-OWNER", task.Owner);
        }
        if (task.ActualEffort.HasValue) AppendLine(output,
            string.Concat("X-OFFICEIMO-ACTUAL-EFFORT:", FormatDuration(task.ActualEffort.Value)));
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
    }

    private static void WriteOrganizerAndAttendees(StringBuilder output, EmailDocument document) {
        if (!string.IsNullOrWhiteSpace(document.From?.Address)) {
            string organizer = string.Concat("ORGANIZER");
            if (!string.IsNullOrWhiteSpace(document.From!.DisplayName)) organizer += string.Concat(";CN=\"",
                EscapeParameter(document.From.DisplayName!), "\"");
            AppendLine(output, string.Concat(organizer, ":mailto:", EscapeUriValue(document.From.Address!)));
        }
        foreach (EmailRecipient recipient in document.Recipients.Where(recipient =>
            (recipient.Kind == EmailRecipientKind.To || recipient.Kind == EmailRecipientKind.Cc ||
             recipient.Kind == EmailRecipientKind.Room || recipient.Kind == EmailRecipientKind.Resource) &&
            !string.IsNullOrWhiteSpace(recipient.Address.Address))) {
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

    private static List<IcsProperty> ParseProperties(string text) {
        string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var unfolded = new List<string>();
        foreach (string line in normalized.Split('\n')) {
            if (line.Length > 0 && (line[0] == ' ' || line[0] == '\t') && unfolded.Count > 0) {
                unfolded[unfolded.Count - 1] += line.Substring(1);
            } else {
                unfolded.Add(line);
            }
        }

        var result = new List<IcsProperty>();
        foreach (string line in unfolded) {
            int colon = FindUnquotedSeparator(line, ':');
            if (colon <= 0) continue;
            IReadOnlyList<string> tokens = SplitUnquoted(line.Substring(0, colon), ';');
            var property = new IcsProperty(tokens[0].Trim().ToUpperInvariant(), line.Substring(colon + 1));
            for (int index = 1; index < tokens.Count; index++) {
                int equals = FindUnquotedSeparator(tokens[index], '=');
                if (equals > 0) property.Parameters[tokens[index].Substring(0, equals).Trim()] =
                    tokens[index].Substring(equals + 1).Trim().Trim('"');
            }
            result.Add(property);
        }
        return result;
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

    private static int FindUnquotedSeparator(string value, char separator) {
        bool quoted = false;
        bool escaped = false;
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (escaped) {
                escaped = false;
            } else if (character == '\\') {
                escaped = true;
            } else if (character == '"') {
                quoted = !quoted;
            } else if (!quoted && character == separator) {
                return index;
            }
        }
        return -1;
    }

    private static IReadOnlyList<string> SplitUnquoted(string value, char separator) {
        var result = new List<string>();
        int start = 0;
        while (start <= value.Length) {
            int relative = FindUnquotedSeparator(value.Substring(start), separator);
            if (relative < 0) {
                result.Add(value.Substring(start));
                break;
            }
            result.Add(value.Substring(start, relative));
            start += relative + 1;
        }
        return result;
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
            } else if (kind == EmailRecipientKind.Room || kind == EmailRecipientKind.Resource) {
                existing.Kind = kind;
            }
        }
    }

    private static string? StripMailTo(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string result = value!.Trim();
        return result.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ? result.Substring(7) : result;
    }

    private static int? ParseInt(string? value) => int.TryParse(value, NumberStyles.Integer,
        CultureInfo.InvariantCulture, out int result) ? result : (int?)null;

    private static TimeSpan? ParseDuration(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string normalized = value!.Trim().ToUpperInvariant();
        bool negative = normalized.StartsWith("-", StringComparison.Ordinal);
        if (negative || normalized.StartsWith("+", StringComparison.Ordinal)) normalized = normalized.Substring(1);
        if (normalized.Length > 2 && normalized[0] == 'P' && normalized[normalized.Length - 1] == 'W' &&
            long.TryParse(normalized.Substring(1, normalized.Length - 2), NumberStyles.None,
                CultureInfo.InvariantCulture, out long weeks)) {
            try {
                long ticks = checked(weeks * 7L * TimeSpan.TicksPerDay);
                return TimeSpan.FromTicks(negative ? checked(-ticks) : ticks);
            } catch (OverflowException) {
                return null;
            }
        }
        try {
            string xmlValue = negative ? string.Concat("-", normalized) : normalized;
            return System.Xml.XmlConvert.ToTimeSpan(xmlValue);
        } catch (FormatException) {
            return null;
        } catch (OverflowException) {
            return null;
        }
    }

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

    internal static string GetMethod(EmailDocument document) {
        string messageClass = document.MessageClass ?? string.Empty;
        if (messageClass.IndexOf("Canceled", StringComparison.OrdinalIgnoreCase) >= 0) return "CANCEL";
        if (messageClass.IndexOf("Request", StringComparison.OrdinalIgnoreCase) >= 0) return "REQUEST";
        if (messageClass.IndexOf("Resp", StringComparison.OrdinalIgnoreCase) >= 0) return "REPLY";
        return document.Recipients.Any(recipient => recipient.Kind == EmailRecipientKind.To ||
            recipient.Kind == EmailRecipientKind.Cc || recipient.Kind == EmailRecipientKind.Room ||
            recipient.Kind == EmailRecipientKind.Resource) ? "REQUEST" : "PUBLISH";
    }

    private static string FormatUtc(DateTimeOffset value) =>
        value.UtcDateTime.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);

    private static string FormatDate(DateTimeOffset value) =>
        value.Date.ToString("yyyyMMdd", CultureInfo.InvariantCulture);

    private static string FormatDuration(TimeSpan value) {
        string sign = value < TimeSpan.Zero ? "-" : string.Empty;
        TimeSpan absolute = value.Duration();
        var result = new StringBuilder(sign).Append('P');
        if (absolute.Days > 0) result.Append(absolute.Days.ToString(CultureInfo.InvariantCulture)).Append('D');
        if (absolute.Hours > 0 || absolute.Minutes > 0 || absolute.Seconds > 0 || absolute.Days == 0) {
            result.Append('T');
            if (absolute.Hours > 0) result.Append(absolute.Hours.ToString(CultureInfo.InvariantCulture)).Append('H');
            if (absolute.Minutes > 0) result.Append(absolute.Minutes.ToString(CultureInfo.InvariantCulture)).Append('M');
            if (absolute.Seconds > 0 || absolute == TimeSpan.Zero) {
                result.Append(absolute.Seconds.ToString(CultureInfo.InvariantCulture)).Append('S');
            }
        }
        return result.ToString();
    }

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
        TimeSpan? value = ParseDuration(trigger.Value);
        return value.HasValue
            ? ConvertDurationMinutes(value.Value, diagnostics, location, ref incomplete, invert: true)
            : null;
    }

    private static int? ConvertDurationMinutes(TimeSpan value, IList<EmailDiagnostic> diagnostics,
        string location, ref bool incomplete, bool invert = false) {
        double minutes = value.TotalMinutes;
        if (invert) minutes = -minutes;
        if (minutes >= int.MinValue && minutes <= int.MaxValue) return (int)minutes;
        ReportDurationOutOfRange(diagnostics, location);
        incomplete = true;
        return null;
    }

    private static void ReportDurationOutOfRange(IList<EmailDiagnostic> diagnostics, string location) {
        if (diagnostics.Any(diagnostic => diagnostic.Code == "EMAIL_ICALENDAR_DURATION_OUT_OF_RANGE" &&
            string.Equals(diagnostic.Location, location, StringComparison.Ordinal))) return;
        diagnostics.Add(new EmailDiagnostic("EMAIL_ICALENDAR_DURATION_OUT_OF_RANGE",
            "An iCalendar duration exceeds the supported whole-minute range and was retained only in the semantic source.",
            EmailDiagnosticSeverity.Warning, location));
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
            FormatDuration(TimeSpan.FromMinutes(-deltaMinutes.Value))));
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

    private static string Unescape(string? value) => (value ?? string.Empty).Replace("\\n", "\n")
        .Replace("\\N", "\n").Replace("\\,", ",").Replace("\\;", ";").Replace("\\\\", "\\");

    private static string? UnescapeOrNull(string? value) => string.IsNullOrWhiteSpace(value)
        ? null
        : Unescape(value);

    private static string EscapeParameter(string value) => value.Replace("\\", "\\\\").Replace("\"", "\\\"")
        .Replace("\r", string.Empty).Replace("\n", " ");

    private static string EscapeUriValue(string value) => value.Replace("\r", string.Empty).Replace("\n", string.Empty)
        .Replace(";", "%3B").Replace(",", "%2C");

    private sealed class IcsProperty {
        internal IcsProperty(string name, string value) { Name = name; Value = value; }
        internal string Name { get; }
        internal string Value { get; }
        internal IDictionary<string, string> Parameters { get; } =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }
}
