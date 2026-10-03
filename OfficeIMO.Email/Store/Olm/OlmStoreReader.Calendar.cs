using System.Xml.Linq;

namespace OfficeIMO.Email.Store;

internal sealed partial class OlmStoreReader {
    private void ProjectAppointmentCalendar(EmailDocument document, XElement item, string location) {
        OutlookAppointment appointment = document.Appointment!;
        string? reminder = Value(item, "OPFCalendarEventCopyReminderDelta");
        if (reminder != null) {
            // Outlook for Mac's delta is seconds; the common Outlook model uses whole minutes.
            if (decimal.TryParse(reminder, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal seconds) &&
                seconds % 60 == 0 && seconds / 60 >= int.MinValue && seconds / 60 <= int.MaxValue) {
                appointment.ReminderDeltaMinutes = (int)(seconds / 60);
            } else RecordProjectionIssue(document, "EMAIL_STORE_OLM_REMINDER_UNSUPPORTED",
                "The OLM reminder offset cannot be represented as a whole number of Outlook reminder minutes.", location);
        }

        string? zoneText = Value(item, "OPFCalendarEventGetStartTimeZoneICSData");
        if (zoneText != null) {
            try {
                _cancellationToken.ThrowIfCancellationRequested();
                IcsDocument calendar = IcsDocument.Parse(zoneText);
                ContentLineComponent[] zones = calendar.GetComponents("VTIMEZONE").ToArray();
                if (calendar.Calendars.Count != 1 || zones.Length != 1)
                    throw new InvalidDataException("The OLM timezone source requires one calendar and one VTIMEZONE.");
                if (!IcsCalendarCodec.TryConvertTimeZone(zones[0], out OutlookTimeZoneDefinition? zone,
                        out string? error, allowMicrosoft1601PlaceholderDates: true))
                    throw new InvalidDataException(error);
                appointment.StartTimeZone = zone;
                appointment.RecurrenceTimeZone = zone;
            } catch (Exception exception) when (exception is InvalidDataException || exception is FormatException ||
                exception is ArgumentException || exception is OverflowException) {
                RecordProjectionIssue(document, "EMAIL_STORE_OLM_TIMEZONE_UNSUPPORTED", exception.Message, location);
            }
        }

        XElement? source = Child(item, "OPFCalendarEventCopyRecurrence");
        if (source == null && appointment.IsRecurring != true) return;
        OutlookRecurrence? recurrence = null;
        string? reason = null;
        if (source == null || !TryProjectRecurrence(source, appointment, out recurrence, out reason)) {
            RecordProjectionIssue(document, "EMAIL_STORE_OLM_RECURRENCE_UNSUPPORTED",
                source == null ? "The OLM item declares recurrence without a recurrence record." : reason!, location);
            return;
        }
        appointment.Recurrence = recurrence;
        appointment.IsRecurring = true;
    }

    private static bool TryProjectRecurrence(XElement source, OutlookAppointment appointment,
        out OutlookRecurrence? result, out string? error) {
        result = null;
        error = null;
        if (appointment.Start == null || appointment.End == null || appointment.RecurrenceTimeZone == null)
            return RecurrenceFailure("OLM recurrence requires valid start/end instants and supported embedded timezone rules.", out error);
        if (!HasOnlyUniqueChildren(source, "OPFRecurrenceCopyEndDate", "OPFRecurrenceCopyStartDate",
                "OPFRecurrenceGetOccurenceCount", "OPFRecurrenceHasEndDate", "OPFRecurrenceIsNoEnd",
                "OPFRecurrenceIsNumbered", "OPFRecurrencePattern"))
            return RecurrenceFailure("The OLM recurrence contains unsupported or repeated fields, including possible exceptions.", out error);
        XElement? pattern = Child(source, "OPFRecurrencePattern");
        if (pattern == null || !HasOnlyUniqueChildren(pattern, "OPFRecurrencePatternDaysOfWeek",
                "OPFRecurrencePatternInterval", "OPFRecurrencePatternType") ||
            Value(pattern, "OPFRecurrencePatternType") != "OPFRecurrencePatternWeekly" ||
            IntegerValue(pattern, "OPFRecurrencePatternInterval") != 1)
            return RecurrenceFailure("The qualified OLM recurrence mapping requires a weekly pattern with interval one.", out error);
        XElement? days = Child(pattern, "OPFRecurrencePatternDaysOfWeek");
        if (days == null || !TryProjectDays(days, out OutlookRecurrenceDays dayMask))
            return RecurrenceFailure("The OLM weekly recurrence has an unavailable or unsupported weekday mask.", out error);

        try {
            OutlookTimeZoneDefinition zone = appointment.RecurrenceTimeZone!;
            DateTime localStart = zone.ConvertUtc(appointment.Start.Value).DateTime;
            DateTime localEnd = zone.ConvertUtc(appointment.End.Value).DateTime;
            if (zone.ResolveLocal(localStart).UtcDateTime != appointment.Start.Value.UtcDateTime ||
                zone.ResolveLocal(localEnd).UtcDateTime != appointment.End.Value.UtcDateTime)
                return RecurrenceFailure("The OLM occurrence uses an ambiguous local time that the common recurrence model resolves to a different instant.", out error);
            DateTimeOffset? rangeStart = DateValue(source, "OPFRecurrenceCopyStartDate");
            if (localEnd < localStart || rangeStart == null || zone.ConvertUtc(rangeStart.Value).Date != localStart.Date)
                return RecurrenceFailure("The OLM recurrence range or duration differs from its first local occurrence.", out error);
            if (((int)dayMask & (1 << (int)localStart.DayOfWeek)) == 0)
                return RecurrenceFailure("The OLM first occurrence does not match its weekly weekday mask.", out error);
            if (!TryReadRecurrenceFlag(source, "OPFRecurrenceIsNumbered", out bool numbered) ||
                !TryReadRecurrenceFlag(source, "OPFRecurrenceIsNoEnd", out bool noEnd) ||
                !TryReadRecurrenceFlag(source, "OPFRecurrenceHasEndDate", out bool hasEnd) || numbered && noEnd)
                return RecurrenceFailure("The OLM recurrence end-condition flags are unavailable or contradictory.", out error);
            var recurrence = new OutlookRecurrence {
                Frequency = OutlookRecurrenceFrequency.Weekly,
                PatternKind = OutlookRecurrencePatternKind.Week,
                Interval = 1,
                Start = localStart,
                Duration = localEnd - localStart,
                DaysOfWeek = dayMask,
                TimeZoneId = zone.KeyName,
                StateDecoded = true
            };
            DateTimeOffset? endDate = DateValue(source, "OPFRecurrenceCopyEndDate");
            DateTime? localEndDate = endDate.HasValue ? zone.ConvertUtc(endDate.Value).Date : (DateTime?)null;
            if (hasEnd && (!localEndDate.HasValue || localEndDate < localStart.Date))
                return RecurrenceFailure("The OLM recurrence end date is unavailable or precedes its start.", out error);
            if (numbered) {
                int? count = IntegerValue(source, "OPFRecurrenceGetOccurenceCount");
                if (count == null || count <= 0)
                    return RecurrenceFailure("The OLM recurrence occurrence count is unavailable or invalid.", out error);
                recurrence.RangeKind = OutlookRecurrenceRangeKind.OccurrenceCount;
                recurrence.OccurrenceCount = count;
            } else if (noEnd && !hasEnd) recurrence.RangeKind = OutlookRecurrenceRangeKind.NoEnd;
            else if (!noEnd && hasEnd) {
                recurrence.RangeKind = OutlookRecurrenceRangeKind.EndDate;
                recurrence.EndDate = localEndDate;
            } else return RecurrenceFailure("The OLM recurrence end condition cannot be represented.", out error);
            result = recurrence;
            return true;
        } catch (Exception exception) when (exception is InvalidOperationException || exception is ArgumentException ||
            exception is OverflowException) {
            return RecurrenceFailure(exception.Message, out error);
        }
    }

    private static bool TryProjectDays(XElement source, out OutlookRecurrenceDays result) {
        result = OutlookRecurrenceDays.None;
        string[] names = { "sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday",
            "weekdays", "weekenddays", "allDays" };
        if (!HasOnlyUniqueChildren(source, names)) return false;
        foreach (XElement field in source.Elements()) {
            if (!TryReadRecurrenceFlag(source, field.Name.LocalName, out bool selected)) return false;
            if (!selected) continue;
            int index = Array.FindIndex(names, name => string.Equals(name, field.Name.LocalName, StringComparison.OrdinalIgnoreCase));
            result |= index < 7 ? (OutlookRecurrenceDays)(1 << index)
                : index == 7 ? OutlookRecurrenceDays.Weekdays
                : index == 8 ? OutlookRecurrenceDays.Weekend : OutlookRecurrenceDays.All;
        }
        return result != OutlookRecurrenceDays.None;
    }

    private static bool TryReadRecurrenceFlag(XElement source, string name, out bool result) {
        string? value = Value(source, name);
        if (bool.TryParse(value, out result)) return true;
        if (decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal number) &&
            (number == 0 || number == 1)) {
            result = number == 1;
            return true;
        }
        result = false;
        return false;
    }

    private static bool HasOnlyUniqueChildren(XElement source, params string[] names) {
        var remaining = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);
        return source.Elements().All(field => remaining.Remove(field.Name.LocalName));
    }

    private static bool RecurrenceFailure(string message, out string? error) {
        error = message;
        return false;
    }

    private void RecordProjectionIssue(EmailDocument document, string code, string message, string location) {
        _diagnostics.Add(new EmailStoreDiagnostic(code, message, EmailStoreDiagnosticSeverity.Warning, location));
        document.MimeSemanticProjectionIsIncomplete = true;
        string[] previous = document.Properties.TryGetValue("Olm:ProjectionIssues", out object? value) && value is string[] issues
            ? issues : Array.Empty<string>();
        document.Properties["Olm:ProjectionIssues"] = previous.Concat(new[] { message }).Distinct(StringComparer.Ordinal).ToArray();
    }
}
