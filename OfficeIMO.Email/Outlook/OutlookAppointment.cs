namespace OfficeIMO.Email;

/// <summary>Typed appointment and meeting fields projected from Outlook named properties.</summary>
public sealed class OutlookAppointment {
    /// <summary>Appointment reminder.</summary>
    public OutlookReminder Reminder { get; } = new OutlookReminder();
    /// <summary>Native PidLidGlobalObjectId used to correlate a meeting and its lifecycle messages.</summary>
    public byte[]? GlobalObjectId { get; set; }
    /// <summary>
    /// Native PidLidCleanGlobalObjectId used to correlate recurrence exceptions with their master meeting.
    /// </summary>
    public byte[]? CleanGlobalObjectId { get; set; }
    /// <summary>Appointment start.</summary>
    public DateTimeOffset? Start { get; set; }
    /// <summary>Appointment end.</summary>
    public DateTimeOffset? End { get; set; }
    /// <summary>Appointment location.</summary>
    public string? Location { get; set; }
    /// <summary>Whether the appointment covers complete days.</summary>
    public bool? IsAllDay { get; set; }
    /// <summary>Busy-status numeric value.</summary>
    public int? BusyStatus { get; set; }
    /// <summary>Meeting-status flags.</summary>
    public int? MeetingStatus { get; set; }
    /// <summary>Response-status numeric value.</summary>
    public int? ResponseStatus { get; set; }
    /// <summary>Appointment update sequence.</summary>
    public int? Sequence { get; set; }
    /// <summary>Appointment duration in minutes.</summary>
    public int? DurationMinutes { get; set; }
    /// <summary>Display list containing every attendee.</summary>
    public string? AllAttendees { get; set; }
    /// <summary>Display list containing required attendees.</summary>
    public string? RequiredAttendees { get; set; }
    /// <summary>Display list containing optional attendees.</summary>
    public string? OptionalAttendees { get; set; }
    /// <summary>Whether attendees are prevented from proposing another time.</summary>
    public bool? NotAllowPropose { get; set; }
    /// <summary>Outlook recurrence type: none, daily, weekly, monthly, or yearly.</summary>
    public int? RecurrenceType { get; set; }
    /// <summary>Human-readable recurrence pattern.</summary>
    public string? RecurrencePattern { get; set; }
    /// <summary>Opaque recurrence-state payload retained for lossless processing.</summary>
    public byte[]? RecurrenceState { get; set; }
    /// <summary>
    /// Typed recurrence. When present, MSG writing encodes this value instead of <see cref="RecurrenceState"/>.
    /// </summary>
    public OutlookRecurrence? Recurrence { get; set; }
    /// <summary>Whether the appointment belongs to a recurring series.</summary>
    public bool? IsRecurring { get; set; }
    /// <summary>Calendar-assistant client-intent flags.</summary>
    public int? ClientIntentFlags { get; set; }
    /// <summary>Whether an appointment reminder is enabled.</summary>
    public bool? ReminderIsSet { get => Reminder.IsSet; set => Reminder.IsSet = value; }
    /// <summary>Reminder lead time in minutes.</summary>
    public int? ReminderDeltaMinutes { get => Reminder.DeltaMinutes; set => Reminder.DeltaMinutes = value; }
    /// <summary>Reminder reference time.</summary>
    public DateTimeOffset? ReminderTime { get => Reminder.Time; set => Reminder.Time = value; }
    /// <summary>Reminder signal time.</summary>
    public DateTimeOffset? ReminderSignalTime { get => Reminder.SignalTime; set => Reminder.SignalTime = value; }
    /// <summary>Legacy appointment time-zone description.</summary>
    public string? TimeZoneDescription { get; set; }
    /// <summary>Legacy appointment time-zone structure.</summary>
    public byte[]? TimeZoneStructure { get; set; }
    /// <summary>Typed legacy PidLidTimeZoneStruct value.</summary>
    public OutlookTimeZoneStructure? LegacyTimeZone { get; set; }
    /// <summary>Start time-zone definition retained in native form.</summary>
    public byte[]? StartTimeZoneDefinition { get; set; }
    /// <summary>Typed start-display time-zone definition.</summary>
    public OutlookTimeZoneDefinition? StartTimeZone { get; set; }
    /// <summary>End time-zone definition retained in native form.</summary>
    public byte[]? EndTimeZoneDefinition { get; set; }
    /// <summary>Typed end-display time-zone definition.</summary>
    public OutlookTimeZoneDefinition? EndTimeZone { get; set; }
    /// <summary>Recurring-series time-zone definition retained in native form.</summary>
    public byte[]? RecurrenceTimeZoneDefinition { get; set; }
    /// <summary>Typed recurring-series time-zone definition.</summary>
    public OutlookTimeZoneDefinition? RecurrenceTimeZone { get; set; }

    /// <summary>Compares the legacy and definition-based time-zone properties without host-zone assumptions.</summary>
    public OutlookTimeZoneConsistencyReport CheckTimeZoneConsistency(int? localYear = null) {
        int year = localYear ?? Recurrence?.Start.Year ?? Start?.Year ?? 1;
        return OutlookTimeZoneConsistency.Compare(LegacyTimeZone, RecurrenceTimeZone ?? StartTimeZone, year);
    }
}
