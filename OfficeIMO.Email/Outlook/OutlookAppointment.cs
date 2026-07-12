namespace OfficeIMO.Email;

/// <summary>Typed appointment and meeting fields projected from Outlook named properties.</summary>
public sealed class OutlookAppointment {
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
    /// <summary>Whether the appointment belongs to a recurring series.</summary>
    public bool? IsRecurring { get; set; }
    /// <summary>Calendar-assistant client-intent flags.</summary>
    public int? ClientIntentFlags { get; set; }
    /// <summary>Whether an appointment reminder is enabled.</summary>
    public bool? ReminderIsSet { get; set; }
    /// <summary>Reminder lead time in minutes.</summary>
    public int? ReminderDeltaMinutes { get; set; }
    /// <summary>Reminder reference time.</summary>
    public DateTimeOffset? ReminderTime { get; set; }
    /// <summary>Reminder signal time.</summary>
    public DateTimeOffset? ReminderSignalTime { get; set; }
    /// <summary>Legacy appointment time-zone description.</summary>
    public string? TimeZoneDescription { get; set; }
    /// <summary>Legacy appointment time-zone structure.</summary>
    public byte[]? TimeZoneStructure { get; set; }
    /// <summary>Start time-zone definition retained in native form.</summary>
    public byte[]? StartTimeZoneDefinition { get; set; }
    /// <summary>End time-zone definition retained in native form.</summary>
    public byte[]? EndTimeZoneDefinition { get; set; }
    /// <summary>Recurring-series time-zone definition retained in native form.</summary>
    public byte[]? RecurrenceTimeZoneDefinition { get; set; }
}
