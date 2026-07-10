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
    /// <summary>Human-readable recurrence pattern.</summary>
    public string? RecurrencePattern { get; set; }
    /// <summary>Opaque recurrence-state payload retained for lossless processing.</summary>
    public byte[]? RecurrenceState { get; set; }
}

/// <summary>Typed Outlook contact fields.</summary>
public sealed class OutlookContact {
    /// <summary>Given name.</summary>
    public string? GivenName { get; set; }
    /// <summary>Surname.</summary>
    public string? Surname { get; set; }
    /// <summary>Company name.</summary>
    public string? CompanyName { get; set; }
    /// <summary>Job title.</summary>
    public string? JobTitle { get; set; }
    /// <summary>Business telephone.</summary>
    public string? BusinessPhone { get; set; }
    /// <summary>Home telephone.</summary>
    public string? HomePhone { get; set; }
    /// <summary>Mobile telephone.</summary>
    public string? MobilePhone { get; set; }
    /// <summary>File-as display value.</summary>
    public string? FileAs { get; set; }
    /// <summary>First electronic-mail address.</summary>
    public string? Email1Address { get; set; }
}

/// <summary>Typed Outlook task fields.</summary>
public sealed class OutlookTask {
    /// <summary>Task start.</summary>
    public DateTimeOffset? Start { get; set; }
    /// <summary>Task due date.</summary>
    public DateTimeOffset? Due { get; set; }
    /// <summary>Task-status numeric value.</summary>
    public int? Status { get; set; }
    /// <summary>Completion ratio from zero through one.</summary>
    public double? PercentComplete { get; set; }
    /// <summary>Whether the task is complete.</summary>
    public bool? IsComplete { get; set; }
    /// <summary>Task owner.</summary>
    public string? Owner { get; set; }
}

/// <summary>Typed Outlook journal fields.</summary>
public sealed class OutlookJournal {
    /// <summary>Journal activity start.</summary>
    public DateTimeOffset? Start { get; set; }
    /// <summary>Journal activity end.</summary>
    public DateTimeOffset? End { get; set; }
    /// <summary>Journal activity type.</summary>
    public string? Type { get; set; }
}

/// <summary>Typed Outlook sticky-note fields.</summary>
public sealed class OutlookNote {
    /// <summary>Outlook note color numeric value.</summary>
    public int? Color { get; set; }
    /// <summary>Saved note window width.</summary>
    public int? Width { get; set; }
    /// <summary>Saved note window height.</summary>
    public int? Height { get; set; }
}
