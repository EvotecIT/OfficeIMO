namespace OfficeIMO.Email;

/// <summary>Kind of Outlook meeting communication represented by a message class.</summary>
public enum OutlookMeetingCommunicationKind {
    /// <summary>The item is not a meeting communication.</summary>
    None = 0,
    /// <summary>A meeting request or update.</summary>
    RequestOrUpdate = 1,
    /// <summary>A meeting cancellation.</summary>
    Cancellation = 2,
    /// <summary>An accepted response.</summary>
    ResponseAccepted = 3,
    /// <summary>A tentative response.</summary>
    ResponseTentative = 4,
    /// <summary>A declined response.</summary>
    ResponseDeclined = 5,
    /// <summary>A meeting-forward notification.</summary>
    ForwardNotification = 6
}

/// <summary>Protocol value describing the type of meeting request or update.</summary>
[Flags]
public enum OutlookMeetingRequestType {
    /// <summary>No meeting request type was stamped.</summary>
    None = 0,
    /// <summary>An initial request.</summary>
    InitialRequest = 0x00000001,
    /// <summary>A significant update such as a time or recurrence change.</summary>
    FullUpdate = 0x00010000,
    /// <summary>An informational update.</summary>
    InformationalUpdate = 0x00020000,
    /// <summary>The request is out of date.</summary>
    OutOfDate = 0x00080000,
    /// <summary>The item is a delegator copy.</summary>
    DelegatorCopy = 0x00100000
}

/// <summary>Typed lifecycle envelope layered over an Outlook appointment payload.</summary>
public sealed class OutlookMeetingCommunication {
    /// <summary>Communication kind derived from, or written to, the message class.</summary>
    public OutlookMeetingCommunicationKind Kind { get; set; }
    /// <summary>Raw PidLidMeetingType flags.</summary>
    public int? RequestTypeValue { get; set; }
    /// <summary>Typed meeting-type flags.</summary>
    public OutlookMeetingRequestType? RequestType {
        get => RequestTypeValue.HasValue ? (OutlookMeetingRequestType)RequestTypeValue.Value : (OutlookMeetingRequestType?)null;
        set => RequestTypeValue = value.HasValue ? (int)value.Value : (int?)null;
    }
    /// <summary>Busy status the organizer intends for attendees.</summary>
    public int? IntendedBusyStatus { get; set; }
    /// <summary>When the organizer sent the request or significant update.</summary>
    public DateTimeOffset? OwnerCriticalChange { get; set; }
    /// <summary>When an attendee last changed a response.</summary>
    public DateTimeOffset? AttendeeCriticalChange { get; set; }
    /// <summary>Whether a response is silent.</summary>
    public bool? IsSilent { get; set; }
    /// <summary>Whether a response proposes another time.</summary>
    public bool? IsCounterProposal { get; set; }
    /// <summary>Proposed replacement start time.</summary>
    public DateTimeOffset? ProposedStart { get; set; }
    /// <summary>Proposed replacement end time.</summary>
    public DateTimeOffset? ProposedEnd { get; set; }
    /// <summary>Proposed duration in minutes.</summary>
    public int? ProposedDurationMinutes { get; set; }
    /// <summary>Time at which the current response was recorded.</summary>
    public DateTimeOffset? ReplyAt { get; set; }
    /// <summary>Name associated with the current response.</summary>
    public string? ReplyName { get; set; }
}
