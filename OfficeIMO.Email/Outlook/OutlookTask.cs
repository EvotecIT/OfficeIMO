namespace OfficeIMO.Email;

/// <summary>Typed Outlook task fields.</summary>
public sealed class OutlookTask {
    private readonly List<string> _contacts = new List<string>();
    private readonly List<string> _companies = new List<string>();
    /// <summary>Task reminder.</summary>
    public OutlookReminder Reminder { get; } = new OutlookReminder();
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
    /// <summary>Estimated effort.</summary>
    public TimeSpan? EstimatedEffort { get; set; }
    /// <summary>Actual effort.</summary>
    public TimeSpan? ActualEffort { get; set; }
    /// <summary>Whether assignment updates should be sent.</summary>
    public bool? SendUpdates { get; set; }
    /// <summary>Whether a completion report should be sent.</summary>
    public bool? SendStatusOnComplete { get; set; }
    /// <summary>Task ownership numeric value.</summary>
    public int? Ownership { get; set; }
    /// <summary>Task acceptance-state numeric value.</summary>
    public int? AcceptanceState { get; set; }
    /// <summary>Task version.</summary>
    public int? Version { get; set; }
    /// <summary>Task state numeric value.</summary>
    public int? State { get; set; }
    /// <summary>Task assigner.</summary>
    public string? Assigner { get; set; }
    /// <summary>Whether this is a team task.</summary>
    public bool? IsTeamTask { get; set; }
    /// <summary>Task ordinal used for ordering.</summary>
    public int? Ordinal { get; set; }
    /// <summary>Whether the task is recurring.</summary>
    public bool? IsRecurring { get; set; }
    /// <summary>Opaque task RecurrencePattern payload retained for lossless processing.</summary>
    public byte[]? RecurrenceState { get; set; }
    /// <summary>
    /// Typed task recurrence. When present, MSG writing encodes this value instead of <see cref="RecurrenceState"/>.
    /// </summary>
    public OutlookRecurrence? Recurrence { get; set; }
    /// <summary>Reminder lead time in minutes.</summary>
    public int? ReminderDeltaMinutes { get => Reminder.DeltaMinutes; set => Reminder.DeltaMinutes = value; }
    /// <summary>Whether a reminder is enabled.</summary>
    public bool? ReminderIsSet { get => Reminder.IsSet; set => Reminder.IsSet = value; }
    /// <summary>Reminder reference time.</summary>
    public DateTimeOffset? ReminderTime { get => Reminder.Time; set => Reminder.Time = value; }
    /// <summary>Reminder signal time.</summary>
    public DateTimeOffset? ReminderSignalTime { get => Reminder.SignalTime; set => Reminder.SignalTime = value; }
    /// <summary>Common start time.</summary>
    public DateTimeOffset? CommonStart { get; set; }
    /// <summary>Common end time.</summary>
    public DateTimeOffset? CommonEnd { get; set; }
    /// <summary>Task mode numeric value.</summary>
    public int? Mode { get; set; }
    /// <summary>Typed assignment mode when the numeric value is defined by the Outlook task protocol.</summary>
    public OutlookTaskCommunicationMode? CommunicationMode {
        get => Mode.HasValue && Enum.IsDefined(typeof(OutlookTaskCommunicationMode), Mode.Value)
            ? (OutlookTaskCommunicationMode)Mode.Value
            : (OutlookTaskCommunicationMode?)null;
        set => Mode = value.HasValue ? (int)value.Value : (int?)null;
    }
    /// <summary>Whether an assignee has replied to the task request.</summary>
    public bool? IsAccepted { get; set; }
    /// <summary>Numeric history value describing the most recent lifecycle change.</summary>
    public int? History { get; set; }
    /// <summary>Typed history value when <see cref="History"/> is defined by the Outlook task protocol.</summary>
    public OutlookTaskHistoryKind? HistoryKind {
        get => History.HasValue && Enum.IsDefined(typeof(OutlookTaskHistoryKind), History.Value)
            ? (OutlookTaskHistoryKind)History.Value
            : (OutlookTaskHistoryKind?)null;
        set => History = value.HasValue ? (int)value.Value : (int?)null;
    }
    /// <summary>Time of the most recent lifecycle change.</summary>
    public DateTimeOffset? LastUpdate { get; set; }
    /// <summary>User who most recently changed the task.</summary>
    public string? LastUser { get; set; }
    /// <summary>User who most recently assigned or was assigned the task.</summary>
    public string? LastDelegate { get; set; }
    /// <summary>Stable identifier used to correlate an assigned task with task communications.</summary>
    public Guid? GlobalId { get; set; }
    /// <summary>To-do ordinal date.</summary>
    public DateTimeOffset? ToDoOrdinalDate { get; set; }
    /// <summary>To-do subordinal tiebreaker.</summary>
    public string? ToDoSubOrdinal { get; set; }
    /// <summary>Contacts associated with the task.</summary>
    public IList<string> Contacts => _contacts;
    /// <summary>Companies associated with the task.</summary>
    public IList<string> Companies => _companies;
    /// <summary>Billing information.</summary>
    public string? BillingInformation { get; set; }
    /// <summary>Mileage information.</summary>
    public string? Mileage { get; set; }
    /// <summary>Completion time.</summary>
    public DateTimeOffset? CompletedAt { get; set; }
}
