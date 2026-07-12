namespace OfficeIMO.Email;

/// <summary>Typed Outlook task fields.</summary>
public sealed class OutlookTask {
    private readonly List<string> _contacts = new List<string>();
    private readonly List<string> _companies = new List<string>();
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
    /// <summary>Reminder lead time in minutes.</summary>
    public int? ReminderDeltaMinutes { get; set; }
    /// <summary>Whether a reminder is enabled.</summary>
    public bool? ReminderIsSet { get; set; }
    /// <summary>Reminder reference time.</summary>
    public DateTimeOffset? ReminderTime { get; set; }
    /// <summary>Reminder signal time.</summary>
    public DateTimeOffset? ReminderSignalTime { get; set; }
    /// <summary>Common start time.</summary>
    public DateTimeOffset? CommonStart { get; set; }
    /// <summary>Common end time.</summary>
    public DateTimeOffset? CommonEnd { get; set; }
    /// <summary>Task mode numeric value.</summary>
    public int? Mode { get; set; }
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
