namespace OfficeIMO.Email;

/// <summary>State of an Outlook informational follow-up flag.</summary>
public enum OutlookFollowUpStatus {
    /// <summary>The item is not flagged. Outlook represents this by omitting PidTagFlagStatus.</summary>
    None = 0,
    /// <summary>The follow-up work is complete.</summary>
    Complete = 1,
    /// <summary>The item is flagged for follow-up.</summary>
    Flagged = 2
}

/// <summary>Color/icon associated with an Outlook follow-up flag.</summary>
public enum OutlookFollowUpIcon {
    /// <summary>No explicit color.</summary>
    None = 0,
    /// <summary>Purple flag.</summary>
    Purple = 1,
    /// <summary>Orange flag.</summary>
    Orange = 2,
    /// <summary>Green flag.</summary>
    Green = 3,
    /// <summary>Yellow flag.</summary>
    Yellow = 4,
    /// <summary>Blue flag.</summary>
    Blue = 5,
    /// <summary>Red flag.</summary>
    Red = 6
}

/// <summary>Typed Outlook follow-up state for a non-task message or contact.</summary>
public sealed class OutlookFollowUp {
    /// <summary>Flag state. <see langword="null"/> retains an absent or unknown source value.</summary>
    public OutlookFollowUpStatus? Status { get; set; }
    /// <summary>Original numeric status, including values not recognized by this OfficeIMO version.</summary>
    public int? RawStatus { get; set; }
    /// <summary>User-facing follow-up request text.</summary>
    public string? Request { get; set; }
    /// <summary>Consolidated to-do-list title.</summary>
    public string? Title { get; set; }
    /// <summary>Start of the follow-up window.</summary>
    public DateTimeOffset? Start { get; set; }
    /// <summary>Due date of the follow-up window.</summary>
    public DateTimeOffset? Due { get; set; }
    /// <summary>Time at which the flag was completed.</summary>
    public DateTimeOffset? CompletedAt { get; set; }
    /// <summary>Optional flag color/icon.</summary>
    public OutlookFollowUpIcon? Icon { get; set; }
    /// <summary>Outlook predefined flag-string index, retained independently of localized request text.</summary>
    public int? FlagString { get; set; }
    /// <summary>Delivery-time proof used by Outlook to assess locally changed request text.</summary>
    public DateTimeOffset? ValidRequestProof { get; set; }
    /// <summary>Raw consolidated to-do flags.</summary>
    public int? ToDoItemFlags { get; set; }

    /// <summary>Sets a follow-up flag while retaining caller-selected text and scheduling.</summary>
    public void SetFlagged(string? request = null, DateTimeOffset? start = null,
        DateTimeOffset? due = null, OutlookFollowUpIcon? icon = null) {
        if (start.HasValue && due.HasValue && due.Value < start.Value) {
            throw new ArgumentException("A follow-up due date cannot precede its start date.", nameof(due));
        }
        Status = OutlookFollowUpStatus.Flagged;
        RawStatus = (int)OutlookFollowUpStatus.Flagged;
        Request = request;
        Start = start;
        Due = due;
        CompletedAt = null;
        Icon = icon;
    }

    /// <summary>Marks the follow-up flag complete.</summary>
    public void MarkComplete(DateTimeOffset completedAt) {
        Status = OutlookFollowUpStatus.Complete;
        RawStatus = (int)OutlookFollowUpStatus.Complete;
        CompletedAt = completedAt;
        Icon = null;
    }

    /// <summary>Clears the flag and its scheduling metadata.</summary>
    public void Clear() {
        Status = OutlookFollowUpStatus.None;
        RawStatus = 0;
        Request = null;
        Title = null;
        Start = null;
        Due = null;
        CompletedAt = null;
        Icon = null;
        FlagString = null;
        ValidRequestProof = null;
        ToDoItemFlags = null;
    }
}
