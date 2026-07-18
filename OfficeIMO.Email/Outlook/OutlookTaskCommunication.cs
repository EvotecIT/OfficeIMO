namespace OfficeIMO.Email;

/// <summary>Kind of Outlook task communication represented by a message class.</summary>
public enum OutlookTaskCommunicationKind {
    /// <summary>The item is not a task communication.</summary>
    None = 0,
    /// <summary>A task assignment request.</summary>
    Request = 1,
    /// <summary>An acceptance response.</summary>
    Accept = 2,
    /// <summary>A rejection response.</summary>
    Decline = 3,
    /// <summary>An update from an assignee.</summary>
    Update = 4
}

/// <summary>PidLidTaskMode values used by task and task-communication objects.</summary>
public enum OutlookTaskCommunicationMode {
    /// <summary>The task is not assigned.</summary>
    Unassigned = 0,
    /// <summary>The task is embedded in a task request.</summary>
    EmbeddedRequest = 1,
    /// <summary>The assignee accepted the task.</summary>
    Accepted = 2,
    /// <summary>The assignee rejected the task.</summary>
    Rejected = 3,
    /// <summary>The task is embedded in a task update.</summary>
    EmbeddedUpdate = 4,
    /// <summary>The task was assigned to the assigner.</summary>
    AssignedToAssigner = 5
}

/// <summary>PidLidTaskHistory values describing the latest task lifecycle transition.</summary>
public enum OutlookTaskHistoryKind {
    /// <summary>No change.</summary>
    None = 0,
    /// <summary>The assignee accepted the task.</summary>
    Accepted = 1,
    /// <summary>The assignee rejected the task.</summary>
    Rejected = 2,
    /// <summary>A task property changed.</summary>
    Updated = 3,
    /// <summary>The due date changed.</summary>
    DueDateChanged = 4,
    /// <summary>The task was assigned.</summary>
    Assigned = 5
}

/// <summary>A validation issue found in the wire shape of a task communication.</summary>
public sealed class OutlookTaskCommunicationValidationIssue {
    internal OutlookTaskCommunicationValidationIssue(string code, string message, bool isError) {
        Code = code;
        Message = message;
        IsError = isError;
    }
    /// <summary>Stable machine-readable issue code.</summary>
    public string Code { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Whether the issue prevents a conforming task communication.</summary>
    public bool IsError { get; }
}

/// <summary>Validation result for an Outlook task communication.</summary>
public sealed class OutlookTaskCommunicationValidationReport {
    internal OutlookTaskCommunicationValidationReport(IReadOnlyList<OutlookTaskCommunicationValidationIssue> issues) => Issues = issues;
    /// <summary>Detected issues.</summary>
    public IReadOnlyList<OutlookTaskCommunicationValidationIssue> Issues { get; }
    /// <summary>Whether no error-level issue was detected.</summary>
    public bool IsValid => !Issues.Any(issue => issue.IsError);
}

/// <summary>Typed task lifecycle envelope and its required embedded task payload.</summary>
public sealed class OutlookTaskCommunication {
    /// <summary>Communication kind derived from, or written to, the message class.</summary>
    public OutlookTaskCommunicationKind Kind { get; set; }
    /// <summary>The task embedded in the communication's first attachment.</summary>
    public EmailDocument? EmbeddedTask { get; set; }
    /// <summary>The source payload attachment, when the communication was read from an artifact.</summary>
    public EmailAttachment? PayloadAttachment { get; internal set; }

    /// <summary>Creates a communication and ensures that its embedded task has a correlation identifier.</summary>
    public static OutlookTaskCommunication Create(OutlookTaskCommunicationKind kind, EmailDocument embeddedTask) {
        if (kind == OutlookTaskCommunicationKind.None) throw new ArgumentOutOfRangeException(nameof(kind));
        if (embeddedTask == null) throw new ArgumentNullException(nameof(embeddedTask));
        if (embeddedTask.Task == null) throw new ArgumentException("The embedded document must contain an Outlook task.", nameof(embeddedTask));
        embeddedTask.OutlookItemKind = OutlookItemKind.Task;
        embeddedTask.MessageClass = embeddedTask.MessageClass ?? "IPM.Task";
        embeddedTask.Task.GlobalId = embeddedTask.Task.GlobalId ?? Guid.NewGuid();
        return new OutlookTaskCommunication { Kind = kind, EmbeddedTask = embeddedTask };
    }

    /// <summary>Validates the semantic payload and the original attachment wire shape, when available.</summary>
    public OutlookTaskCommunicationValidationReport Validate() {
        var issues = new List<OutlookTaskCommunicationValidationIssue>();
        if (Kind == OutlookTaskCommunicationKind.None)
            issues.Add(new OutlookTaskCommunicationValidationIssue("TASK_COMMUNICATION_KIND_NONE", "A task communication kind is required.", true));
        if (EmbeddedTask == null)
            issues.Add(new OutlookTaskCommunicationValidationIssue("TASK_COMMUNICATION_PAYLOAD_MISSING", "The first attachment does not contain an embedded task.", true));
        else if (EmbeddedTask.Task == null || !((EmbeddedTask.MessageClass ?? "IPM.Task").StartsWith("IPM.Task", StringComparison.OrdinalIgnoreCase)))
            issues.Add(new OutlookTaskCommunicationValidationIssue("TASK_COMMUNICATION_PAYLOAD_NOT_TASK", "The embedded payload is not an Outlook task.", true));
        if (PayloadAttachment != null) {
            if (PayloadAttachment.MapiAttachMethod != 5)
                issues.Add(new OutlookTaskCommunicationValidationIssue("TASK_COMMUNICATION_ATTACH_METHOD", "The task payload attachment must use embedded-message method 5.", true));
            if (!PayloadAttachment.IsHidden)
                issues.Add(new OutlookTaskCommunicationValidationIssue("TASK_COMMUNICATION_ATTACHMENT_VISIBLE", "The task payload attachment must be hidden.", true));
            if (PayloadAttachment.RenderingPosition != -1)
                issues.Add(new OutlookTaskCommunicationValidationIssue("TASK_COMMUNICATION_RENDERING_POSITION", "The task payload attachment must use rendering position -1.", true));
        }
        return new OutlookTaskCommunicationValidationReport(issues);
    }
}
