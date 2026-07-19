using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>State of one reminder relative to the query's explicit as-of instant.</summary>
public enum EmailStoreReminderState {
    /// <summary>Reminder properties exist but PidLidReminderSet is not true.</summary>
    Disabled = 0,
    /// <summary>Active reminder whose signal time is in the future.</summary>
    Pending = 1,
    /// <summary>Active reminder whose signal time has passed.</summary>
    Overdue = 2,
    /// <summary>Active reminder without enough evidence to establish a signal time.</summary>
    ActiveWithoutSignalTime = 3
}
/// <summary>Evidence used to establish the queue signal time.</summary>
public enum EmailStoreReminderSignalSource {
    /// <summary>No signal time could be established.</summary>
    Missing = 0,
    /// <summary>PidLidReminderSignalTime.</summary>
    ReminderSignalTime = 1,
    /// <summary>PidLidReminderTime for a non-calendar item.</summary>
    NonCalendarReminderTime = 2,
    /// <summary>Single-instance appointment start minus PidLidReminderDelta.</summary>
    AppointmentStartMinusDelta = 3
}

/// <summary>One sorted reminder-queue row.</summary>
public sealed class EmailStoreReminderQueueItem {
    internal EmailStoreReminderQueueItem(EmailStoreItemReference reference, EmailStoreFolderInfo folder,
        EmailStoreItemSummary summary, OutlookReminder reminder, DateTimeOffset? signalTime,
        EmailStoreReminderSignalSource signalSource, EmailStoreReminderState state) {
        Reference = reference;
        Folder = folder;
        Summary = summary;
        Reminder = reminder;
        SignalTime = signalTime;
        SignalSource = signalSource;
        State = state;
    }

    /// <summary>Stable Store reference.</summary>
    public EmailStoreItemReference Reference { get; }
    /// <summary>Containing folder.</summary>
    public EmailStoreFolderInfo Folder { get; }
    /// <summary>Projected lightweight identity and subject.</summary>
    public EmailStoreItemSummary Summary { get; }
    /// <summary>Detached reminder property snapshot.</summary>
    public OutlookReminder Reminder { get; }
    /// <summary>Effective signal time, if it can be established without guessing.</summary>
    public DateTimeOffset? SignalTime { get; }
    /// <summary>Evidence used for <see cref="SignalTime"/>.</summary>
    public EmailStoreReminderSignalSource SignalSource { get; }
    /// <summary>Reminder state at the queue's explicit as-of instant.</summary>
    public EmailStoreReminderState State { get; }
}

/// <summary>Bounded, sorted reminder queue and completeness evidence.</summary>
public sealed class EmailStoreReminderQueue {
    internal EmailStoreReminderQueue(IReadOnlyList<EmailStoreReminderQueueItem> items,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics, DateTimeOffset asOf,
        int scannedItems, int excludedFolders, bool isComplete) {
        Items = items;
        Diagnostics = diagnostics;
        AsOf = asOf;
        ScannedItems = scannedItems;
        ExcludedFolders = excludedFolders;
        IsComplete = isComplete;
    }

    /// <summary>Rows sorted by signal time, folder ID, then item ID.</summary>
    public IReadOnlyList<EmailStoreReminderQueueItem> Items { get; }
    /// <summary>Per-item and bound diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>UTC instant used to classify state.</summary>
    public DateTimeOffset AsOf { get; }
    /// <summary>Normal item references examined.</summary>
    public int ScannedItems { get; }
    /// <summary>Folders skipped because Outlook excludes them from the reminder domain.</summary>
    public int ExcludedFolders { get; }
    /// <summary>True when the complete selected domain was inspected within all bounds.</summary>
    public bool IsComplete { get; }
}
