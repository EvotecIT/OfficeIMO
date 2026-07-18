using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>
/// Composable typed item changes for one or many PST items. The patch is reusable and does not retain an item.
/// </summary>
public sealed class EmailStoreItemPatch {
    private bool _hasReadState;
    private bool _readState;
    private bool _hasImportance;
    private EmailMessageImportance? _importance;
    private bool _hasSensitivity;
    private int? _sensitivity;
    private string[]? _categories;
    private FollowUpChange _followUp;
    private string? _followUpRequest;
    private DateTimeOffset? _followUpStart;
    private DateTimeOffset? _followUpDue;
    private DateTimeOffset? _followUpCompletedAt;
    private OutlookFollowUpIcon? _followUpIcon;
    private OutlookReminder? _reminder;

    /// <summary>Exact MAPI property changes applied after typed projection.</summary>
    public MapiPropertyPatch Properties { get; } = new MapiPropertyPatch();
    /// <summary>Ordered attachment-list changes.</summary>
    public EmailAttachmentPatch Attachments { get; } = new EmailAttachmentPatch();
    /// <summary>Number of typed, property, and attachment operations in this patch.</summary>
    public int ChangeCount { get; private set; }
    /// <summary>Whether no operation is staged.</summary>
    public bool IsEmpty => ChangeCount == 0 && Properties.IsEmpty && Attachments.IsEmpty;

    /// <summary>Marks the message read or unread.</summary>
    public EmailStoreItemPatch SetReadState(bool isRead) {
        _hasReadState = true;
        _readState = isRead;
        ChangeCount++;
        return this;
    }

    /// <summary>Sets or clears message importance.</summary>
    public EmailStoreItemPatch SetImportance(EmailMessageImportance? importance) {
        _hasImportance = true;
        _importance = importance;
        ChangeCount++;
        return this;
    }

    /// <summary>Sets or clears the raw MAPI sensitivity value.</summary>
    public EmailStoreItemPatch SetSensitivity(int? sensitivity) {
        if (sensitivity.HasValue && (sensitivity.Value < 0 || sensitivity.Value > 3))
            throw new ArgumentOutOfRangeException(nameof(sensitivity), "MAPI sensitivity must be from 0 through 3.");
        _hasSensitivity = true;
        _sensitivity = sensitivity;
        ChangeCount++;
        return this;
    }

    /// <summary>Replaces item category names using Outlook's case-insensitive normalization.</summary>
    public EmailStoreItemPatch SetCategories(IEnumerable<string> categories) {
        if (categories == null) throw new ArgumentNullException(nameof(categories));
        _categories = categories.ToArray();
        var validation = new OutlookCategoryCollection();
        validation.ReplaceWith(_categories);
        _categories = validation.ToArray();
        ChangeCount++;
        return this;
    }

    /// <summary>Sets a follow-up flag.</summary>
    public EmailStoreItemPatch SetFollowUp(string? request = null, DateTimeOffset? start = null,
        DateTimeOffset? due = null, OutlookFollowUpIcon? icon = null) {
        if (start.HasValue && due.HasValue && due.Value < start.Value)
            throw new ArgumentException("A follow-up due date cannot precede its start date.", nameof(due));
        _followUp = FollowUpChange.Flag;
        _followUpRequest = request;
        _followUpStart = start;
        _followUpDue = due;
        _followUpIcon = icon;
        ChangeCount++;
        return this;
    }

    /// <summary>Marks a follow-up complete.</summary>
    public EmailStoreItemPatch CompleteFollowUp(DateTimeOffset completedAt) {
        _followUp = FollowUpChange.Complete;
        _followUpCompletedAt = completedAt;
        ChangeCount++;
        return this;
    }

    /// <summary>Clears follow-up state.</summary>
    public EmailStoreItemPatch ClearFollowUp() {
        _followUp = FollowUpChange.Clear;
        ChangeCount++;
        return this;
    }

    /// <summary>Replaces common message, appointment, or task reminder fields.</summary>
    public EmailStoreItemPatch SetReminder(OutlookReminder reminder) {
        if (reminder == null) throw new ArgumentNullException(nameof(reminder));
        _reminder = Clone(reminder);
        ChangeCount++;
        return this;
    }

    /// <summary>Appends exact MAPI changes to this patch.</summary>
    public EmailStoreItemPatch PatchProperties(MapiPropertyPatch patch) {
        Properties.Append(patch ?? throw new ArgumentNullException(nameof(patch)));
        return this;
    }

    /// <summary>Appends attachment operations to this patch.</summary>
    public EmailStoreItemPatch PatchAttachments(EmailAttachmentPatch patch) {
        if (patch == null) throw new ArgumentNullException(nameof(patch));
        foreach (EmailAttachmentPatchChange change in patch.Changes) {
            switch (change.Operation) {
                case EmailAttachmentPatchOperation.Add: Attachments.Add(change.Attachment!); break;
                case EmailAttachmentPatchOperation.RemoveAt: Attachments.RemoveAt(change.Index!.Value); break;
                case EmailAttachmentPatchOperation.ReplaceAt:
                    Attachments.ReplaceAt(change.Index!.Value, change.Attachment!); break;
            }
        }
        return this;
    }

    internal int Apply(EmailDocument document) {
        Validate(document);
        if (_hasReadState) document.MessageMetadata.IsRead = _readState;
        if (_hasImportance) document.MessageMetadata.Importance = _importance;
        if (_hasSensitivity) document.MessageMetadata.Sensitivity = _sensitivity;
        if (_categories != null) document.MessageMetadata.Categories.ReplaceWith(_categories);
        switch (_followUp) {
            case FollowUpChange.Flag:
                document.MessageMetadata.FollowUp.SetFlagged(
                    _followUpRequest, _followUpStart, _followUpDue, _followUpIcon);
                break;
            case FollowUpChange.Complete:
                document.MessageMetadata.FollowUp.MarkComplete(_followUpCompletedAt!.Value);
                break;
            case FollowUpChange.Clear:
                document.MessageMetadata.FollowUp.Clear();
                break;
        }
        if (_reminder != null) Copy(_reminder, SelectReminder(document));
        if (!Properties.IsEmpty) document.MapiWritePatch.Append(Properties);
        if (!Attachments.IsEmpty) Attachments.Apply(document.Attachments);
        return checked(ChangeCount + Properties.Changes.Count + Attachments.Changes.Count);
    }

    internal void Validate(EmailDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (!Attachments.IsEmpty) Attachments.Validate(document.Attachments);
    }

    private static OutlookReminder SelectReminder(EmailDocument document) =>
        document.Appointment?.Reminder ?? document.Task?.Reminder ?? document.MessageMetadata.Reminder;

    private static OutlookReminder Clone(OutlookReminder source) {
        var result = new OutlookReminder();
        Copy(source, result);
        return result;
    }

    private static void Copy(OutlookReminder source, OutlookReminder destination) {
        destination.IsSet = source.IsSet;
        destination.DeltaMinutes = source.DeltaMinutes;
        destination.Time = source.Time;
        destination.SignalTime = source.SignalTime;
        destination.Override = source.Override;
        destination.PlaySound = source.PlaySound;
        destination.SoundFile = source.SoundFile;
    }

    private enum FollowUpChange { None, Flag, Complete, Clear }
}

/// <summary>Result of applying one patch to a bounded typed query selection.</summary>
public sealed class EmailStorePstMutationSelectionReport {
    internal EmailStorePstMutationSelectionReport(IReadOnlyList<EmailStoreItemId> itemIds,
        int itemsScanned) {
        ItemIds = itemIds;
        ItemsScanned = itemsScanned;
    }
    /// <summary>Typed stable identifiers selected and patched.</summary>
    public IReadOnlyList<EmailStoreItemId> ItemIds { get; }
    /// <summary>Number of lightweight source rows evaluated.</summary>
    public int ItemsScanned { get; }
    /// <summary>Number of selected and patched items.</summary>
    public int PatchedItems => ItemIds.Count;
}
