namespace OfficeIMO.Email;

/// <summary>Kind of ordered attachment-list patch operation.</summary>
public enum EmailAttachmentPatchOperation {
    /// <summary>Append an attachment.</summary>
    Add = 0,
    /// <summary>Remove the attachment currently at an index.</summary>
    RemoveAt = 1,
    /// <summary>Replace the attachment currently at an index.</summary>
    ReplaceAt = 2
}

/// <summary>One immutable attachment-list patch operation.</summary>
public sealed class EmailAttachmentPatchChange {
    internal EmailAttachmentPatchChange(EmailAttachmentPatchOperation operation, int? index,
        EmailAttachment? attachment) {
        Operation = operation;
        Index = index;
        Attachment = attachment;
    }
    /// <summary>Operation to apply.</summary>
    public EmailAttachmentPatchOperation Operation { get; }
    /// <summary>Zero-based index for remove and replace operations.</summary>
    public int? Index { get; }
    /// <summary>Attachment for add and replace operations.</summary>
    public EmailAttachment? Attachment { get; }
}

/// <summary>Ordered, bounds-checked changes to an email attachment collection.</summary>
public sealed class EmailAttachmentPatch {
    private readonly List<EmailAttachmentPatchChange> _changes = new List<EmailAttachmentPatchChange>();
    /// <summary>Ordered immutable view of staged operations.</summary>
    public IReadOnlyList<EmailAttachmentPatchChange> Changes => _changes;
    /// <summary>Whether no attachment operation is staged.</summary>
    public bool IsEmpty => _changes.Count == 0;

    /// <summary>Stages an attachment append.</summary>
    public EmailAttachmentPatch Add(EmailAttachment attachment) {
        if (attachment == null) throw new ArgumentNullException(nameof(attachment));
        _changes.Add(new EmailAttachmentPatchChange(EmailAttachmentPatchOperation.Add, null, attachment));
        return this;
    }

    /// <summary>Stages removal at the current zero-based index.</summary>
    public EmailAttachmentPatch RemoveAt(int index) {
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
        _changes.Add(new EmailAttachmentPatchChange(EmailAttachmentPatchOperation.RemoveAt, index, null));
        return this;
    }

    /// <summary>Stages replacement at the current zero-based index.</summary>
    public EmailAttachmentPatch ReplaceAt(int index, EmailAttachment attachment) {
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
        if (attachment == null) throw new ArgumentNullException(nameof(attachment));
        _changes.Add(new EmailAttachmentPatchChange(EmailAttachmentPatchOperation.ReplaceAt, index, attachment));
        return this;
    }

    /// <summary>Validates every operation against the collection state produced by preceding operations.</summary>
    public void Validate(IList<EmailAttachment> attachments) {
        if (attachments == null) throw new ArgumentNullException(nameof(attachments));
        if (attachments.IsReadOnly && _changes.Count > 0)
            throw new NotSupportedException("The attachment collection is read-only.");
        int count = attachments.Count;
        foreach (EmailAttachmentPatchChange change in _changes) {
            switch (change.Operation) {
                case EmailAttachmentPatchOperation.Add:
                    count = checked(count + 1);
                    break;
                case EmailAttachmentPatchOperation.RemoveAt:
                    EnsureIndex(count, change.Index!.Value);
                    count--;
                    break;
                case EmailAttachmentPatchOperation.ReplaceAt:
                    EnsureIndex(count, change.Index!.Value);
                    break;
                default:
                    throw new InvalidOperationException("The attachment patch contains an unsupported operation.");
            }
        }
    }

    /// <summary>Applies all operations only after the complete ordered patch has passed validation.</summary>
    public void Apply(IList<EmailAttachment> attachments) {
        if (attachments == null) throw new ArgumentNullException(nameof(attachments));
        Validate(attachments);
        foreach (EmailAttachmentPatchChange change in _changes) {
            switch (change.Operation) {
                case EmailAttachmentPatchOperation.Add:
                    attachments.Add(change.Attachment!);
                    break;
                case EmailAttachmentPatchOperation.RemoveAt:
                    attachments.RemoveAt(change.Index!.Value);
                    break;
                case EmailAttachmentPatchOperation.ReplaceAt:
                    attachments[change.Index!.Value] = change.Attachment!;
                    break;
            }
        }
    }

    private static void EnsureIndex(int count, int index) {
        if (index >= count) throw new ArgumentOutOfRangeException(nameof(index),
            "The attachment patch index does not exist in the current collection.");
    }
}
