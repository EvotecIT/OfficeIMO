namespace OfficeIMO.Email.Store;

/// <summary>Immutable typed column selection for a Store table query.</summary>
public sealed class EmailStoreProjection {
    private readonly IReadOnlyList<EmailStoreField> _fields;

    /// <summary>Creates a projection from one or more canonical fields.</summary>
    public EmailStoreProjection(params EmailStoreField[] fields) {
        if (fields == null) throw new ArgumentNullException(nameof(fields));
        if (fields.Length == 0) throw new ArgumentException("A projection must contain at least one field.", nameof(fields));
        if (fields.Any(field => field == null)) throw new ArgumentException("A projection cannot contain null fields.", nameof(fields));
        string? duplicate = fields.GroupBy(field => field.Key, StringComparer.Ordinal)
            .FirstOrDefault(group => group.Count() > 1)?.Key;
        if (duplicate != null) throw new ArgumentException(string.Concat("Projection field '", duplicate, "' is duplicated."), nameof(fields));
        _fields = Array.AsReadOnly(fields.ToArray());
    }

    /// <summary>Default browsing projection.</summary>
    public static EmailStoreProjection Default { get; } = new EmailStoreProjection(
        EmailStoreFields.ItemId,
        EmailStoreFields.FolderId,
        EmailStoreFields.OutlookItemKind,
        EmailStoreFields.Subject,
        EmailStoreFields.FromAddress,
        EmailStoreFields.SentAt,
        EmailStoreFields.ReceivedAt,
        EmailStoreFields.HasAttachments,
        EmailStoreFields.IsRead);

    /// <summary>Selected columns in caller order.</summary>
    public IReadOnlyList<EmailStoreField> Fields => _fields;
}
