using OfficeIMO.Email.Data;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email;

/// <summary>Metadata query options for the message-centric local-file workflow.</summary>
public sealed class EmailMessageQuery {
    /// <summary>Exact case-insensitive folder path, or a unique folder name.</summary>
    public string? Folder { get; set; }
    /// <summary>Include descendant folders when a folder was specified.</summary>
    public bool Recurse { get; set; }
    /// <summary>Case-insensitive subject fragment.</summary>
    public string? SubjectContains { get; set; }
    /// <summary>Case-insensitive author name/address fragment.</summary>
    public string? SenderContains { get; set; }
    /// <summary>Inclusive received-date lower bound, falling back to sent date.</summary>
    public DateTimeOffset? Since { get; set; }
    /// <summary>Exclusive received-date upper bound, falling back to sent date.</summary>
    public DateTimeOffset? Before { get; set; }
    /// <summary>Maximum matches returned, independently of the scan budget.</summary>
    public int First { get; set; } = 1000;
    /// <summary>Maximum summaries examined. A reached limit is reported explicitly.</summary>
    public int MaxItemsScanned { get; set; } = 1_000_000;
    internal EmailStoreQuery ToStoreQuery(string? folderId = null) => new EmailStoreQuery(folderId,
        includeDescendants: Recurse, itemKind: OutlookItemKind.Message, subjectContains: SubjectContains,
        senderContains: SenderContains, since: Since, before: Before, maxItemsScanned: MaxItemsScanned, maxResults: First);
}

/// <summary>Handle-free messages and explicit scan completion evidence.</summary>
public sealed class EmailMessageReadResult {
    internal EmailMessageReadResult(IReadOnlyList<EmailMessage> messages, int scanned, bool scanLimit) {
        Messages = messages; ItemsScanned = scanned; StoppedAtScanLimit = scanLimit;
    }
    /// <summary>Matching messages in source order.</summary>
    public IReadOnlyList<EmailMessage> Messages { get; }
    /// <summary>Summaries examined.</summary>
    public int ItemsScanned { get; }
    /// <summary>More references existed after the scan budget was consumed.</summary>
    public bool StoppedAtScanLimit { get; }
}

/// <summary>Reads local email files and offline archives without transferring a resource lifetime to the caller.</summary>
public static class EmailMessageReader {
    /// <summary>Reads matching messages and closes the source before returning.</summary>
    public static EmailMessageReadResult Read(string path, EmailMessageQuery? query = null, CancellationToken cancellationToken = default) {
        var source = new EmailMessageSource(path);
        using var data = source.Open(false, cancellationToken);
        return Read(source, data, query ?? new EmailMessageQuery(), cancellationToken);
    }

    internal static EmailMessageReadResult Read(EmailMessageSource source, EmailDataOpenResult data,
        EmailMessageQuery query, CancellationToken token) {
        source.Validate();
        EmailStoreQuery nativeQuery = query.ToStoreQuery(); // also validates dates and budgets for individual files
        var messages = new List<EmailMessage>();
        if (data.Store != null) {
            EmailStoreSession store = data.Store;
            string? folderId = ResolveFolder(store, query.Folder);
            EmailStoreSearchReport report = store.SearchWithReport(query.ToStoreQuery(folderId), token);
            foreach (EmailStoreSearchResult match in report.Results) {
                token.ThrowIfCancellationRequested();
                EmailStoreItem item = store.ReadItem(match.Reference, new EmailStoreItemReadOptions(
                    EmailStoreItemReadParts.Metadata | EmailStoreItemReadParts.Bodies |
                    EmailStoreItemReadParts.Recipients | EmailStoreItemReadParts.AttachmentMetadata), token);
                string folderPath = string.Join("/", store.FolderCatalog.GetPath(item.FolderKey).Select(f => f.Name));
                messages.Add(new EmailMessage(source, item.Document, match.Reference, folderPath, item.ContentAvailability));
            }
            source.Validate();
            return new EmailMessageReadResult(messages.AsReadOnly(), report.ItemsScanned, report.StoppedAtItemLimit);
        }
        EmailDocument document = data.EmailDocument ?? throw new InvalidDataException("The source is not an email or mail store.");
        if (!string.IsNullOrWhiteSpace(query.Folder)) throw new ArgumentException("Folder filtering requires a mail store.", nameof(query));
        DateTimeOffset? date = document.ReceivedDate ?? document.Date;
        bool contains(string? value, string? fragment) => string.IsNullOrWhiteSpace(fragment) ||
            (value ?? string.Empty).IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0;
        if (contains(document.Subject, nativeQuery.SubjectContains) &&
            contains(document.From?.ToString(), nativeQuery.SenderContains) &&
            (!query.Since.HasValue || date >= query.Since) && (!query.Before.HasValue || date < query.Before))
            messages.Add(new EmailMessage(source, document));
        source.Validate();
        return new EmailMessageReadResult(messages.AsReadOnly(), 1, false);
    }

    private static string? ResolveFolder(EmailStoreSession store, string? path) {
        if (string.IsNullOrWhiteSpace(path)) return null;
        string normalized = path!.Replace('\\', '/').Trim('/');
        var exact = store.Folders.Where(f => string.Equals(string.Join("/", store.FolderCatalog.GetPath(f.Key)
            .Select(p => p.Name)), normalized, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (exact.Length == 1) return exact[0].Id;
        var candidates = store.Folders.Where(f => string.Equals(f.Name, normalized, StringComparison.OrdinalIgnoreCase) ||
            string.Join("/", store.FolderCatalog.GetPath(f.Key).Select(p => p.Name))
                .EndsWith("/" + normalized, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (candidates.Length == 1) return candidates[0].Id;
        throw new ArgumentException(candidates.Length == 0 ? $"Folder '{path}' was not found." :
            $"Folder '{path}' is ambiguous. Use its full folder path.", nameof(path));
    }
}

/// <summary>Explicit scope for repeated message reads from one open local archive.</summary>
public sealed class EmailMessageStore : IDisposable {
    private readonly EmailMessageSource _source;
    private readonly EmailDataOpenResult _data;
    private bool _disposed;
    private EmailMessageStore(EmailMessageSource source, EmailDataOpenResult data) { _source = source; _data = data; }
    /// <summary>Opens one archive. The scope owns its reader.</summary>
    public static EmailMessageStore Open(string path, CancellationToken cancellationToken = default) {
        var source = new EmailMessageSource(path);
        EmailDataOpenResult data = source.Open(false, cancellationToken);
        if (data.Store == null) { data.Dispose(); throw new ArgumentException("A mailbox store file is required.", nameof(path)); }
        return new EmailMessageStore(source, data);
    }
    /// <summary>Source archive path.</summary>
    public string SourcePath => _source.Path;
    /// <summary>Reads messages while the explicit scope remains open.</summary>
    public EmailMessageReadResult Read(EmailMessageQuery? query = null, CancellationToken cancellationToken = default) {
        if (_disposed) throw new ObjectDisposedException(nameof(EmailMessageStore));
        return EmailMessageReader.Read(_source, _data, query ?? new EmailMessageQuery(), cancellationToken);
    }
    /// <summary>Closes the archive. Returned message views remain usable while their source is unchanged.</summary>
    public void Dispose() { if (_disposed) return; _disposed = true; _data.Dispose(); }
}
