using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Adapts the OfficeIMO.Email mbox engine to the shared email-store session model.</summary>
internal sealed class MboxStoreReader {
    private const string FolderId = "mbox:folder:root";
    private readonly EmailStoreReaderOptions _options;

    internal MboxStoreReader(EmailStoreReaderOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    internal EmailStoreReadResult Read(Stream stream, string? sourceName,
        CancellationToken cancellationToken) {
        var mailboxOptions = new EmailMailboxReaderOptions(
            _options.MaxInputBytes,
            EmailStoreMessageReader.CreateOptions(_options),
            MboxVariant.Auto,
            _options.MaxItemCount);
        EmailMailboxReadResult result;
        try {
            result = new EmailMailboxReader(mailboxOptions).Read(stream, cancellationToken);
        } catch (EmailLimitExceededException exception) {
            throw ConvertLimit(exception);
        }

        var store = new EmailStore {
            Format = EmailStoreFormat.Mbox,
            DisplayName = GetDisplayName(sourceName)
        };
        var folder = new EmailStoreFolder(FolderId, null, store.DisplayName ?? "Mailbox");
        store.MutableFolders.Add(folder);
        for (int index = 0; index < result.Mailbox.Messages.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            EmailMailboxEntry entry = result.Mailbox.Messages[index];
            string itemId = "mbox:item:" + index.ToString("D8", CultureInfo.InvariantCulture);
            entry.Document.Properties["EmailStore:Format"] = EmailStoreFormat.Mbox.ToString();
            entry.Document.Properties["EmailStore:ItemId"] = itemId;
            entry.Document.Properties["EmailStore:FolderId"] = FolderId;
            if (entry.EnvelopeSender != null) entry.Document.Properties["Mbox:EnvelopeSender"] = entry.EnvelopeSender;
            if (entry.EnvelopeDate.HasValue) entry.Document.Properties["Mbox:EnvelopeDate"] = entry.EnvelopeDate.Value;
            if (entry.RawFromLine != null) entry.Document.Properties["Mbox:RawFromLine"] = entry.RawFromLine;
            folder.MutableItems.Add(new EmailStoreItem(itemId, FolderId, entry.Document,
                format: EmailStoreFormat.Mbox));
        }

        var diagnostics = result.Diagnostics.Select(ConvertDiagnostic).ToArray();
        return new EmailStoreReadResult(store, diagnostics, result.BytesRead);
    }

    private static string? GetDisplayName(string? sourceName) {
        if (string.IsNullOrWhiteSpace(sourceName)) return "Mailbox";
        try { return Path.GetFileNameWithoutExtension(sourceName); }
        catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException) {
            return sourceName;
        }
    }

    private static EmailStoreDiagnostic ConvertDiagnostic(EmailDiagnostic diagnostic) =>
        new EmailStoreDiagnostic(diagnostic.Code, diagnostic.Message,
            diagnostic.Severity == EmailDiagnosticSeverity.Error
                ? EmailStoreDiagnosticSeverity.Error
                : diagnostic.Severity == EmailDiagnosticSeverity.Information
                    ? EmailStoreDiagnosticSeverity.Information
                    : EmailStoreDiagnosticSeverity.Warning,
            diagnostic.Location);

    private static EmailStoreLimitExceededException ConvertLimit(EmailLimitExceededException exception) {
        string name = exception.LimitName == nameof(EmailMailboxReaderOptions.MaxMailboxBytes)
            ? nameof(EmailStoreReaderOptions.MaxInputBytes)
            : exception.LimitName == nameof(EmailMailboxReaderOptions.MaxMessageCount)
                ? nameof(EmailStoreReaderOptions.MaxItemCount)
                : exception.LimitName == nameof(EmailReaderOptions.MaxInputBytes)
                    ? nameof(EmailStoreReaderOptions.MaxMessageBytes)
                : exception.LimitName;
        return new EmailStoreLimitExceededException(name, exception.ActualValue, exception.MaximumValue);
    }
}
