using OfficeIMO.Email.Store;

namespace OfficeIMO.Email;

/// <summary>A handle-free message view. Payload operations reopen its unchanged local source for their duration.</summary>
public sealed class EmailMessage {
    private readonly EmailMessageSource _source;
    private readonly EmailStoreItemReference? _reference;

    internal EmailMessage(EmailMessageSource source, EmailDocument document,
        EmailStoreItemReference? reference = null, string? folderPath = null,
        EmailStoreItemContentAvailability? availability = null) {
        _source = source;
        _reference = reference;
        Id = reference?.Id ?? document.MessageId ?? "message";
        FolderPath = folderPath;
        Subject = document.Subject;
        From = document.From;
        To = document.Recipients.Where(r => r.Kind == EmailRecipientKind.To).Select(r => r.Address).ToArray();
        Cc = document.Recipients.Where(r => r.Kind == EmailRecipientKind.Cc).Select(r => r.Address).ToArray();
        Bcc = document.Recipients.Where(r => r.Kind == EmailRecipientKind.Bcc).Select(r => r.Address).ToArray();
        Date = document.Date;
        ReceivedDate = document.ReceivedDate;
        TextBody = document.Body.Text;
        HtmlBody = document.Body.Html;
        RtfBody = document.Body.Rtf;
        ContentAvailability = availability;
        Attachments = document.Attachments.Select((a, i) => new EmailMessageAttachment(this, a, i)).ToArray();
        ExportName = EmailStoreExportPathBuilder.SanitizeSegmentByUtf8Bytes(Subject, 100, 120, "message") + "-" +
            EmailStoreExportPathBuilder.GetStableHash(SourcePath + "\n" + Id);
    }

    /// <summary>Stable item identifier within the source.</summary>
    public string Id { get; }
    /// <summary>Absolute local source path.</summary>
    public string SourcePath => _source.Path;
    /// <summary>Store folder path, or null for an individual file.</summary>
    public string? FolderPath { get; }
    /// <summary>Message subject.</summary>
    public string? Subject { get; }
    /// <summary>Represented author.</summary>
    public EmailAddress? From { get; }
    /// <summary>Primary recipients.</summary>
    public IReadOnlyList<EmailAddress> To { get; }
    /// <summary>Carbon-copy recipients.</summary>
    public IReadOnlyList<EmailAddress> Cc { get; }
    /// <summary>Blind-carbon-copy recipients.</summary>
    public IReadOnlyList<EmailAddress> Bcc { get; }
    /// <summary>Declared sent date.</summary>
    public DateTimeOffset? Date { get; }
    /// <summary>Received date when available.</summary>
    public DateTimeOffset? ReceivedDate { get; }
    /// <summary>Plain-text body alternative, when present.</summary>
    public string? TextBody { get; }
    /// <summary>Original HTML body alternative, when present. Use the HTML exporter for a safe projection.</summary>
    public string? HtmlBody { get; }
    /// <summary>RTF body alternative, when present.</summary>
    public string? RtfBody { get; }
    /// <summary>Attachment metadata; reading this collection does not retain payloads.</summary>
    public IReadOnlyList<EmailMessageAttachment> Attachments { get; }
    /// <summary>Local-cache availability reported by the store reader; null for individual files.</summary>
    public EmailStoreItemContentAvailability? ContentAvailability { get; }
    /// <summary>Portable subject and source-identity filename stem for batch exports.</summary>
    public string ExportName { get; }

    /// <summary>Runs an owner operation with a full document and releases its source immediately afterward.</summary>
    /// <remarks>Do not return content streams or the document from the callback. Keep the source unchanged between reading and exporting.</remarks>
    public T UseDocument<T>(Func<EmailDocument, T> operation, CancellationToken cancellationToken = default) {
        if (operation == null) throw new ArgumentNullException(nameof(operation));
        _source.Validate();
        using var data = _source.Open(true, cancellationToken);
        EmailDocument document = _reference == null
            ? data.EmailDocument ?? throw new InvalidDataException("The source is no longer an email file.")
            : (data.Store ?? throw new InvalidDataException("The source is no longer a mail store."))
                .ReadItem(_reference, new EmailStoreItemReadOptions(EmailStoreItemReadParts.All,
                    preferStreamingAttachmentContent: true), cancellationToken).Document;
        _source.Validate();
        if (document.Subject != Subject || document.Body.Text != TextBody || document.Body.Html != HtmlBody ||
            document.Attachments.Count != Attachments.Count)
            throw new IOException("The source message changed after it was read. Read it again before exporting.");
        return operation(document);
    }

    /// <summary>Exports a native email artifact with an atomic destination conflict policy.</summary>
    public EmailWriteResult Save(string path, bool overwrite = false, CancellationToken cancellationToken = default,
        EmailWriterOptions? options = null) =>
        UseDocument(d => Task.Run(() => d.SaveWithConflictPolicyAsync(path,
            overwrite ? EmailFileConflictPolicy.Replace : EmailFileConflictPolicy.FailIfExists, options, cancellationToken),
            cancellationToken).GetAwaiter().GetResult(), cancellationToken);

    /// <summary>Extracts selected top-level attachments through the bounded shared extractor.</summary>
    public EmailAttachmentExtractionResult SaveAttachments(string directory, IEnumerable<int> indexes,
        CancellationToken cancellationToken = default, string? fileNamePrefix = null) => UseDocument(d => EmailAttachmentExtractor.Extract(d, directory,
            new EmailAttachmentExtractionOptions(includeHidden: true, selectedAttachmentIndexes: indexes,
                fileNamePrefix: fileNamePrefix), cancellationToken), cancellationToken);
}

/// <summary>Handle-free attachment metadata tied to one source message.</summary>
public sealed class EmailMessageAttachment {
    internal EmailMessageAttachment(EmailMessage message, EmailAttachment attachment, int index) {
        Message = message; Index = index; FileName = attachment.FileName; ContentType = attachment.ContentType;
        ContentId = attachment.ContentId; ContentLocation = attachment.ContentLocation;
        IsInline = attachment.IsInline; IsHidden = attachment.IsHidden; Size = attachment.Length;
    }
    /// <summary>Owning source message.</summary>
    public EmailMessage Message { get; }
    /// <summary>Zero-based position in the source attachment collection.</summary>
    public int Index { get; }
    /// <summary>Original, untrusted filename.</summary>
    public string? FileName { get; }
    /// <summary>Declared media type.</summary>
    public string? ContentType { get; }
    /// <summary>Embedded resource identifier.</summary>
    public string? ContentId { get; }
    /// <summary>Embedded resource location.</summary>
    public string? ContentLocation { get; }
    /// <summary>Whether marked inline by the source.</summary>
    public bool IsInline { get; }
    /// <summary>Whether hidden by Outlook.</summary>
    public bool IsHidden { get; }
    /// <summary>Declared decoded size.</summary>
    public long Size { get; }
}
