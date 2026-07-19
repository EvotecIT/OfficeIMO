using OfficeIMO.Email;

namespace OfficeIMO.Reader.Email;

internal static class EmailArtifactReaderAdapter {
    internal static ReaderEmailOptions Clone(ReaderEmailOptions? source) {
        EmailReaderOptions message = CloneMessage(source?.MessageOptions ?? EmailReaderOptions.Default,
            source?.IncludeAttachmentContent ?? true);
        EmailMailboxReaderOptions mailboxSource = source?.MailboxOptions ?? EmailMailboxReaderOptions.Default;
        EmailMailboxReaderOptions mailbox = new EmailMailboxReaderOptions(
            mailboxSource.MaxMailboxBytes,
            CloneMessage(mailboxSource.MessageOptions, source?.IncludeAttachmentContent ?? true),
            mailboxSource.Variant,
            mailboxSource.MaxMessageCount);
        ContentLineReaderOptions lines = CloneContentLines(source?.ContentLineOptions ?? ContentLineReaderOptions.Default);
        return new ReaderEmailOptions {
            MessageOptions = message,
            MailboxOptions = mailbox,
            ContentLineOptions = lines,
            IncludeAttachmentContent = source?.IncludeAttachmentContent ?? true
        };
    }

    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions readerOptions, ReaderEmailOptions options, CancellationToken cancellationToken) {
        string extension = Path.GetExtension(path).ToLowerInvariant();
        if (IsCalendar(extension) || IsVCard(extension)) {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return ReadContentLines(stream, path, extension, readerOptions, options, cancellationToken);
        }
        if (IsMailbox(extension)) {
            EmailMailboxReadResult mailbox = new EmailMailboxReader(EffectiveMailboxOptions(options, readerOptions)).Read(path, cancellationToken);
            return EmailReaderProjection.ProjectMailboxToPathResult(mailbox, path, readerOptions, cancellationToken);
        }
        using EmailReadResult result = new EmailDocumentReader(EffectiveMessageOptions(options, readerOptions)).Read(path, cancellationToken);
        return EmailReaderProjection.ProjectEmailDocumentsToPathResult(
            new[] { result.Document }, new string?[] { path }, result.Diagnostics, result.Document.Format,
            path, path, readerOptions, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName, ReaderOptions readerOptions, ReaderEmailOptions options, CancellationToken cancellationToken) {
        string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "message.eml" : sourceName!.Trim();
        string extension = Path.GetExtension(logicalName).ToLowerInvariant();
        if (IsCalendar(extension) || IsVCard(extension)) {
            return ReadContentLines(stream, logicalName, extension, readerOptions, options, cancellationToken);
        }
        if (IsMailbox(extension)) {
            EmailMailboxReadResult mailbox = new EmailMailboxReader(EffectiveMailboxOptions(options, readerOptions)).Read(stream, cancellationToken);
            return EmailReaderProjection.ProjectMailboxToStreamResult(mailbox, logicalName, stream, readerOptions, cancellationToken);
        }
        using EmailReadResult result = new EmailDocumentReader(EffectiveMessageOptions(options, readerOptions)).Read(stream, logicalName, cancellationToken);
        return EmailReaderProjection.ProjectEmailDocumentsToStreamResult(
            new[] { result.Document }, new string?[] { logicalName }, result.Diagnostics, result.Document.Format,
            logicalName, stream, readerOptions, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadCalendarDocument(string path, ReaderOptions readerOptions, ReaderEmailOptions options, CancellationToken cancellationToken) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return ReadContentLines(stream, path, ".ics", readerOptions, options, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadCalendarDocument(Stream stream, string? sourceName, ReaderOptions readerOptions, ReaderEmailOptions options, CancellationToken cancellationToken) =>
        ReadContentLines(stream, string.IsNullOrWhiteSpace(sourceName) ? "calendar.ics" : sourceName!, ".ics", readerOptions, options, cancellationToken);

    internal static OfficeDocumentReadResult ReadVCardDocument(string path, ReaderOptions readerOptions, ReaderEmailOptions options, CancellationToken cancellationToken) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return ReadContentLines(stream, path, ".vcf", readerOptions, options, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadVCardDocument(Stream stream, string? sourceName, ReaderOptions readerOptions, ReaderEmailOptions options, CancellationToken cancellationToken) =>
        ReadContentLines(stream, string.IsNullOrWhiteSpace(sourceName) ? "contact.vcf" : sourceName!, ".vcf", readerOptions, options, cancellationToken);

    private static OfficeDocumentReadResult ReadContentLines(Stream stream, string sourceName, string extension, ReaderOptions readerOptions, ReaderEmailOptions options, CancellationToken cancellationToken) {
        ContentLineReaderOptions contentOptions = EffectiveContentLineOptions(options, readerOptions);
        string text;
        ReaderInputKind kind;
        string capability;
        if (IsVCard(extension)) {
            text = VCardDocument.Load(stream, contentOptions, cancellationToken).Serialize();
            kind = ReaderInputKind.VCard;
            capability = OfficeDocumentReaderBuilderEmailExtensions.VCardHandlerId;
        } else {
            text = IcsDocument.Load(stream, contentOptions, cancellationToken).Serialize();
            kind = ReaderInputKind.Calendar;
            capability = OfficeDocumentReaderBuilderEmailExtensions.CalendarHandlerId;
        }
        ReaderChunk[] chunks = ChunkText(text, sourceName, kind, readerOptions.MaxChars).ToArray();
        return DocumentReaderEngine.CreateDocumentResult(chunks, kind, null,
            new[] { capability });
    }

    private static IEnumerable<ReaderChunk> ChunkText(string text, string sourceName, ReaderInputKind kind, int maxChars) {
        int limit = Math.Max(256, maxChars);
        string normalized = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        int index = 0;
        for (int offset = 0; offset < normalized.Length; offset += limit) {
            string part = normalized.Substring(offset, Math.Min(limit, normalized.Length - offset));
            yield return new ReaderChunk {
                Id = $"{(kind == ReaderInputKind.VCard ? "vcard" : "calendar")}:{Path.GetFileName(sourceName)}:{index.ToString("D4", CultureInfo.InvariantCulture)}",
                Kind = kind,
                Location = new ReaderLocation { Path = sourceName, BlockIndex = index, SourceBlockIndex = index, SourceBlockKind = kind == ReaderInputKind.VCard ? "vcard" : "icalendar" },
                Text = part,
                Markdown = "```text\n" + part + "\n```"
            };
            index++;
        }
    }

    private static EmailReaderOptions EffectiveMessageOptions(ReaderEmailOptions options, ReaderOptions readerOptions) {
        EmailReaderOptions source = options.MessageOptions ?? EmailReaderOptions.Default;
        long max = readerOptions.MaxInputBytes.HasValue ? Math.Min(source.MaxInputBytes, readerOptions.MaxInputBytes.Value) : source.MaxInputBytes;
        return CloneMessage(source, options.IncludeAttachmentContent, max);
    }

    private static EmailMailboxReaderOptions EffectiveMailboxOptions(ReaderEmailOptions options, ReaderOptions readerOptions) {
        EmailMailboxReaderOptions source = options.MailboxOptions ?? EmailMailboxReaderOptions.Default;
        long max = readerOptions.MaxInputBytes.HasValue ? Math.Min(source.MaxMailboxBytes, readerOptions.MaxInputBytes.Value) : source.MaxMailboxBytes;
        return new EmailMailboxReaderOptions(max, EffectiveMessageOptions(options, readerOptions), source.Variant, source.MaxMessageCount);
    }

    private static ContentLineReaderOptions EffectiveContentLineOptions(ReaderEmailOptions options, ReaderOptions readerOptions) {
        ContentLineReaderOptions source = options.ContentLineOptions ?? ContentLineReaderOptions.Default;
        long max = readerOptions.MaxInputBytes.HasValue ? Math.Min(source.MaxInputBytes, readerOptions.MaxInputBytes.Value) : source.MaxInputBytes;
        return new ContentLineReaderOptions(max, source.MaxUnfoldedLineBytes, source.MaxComponents, source.MaxProperties, source.MaxNestingDepth, source.Encoding);
    }

    private static EmailReaderOptions CloneMessage(EmailReaderOptions source, bool includeAttachmentContent, long? maxInputBytes = null) => new EmailReaderOptions(
        maxInputBytes ?? source.MaxInputBytes, source.MaxHeaderBytes, source.MaxHeaderCount, source.MaxPartCount,
        source.MaxMimeDepth, source.MaxAttachmentBytes, source.MaxTotalAttachmentBytes, source.MaxNestedMessageDepth,
        includeAttachmentContent, source.PreserveRawSource, source.MaxCompoundDirectoryEntries, source.MaxMapiPropertyCount,
        source.MaxDecodedPropertyBytes, source.MaxTnefAttributeCount);

    private static ContentLineReaderOptions CloneContentLines(ContentLineReaderOptions source) => new ContentLineReaderOptions(
        source.MaxInputBytes, source.MaxUnfoldedLineBytes, source.MaxComponents, source.MaxProperties,
        source.MaxNestingDepth, source.Encoding);

    private static bool IsMailbox(string extension) => extension is ".mbox" or ".mbx";
    private static bool IsCalendar(string extension) => extension is ".ics" or ".ical" or ".ifb" or ".vcs";
    private static bool IsVCard(string extension) => extension is ".vcf" or ".vcard";
}
