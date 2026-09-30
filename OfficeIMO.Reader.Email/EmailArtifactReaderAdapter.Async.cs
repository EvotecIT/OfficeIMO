using OfficeIMO.Email;
using System.Threading.Tasks;

namespace OfficeIMO.Reader.Email;

internal static partial class EmailArtifactReaderAdapter {
    internal static Task<OfficeDocumentReadResult> ReadDocumentAsync(string path, ReaderOptions readerOptions,
        ReaderEmailOptions options, CancellationToken cancellationToken) =>
        ReadDocumentPathAsync(path, readerOptions, options, null, cancellationToken);

    internal static Task<OfficeDocumentReadResult> ReadCalendarDocumentAsync(string path, ReaderOptions readerOptions,
        ReaderEmailOptions options, CancellationToken cancellationToken) =>
        ReadDocumentPathAsync(path, readerOptions, options, ReaderInputKind.Calendar, cancellationToken);

    internal static Task<OfficeDocumentReadResult> ReadVCardDocumentAsync(string path, ReaderOptions readerOptions,
        ReaderEmailOptions options, CancellationToken cancellationToken) =>
        ReadDocumentPathAsync(path, readerOptions, options, ReaderInputKind.VCard, cancellationToken);

    internal static Task<OfficeDocumentReadResult> ReadCalendarDocumentAsync(Stream stream, string? sourceName,
        ReaderOptions readerOptions, ReaderEmailOptions options, CancellationToken cancellationToken) =>
        ReadDocumentAsyncCore(stream, sourceName ?? "calendar.ics", readerOptions, options, false, ReaderInputKind.Calendar, cancellationToken);

    internal static Task<OfficeDocumentReadResult> ReadVCardDocumentAsync(Stream stream, string? sourceName,
        ReaderOptions readerOptions, ReaderEmailOptions options, CancellationToken cancellationToken) =>
        ReadDocumentAsyncCore(stream, sourceName ?? "contact.vcf", readerOptions, options, false, ReaderInputKind.VCard, cancellationToken);

    private static async Task<OfficeDocumentReadResult> ReadDocumentPathAsync(string path, ReaderOptions readerOptions,
        ReaderEmailOptions options, ReaderInputKind? selectedContentLineKind, CancellationToken cancellationToken) {
        string extension = Path.GetExtension(path).ToLowerInvariant();
        bool contentLines = selectedContentLineKind.HasValue || IsCalendar(extension) || IsVCard(extension);
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
            contentLines ? FileShare.ReadWrite | FileShare.Delete : FileShare.Read,
            81920, FileOptions.Asynchronous | FileOptions.SequentialScan);
        return await ReadDocumentAsyncCore(stream, path, readerOptions, options, true, selectedContentLineKind, cancellationToken).ConfigureAwait(false);
    }

    internal static Task<OfficeDocumentReadResult> ReadDocumentAsync(Stream stream, string? sourceName,
        ReaderOptions readerOptions, ReaderEmailOptions options, CancellationToken cancellationToken) =>
        ReadDocumentAsyncCore(stream, sourceName, readerOptions, options, false, null, cancellationToken);

    private static async Task<OfficeDocumentReadResult> ReadDocumentAsyncCore(Stream stream, string? sourceName,
        ReaderOptions readerOptions, ReaderEmailOptions options, bool pathSource,
        ReaderInputKind? selectedContentLineKind, CancellationToken cancellationToken) {
        string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "message.eml" : sourceName!.Trim();
        string extension = Path.GetExtension(logicalName).ToLowerInvariant();
        OfficeDocumentReadResult result;
        if (selectedContentLineKind.HasValue || IsCalendar(extension) || IsVCard(extension)) {
            ContentLineReaderOptions contentOptions = EffectiveContentLineOptions(options, readerOptions);
            bool vcard = selectedContentLineKind.HasValue ? selectedContentLineKind == ReaderInputKind.VCard : IsVCard(extension);
            string text = vcard
                ? (await VCardDocument.LoadAsync(stream, contentOptions, cancellationToken).ConfigureAwait(false)).Serialize()
                : (await IcsDocument.LoadAsync(stream, contentOptions, cancellationToken).ConfigureAwait(false)).Serialize();
            ReaderInputKind kind = vcard ? ReaderInputKind.VCard : ReaderInputKind.Calendar;
            result = DocumentReaderEngine.CreateDocumentResult(ChunkText(text, logicalName, kind, readerOptions.MaxChars).ToArray(), kind, null,
                new[] { vcard ? OfficeDocumentReaderBuilderEmailExtensions.VCardHandlerId : OfficeDocumentReaderBuilderEmailExtensions.CalendarHandlerId });
        } else if (IsMailbox(extension)) {
            EmailMailboxReadResult mailbox = await new EmailMailboxReader(EffectiveMailboxOptions(options, readerOptions))
                .ReadAsync(stream, cancellationToken).ConfigureAwait(false);
            result = pathSource
                ? EmailReaderProjection.ProjectMailboxToPathResult(mailbox, logicalName, readerOptions, cancellationToken, computeSourceHash: false,
                    includeEmbeddedMessageContent: options.MailboxOptions!.MessageOptions.IncludeEmbeddedMessages)
                : EmailReaderProjection.ProjectMailboxToStreamResult(mailbox, logicalName, stream, readerOptions, cancellationToken, computeSourceHash: false,
                    includeEmbeddedMessageContent: options.MailboxOptions!.MessageOptions.IncludeEmbeddedMessages);
        } else {
            using EmailReadResult read = await new EmailDocumentReader(EffectiveMessageOptions(options, readerOptions))
                .ReadAsync(stream, logicalName, cancellationToken).ConfigureAwait(false);
            result = pathSource
                ? EmailReaderProjection.ProjectEmailDocumentsToPathResult(new[] { read.Document },
                    new string?[] { logicalName }, read.Diagnostics, read.Document.Format, logicalName, logicalName,
                    readerOptions, cancellationToken, computeSourceHash: false,
                    includeEmbeddedMessageContent: options.MessageOptions!.IncludeEmbeddedMessages)
                : EmailReaderProjection.ProjectEmailDocumentsToStreamResult(new[] { read.Document },
                    new string?[] { logicalName }, read.Diagnostics, read.Document.Format, logicalName, stream,
                    readerOptions, cancellationToken, computeSourceHash: false,
                    includeEmbeddedMessageContent: options.MessageOptions!.IncludeEmbeddedMessages);
        }
        return result;
    }

}
