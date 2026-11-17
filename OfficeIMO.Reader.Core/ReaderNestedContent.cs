namespace OfficeIMO.Reader;

/// <summary>
/// Lets a Reader handler delegate nested content to the same immutable handler set that selected it.
/// </summary>
/// <remarks>
/// Archive, mailbox, and attachment adapters use this context to preserve selective package composition.
/// The context is available while a handler is executing; callers outside a Reader operation have no
/// registered nested handlers.
/// </remarks>
public static class ReaderNestedContent {
    /// <summary>Returns true when the active Reader has a stream handler registered for the source name.</summary>
    public static bool CanRead(string sourceName) {
        if (string.IsNullOrWhiteSpace(sourceName)) return false;
        return DocumentReaderEngine.CanReadNestedSource(sourceName);
    }

    /// <summary>Reads nested content through the active Reader handler set.</summary>
    public static IReadOnlyList<ReaderChunk> Read(
        Stream stream,
        string sourceName,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceName);
        return ReadDocument(stream, sourceName, options, cancellationToken).Chunks;
    }

    /// <summary>Reads a nested source without discarding its rich links, forms, assets or metadata.</summary>
    /// <remarks>The containing rich read captures the result in NestedDocuments. SourceName also identifies
    /// its container-relative or virtual path. The caller retains ownership of the stream.</remarks>
    public static OfficeDocumentReadResult ReadDocument(Stream stream, string sourceName,
        ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        return ReadDocumentInContainer(stream, sourceName, sourceName, options, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocumentInContainer(Stream stream, string sourceName, string containerPath,
        ReaderOptions? options, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceName);
        OfficeDocumentReadResult result;
        using (ReaderReadScope.Enter(options, nested: true)) {
            var budget = ReaderReadScope.Current?.Budget;
            Stream input = stream;
            bool ownsInput = false;
            try {
                if (budget != null && !budget.IsReserved(stream)) {
                    if (!stream.CanSeek) {
                        budget.ReserveNestedInput(0);
                        long? remaining = budget.RemainingNestedBytes;
                        using var bounded = new ReaderIncrementalInputStream(stream, remaining, cancellationToken,
                            nameof(ReaderResourceLimits.MaxNestedInputBytes), budget.NestedInputMaximum);
                        var effective = DocumentReaderEngine.NormalizeOptions(options);
                        input = ReaderInputLimits.EnsureSeekableReadStream(bounded,
                            DocumentReaderEngine.ResolveStreamMaxInputBytes(sourceName, effective, streamCanSeek: false),
                            DocumentReaderEngine.ResolveStreamInputLimitProbe(sourceName, effective), cancellationToken, out ownsInput);
                        budget.ReserveNestedBytes(input.Length);
                    } else {
                        budget.ReserveNestedInput(input.Length);
                    }
                }
                result = DocumentReaderEngine.ReadDocument(input, sourceName, options, cancellationToken);
            } finally { if (ownsInput) input.Dispose(); }
        }
        ReaderReadScope.RecordNested(containerPath, result);
        return result;
    }

    /// <summary>Appends an extension implied by the declared media type when an attachment name has none, so the
    /// active handler set can be selected by name.</summary>
    internal static string ResolveSourceName(string fileName, string? mediaType) {
        string extension;
        try { extension = Path.GetExtension(fileName); } catch (ArgumentException) { extension = string.Empty; }
        if (!string.IsNullOrWhiteSpace(extension)) return fileName;
        string? implied = mediaType?.Trim().ToLowerInvariant() switch {
            "application/pdf" => ".pdf",
            "text/plain" => ".txt",
            "text/html" => ".html",
            "text/rtf" or "application/rtf" => ".rtf",
            "text/csv" => ".csv",
            "text/markdown" => ".md",
            "message/rfc822" => ".eml",
            "application/vnd.ms-outlook" => ".msg",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => ".docx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => ".xlsx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation" => ".pptx",
            _ => null
        };
        return implied == null ? fileName : fileName + implied;
    }

    /// <summary>Reserves decoded input bytes before an archive adapter allocates an entry payload.</summary>
    internal static void ReserveDecodedInput(long bytes) => ReaderReadScope.Current?.Budget?.ReserveNestedInput(bytes);
    internal static void MarkReservedPayload(byte[] bytes) => ReaderReadScope.Current?.Budget?.MarkReservedPayload(bytes);
}
