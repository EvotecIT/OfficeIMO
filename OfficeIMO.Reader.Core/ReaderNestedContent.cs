using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

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
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (string.IsNullOrWhiteSpace(sourceName)) throw new ArgumentException("A nested source name is required.", nameof(sourceName));
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
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (string.IsNullOrWhiteSpace(sourceName)) throw new ArgumentException("A nested source name is required.", nameof(sourceName));
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

    /// <summary>Reserves decoded input bytes before an archive adapter allocates an entry payload.</summary>
    internal static void ReserveDecodedInput(long bytes) => ReaderReadScope.Current?.Budget?.ReserveNestedInput(bytes);
    internal static void MarkReservedPayload(byte[] bytes) => ReaderReadScope.Current?.Budget?.MarkReservedPayload(bytes);
}
