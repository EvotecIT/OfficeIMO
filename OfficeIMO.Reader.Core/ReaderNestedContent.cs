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
        return DocumentReaderEngine.Read(stream, sourceName, options, cancellationToken).ToArray();
    }
}
