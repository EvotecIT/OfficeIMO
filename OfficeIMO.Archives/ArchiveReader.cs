using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Archives;

/// <summary>Reads selected archive content without filesystem publication or executing its contents.</summary>
public static partial class ArchiveReader {
    /// <summary>Reads a ZIP from bytes using the same strict member validation as extraction.
    /// Returns only after every selected payload is fully verified. Do not modify the input during this call.</summary>
    public static IReadOnlyList<ArchiveEntryData> ReadZip(byte[] bytes, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        var limits = (options ?? new ArchiveReadOptions()).Limits();
        if (bytes.LongLength > limits.Input) throw new InvalidDataException("ZIP exceeds the configured input-byte limit.");
        using var source = new MemoryStream(bytes, writable: false);
        return Read(source, limits, cancellationToken);
    }

    /// <summary>Reads from the stream's current position to EOF into a bounded snapshot.
    /// Supports nonseekable streams, leaves the caller's stream open, and advances its position by the bytes read.
    /// Selected content is independent of the stream after return.</summary>
    public static IReadOnlyList<ArchiveEntryData> ReadZip(Stream source, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead) throw new ArgumentException("A readable ZIP stream is required.", nameof(source));
        var limits = (options ?? new ArchiveReadOptions()).Limits();
        cancellationToken.ThrowIfCancellationRequested();
        using var snapshot = new MemoryStream();
        var buffer = new byte[81920];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            long remaining = limits.Input - snapshot.Length;
            int read = source.Read(buffer, 0, remaining >= buffer.Length ? buffer.Length : (int)remaining + 1);
            if (read == 0) break;
            if (read > limits.Input - snapshot.Length) throw new InvalidDataException("ZIP exceeds the configured input-byte limit.");
            snapshot.Write(buffer, 0, read);
        }
        return Read(snapshot, limits, cancellationToken);
    }

    private static IReadOnlyList<ArchiveEntryData> Read(Stream source, ArchiveLimits limits, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (ArchiveExtractor.Detect(source, token) != ArchiveFormat.Zip) throw new NotSupportedException("A ZIP archive is required.");
        var files = new List<ArchiveEntryData>();
        ZipArchiveContent.Read(source, limits, (path, payload, length) => {
            if (length > int.MaxValue) throw new InvalidDataException("Selected ZIP content is too large for an in-memory result.");
            using var content = new MemoryStream();
            payload.CopyTo(content);
            files.Add(new ArchiveEntryData(path, content.ToArray()));
        }, token);
        token.ThrowIfCancellationRequested();
        return Array.AsReadOnly(files.OrderBy(file => file.RelativePath, StringComparer.Ordinal).ToArray());
    }
}
