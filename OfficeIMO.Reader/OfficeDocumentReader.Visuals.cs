using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Reads a file and returns visual payloads in source order.</summary>
    public IReadOnlyList<ReaderVisual> ReadVisuals(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReader.ExtractVisuals(Read(path, options, cancellationToken), cancellationToken);
    }

    /// <summary>Reads a caller-owned stream and returns visual payloads in source order.</summary>
    public IReadOnlyList<ReaderVisual> ReadVisuals(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReader.ExtractVisuals(Read(stream, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>Reads bytes and returns visual payloads in source order.</summary>
    public IReadOnlyList<ReaderVisual> ReadVisuals(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        return DocumentReader.ExtractVisuals(Read(bytes, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>Reads a file and exports its visuals with deterministic payload and JSON sidecars.</summary>
    public IReadOnlyList<ReaderVisualExportBundle> ReadVisualExports(
        string path,
        ReaderOptions? options = null,
        bool indentedJson = false,
        CancellationToken cancellationToken = default) {
        return DocumentReader.ExportVisuals(ReadVisuals(path, options, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>Reads a caller-owned stream and exports its visuals with deterministic payload and JSON sidecars.</summary>
    public IReadOnlyList<ReaderVisualExportBundle> ReadVisualExports(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        bool indentedJson = false,
        CancellationToken cancellationToken = default) {
        return DocumentReader.ExportVisuals(
            ReadVisuals(stream, sourceName, options, cancellationToken),
            indentedJson,
            cancellationToken);
    }

    /// <summary>Reads bytes and exports their visuals with deterministic payload and JSON sidecars.</summary>
    public IReadOnlyList<ReaderVisualExportBundle> ReadVisualExports(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        bool indentedJson = false,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        return DocumentReader.ExportVisuals(
            ReadVisuals(bytes, sourceName, options, cancellationToken),
            indentedJson,
            cancellationToken);
    }
}
