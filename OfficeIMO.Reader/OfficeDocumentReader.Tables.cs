using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Reads a file and returns discovered tables in source order.</summary>
    public IReadOnlyList<ReaderTable> ReadTables(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReader.ExtractTables(Read(path, options, cancellationToken), cancellationToken);
    }

    /// <summary>Reads a caller-owned stream and returns discovered tables in source order.</summary>
    public IReadOnlyList<ReaderTable> ReadTables(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReader.ExtractTables(Read(stream, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>Reads bytes and returns discovered tables in source order.</summary>
    public IReadOnlyList<ReaderTable> ReadTables(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        return DocumentReader.ExtractTables(Read(bytes, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>Reads a file and exports its tables as deterministic CSV, Markdown, and JSON payloads.</summary>
    public IReadOnlyList<ReaderTableExportBundle> ReadTableExports(
        string path,
        ReaderOptions? options = null,
        bool indentedJson = false,
        CancellationToken cancellationToken = default) {
        return DocumentReader.ExportTables(ReadTables(path, options, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>Reads a caller-owned stream and exports its tables as deterministic CSV, Markdown, and JSON payloads.</summary>
    public IReadOnlyList<ReaderTableExportBundle> ReadTableExports(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        bool indentedJson = false,
        CancellationToken cancellationToken = default) {
        return DocumentReader.ExportTables(
            ReadTables(stream, sourceName, options, cancellationToken),
            indentedJson,
            cancellationToken);
    }

    /// <summary>Reads bytes and exports their tables as deterministic CSV, Markdown, and JSON payloads.</summary>
    public IReadOnlyList<ReaderTableExportBundle> ReadTableExports(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        bool indentedJson = false,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        return DocumentReader.ExportTables(
            ReadTables(bytes, sourceName, options, cancellationToken),
            indentedJson,
            cancellationToken);
    }
}
